const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');

let Automizer;
let modify;
try {
  const mod = require('pptx-automizer');
  Automizer = mod.default || mod;
  modify = mod.modify;
} catch (e) {
  console.error("Kunne ikke loade pptx-automizer:", e);
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  let templatePath = path.join('/tmp', `template_${Date.now()}.pptx`);
  let outputPath = path.join('/tmp', `output_${Date.now()}.pptx`);
  let uploadUrl = "";

  try {
    let body = req.body;
    if (typeof body === 'string') {
      body = JSON.parse(body);
    }

    const { template_url, placeholders } = body;
    if (!template_url || !placeholders) {
      return res.status(400).json({ error: 'Manglende template_url eller placeholders i JSON.' });
    }

    // --- SKRIDT 1: DOWNLOAD SKABELON ---
    let finalUrl = template_url.trim();
    if (finalUrl.startsWith('//')) {
      finalUrl = 'https:' + finalUrl;
    }

    try {
      const templateResponse = await axios.get(finalUrl, { responseType: 'arraybuffer' });
      fs.writeFileSync(templatePath, Buffer.from(templateResponse.data));
    } catch (downloadError) {
      return res.status(500).json({
        error: 'Fejl under download af din PPTX skabelon fra Vercel/GitHub',
        message: downloadError.message
      });
    }

    // --- SKRIDT 2: FLET POWERPOINT (PLACEHOLDERS I TEKST & TABELLER) ---
    if (Automizer && modify) {
      try {
        const automizer = new Automizer({ 
          templateDir: '/tmp', 
          outputDir: '/tmp',
          removeExistingSlides: true 
        });
        
        const templateFilename = path.basename(templatePath);
        let pres = automizer.loadRoot(templateFilename);
        pres.load(templateFilename, 'base');

        const info = await pres.getInfo();
        const slides = info.slidesByTemplate('base');

        // Gør erstatnings-parametrene klar (både rå tekst og {{tekst}})
        const replaceParams = [];
        for (const [key, value] of Object.entries(placeholders)) {
          replaceParams.push({ replace: key, by: { text: String(value) } });
          replaceParams.push({ replace: `{{${key}}}`, by: { text: String(value) } });
        }

        const shapeModCb = modify.replaceText(replaceParams);

        for (const slide of slides) {
          pres.addSlide('base', slide.number, async (s) => {
            // 1. Hent alle almindelige tekst-elementer
            const textElements = await s.getAllTextElementIds();
            
            // 2. Hent alle tabel-elementer (så vi kan ramme placeholders inde i tabeller)
            let tableElements = [];
            try {
              if (typeof s.getAllElements === 'function') {
                const allElements = await s.getAllElements();
                tableElements = allElements
                  .filter(el => el && el.type === 'table' && el.name)
                  .map(el => el.name);
              }
            } catch (tableError) {
              console.error("Kunne ikke scanne efter tabeller på slide " + slide.number, tableError);
            }

            // 3. Kombiner tekstbokse og tabeller til én samlet liste uden dubletter
            const combinedElements = Array.from(new Set([...textElements, ...tableElements]));

            // 4. Kør Søg & Erstat på alle fundne elementer
            for (const element of combinedElements) {
              s.modifyElement(element, shapeModCb);
            }
          });
        }

        await pres.write(path.basename(outputPath));
      } catch (fletFejl) {
        console.error("Fletfejl, sender rå skabelon videre:", fletFejl);
        fs.copyFileSync(templatePath, outputPath);
      }
    } else {
      fs.copyFileSync(templatePath, outputPath);
    }

    // --- SKRIDT 3: UPLOAD TIL ONLYOFFICE ---
    const docSpaceUrl = process.env.DOCSPACE_URL;
    const docSpaceToken = process.env.DOCSPACE_TOKEN;
    const folderId = process.env.DOCSPACE_FOLDER_ID;

    if (!docSpaceUrl || !docSpaceToken || !folderId) {
      return res.status(500).json({ error: 'Vercel mangler DOCSPACE_URL, DOCSPACE_TOKEN eller DOCSPACE_FOLDER_ID i indstillingerne.' });
    }

    let baseUrl = docSpaceUrl.trim().replace(/\/$/, '');
    if (baseUrl.endsWith('/api/2.0')) {
      baseUrl = baseUrl.replace('/api/2.0', '');
    }
    
    uploadUrl = `${baseUrl}/api/2.0/files/${folderId}/upload`;

    const form = new FormData();
    form.append('file', fs.createReadStream(outputPath));

    let onlyOfficeResponse;
    try {
      onlyOfficeResponse = await axios.post(uploadUrl, form, {
        headers: {
          'Authorization': `Bearer ${docSpaceToken}`,
          ...form.getHeaders()
        }
      });
    } catch (uploadError) {
      return res.status(500).json({
        error: 'Fejl under upload til ONLYOFFICE DocSpace API',
        message: uploadError.message,
        details: uploadError.response?.data
      });
    }

    // --- SKRIDT 3.2: FIND FILE ID & FALLBACK URL ---
    const ooData = onlyOfficeResponse.data;
    let onlyOfficeFileId = "ukendt-id";
    let fallbackWebUrl = "";
    
    if (ooData) {
      if (ooData.id) onlyOfficeFileId = ooData.id;
      else if (ooData.response?.id) onlyOfficeFileId = ooData.response.id;
      else if (ooData.response?.Id) onlyOfficeFileId = ooData.response.Id;
      else if (Array.isArray(ooData.response) && ooData.response[0]?.id) onlyOfficeFileId = ooData.response[0].id;
      else if (Array.isArray(ooData.response) && ooData.response[0]?.Id) onlyOfficeFileId = ooData.response[0].Id;
      else if (ooData.response?.file?.id) onlyOfficeFileId = ooData.response.file.id;

      const fileObj = ooData.response?.file || (Array.isArray(ooData.response) ? ooData.response[0] : ooData.response);
      if (fileObj) {
        fallbackWebUrl = fileObj.webUrl || fileObj.WebUrl || fileObj.viewUrl || fileObj.ViewUrl || "";
      }
    }

    // --- SKRIDT 3.5: GENERER DELINGSLINK ---
    let secureFileUrl = "";
    let linkDebugInfo = "OK - Link oprettet succesfuldt via API";

    if (onlyOfficeFileId !== "ukendt-id") {
      try {
        const linkEndpoint = `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/link`;
        const linkResponse = await axios.post(linkEndpoint, {}, {
          headers: {
            'Authorization': `Bearer ${docSpaceToken}`,
            'Content-Type': 'application/json'
          }
        });
        
        const resData = linkResponse.data;
        secureFileUrl = resData?.response?.shareLink || 
                        resData?.response?.link || 
                        resData?.shareLink || 
                        resData?.link || "";
                        
        if (!secureFileUrl) {
          linkDebugInfo = "API svarede 200, men kunne ikke matche JSON-stien. Svar-data: " + JSON.stringify(resData);
        }
      } catch (linkError) {
        linkDebugInfo = "Fejl på /link endpoint: " + linkError.message + 
                        (linkError.response?.data ? " - API svar: " + JSON.stringify(linkError.response.data) : "");
      }
    } else {
      linkDebugInfo = "Kunne ikke kalde /link, da fileId ikke blev fundet i upload-svaret.";
    }

    if (!secureFileUrl && fallbackWebUrl) {
      secureFileUrl = fallbackWebUrl;
    }

    if (secureFileUrl && secureFileUrl.startsWith('/')) {
      secureFileUrl = baseUrl + secureFileUrl;
    }

    // --- SKRIDT 4: OPRYDNING & SVAR ---
    if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

    return res.status(200).json({
      success: true,
      fileId: String(onlyOfficeFileId),
      fileUrl: secureFileUrl,
      debugInfo: linkDebugInfo
    });

  } catch (globalError) {
    return res.status(500).json({
      error: 'Uventet global fejl i backenden',
      message: globalError.message
    });
  }
};
