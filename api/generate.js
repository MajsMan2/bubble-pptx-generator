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

    // --- SKRIDT 2: FLET POWERPOINT ---
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

        const replaceParams = [];
        for (const [key, value] of Object.entries(placeholders)) {
          replaceParams.push({ replace: key, by: { text: String(value) } });
          replaceParams.push({ replace: `{{${key}}}`, by: { text: String(value) } });
        }

        const shapeModCb = modify.replaceText(replaceParams);

        for (const slide of slides) {
          pres.addSlide('base', slide.number, async (s) => {
            const textElements = await s.getAllTextElementIds();
            
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

            const combinedElements = Array.from(new Set([...textElements, ...tableElements]));

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

    // --- SKRIDT 3.2: FIND FILE ID & FALLBACK ---
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

    // --- SKRIDT 3.5: OFFENTLIG DELING VIA UNIVERSAL ENDPOINT ---
    let secureFileUrl = "";
    let linkDebugInfo = "OK - Offentligt redigeringslink oprettet";

    if (onlyOfficeFileId !== "ukendt-id") {
      try {
        // Universelt DocSpace share-endpoint
        const shareEndpoint = `${baseUrl}/api/2.0/files/share`;
        
        const shareResponse = await axios.post(shareEndpoint, {
          fileId: Number(onlyOfficeFileId), // Sender id'et med i JSON body
          shareType: 1,                    // 1 = Offentligt link (Public link)
          access: 2                        // 2 = Redigeringsadgang (Edit)
        }, {
          headers: {
            'Authorization': `Bearer ${docSpaceToken}`,
            'Content-Type': 'application/json'
          }
        });
        
        const resData = shareResponse.data;
        secureFileUrl = resData?.response?.link || 
                        resData?.response?.shareLink || 
                        resData?.link || 
                        resData?.shareLink || "";
                        
        if (!secureFileUrl) {
          linkDebugInfo = "API svarede 200, men manglede URL i JSON: " + JSON.stringify(resData);
        }
      } catch (linkError) {
        // Hvis universelt fejler, prøver vi det sekundære "invitation" link-endpoint
        try {
          const fallbackEndpoint = `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/shares`;
          const altResponse = await axios.post(fallbackEndpoint, {
            share: { shareType: 1, access: 2 }
          }, {
            headers: { 'Authorization': `Bearer ${docSpaceToken}`, 'Content-Type': 'application/json' }
          });
          secureFileUrl = altResponse.data?.response?.link || altResponse.data?.link || "";
        } catch (altError) {
          linkDebugInfo = `404 afhjulpet, men ny fejl opstod. Besked: ${linkError.message}. Detaljer: ` + JSON.stringify(linkError.response?.data || {});
        }
      }
    } else {
      linkDebugInfo = "Kunne ikke oprette link: Fandt ikke et gyldigt fileId i upload-svaret.";
    }

    // Hvis alt fejler, giv dem i det mindste fallback web-url'en så systemet ikke crasher
    if (!secureFileUrl && fallbackWebUrl) {
      secureFileUrl = fallbackWebUrl;
      linkDebugInfo += " -> Brugte intern fallback URL.";
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
