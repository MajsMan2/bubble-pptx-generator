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

    // --- SKRIDT 2: FLET POWERPOINT (PLACEHOLDERS) ---
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
            const elements = await s.getAllTextElementIds();
            for (const element of elements) {
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

    // --- FIND DET RIGTIGE FILE ID ---
    const ooData = onlyOfficeResponse.data;
    let onlyOfficeFileId = "ukendt-id";
    
    if (ooData) {
      if (ooData.id) onlyOfficeFileId = ooData.id;
      else if (ooData.response?.id) onlyOfficeFileId = ooData.response.id;
      else if (ooData.response?.Id) onlyOfficeFileId = ooData.response.Id;
      else if (Array.isArray(ooData.response) && ooData.response[0]?.id) onlyOfficeFileId = ooData.response[0].id;
      else if (Array.isArray(ooData.response) && ooData.response[0]?.Id) onlyOfficeFileId = ooData.response[0].Id;
      else if (ooData.response?.file?.id) onlyOfficeFileId = ooData.response.file.id;
    }

    // --- SKRIDT 3.5: GENERER ISOLERET EKSTERNT DELINGSLINK ---
    let secureFileUrl = "";
    if (onlyOfficeFileId !== "ukendt-id") {
      try {
        const linkEndpoint = `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/link`;
        const linkResponse = await axios.post(linkEndpoint, {}, {
          headers: {
            'Authorization': `Bearer ${docSpaceToken}`,
            'Content-Type': 'application/json'
          }
        });
        // Gemmer det direkte standalone-link fra API-svaret
        secureFileUrl = linkResponse.data?.response?.shareLink || linkResponse.data?.response || "";
      } catch (linkError) {
        console.error("Kunne ikke generere delingslink:", linkError.message);
      }
    }

    // --- SKRIDT 4: OPRYDNING & SVAR ---
    if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

    return res.status(200).json({
      success: true,
      fileId: String(onlyOfficeFileId),
      fileUrl: secureFileUrl // <--- DETTE LINK SENDES NU RETUR TIL BUBBLE!
    });

  } catch (globalError) {
    return res.status(500).json({
      error: 'Uventet global fejl i backenden',
      message: globalError.message
    });
  }
};
