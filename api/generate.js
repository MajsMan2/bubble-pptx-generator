const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');

let Automizer;
try {
  const mod = require('pptx-automizer');
  Automizer = mod.default || mod;
} catch (e) {
  console.error("Kunne ikke loade pptx-automizer:", e);
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  let templatePath = path.join('/tmp', `template_${Date.now()}.pptx`);
  let outputPath = path.join('/tmp', `output_${Date.now()}.pptx`);
  let uploadUrl = ""; // Defineres globalt så den kan bruges i fejlhåndtering

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
    if (Automizer) {
      try {
        const automizer = new Automizer({ templateDir: '/tmp', outputDir: '/tmp' });
        let pres = automizer.loadRoot(path.basename(templatePath));
        await pres.write(path.basename(outputPath));
      } catch (fletFejl) {
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

    // SIKKERHEDS-RENSNING AF URL:
    let baseUrl = docSpaceUrl.trim().replace(/\/$/, '');
    // Hvis du er kommet til at skrive /api/2.0 i din Vercel variabel, fjerner vi det her for at undgå dubletter
    if (baseUrl.endsWith('/api/2.0')) {
      baseUrl = baseUrl.replace('/api/2.0', '');
    }
    
    // Vi stykker den officielle upload-URL sammen
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
      // Her sender vi den præcise URL med ud i Bubble, så vi kan se fejlen sort på hvidt
      return res.status(500).json({
        error: 'Fejl under upload til ONLYOFFICE DocSpace API',
        message: uploadError.message,
        status: uploadError.response?.status,
        details: {
          URL_vi_proevede_at_ramme: uploadUrl,
          onlyOfficeSvar: uploadError.response?.data || "Ingen data sendt retur fra ONLYOFFICE"
        }
      });
    }

    // --- SKRIDT 4: OPRYDNING & SVAR ---
    if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

    const onlyOfficeFileId = onlyOfficeResponse.data?.response?.id || "ukendt-id";

    return res.status(200).json({
      success: true,
      fileId: onlyOfficeFileId
    });

  } catch (globalError) {
    return res.status(500).json({
      error: 'Uventet global fejl i backenden',
      message: globalError.message,
      details: { URL_under_fejl: uploadUrl }
    });
  }
};
