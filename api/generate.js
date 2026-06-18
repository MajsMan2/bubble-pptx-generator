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

  try {
    let body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    const { template_url, placeholders } = body;
    
    if (!template_url || !placeholders) {
      return res.status(400).json({ error: 'Manglende template_url eller placeholders.' });
    }

    // 1. DOWNLOAD
    const templateResponse = await axios.get(template_url.trim().startsWith('//') ? 'https:' + template_url.trim() : template_url.trim(), { responseType: 'arraybuffer' });
    fs.writeFileSync(templatePath, Buffer.from(templateResponse.data));

    // 2. FLET
    if (Automizer && modify) {
      try {
        const automizer = new Automizer({ templateDir: '/tmp', outputDir: '/tmp', removeExistingSlides: true });
        const templateFilename = path.basename(templatePath);
        let pres = automizer.loadRoot(templateFilename);
        pres.load(templateFilename, 'base');

        const slides = (await pres.getInfo()).slidesByTemplate('base');
        const replaceParams = Object.entries(placeholders).flatMap(([k, v]) => [
          { replace: k, by: { text: String(v) } },
          { replace: `{{${k}}}`, by: { text: String(v) } }
        ]);
        const shapeModCb = modify.replaceText(replaceParams);

        for (const slide of slides) {
          pres.addSlide('base', slide.number, async (s) => {
            const textIds = await s.getAllTextElementIds();
            const allElements = typeof s.getAllElements === 'function' ? await s.getAllElements() : [];
            const tableIds = allElements.filter(el => el?.type === 'table').map(el => el.name);
            [...new Set([...textIds, ...tableIds])].forEach(id => s.modifyElement(id, shapeModCb));
          });
        }
        await pres.write(path.basename(outputPath));
      } catch (e) {
        fs.copyFileSync(templatePath, outputPath);
      }
    } else {
      fs.copyFileSync(templatePath, outputPath);
    }

    // 3. UPLOAD TIL ONLYOFFICE
    const { DOCSPACE_URL, DOCSPACE_TOKEN, DOCSPACE_FOLDER_ID } = process.env;
    let baseUrl = DOCSPACE_URL.trim().replace(/\/$/, '').replace('/api/2.0', '');
    
    const form = new FormData();
    form.append('file', fs.createReadStream(outputPath));
    const uploadRes = await axios.post(`${baseUrl}/api/2.0/files/${DOCSPACE_FOLDER_ID}/upload`, form, {
      headers: { 'Authorization': `Bearer ${DOCSPACE_TOKEN}`, ...form.getHeaders() }
    });

    const fileId = uploadRes.data?.response?.file?.id || uploadRes.data?.response?.id || uploadRes.data?.id;

    // 4. GENERER OFFENTLIGT DELINGSLINK (Med korrekt API-håndtering)
    let secureFileUrl = "";
    let debugInfo = "Forsøgte deling...";

    try {
      // Sendes som array 'files' jf. DocSpace API krav
      await axios.post(`${baseUrl}/api/2.0/files/share`, {
        files: [Number(fileId)],
        access: 2,
        shareType: 1
      }, { headers: { 'Authorization': `Bearer ${DOCSPACE_TOKEN}` } });

      // Hent delinger med GET for at finde linket
      const shareList = await axios.get(`${baseUrl}/api/2.0/files/file/${fileId}/shares`, {
        headers: { 'Authorization': `Bearer ${DOCSPACE_TOKEN}` }
      });
      
      const publicShare = shareList.data?.response?.find(s => s.shareType === 1);
      secureFileUrl = publicShare?.link || "";
      debugInfo = "Succes: Link genereret via shares-liste.";
    } catch (e) {
      debugInfo = "Deling fejlede: " + e.message;
    }

    // OPRYDNING
    [templatePath, outputPath].forEach(p => fs.existsSync(p) && fs.unlinkSync(p));

    return res.status(200).json({ success: true, fileId, fileUrl: secureFileUrl, debugInfo });

  } catch (err) {
    return res.status(500).json({ error: 'Global fejl', message: err.message });
  }
};
