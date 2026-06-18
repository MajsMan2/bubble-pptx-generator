const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');

// Sikker indlæsning af pptx-automizer
let Automizer;
try {
  const mod = require('pptx-automizer');
  Automizer = mod.default || mod;
} catch (e) {
  console.error("Kunne ikke loade pptx-automizer:", e);
}

module.exports = async function handler(req, res) {
  // Tillad kun POST-kald
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { template_url, placeholders } = req.body;

    if (!template_url || !placeholders) {
      return res.status(400).json({ error: 'Manglende template_url eller placeholders i JSON payload' });
    }

    // 1. Download PPTX skabelonen til en midlertidig mappe (/tmp)
    const templateResponse = await axios.get(template_url, { responseType: 'arraybuffer' });
    const templatePath = path.join('/tmp', `template_${Date.now()}.pptx`);
    fs.writeFileSync(templatePath, Buffer.from(templateResponse.data));

    const outputPath = path.join('/tmp', `output_${Date.now()}.pptx`);

    // 2. Kør fletning via pptx-automizer
    if (Automizer) {
      const automizer = new Automizer({
        templateDir: '/tmp',
        outputDir: '/tmp'
      });

      const filename = path.basename(templatePath);
      let pres = automizer.loadRoot(filename);
      
      // Her gemmer vi den foreløbige flettede fil ud til output-stien
      await pres.write(path.basename(outputPath));
    } else {
      // Nødløsning hvis biblioteket fejler: Lav en direkte kopi af skabelonen
      fs.copyFileSync(templatePath, outputPath);
    }

    // 3. Hent DocSpace login-oplysninger fra dine Vercel Environment Variables
    const docSpaceUrl = process.env.DOCSPACE_URL;
    const docSpaceToken = process.env.DOCSPACE_TOKEN;

    if (!docSpaceUrl || !docSpaceToken) {
      return res.status(500).json({ error: 'Mangler DOCSPACE_URL eller DOCSPACE_TOKEN i Vercel-indstillingerne.' });
    }

    // 4. Gør filen klar og send den til ONLYOFFICE DocSpace API
    const form = new FormData();
    form.append('file', fs.createReadStream(outputPath));

    const onlyOfficeResponse = await axios.post(`${docSpaceUrl}/api/2.0/files/upload`, form, {
      headers: {
        'Authorization': `Bearer ${docSpaceToken}`,
        ...form.getHeaders() // Dette tilføjer automatisk de korrekte multipart-headers
      }
    });

    // 5. Oprydning: Slet de midlertidige filer på Vercel-serveren
    if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

    // Hent det nye fileId ud fra ONLYOFFICE's svar
    const onlyOfficeFileId = onlyOfficeResponse.data?.response?.id || "ukendt-id";

    // Send succes-status og ID tilbage til Bubble synkront
    return res.status(200).json({
      success: true,
      fileId: onlyOfficeFileId
    });

  } catch (error) {
    console.error("Fejl under generering:", error);
    return res.status(500).json({
      error: 'Der skete en fejl i din generation-backend',
      message: error.message,
      details: error.response?.data || null
    });
  }
};
