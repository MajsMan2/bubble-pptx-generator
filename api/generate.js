import { PptxAutomizer } from 'pptx-automizer';
import axios from 'axios';
import fs from 'fs';
import path from 'path';

export default async function handler(req, res) {
  // Tillad kun POST-kald (hvor man sender data med)
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { template_url, placeholders } = req.body;
    
    // 1. Download PPTX skabelonen fra URL'en (f.eks. fra Bubble Storage)
    const templateResponse = await axios.get(template_url, { responseType: 'arraybuffer' });
    const inputBuffer = Buffer.from(templateResponse.data);
    
    const outputFilePath = path.join('/tmp', `output_${Date.now()}.pptx`);

    // 2. Start pptx-automizer til at flette placeholders
    const automizer = new PptxAutomizer();
    
    // Her kører fletningen. Koden looper igennem dine placeholders fra Bubble
    // og erstatter dem i PowerPoint-filen.
    let result = await automizer
      .load(inputBuffer)
      .modifyCell((cell) => {
        // En simpel placeholder-erstatning fx {{customer_name}}
        for (const [key, value] of Object.entries(placeholders)) {
          if (cell.text && cell.text.includes(`{{${key}}}`)) {
            cell.text = cell.text.replace(`{{${key}}}` , value);
          }
        }
      })
      .write(outputFilePath);

    // 3. Upload den flettede fil til ONLYOFFICE DocSpace API
    const fileStream = fs.createReadStream(outputFilePath);
    
    // OBS: Du skal bruge din egen ONLYOFFICE DocSpace URL og dit API-token/JWT
    const docSpaceUrl = process.env.DOCSPACE_URL; 
    const docSpaceToken = process.env.DOCSPACE_TOKEN;

    const formData = new FormData();
    formData.append('file', fileStream);

    const onlyOfficeResponse = await axios.post(`${docSpaceUrl}/api/2.0/files/upload`, formData, {
      headers: {
        'Authorization': `Bearer ${docSpaceToken}`,
        ...formData.getHeaders()
      }
    });

    // API'et fra ONLYOFFICE returnerer et ID på filen
    const onlyOfficeFileId = onlyOfficeResponse.data.response.id;

    // Fjern den midlertidige fil fra Vercel-serveren igen
    fs.unlinkSync(outputFilePath);

    // 4. Send succes og fileId direkte tilbage til Bubble synkront
    return res.status(200).json({
      success: true,
      fileId: onlyOfficeFileId
    });

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: 'Noget gik galt under fletningen: ' + error.message });
  }
}
