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

    // --- SKRIDT 3.5: OFFENTLIG DELING ---
    let secureFileUrl = "";
    let linkDebugInfo = "OK - Offentligt link oprettet";

    if (onlyOfficeFileId !== "ukendt-id") {

      // Hjælpefunktion: udtræk shareLink fra et response-objekt
      const extractLink = (data) => {
        if (!data) return "";
        const r = data.response ?? data;
        // Direkte på response-objektet
        if (r?.sharedTo?.shareLink) return r.sharedTo.shareLink;
        if (r?.shareLink)           return r.shareLink;
        if (r?.link)                return r.link;
        if (r?.url)                 return r.url;
        // Første element i et array
        const first = Array.isArray(r) ? r[0] : null;
        if (first?.sharedTo?.shareLink) return first.sharedTo.shareLink;
        if (first?.shareLink)           return first.shareLink;
        if (first?.link)                return first.link;
        return "";
      };

      // ---
      // Forsøg 1: POST /api/2.0/files/file/{id}/link
      // Opretter det primære eksterne link (officielt DocSpace-endpoint).
      // ---
      try {
        const createLinkRes = await axios.post(
          `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/link`,
          {},   // tom body — API'et kræver ingen parametre her
          { headers: { 'Authorization': `Bearer ${docSpaceToken}`, 'Content-Type': 'application/json' } }
        );
        secureFileUrl = extractLink(createLinkRes.data);
        if (secureFileUrl) {
          linkDebugInfo = "OK - primært eksternt link oprettet via POST /file/:id/link";
        }
      } catch (e1) {
        linkDebugInfo = `POST /link fejlede (${e1.response?.status ?? e1.message})`;
      }

      // ---
      // Forsøg 2: GET /api/2.0/files/file/{id}/link
      // Henter det primære eksterne link, hvis det allerede eksisterer.
      // ---
      if (!secureFileUrl) {
        try {
          const getLinkRes = await axios.get(
            `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/link`,
            { headers: { 'Authorization': `Bearer ${docSpaceToken}` } }
          );
          secureFileUrl = extractLink(getLinkRes.data);
          if (secureFileUrl) {
            linkDebugInfo = "OK - primært eksternt link hentet via GET /file/:id/link";
          }
        } catch (e2) {
          linkDebugInfo += ` | GET /link fejlede (${e2.response?.status ?? e2.message})`;
        }
      }

      // ---
      // Forsøg 3: PUT /api/2.0/files/file/{id}/links
      // Sætter et eksternt link med adgangsniveau (det officielle "Set an external link"-endpoint).
      // ---
      if (!secureFileUrl) {
        try {
          const putLinksRes = await axios.put(
            `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/links`,
            { access: 2, linkType: 2, denyDownload: false },
            { headers: { 'Authorization': `Bearer ${docSpaceToken}`, 'Content-Type': 'application/json' } }
          );
          secureFileUrl = extractLink(putLinksRes.data);
          if (secureFileUrl) {
            linkDebugInfo = "OK - eksternt link oprettet via PUT /file/:id/links";
          } else {
            linkDebugInfo += ` | PUT /links svarede 200 men uden link: ${JSON.stringify(putLinksRes.data).substring(0, 200)}`;
          }
        } catch (e3) {
          linkDebugInfo += ` | PUT /links fejlede (${e3.response?.status ?? e3.message})`;
        }
      }

      // ---
      // Forsøg 4: GET /api/2.0/files/file/{id}/links
      // Henter alle eksisterende eksterne links på filen.
      // ---
      if (!secureFileUrl) {
        try {
          const getLinksRes = await axios.get(
            `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/links`,
            { headers: { 'Authorization': `Bearer ${docSpaceToken}` } }
          );
          secureFileUrl = extractLink(getLinksRes.data);
          if (secureFileUrl) {
            linkDebugInfo = "OK - eksternt link hentet via GET /file/:id/links";
          } else {
            linkDebugInfo += ` | GET /links svarede uden link: ${JSON.stringify(getLinksRes.data).substring(0, 200)}`;
          }
        } catch (e4) {
          linkDebugInfo += ` | GET /links fejlede (${e4.response?.status ?? e4.message})`;
        }
      }

      // ---
      // Forsøg 5: Byg direkte editor-URL som absolut nødløsning
      // ---
      if (!secureFileUrl) {
        secureFileUrl = `${baseUrl}/doceditor?fileId=${onlyOfficeFileId}&action=view`;
        linkDebugInfo += " | Brugte direkte editor-URL som nødløsning";
      }

    } else {
      linkDebugInfo = "Kunne ikke oprette link: Fandt ikke et gyldigt fileId i upload-svaret.";
    }

    // Sikr at relative URL'er bliver absolutte
    if (secureFileUrl && secureFileUrl.startsWith('/')) {
      secureFileUrl = baseUrl + secureFileUrl;
    }

    // Fallback til intern webUrl fra upload-svaret, hvis alt andet er tomt
    if (!secureFileUrl && fallbackWebUrl) {
      secureFileUrl = fallbackWebUrl;
      linkDebugInfo += " -> Brugte intern fallback URL fra upload-svar.";
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
