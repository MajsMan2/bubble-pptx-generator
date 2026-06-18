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
    let linkDebugInfo = "OK - Offentligt redigeringslink oprettet";

    if (onlyOfficeFileId !== "ukendt-id") {

      // Forsøg 1: Fil-specifikt share-endpoint med PUT (korrekt DocSpace endpoint)
      try {
        const shareEndpoint = `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/share`;

        const shareResponse = await axios.put(shareEndpoint, {
          share: [
            {
              shareTo: "00000000-0000-0000-0000-000000000000", // "Everyone" / public
              access: "Edit"
            }
          ],
          notify: false,
          sharingMessage: ""
        }, {
          headers: {
            'Authorization': `Bearer ${docSpaceToken}`,
            'Content-Type': 'application/json'
          }
        });

        const resData = shareResponse.data;
        const entries = resData?.response;

        if (Array.isArray(entries) && entries.length > 0) {
          secureFileUrl = entries[0]?.sharedTo?.shareLink
                       || entries[0]?.link
                       || entries[0]?.shareLink
                       || "";
        }

        // Forsøg 2: Hvis PUT ikke returnerede et link, hent eksisterende shares via GET
        if (!secureFileUrl) {
          const getSharesResponse = await axios.get(shareEndpoint, {
            headers: { 'Authorization': `Bearer ${docSpaceToken}` }
          });
          const shares = getSharesResponse.data?.response;
          if (Array.isArray(shares) && shares.length > 0) {
            secureFileUrl = shares[0]?.sharedTo?.shareLink
                         || shares[0]?.link
                         || shares[0]?.shareLink
                         || "";
          }
        }

        if (!secureFileUrl) {
          linkDebugInfo = "Share oprettet, men intet link i svar: " + JSON.stringify(resData).substring(0, 300);
        }

      } catch (shareError) {

        // Forsøg 3: DocSpace "external link" endpoint
        try {
          const externalLinkEndpoint = `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/links`;

          const linkResponse = await axios.post(externalLinkEndpoint, {
            access: 2,          // 2 = Edit
            linkType: 2,        // 2 = External/public link
            password: "",
            expirationDate: null,
            denyDownload: false
          }, {
            headers: {
              'Authorization': `Bearer ${docSpaceToken}`,
              'Content-Type': 'application/json'
            }
          });

          const linkData = linkResponse.data?.response;
          secureFileUrl = linkData?.sharedTo?.shareLink
                       || linkData?.link
                       || linkData?.shareLink
                       || linkData?.url
                       || "";

          if (!secureFileUrl) {
            linkDebugInfo = "links-endpoint svarede, men intet link: " + JSON.stringify(linkData).substring(0, 300);
          }

        } catch (linkError) {

          // Forsøg 4: Byg direkte editor-URL som nødløsning
          const viewerUrl = `${baseUrl}/doceditor?fileId=${onlyOfficeFileId}&action=view`;
          secureFileUrl = viewerUrl;
          linkDebugInfo = `Brugte direkte editor-URL som nødløsning. Share-fejl: ${shareError.message}. Links-fejl: ${linkError.message}`;
        }
      }
    } else {
      linkDebugInfo = "Kunne ikke oprette link: Fandt ikke et gyldigt fileId i upload-svaret.";
    }

    // Hvis alt fejler, giv fallback web-url så systemet ikke crasher
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
