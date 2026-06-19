const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');
const AdmZip = require('adm-zip');

let Automizer;
let modify;
try {
  const mod = require('pptx-automizer');
  Automizer = mod.default || mod;
  modify = mod.modify;
} catch (e) {
  console.error("Kunne ikke loade pptx-automizer:", e);
}

function randomId() {
  return Math.random().toString(36).substring(2, 6);
}

function safeFilename(name) {
  return String(name).replace(/[^a-zA-Z0-9æøåÆØÅ\-_]/g, '_').substring(0, 60);
}

function tryParseArray(value) {
  if (Array.isArray(value)) return value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed.startsWith('[')) {
      try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) return parsed;
      } catch (e) {}
    }
  }
  return null;
}

function isEmpty(value) {
  if (value === null || value === undefined) return true;
  const str = String(value).trim();
  if (str === '') return true;
  if (str.toLowerCase() === 'null') return true;
  if (str.toLowerCase() === 'undefined') return true;
  if (/^<[^>]+>$/.test(str)) return true;
  return false;
}

function buildReplaceParams(placeholders) {
  const params = [];
  for (const [key, value] of Object.entries(placeholders)) {
    if (isEmpty(value)) {
      params.push({ replace: key,          by: { text: '' } });
      params.push({ replace: `{{${key}}}`, by: { text: '' } });
      continue;
    }
    const arr = tryParseArray(value);
    const text = arr ? arr.join('\n') : String(value);
    params.push({ replace: key,          by: { text } });
    params.push({ replace: `{{${key}}}`, by: { text } });
  }
  return params;
}

// --- POST-PROCESSING: XML-niveau cleanup ---
// Efter automizer er færdig, åbner vi PPTX som ZIP og erstatter
// alle tilbageværende {{...}} og <placeholder> mønstre direkte i XML.
function cleanupResidualPlaceholders(pptxPath, placeholders) {
  try {
    const zip = new AdmZip(pptxPath);
    const entries = zip.getEntries();

    for (const entry of entries) {
      // Kun slide XML-filer
      if (!entry.entryName.match(/^ppt\/slides\/slide\d+\.xml$/)) continue;

      let xml = entry.getData().toString('utf8');
      let changed = false;

      // 1) Erstat kendte placeholders med deres værdier (eller tom streng)
      for (const [key, value] of Object.entries(placeholders)) {
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const replacement = isEmpty(value) ? '' : String(tryParseArray(value) ? tryParseArray(value).join('\n') : value);

        // {{key}} format
        const re1 = new RegExp(`\\{\\{${escapedKey}\\}\\}`, 'g');
        if (re1.test(xml)) { xml = xml.replace(re1, replacement); changed = true; }

        // <key> format (kun hvis det ligner en placeholder, ikke HTML-tags)
        const re2 = new RegExp(`<${escapedKey}>`, 'g');
        if (re2.test(xml)) { xml = xml.replace(re2, replacement); changed = true; }
      }

      // 2) Catch-all: fjern alle tilbageværende {{...}} mønstre
      if (/\{\{[^}]+\}\}/.test(xml)) {
        xml = xml.replace(/\{\{[^}]+\}\}/g, '');
        changed = true;
      }

      // 3) Erstat "null" der står alene som tekstindhold i en XML-celle
      // Matcher >null< og >null < og lignende
      if (/>[  ]*null[  ]*</.test(xml)) {
        xml = xml.replace(/(?<=>[ ]*)null(?=[ ]*<)/g, '');
        changed = true;
      }

      if (changed) {
        zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
      }
    }

    zip.writeZip(pptxPath);
  } catch (cleanupError) {
    console.error("Post-processing cleanup fejlede:", cleanupError);
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  let templatePath = path.join('/tmp', `template_${Date.now()}.pptx`);
  let outputPath;

  try {
    let body = req.body;

    if (typeof body === 'string') {
      try {
        body = JSON.parse(body);
      } catch (parseError) {
        return res.status(400).json({
          error: 'Ugyldig JSON i request body',
          message: parseError.message,
          rawBody: body.substring(0, 300)
        });
      }
    }

    const { template_url, placeholders, company_unique_id, company_name } = body;

    if (!template_url || !placeholders) {
      return res.status(400).json({ error: 'Manglende template_url eller placeholders i JSON.' });
    }

    const filePrefix = company_name
      ? `${randomId()}_${safeFilename(company_name)}`
      : `${randomId()}_rapport`;
    outputPath = path.join('/tmp', `${filePrefix}.pptx`);

    // --- SKRIDT 1: DOWNLOAD SKABELON ---
    let finalUrl = template_url.trim();
    if (finalUrl.startsWith('//')) finalUrl = 'https:' + finalUrl;

    try {
      const templateResponse = await axios.get(finalUrl, { responseType: 'arraybuffer' });
      fs.writeFileSync(templatePath, Buffer.from(templateResponse.data));
    } catch (downloadError) {
      return res.status(500).json({
        error: 'Fejl under download af din PPTX skabelon fra Vercel/GitHub',
        message: downloadError.message
      });
    }

    // --- SKRIDT 2: FLET POWERPOINT VIA AUTOMIZER ---
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

        const replaceParams = buildReplaceParams(placeholders);
        const shapeModCb = modify.replaceText(replaceParams);

        const arrayPlaceholders = {};
        for (const [key, value] of Object.entries(placeholders)) {
          if (isEmpty(value)) continue;
          const arr = tryParseArray(value);
          if (arr && arr.length > 1) {
            arrayPlaceholders[key] = arr;
          }
        }

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

            for (const tableName of tableElements) {
              try {
                s.modifyElement(tableName, async (element, xmlData) => {
                  const xmlStr = typeof xmlData === 'string' ? xmlData : JSON.stringify(xmlData);

                  let arrayKey = null;
                  let arrayValues = null;

                  for (const [key, values] of Object.entries(arrayPlaceholders)) {
                    if (xmlStr.includes(key) || xmlStr.includes(`{{${key}}}`)) {
                      arrayKey = key;
                      arrayValues = values;
                      break;
                    }
                  }

                  if (!arrayKey || !arrayValues) return element;

                  if (xmlData && xmlData.elements) {
                    const tblEl = xmlData.elements.find(el => el.name === 'a:tbl' || el.name === 'tbl');
                    if (tblEl && tblEl.elements) {
                      const rows = tblEl.elements.filter(el => el.name === 'a:tr' || el.name === 'tr');

                      let templateRowIndex = -1;
                      for (let i = 0; i < rows.length; i++) {
                        const rowStr = JSON.stringify(rows[i]);
                        if (rowStr.includes(arrayKey) || rowStr.includes(`{{${arrayKey}}}`)) {
                          templateRowIndex = i;
                          break;
                        }
                      }

                      if (templateRowIndex >= 0) {
                        const templateRow = rows[templateRowIndex];
                        const newRows = [];
                        const escapedKey = arrayKey.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

                        for (const arrayValue of arrayValues) {
                          const rowCopy = JSON.parse(JSON.stringify(templateRow));
                          const rowStr = JSON.stringify(rowCopy);
                          const replacedStr = rowStr
                            .replace(new RegExp(escapedKey, 'g'), String(arrayValue))
                            .replace(new RegExp(`\\{\\{${escapedKey}\\}\\}`, 'g'), String(arrayValue));
                          newRows.push(JSON.parse(replacedStr));
                        }

                        tblEl.elements.splice(
                          tblEl.elements.indexOf(templateRow),
                          1,
                          ...newRows
                        );
                      }
                    }
                  }

                  return element;
                });
              } catch (tableModError) {
                console.error(`Fejl ved tabelmanipulation af ${tableName}:`, tableModError);
              }

              s.modifyElement(tableName, shapeModCb);
            }

            const combinedElements = Array.from(new Set([...textElements]));
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

    // --- SKRIDT 2.5: XML-NIVEAU CLEANUP ---
    // Fjerner alle tilbageværende {{...}} og "null"-værdier direkte i PPTX-XML
    cleanupResidualPlaceholders(outputPath, placeholders);

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

    const form = new FormData();
    form.append('file', fs.createReadStream(outputPath), path.basename(outputPath));

    let onlyOfficeResponse;
    try {
      onlyOfficeResponse = await axios.post(
        `${baseUrl}/api/2.0/files/${folderId}/upload`,
        form,
        {
          headers: {
            'Authorization': `Bearer ${docSpaceToken}`,
            ...form.getHeaders()
          }
        }
      );
    } catch (uploadError) {
      return res.status(500).json({
        error: 'Fejl under upload til ONLYOFFICE DocSpace API',
        message: uploadError.message,
        details: uploadError.response?.data
      });
    }

    // --- SKRIDT 3.2: FIND FILE ID ---
    const ooData = onlyOfficeResponse.data;
    let onlyOfficeFileId = "ukendt-id";

    if (ooData) {
      if (ooData.id)                                                        onlyOfficeFileId = ooData.id;
      else if (ooData.response?.id)                                         onlyOfficeFileId = ooData.response.id;
      else if (ooData.response?.Id)                                         onlyOfficeFileId = ooData.response.Id;
      else if (Array.isArray(ooData.response) && ooData.response[0]?.id)   onlyOfficeFileId = ooData.response[0].id;
      else if (Array.isArray(ooData.response) && ooData.response[0]?.Id)   onlyOfficeFileId = ooData.response[0].Id;
      else if (ooData.response?.file?.id)                                   onlyOfficeFileId = ooData.response.file.id;
    }

    // --- SKRIDT 3.5: OPRET OFFENTLIGT EKSTERNT LINK MED EDIT-ADGANG ---
    let shareToken = "";
    let linkDebugInfo = "";

    const extractShareToken = (data) => {
      if (!data) return "";
      const r = data.response ?? data;
      const candidates = [r, ...(Array.isArray(r) ? r : [])];
      for (const c of candidates) {
        const raw = c?.sharedTo?.shareLink || c?.shareLink || c?.link || c?.url || "";
        if (raw) return raw;
      }
      return "";
    };

    if (onlyOfficeFileId !== "ukendt-id") {
      try {
        const r = await axios.post(
          `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/link`,
          { access: 2 },
          { headers: { 'Authorization': `Bearer ${docSpaceToken}`, 'Content-Type': 'application/json' } }
        );
        shareToken = extractShareToken(r.data);
        linkDebugInfo = shareToken
          ? "POST /file/:id/link (Edit)"
          : `POST /link 200 men tomt: ${JSON.stringify(r.data).substring(0, 150)}`;
      } catch (e1) {
        linkDebugInfo = `POST /link fejl (${e1.response?.status ?? e1.message})`;
      }

      if (!shareToken) {
        try {
          const r = await axios.get(
            `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/link`,
            { headers: { 'Authorization': `Bearer ${docSpaceToken}` } }
          );
          shareToken = extractShareToken(r.data);
          linkDebugInfo += shareToken
            ? " | GET /file/:id/link OK"
            : ` | GET /link tomt: ${JSON.stringify(r.data).substring(0, 150)}`;
        } catch (e2) {
          linkDebugInfo += ` | GET /link fejl (${e2.response?.status ?? e2.message})`;
        }
      }

      if (!shareToken) {
        try {
          const r = await axios.put(
            `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/links`,
            { access: 2, linkType: 2, denyDownload: false },
            { headers: { 'Authorization': `Bearer ${docSpaceToken}`, 'Content-Type': 'application/json' } }
          );
          shareToken = extractShareToken(r.data);
          linkDebugInfo += shareToken
            ? " | PUT /links OK"
            : ` | PUT /links tomt: ${JSON.stringify(r.data).substring(0, 150)}`;
        } catch (e3) {
          linkDebugInfo += ` | PUT /links fejl (${e3.response?.status ?? e3.message})`;
        }
      }

      if (!shareToken) {
        try {
          const r = await axios.get(
            `${baseUrl}/api/2.0/files/file/${onlyOfficeFileId}/links`,
            { headers: { 'Authorization': `Bearer ${docSpaceToken}` } }
          );
          shareToken = extractShareToken(r.data);
          linkDebugInfo += shareToken
            ? " | GET /links OK"
            : ` | GET /links tomt: ${JSON.stringify(r.data).substring(0, 150)}`;
        } catch (e4) {
          linkDebugInfo += ` | GET /links fejl (${e4.response?.status ?? e4.message})`;
        }
      }
    }

    // --- SKRIDT 3.6: BYGG EDITOR-URL ---
    let editorUrl = "";

    if (shareToken) {
      const tokenMatch = shareToken.match(/\/s\/([^/?#]+)/);
      const token = tokenMatch ? tokenMatch[1] : "";
      editorUrl = token
        ? `${baseUrl}/doceditor?fileId=${onlyOfficeFileId}&share=${token}&action=edit&type=desktop`
        : shareToken;
      linkDebugInfo = `OK (${linkDebugInfo.trim()})`;
    } else {
      editorUrl = `${baseUrl}/doceditor?fileId=${onlyOfficeFileId}&action=edit&type=desktop`;
      linkDebugInfo = `Ingen share-token — editor-URL kræver login. Debug: ${linkDebugInfo}`;
    }

    // --- SKRIDT 4: OPRYDNING & SVAR ---
    if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
    if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

    return res.status(200).json({
      success: true,
      company_unique_id: company_unique_id ?? null,
      fileId: String(onlyOfficeFileId),
      fileName: path.basename(outputPath),
      fileUrl: editorUrl,
      shareLink: shareToken,
      debugInfo: linkDebugInfo
    });

  } catch (globalError) {
    return res.status(500).json({
      error: 'Uventet global fejl i backenden',
      message: globalError.message
    });
  }
};
