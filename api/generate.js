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

function parseJsonValue(value) {
  let parsed = value;

  for (let attempt = 0; attempt < 3 && typeof parsed === 'string'; attempt++) {
    const trimmed = parsed.trim();
    try {
      const nextValue = JSON.parse(trimmed);
      if (nextValue === parsed) break;
      parsed = nextValue;
    } catch (e) {
      break;
    }
  }

  return parsed;
}

function tryParseArray(value) {
  const parsed = parseJsonValue(value);
  return Array.isArray(parsed) ? parsed : null;
}

function tryParseObject(value) {
  const parsed = parseJsonValue(value);
  return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
}

function tryParseList(value) {
  return tryParseArray(value);
}

function splitDelimitedValue(value) {
  if (typeof value !== 'string' || !value.includes(',')) return null;

  const values = value.split(',').map(item => item.trim());
  if (values.length < 2 || values.some(item => item === '')) return null;
  return values;
}

function findSharedDelimitedCount(entries) {
  const counts = new Map();
  for (const [, values] of entries) {
    counts.set(values.length, (counts.get(values.length) || 0) + 1);
  }

  return [...counts.entries()]
    .filter(([count, columnCount]) => count > 1 && columnCount > 1)
    .sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]?.[0] ?? null;
}

function expandDelimitedObject(object) {
  if (!object || typeof object !== 'object' || Array.isArray(object)) return null;

  const entries = Object.entries(object);
  const delimitedEntries = entries
    .map(([key, value]) => [key, splitDelimitedValue(value)])
    .filter(([, values]) => values);
  const sharedCount = findSharedDelimitedCount(delimitedEntries);

  if (!sharedCount) return null;

  return Array.from({ length: sharedCount }, (_, index) => Object.fromEntries(
    entries.map(([key, value]) => {
      const values = splitDelimitedValue(value);
      return [key, values?.length === sharedCount ? values[index] : value];
    })
  ));
}

function expandDelimitedRows(rows) {
  if (!Array.isArray(rows)) return rows;

  return rows.flatMap(row => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) return [row];
    return expandDelimitedObject(row) || [row];
  });
}

function normalizeRows(rows) {
  return expandDelimitedRows(tryParseArray(rows));
}

function normalizeDelimitedPlaceholders(placeholders) {
  if (!placeholders || typeof placeholders !== 'object') return placeholders;
  const expandedRows = expandDelimitedObject(placeholders);
  if (!expandedRows) return placeholders;

  return Object.fromEntries(Object.keys(placeholders).map(key => [
    key,
    splitDelimitedValue(placeholders[key])?.length === expandedRows.length
      ? expandedRows.map(row => row[key])
      : placeholders[key]
  ]));
}

function buildPlaceholdersFromRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return null;

  const keys = new Set();
  for (const row of rows) {
    if (!row || typeof row !== 'object' || Array.isArray(row)) continue;
    Object.keys(row).forEach(key => keys.add(key));
  }

  if (keys.size === 0) return null;
  return Object.fromEntries(
    [...keys].map(key => {
      const values = rows.map(row => row?.[key] ?? '');
      const meaningfulValues = values.filter(value => !isEmpty(value));
      const hasOneRepeatedValue = meaningfulValues.length > 1 && meaningfulValues.every(
        value => String(value).trim() === String(meaningfulValues[0]).trim()
      );

      return [key, hasOneRepeatedValue ? [meaningfulValues[0]] : values];
    })
  );
}

function isNumericValue(value) {
  if (typeof value === 'number') return Number.isFinite(value);
  if (typeof value !== 'string') return false;
  return value.trim() !== '' && Number.isFinite(Number(value));
}

function defaultPlaceholderValue(key, value, numericKeys) {
  if (!isEmpty(value)) return value;
  return numericKeys.has(key) ? 0 : '';
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
    const arr = tryParseArray(value);
    const text = isEmpty(value) ? '' : (arr ? arr.join('\n') : String(value));

    // Erstat KUN {{key}}-format for at undgå at ramme normale ord i brødtekst.
    // Fx vil "domme", "total", "men" ellers slette ord i løbende tekst.
    params.push({ replace: `{{${key}}}`, by: { text } });
  }
  return params;
}

function findXmlElement(root, names) {
  if (!root || !root.elements) return null;
  for (const element of root.elements) {
    if (names.includes(element.name)) return element;
    const found = findXmlElement(element, names);
    if (found) return found;
  }
  return null;
}

function resizeExpandedTable(xmlData, rows) {
  const table = findXmlElement(xmlData, ['a:tbl', 'tbl']);
  const transform = findXmlElement(xmlData, ['a:xfrm', 'xfrm']);
  if (!table || !transform || !table.elements || !transform.elements) return;

  const rowHeights = rows
    .map(row => Number(row.attributes?.h || row.attributes?.height || 0))
    .filter(height => Number.isFinite(height) && height > 0);
  if (rowHeights.length === 0) return;

  const ext = transform.elements.find(element => element.name === 'a:ext' || element.name === 'ext');
  const off = transform.elements.find(element => element.name === 'a:off' || element.name === 'off');
  if (!ext || !ext.attributes) return;

  const tableHeight = rowHeights.reduce((total, height) => total + height, 0);
  const slideHeight = 5143500;
  const top = Number(off?.attributes?.y || 0);
  const availableHeight = Math.max(0, slideHeight - top);
  const targetHeight = Math.min(tableHeight, availableHeight);
  const scale = tableHeight > 0 ? targetHeight / tableHeight : 1;

  if (scale < 1) {
    for (const row of rows) {
      const height = Number(row.attributes?.h || row.attributes?.height || 0);
      if (!Number.isFinite(height) || height <= 0) continue;
      const scaledHeight = String(Math.max(1, Math.round(height * scale)));
      if (row.attributes.h !== undefined) row.attributes.h = scaledHeight;
      else row.attributes.height = scaledHeight;
    }
  }

  ext.attributes.cy = String(Math.round(targetHeight));
}

function escapeXmlText(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function placeholderPattern(key) {
  const token = `{{${key}}}`;
  return new RegExp(
    Array.from(token).map(character => escapeRegExp(character)).join('(?:<[^>]+>)*'),
    'g'
  );
}

function replacePlaceholderInXml(xml, key, value) {
  const replacement = escapeXmlText(value);
  const tokenPattern = placeholderPattern(key);
  if (tokenPattern.test(xml)) {
    tokenPattern.lastIndex = 0;
    return xml.replace(tokenPattern, () => replacement);
  }

  const barePattern = new RegExp(
    Array.from(String(key)).map(character => escapeRegExp(character)).join('(?:<[^>]+>)*'),
    'g'
  );
  return xml.replace(barePattern, () => replacement);
}

function expandArrayTablesInXml(pptxPath, placeholders) {
  const arrayPlaceholders = Object.fromEntries(
    Object.entries(placeholders).filter(([, value]) => {
      const values = tryParseList(value);
      return values && values.length > 1;
    })
  );
  if (Object.keys(arrayPlaceholders).length === 0) return;

  try {
    const zip = new AdmZip(pptxPath);
    for (const entry of zip.getEntries()) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName)) continue;

      let xml = entry.getData().toString('utf8');
      let changed = false;

      xml = xml.replace(/<a:tbl(?:\s[^>]*)?>[\s\S]*?<\/a:tbl>/g, (tableXml) => {
        let updatedTable = tableXml;

        const arrayEntries = Object.entries(arrayPlaceholders);
        const placeholderEntries = Object.entries(placeholders).map(([key, value]) => {
          const values = tryParseList(value);
          return [key, values || [value]];
        });
        const rows = updatedTable.match(/<a:tr(?:\s[^>]*)?>[\s\S]*?<\/a:tr>/g) || [];
        const templateRow = rows.find(row => {
          const textOnlyRow = row
            .replace(/<[^>]+>/g, '')
            .replace(/\s+/g, '');
          return arrayEntries.some(([key]) =>
            textOnlyRow.includes(`{{${key}}}`) || textOnlyRow.includes(key)
          );
        });

        if (!templateRow) return updatedTable;

        const templateText = templateRow
          .replace(/<[^>]+>/g, '')
          .replace(/\s+/g, '');
        const rowArrays = arrayEntries.filter(([key]) =>
          templateText.includes(`{{${key}}}`) || templateText.includes(key)
        );
        if (rowArrays.length === 0) return updatedTable;

        const rowCount = Math.max(...rowArrays.map(([, values]) => values.length));
        const expandedRows = Array.from({ length: rowCount }, (_, index) => {
          return placeholderEntries.reduce((row, [key, values]) => {
            const value = values[index] ?? '';
            return replacePlaceholderInXml(row, key, value);
          }, templateRow);
        }).join('');

        updatedTable = updatedTable.replace(templateRow, expandedRows);
        changed = true;

        return updatedTable;
      });

      if (changed) {
        zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
      }
    }

    zip.writeZip(pptxPath);
  } catch (error) {
    console.error('Tabeludvidelse via XML fejlede:', error);
  }
}

function cleanupResidualPlaceholders(pptxPath, placeholders, numericKeys) {
  try {
    const zip = new AdmZip(pptxPath);
    const replacements = Object.entries(placeholders).map(([key, value]) => ({
      pattern: placeholderPattern(key),
      value: isEmpty(value)
        ? defaultPlaceholderValue(key, value, numericKeys)
        : escapeXmlText(tryParseList(value)?.join('\n') ?? value)
    }));

    for (const entry of zip.getEntries()) {
      if (!/^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName)) continue;

      let xml = entry.getData().toString('utf8');
      let changed = false;

      for (const replacement of replacements) {
        const updatedXml = xml.replace(replacement.pattern, () => replacement.value);
        if (updatedXml !== xml) changed = true;
        xml = updatedXml;
      }

      const withoutUnknownPlaceholders = xml.replace(/\{\{[^}]+\}\}/g, '');
      if (withoutUnknownPlaceholders !== xml) changed = true;
      xml = withoutUnknownPlaceholders;

      if (changed) {
        zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
      }
    }

    zip.writeZip(pptxPath);
  } catch (error) {
    console.error('Placeholder-cleanup fejlede:', error);
  }
}

// --- SLET SLIDES VIA ZIP/XML ---
// slidesToDelete: array af 1-baserede slide-numre, fx [1, 3, 5]
function deleteSlides(pptxPath, slidesToDelete) {
  if (!slidesToDelete || slidesToDelete.length === 0) return;
  try {
    const zip = new AdmZip(pptxPath);

    // Sorter faldende så vi sletter bagfra og undgår indeks-forskydning
    const sorted = [...slidesToDelete].map(Number).sort((a, b) => b - a);

    // Læs presentation.xml for at finde slide-referencer
    const presEntry = zip.getEntry('ppt/presentation.xml');
    if (!presEntry) return;
    let presXml = presEntry.getData().toString('utf8');

    // Find alle slide-id'er i rækkefølge: <p:sldId id="..." r:id="rId..."/>
    const sldIdRegex = /<p:sldId[^/]* r:id="(rId\d+)"[^/]*\/>/g;
    const slideRefs = [];
    let m;
    while ((m = sldIdRegex.exec(presXml)) !== null) {
      slideRefs.push({ full: m[0], rId: m[1] });
    }

    // Læs .rels for at finde filnavne
    const relsEntry = zip.getEntry('ppt/_rels/presentation.xml.rels');
    if (!relsEntry) return;
    let relsXml = relsEntry.getData().toString('utf8');

    for (const slideNum of sorted) {
      const idx = slideNum - 1;
      if (idx < 0 || idx >= slideRefs.length) continue;
      const { full, rId } = slideRefs[idx];

      // Find filnavn fra rels: Target="slides/slideN.xml"
      const relMatch = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"[^>]*/>`));
      if (!relMatch) continue;
      const target = relMatch[1]; // fx "slides/slide3.xml"
      const slidePath = `ppt/${target}`;
      const slideRelsPath = `ppt/slides/_rels/${path.basename(target)}.rels`;

      // Fjern slide-filen og dens .rels
      zip.deleteFile(slidePath);
      zip.deleteFile(slideRelsPath);

      // Fjern reference i presentation.xml
      presXml = presXml.replace(full, '');

      // Fjern reference i .rels
      relsXml = relsXml.replace(relMatch[0], '');

      // Fjern fra [Content_Types].xml
      const ctEntry = zip.getEntry('[Content_Types].xml');
      if (ctEntry) {
        let ctXml = ctEntry.getData().toString('utf8');
        const ctPath = `/${slidePath}`;
        ctXml = ctXml.replace(new RegExp(`<Override[^>]*PartName="${ctPath.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
        zip.updateFile('[Content_Types].xml', Buffer.from(ctXml, 'utf8'));
      }

      slideRefs.splice(idx, 1); // Opdater lokalt array
    }

    zip.updateFile('ppt/presentation.xml', Buffer.from(presXml, 'utf8'));
    zip.updateFile('ppt/_rels/presentation.xml.rels', Buffer.from(relsXml, 'utf8'));
    zip.writeZip(pptxPath);
    console.log(`Slettede slides: ${slidesToDelete.join(', ')}`);
  } catch (e) {
    console.error("deleteSlides fejlede:", e);
  }
}

// --- SLET TABELLER VIA NAVN I XML ---
// tablesToDelete: array af tabelnavne som de er sat i PowerPoint, fx ["tabel_affald", "tabel_bio"]
// Tabeller navngives i PowerPoint: Hjem → Arranger → Markeringsrude → omdøb elementet
function deleteTables(pptxPath, tablesToDelete) {
  if (!tablesToDelete || tablesToDelete.length === 0) return;
  try {
    const zip = new AdmZip(pptxPath);
    const entries = zip.getEntries();

    for (const entry of entries) {
      if (!entry.entryName.match(/^ppt\/slides\/slide\d+\.xml$/)) continue;

      let xml = entry.getData().toString('utf8');
      let changed = false;

      for (const tableName of tablesToDelete) {
        const escaped = tableName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        // En tabel i PPTX sidder i et <p:sp> eller <p:graphicFrame> element
        // med et <p:cNvPr name="TABELNAVN"> attribut
        // Vi matcher hele det omsluttende element og fjerner det

        // Forsøg 1: <p:graphicFrame> (tabeller bruger typisk dette)
        const gfRegex = new RegExp(
          `<p:graphicFrame>(?:(?!<p:graphicFrame>)[\\s\\S])*?<p:cNvPr[^>]*name="${escaped}"[\\s\\S]*?<\\/p:graphicFrame>`,
          'g'
        );
        if (gfRegex.test(xml)) {
          xml = xml.replace(gfRegex, '');
          changed = true;
          continue;
        }

        // Forsøg 2: <p:sp> (shapes/tekstbokse med tabel-lignende indhold)
        const spRegex = new RegExp(
          `<p:sp>(?:(?!<p:sp>)[\\s\\S])*?<p:cNvPr[^>]*name="${escaped}"[\\s\\S]*?<\\/p:sp>`,
          'g'
        );
        if (spRegex.test(xml)) {
          xml = xml.replace(spRegex, '');
          changed = true;
        }
      }

      if (changed) {
        zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
      }
    }

    zip.writeZip(pptxPath);
    console.log(`Slettede tabeller: ${tablesToDelete.join(', ')}`);
  } catch (e) {
    console.error("deleteTables fejlede:", e);
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

    const { template_url, placeholders: requestPlaceholders, rows, company_unique_id, company_name, delete_slides, delete_tables, enable_table_deletion } = body;
    const normalizedRows = normalizeRows(rows);
    const parsedPlaceholderValue = parseJsonValue(requestPlaceholders);
    const placeholderRows = normalizeRows(parsedPlaceholderValue);
    const placeholders = buildPlaceholdersFromRows(normalizedRows)
      || buildPlaceholdersFromRows(placeholderRows)
      || normalizeDelimitedPlaceholders(tryParseObject(parsedPlaceholderValue));

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

    const arrayPlaceholders = {};
    const numericKeys = new Set();
    for (const [key, value] of Object.entries(placeholders)) {
      if (isEmpty(value)) continue;
      const arr = tryParseList(value);
      if (arr && arr.length > 1) {
        arrayPlaceholders[key] = arr;
        if (arr.some(isNumericValue)) numericKeys.add(key);
      } else if (isNumericValue(value)) {
        numericKeys.add(key);
      }
    }

    // --- SKRIDT 2: FLET POWERPOINT VIA AUTOMIZER ---
    // Keep the original PPTX when no table needs row expansion. This preserves
    // tables, charts, and other unsupported PowerPoint elements exactly.
    if (false) {
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

        for (const slide of slides) {
          pres.addSlide('base', slide.number, async (s) => {
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
                        const allRows = tblEl.elements.filter(el => el.name === 'a:tr' || el.name === 'tr');
                        resizeExpandedTable(xmlData, allRows);
                      }
                    }
                  }

                  return element;
                });
              } catch (tableModError) {
                console.error(`Fejl ved tabelmanipulation af ${tableName}:`, tableModError);
              }

              // Table XML is modified above. Do not pass the entire table through
              // the text-only modifier, because it can remove table structure.
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

    expandArrayTablesInXml(outputPath, placeholders);

    // --- SKRIDT 2.5: XML-NIVEAU CLEANUP ---
    // Fjerner alle tilbageværende {{...}} og "null"-værdier direkte i PPTX-XML
    cleanupResidualPlaceholders(outputPath, placeholders, numericKeys);

    // --- SKRIDT 2.6: SLET SLIDES OG TABELLER ---
    // delete_slides: [1, 3, 5]  — 1-baserede slide-numre
    // delete_tables: ["tabel_affald", "tabel_bio"] — navne sat i PowerPoint
    const slidesToDelete = tryParseArray(delete_slides) || (Array.isArray(delete_slides) ? delete_slides : []);
    // Tables must be preserved; table deletion is disabled for this generator.
    const tablesToDelete = [];
    deleteSlides(outputPath, slidesToDelete);
    deleteTables(outputPath, tablesToDelete);

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
        const raw = c?.sharedTo?.shareLink || c?.shareLink || c?.link || "";
        if (raw && /\/s\/[^/?#]+/.test(raw) && !/^about:blank$/i.test(raw)) return raw;
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
