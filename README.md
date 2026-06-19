# bubble-pptx-generator — API Dokumentation

Backend til automatisk generering af ESG-rapporter i PPTX-format. Modtager data fra Bubble, fletter det ind i en PowerPoint-skabelon, uploader til OnlyOffice DocSpace og returnerer et direkte editor-link.

---

## Opsætning

### Dependencies (`package.json`)
```json
{
  "dependencies": {
    "pptx-automizer": "latest",
    "axios": "^1.6.0",
    "form-data": "^4.0.0",
    "adm-zip": "^0.5.10"
  }
}
```

### Miljøvariabler (Vercel)
| Variabel | Beskrivelse |
|---|---|
| `DOCSPACE_URL` | Base-URL til OnlyOffice DocSpace, fx `https://ditdomain.onlyoffice.com` |
| `DOCSPACE_TOKEN` | Bearer token til DocSpace API |
| `DOCSPACE_FOLDER_ID` | ID på den mappe rapporterne uploades til |

---

## Endpoint

```
POST /api/generate
Content-Type: application/json
```

---

## Request Body

### Top-niveau felter

| Felt | Type | Påkrævet | Beskrivelse |
|---|---|---|---|
| `template_url` | string | ✅ | URL til PPTX-skabelon, fx raw GitHub URL |
| `placeholders` | object | ✅ | Nøgle/værdi-par der erstattes i slides |
| `company_unique_id` | string | ❌ | Firmaets unikke ID — returneres uændret i svaret |
| `company_name` | string | ❌ | Bruges i filnavnet, fx `f3r5_TestCompany.pptx` |
| `delete_slides` | array/string | ❌ | Liste af slide-numre der skal slettes |
| `delete_tables` | array/string | ❌ | Liste af tabelnavne der skal slettes |

### Eksempel på fuld request
```json
{
  "template_url": "https://raw.githubusercontent.com/MajsMan2/bubble-pptx-generator/main/report_basic_dk_standard.pptx",
  "company_unique_id": "firma-abc-123",
  "company_name": "TestCompany",
  "delete_slides": "[3, 7]",
  "delete_tables": "[\"tabel_affald\", \"tabel_bio\"]",
  "placeholders": {
    "company_name": "TestCompany ApS",
    "insert_year": "2024",
    "revenue": "5000000",
    "total_employees": "42"
  }
}
```

---

## Placeholders

Placeholders i skabelonen kan skrives på to måder — begge understøttes:
- `company_name` (uden tuborg-klammer)
- `{{company_name}}` (med tuborg-klammer)

### Array-værdier (tabelrækker)
Hvis en placeholder skal fylde flere rækker i en tabel, sendes værdien som en JSON-array:

```json
"Type_af_affald": "[\"Papir\", \"Plast\", \"Metal\"]"
```

Backenden duplikerer automatisk template-rækken i tabellen — én række per element.

### Tomme/manglende værdier
Følgende værdier behandles som tomme og gør placeholderen usynlig i den færdige rapport:
- `null`
- `"null"`
- `""` (tom streng)
- `undefined`
- `"<feltnavn>"` (uudfyldt Bubble-placeholder)

Placeholders der slet ikke er med i `placeholders`-objektet fjernes også automatisk via catch-all cleanup.

---

## Slet slides (`delete_slides`)

Sender en liste af **1-baserede** slide-numre der skal fjernes fra rapporten.

```json
"delete_slides": "[1, 3, 5]"
```

- Slides slettes bagfra så nummereringen ikke forskydes
- Slide-filen, dens `.rels`-fil og referencerne i `presentation.xml` og `[Content_Types].xml` fjernes alle
- Valgfrit felt — udelad det hvis ingen slides skal slettes

---

## Slet tabeller (`delete_tables`)

Sender en liste af **tabelnavne** der skal fjernes fra slides.

```json
"delete_tables": "[\"tabel_affald\", \"tabel_bio\"]"
```

Tabeller navngives i PowerPoint:
> **Hjem → Arranger → Markeringsrude → dobbeltklik på elementet og omdøb**

- Fungerer på både `<p:graphicFrame>` (tabeller) og `<p:sp>` (shapes)
- Søger på tværs af alle slides
- Valgfrit felt — udelad det hvis ingen tabeller skal slettes

---

## Svar (Response)

```json
{
  "success": true,
  "company_unique_id": "firma-abc-123",
  "fileId": "4399296",
  "fileName": "f3r5_TestCompany.pptx",
  "fileUrl": "https://docspace.../doceditor?fileId=4399296&share=TOKEN&action=edit&type=desktop",
  "shareLink": "https://docspace.../s/TOKEN",
  "debugInfo": "OK (POST /file/:id/link (Edit))"
}
```

| Felt | Beskrivelse |
|---|---|
| `success` | `true` hvis alt gik godt |
| `company_unique_id` | Samme ID som sendt i request — bruges til at matche svaret til det rigtige firma i Bubble |
| `fileId` | OnlyOffice fil-ID |
| `fileName` | Filnavnet på den genererede rapport |
| `fileUrl` | Direkte editor-link — åbner filen i OnlyOffice uden login |
| `shareLink` | Råt share-link fra DocSpace |
| `debugInfo` | Status/fejlbesked fra share-link-oprettelsen |

### Fejlsvar
```json
{
  "error": "Beskrivelse af fejlen",
  "message": "Teknisk fejlbesked",
  "rawBody": "..." 
}
```
`rawBody` vises kun ved JSON-parse-fejl og viser de første 300 tegn af det der kom ind — nyttigt til at debugge Bubble-request.

---

## Pipeline (rækkefølge internt)

```
1. Download PPTX-skabelon fra template_url
2. Flet placeholders via pptx-automizer
   └─ Array-værdier duplikerer tabelrækker automatisk
3. XML-niveau cleanup (adm-zip)
   ├─ Saml splittede placeholders på tværs af XML-noder
   ├─ Erstat alle kendte placeholders
   ├─ Catch-all: fjern resterende {{...}} mønstre
   └─ Fjern "null" der står alene som celleindhold
4. Slet slides (delete_slides)
5. Slet tabeller (delete_tables)
6. Upload til OnlyOffice DocSpace
7. Opret offentligt eksternt link med edit-adgang
   ├─ Forsøg 1: POST /files/file/{id}/link
   ├─ Forsøg 2: GET /files/file/{id}/link
   ├─ Forsøg 3: PUT /files/file/{id}/links
   └─ Forsøg 4: GET /files/file/{id}/links
8. Byg editor-URL med share-token
9. Returner JSON-svar til Bubble
```

---

## Skabelon-URL (GitHub)

Raw GitHub URL til skabelon:
```
https://raw.githubusercontent.com/MajsMan2/bubble-pptx-generator/main/report_basic_dk_standard.pptx
```

---

## Kendte begrænsninger

- PowerPoint splitter sommetider placeholdernavne over flere XML-noder internt. Cleanup-funktionen håndterer dette, men meget kompleks formatering kan i sjældne tilfælde give uventede resultater.
- Tabelrække-duplikering understøtter kun ét array-felt per tabel ad gangen.
- Sletning af slides opdaterer ikke automatisk interne krydsreferencer (fx slide-numre i tekstfelter i skabelonen).
