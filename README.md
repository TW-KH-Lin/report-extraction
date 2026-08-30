# Report Extraction — Pure Web Version

This version requires **no Python and no software installation on the user's laptop**.

It is a static browser application. It can be hosted on:

- GitHub Pages
- Cloudflare Pages
- Netlify
- Vercel static hosting
- any ordinary internal web server

## Privacy model

The application has no backend and uses a restrictive browser content-security policy.

Complaint PDF / MSG / ZIP files and the Excel workbook are processed with JavaScript inside the browser. The code does not intentionally upload those files to GitHub, the hosting service, or an API.

All required browser libraries are pinned and stored in the repository under `vendor/`. The deployed app makes no third-party library or API connections.

## How to deploy with GitHub Pages

1. Create a new GitHub repository, for example `membrane-complaint-web`.
2. Upload these files and folders to the repository root:
   - `index.html`
   - `app.js`
   - `styles.css`
   - `vendor/`
3. In GitHub, open **Settings → Pages**.
4. Choose deployment from the repository branch (normally `main`) and root folder.
5. Save.
6. GitHub will provide a web address such as:
   `https://<your-account>.github.io/membrane-complaint-web/`

No Python is required to use the deployed page.

## Recommended company-safe deployment

Do **not** put a workbook containing real complaint/customer information in the GitHub repository.

Instead:
1. host only the app source files and bundled `vendor/` libraries;
2. open the webpage;
3. optionally upload an existing Excel workbook from your computer;
4. upload one or more complaint reports;
5. review extracted information;
6. click **Create Excel file**.

The Excel file is generated locally in the browser and downloaded back to the user's computer. If no workbook is selected, the app creates a new workbook from the reviewed report extractions.

## Current features

- PDF complaint report extraction
- Outlook `.msg` extraction in the browser
- ZIP containing PDF/MSG
- Complaint number
- Complaint registered date, report date and elapsed days
- Material number
- Membrane type (including backed/unbacked when recoverable)
- product-family mapping:
  - `1UN95` → CN95
  - `1UN14AR` → CN140ub
  - `1UN14ER` → CN140
  - `1UN11` → CN110
  - `1UN18` → CN180
- Lot
- Problem
- Standardized symptoms, problem type and LFA relevance
- Confirmed / Not confirmed / Not conclusive when recoverable from PDF text
- Master Roll / Final Roll / MR-FR parsing
- Zone and combined MR-FR fields for summary-workbook transfer
- Product description and customer-reported failure
- Coordinator / report author
- Recurrence, containment and corrective-action decisions
- Failure reproduction and process-related root-cause decisions
- Final assessment and scope decision
- Root-cause conclusion cross-check for the problem description
- Broad formal issue labels are enriched with the customer statement and root-cause terminology
- Structured test evidence with result, outcome, specification status and source page
- Dedicated `Extracted Test Evidence` sheet aligned to the detailed investigation columns
- Manual review/edit before export
- Lot-history lookup in the uploaded Excel and current extraction
- Unique complaint count for the same lot
- Symptom-frequency summary without double-counting complaints repeated across workbook sheets
- `Lot & Symptom Summary` worksheet in generated Excel files
- One complaint number = one row
- Updates three category sheets
- Rebuilds CN95 / CN140ub / CN140 / CN110 / CN180 sheets
- Product-family sheets sorted by Problem

## Limitations of this first browser-only prototype

1. Scanned image-only PDFs need browser OCR; OCR is not yet included.
2. PDF checkbox extraction can vary depending on how the report was generated.
3. Email threads can contain old quoted complaint text; `.msg` results must be reviewed.
4. Excel formatting is intentionally simplified when the app rewrites the managed sheets.
5. It is rule-based extraction, not an AI model.

## Best next upgrade

Add a browser-side review screen showing:
- extracted value
- confidence
- report page / source sentence

Then add optional browser OCR for scanned PDFs.

## Bundled third-party libraries

Pinned browser builds are stored under `vendor/`. See `THIRD_PARTY_NOTICES.md` and `vendor/licenses/` for versions and license texts.
