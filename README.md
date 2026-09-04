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

1. Create a new GitHub repository, for example `report-extraction`.
2. Upload these files and folders to the repository root:
   - `index.html`
   - `app.js`
   - `styles.css`
   - `vendor/`
3. In GitHub, open **Settings → Pages**.
4. Choose deployment from the repository branch (normally `main`) and root folder.
5. Save.
6. GitHub will provide a web address such as:
   `https://<your-account>.github.io/report-extraction/`

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
- Add reports in later selections without replacing earlier extracted complaints
- Upload an existing Excel workbook and several reports together, then download an updated workbook
- Review every extracted complaint as one complete editable row across three focused tabs: Complaint overview, Complaint investigation, and Tests & root cause
- The Complaint overview uses compact two-line headings and narrower columns; longer values wrap and expand so information is not hidden
- The Complaint overview omits Source Group; records from `Ongoing - Email` display `Ongoing` in Final Result / Status
- Compact typography is applied consistently across all app tabs
- Complaint Number, Lot Number and Customer keep the same first-three-column layout in all review tabs
- Soft alternating case colors make separate complaints easier to follow when switching tabs
- Customer country codes include `GR` for Greece, `UK` for the United Kingdom and `CA` for Canada
- End-customer names are standardized to `Key name, City, Country code` where the location can be identified
- Membrane labels use only the normalized family name, including `CN140ub` for the unbacked CN140 family
- Structured Test Evidence appears as a compact editable one-test-per-row table; complaint, lot and customer are stacked in one shared details column, and test methods are always formatted as bullet points
- In-process data review and batch record review are standardized as `IPW review` throughout the app and exported workbooks
- Click **Save changes** to keep edited review and test-evidence values for the next Excel export
- Dates in the organized overview use year-month-day format
- Export the three organized review tabs to separate Excel worksheets
- Re-extract the same complaint to update its existing row instead of creating a duplicate
- Save reviewed results temporarily in the same browser and restore them after refresh
- Review a live lot summary table before export
- Choose individual app-managed worksheets to include in the Excel download
- Separate Workbook Summary tab that combines complaint information across relevant worksheets in multiple uploaded Excel files
- Summary and Quick Search can also read PDF, MSG and ZIP report files without adding those files to the repository
- Selectable summary columns, including complaint, lot, related-lot count, customer, reason, final result, tests and dates
- **Show lot details** automatically opens a grouped view showing each lot's complaint count and the related complaint numbers, customers, reasons and results
- Lot-detail groups use alternating background colors, show product family after the lot, and can be filtered instantly by CN type
- One-row-per-complaint consolidation when the same complaint appears on several product or category sheets
- Original Quick Search for lot or complaint number remains available, with optional product-family filtering
- Additional linked Quick Search in the order: customer, membrane type, complaint number
- Quick Search complaint-number choices show the associated end-customer name beside each number
- Quick Search also finds similar complaints by membrane family and symptom, then shows related complaint numbers, lots, historical tests and final results
- Built-in, read-only Master–Final Roll Plan reference with seven ZM worksheet choices, 50 master-roll rows, final-roll positions and colored zones
- Workbook summary and search also include reports currently extracted in the browser
- Complaint number
- Complaint registered date, report date and elapsed days
- Material number
- Membrane type normalized to CN95 / CN110 / CN140 / CN140ub / CN180
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
- Dedicated `Complaint Summary` sheet with one reusable row per complaint for later overview and search
- Legacy-format CN sheets are preserved unchanged so historical rows are not removed; new records remain available in the standardized summary and category sheets
- Manual review/edit before export
- Matching Complaint Number, Lot Number and Customer column widths across all three review tabs
- Compact two-line headers across all three review tables
- Contained investigation fields with narrow Rolls Implicated and Samples Received columns
- Full-width investigation and root-cause layouts without horizontal table scrolling
- Auto-growing multi-line MR-FR Area(s) fields for longer entries
- Auto-growing multi-line Standardized Symptom(s) fields
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
