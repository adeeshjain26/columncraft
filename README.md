# ColumnCraft — Excel Column Extractor

A browser-based spreadsheet column extractor built with Next.js, TypeScript, and SheetJS. All workbook parsing and export happens in the browser; spreadsheet contents are not sent to a server.

## Supported input formats

- `.xlsx` — Excel Open XML workbook
- `.xls` — legacy Excel workbook
- `.xlsm` — macro-enabled Excel workbook (macros are not preserved in the extracted output)
- `.xlsb` — Excel binary workbook
- `.csv` — comma-separated values
- `.tsv` — tab-separated values
- `.ods` — OpenDocument spreadsheet

## Output formats

- `.xlsx` (default)
- `.xls` (Excel 97–2003)
- `.csv`
- `.tsv`

CSV and TSV contain a single table, so they do not have multiple worksheets. Workbook inputs with multiple sheets show a worksheet selector. Output includes the selected columns and non-empty data rows. Complex workbook formatting, macros, pivot tables, and other advanced Excel features are not preserved by this extraction workflow.

## Run locally

Requirements: Node.js 20 or newer recommended.

```bash
npm install
npm run dev
```

Open http://localhost:3000.

## Build

```bash
npm run build
npm start
```

## Deploy to Vercel

Push the project to a Git repository, import it at https://vercel.com/new, and deploy with the default Next.js settings. No environment variables or backend are required.

## Privacy

Files are read with the browser File API and parsed in client-side JavaScript. No upload endpoint is included. Hosting still requires internet to load the site unless a separate PWA/offline-caching layer is added.
