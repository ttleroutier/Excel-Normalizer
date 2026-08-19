# Excel File Normalization

A lightweight web tool to **normalize the content of Excel/CSV files** directly in the browser.
No data is sent to a server: all processing happens **locally**.

## Features

- Import `.xlsx`, `.xls`, `.xlsm`, `.csv` files
- Fine-grained selection of **sheets** and **columns** to modify (nothing is modified by default)
- Normalization options:
  - whitespace handling (keep / trim edges / replace with `_`)
  - capitalize the first letter
  - keep the first row (headers) unchanged
  - terms always kept uppercase (e.g. `LHD`, `RHD`)
- Per-column type conversion: Automatic, Text, Number, Date, Boolean
- Per-column case handling (keep / lowercase / UPPERCASE)
- **Business rules** for find & replace (with "whole word" and "case sensitive" options)
- Result preview (first 100 rows) before export
- Export the normalized file as `.xlsx`

## Usage

1. Open `index.html` in a modern browser.
2. Select a source file.
3. Check the sheets and columns to process, then adjust the options and rules.
4. Click **Traiter les colonnes sélectionnées** (Process selected columns).
5. Review the preview, then **download the normalized file**.

## Project Structure
normalisation-excel/
├── index.html # Page structure
├── css/
│ └── styles.css # Styles
└── js/
└── app.js # Application logic

## Dependencies

- [SheetJS (xlsx)](https://sheetjs.com/) — loaded via CDN, no installation required.

## Hosting (GitHub Pages)

This is a fully static project, so it can be published for free.
1. Push the repository to GitHub.
2. **Settings → Pages → Source: branch `main` / folder `/root`**.
3. The site becomes available at the URL provided by GitHub Pages.

## Privacy

Files never leave your machine: reading, transformation, and export
are all performed in the browser.

## License

To be defined (e.g. MIT).

