# Конвертор на поръчки Excel → CSV

Web app (React + Vite + Tailwind + SheetJS) that converts shop order files (`.xls` / `.xlsx`) into CSV rows of the form:

```
supplier_code,,quantity
```

Everything runs in the browser. There is no backend, and the original file is never modified.

## Running locally

Requires Node.js 20+.

```bash
npm install
npm run dev
```

Open the URL Vite prints (usually http://localhost:5173).

Other commands:

```bash
npm test         # converter tests (uses samples/*.xls)
npm run build    # production build in dist/
npm run preview  # serve the production build
```

## Deploying to GitHub Pages

The site deploys automatically with GitHub Actions ([.github/workflows/deploy.yml](.github/workflows/deploy.yml)) on every push to `main`. You can also start it manually from the **Actions** tab ("Run workflow").

### One-time setup

1. Create a repository on GitHub, `FuewI/excel-to-csv`, and push this project to the `main` branch.
2. In the repository, open **Settings → Pages**.
3. Under **Build and deployment → Source**, choose **GitHub Actions**.
4. Open the **Actions** tab. The "Deploy to GitHub Pages" workflow runs on the next push; re-run it if it ran before Pages was enabled.
5. When it finishes, the site is live at **https://fuewi.github.io/excel-to-csv/**. The link is also shown on the workflow run.

### How it works

- **Base path:** `vite.config.js` sets `base: '/excel-to-csv/'` for production builds, matching the repository name. If the repository is renamed, change it there too. `npm run dev` still uses `/`.
- **No 404 on refresh:** this is a single page without client-side routes. The workflow also copies `index.html` to `404.html`, so any unknown URL under the site still loads the app instead of GitHub's 404 page. `.nojekyll` is added so Pages serves the files as they are.
- **Client-side only:** GitHub Pages hosts static files only. Uploading, parsing, CSV and XLSX generation and downloads all run in the browser, and order files never leave the computer.
- **Tests:** the workflow runs `npm test` before building, and a failing test stops the deploy.

### Sample order files are not committed

`.gitignore` excludes `*.xls` and `*.xlsx`, because real orders contain company names, VAT numbers and addresses, and a GitHub Pages repository on a free plan is public. Tests that need `samples/Order…039011.xls` are skipped when the file is missing, for example in GitHub Actions. They still run locally.

## Sample file structure (`samples/Order…039011.xls`)

One sheet (`Sheet1`), range A1:I25:

| Location | Content |
|---|---|
| A1 (merged A1:G1) | `ЗАЯВКА N 3900074104/05.11.2026` |
| A2–A6 | Ordering company (`От фирма: КОМЕ ООД`, VAT, address…) |
| F2 | Supplier: `Към фирма: ЕКСПРЕС ДИСТРИБЮШЪН-Д.МАДЖАРОВ-2-МЛЕЧНИ` ← **category** |
| A7 | `Магазин за доставка: КОМЕ КНЯЖЕВО` ← **shop name** |
| Row 10 | Table header: `№, Код, Баркод, Име, МЕ, Код на дост., Пакетаж, Кол., ТМ` |
| F11:F19 | **Supplier code** (text, trailing spaces) |
| H11:H19 | **Quantity** (number, format `0.0000`, e.g. `3.0000` → `3`) |
| A21 | `Общо: 135.46` (total, excluded) |
| A22–A25 | `Забележка:`, `Референция:` (notes, excluded) |

The source file labels the supplier-code column `Код на дост.`. The converter accepts this, the full `Код на доставчика`, and other abbreviations. It never accepts plain `Код`, which is the internal article code.

## How detection works (`src/lib/converter.js`)

- **Header row:** found by scanning the whole sheet for the supplier-code header, then looking for `Кол.` in the same row or up to 2 rows above or below. Merged header cells are supported. If the workbook has several sheets, the first one containing the header is used.
- **Product rows:** rows below the header that have both a supplier code and a numeric quantity. Blank rows, repeated headers, total rows (`Общо`, `Сума`, …) and note rows are skipped. A row with a code but no quantity, or a quantity but no code, is **not** written. Such rows are listed as warnings in the UI.
- **Values:** supplier codes are kept as text, so leading zeros survive and surrounding spaces are trimmed. Quantities keep their numeric value (`3.0000` → `3`, `1.5` → `1.5`).
- **Shop name:** a label such as `Магазин за доставка:`, `Магазин:`, `Обект:` or `Получател:`, with the value either after the colon or in the next cell. If no label is found, the first descriptive text above the table is offered and marked as uncertain. The field can always be edited manually.
- **Category:** `Млечни`, or `Месни`/`Местни` (both mapped to `Месни`), found anywhere outside the product rows. The match covers standalone cells and suffixes such as `…-МЛЕЧНИ`. If neither appears, no category is used. If both appear, the user has to choose.

## Output formats

Every format has no header row and exactly three fields per row: A = supplier code, B = empty, C = quantity.

| Format | Delimiter | Encoding | First line | Use for |
|---|---|---|---|---|
| **Excel CSV** (default) | `;` | UTF-8 with BOM | `sep=;` | Double-clicking in Excel (Windows/Mac) |
| Standard CSV | `,` | UTF-8, no BOM | — | Importing into other systems |
| Excel (.xlsx) | — | — | — | Exact display in Excel. Column A is a text cell (`@` format), column C is a number |

```
sep=;            107106,,3
107106;;3        116125,,8
116125;;8
```

- Line endings are CRLF. A value is quoted only if it contains the delimiter, a quote or a line break.
- The preview on the website shows data rows only, without the `sep=;` line.
- **Note:** when Excel opens *any* CSV it treats numeric-looking codes as numbers. It hides leading zeros and shows 13-digit codes as `3,8E+12`. The file itself keeps the exact codes. For a correct view inside Excel, use **„Изтегли Excel (.xlsx)“**.
- File names: `Магазин.csv`, `Магазин_Млечни.csv`, `Магазин_Месни.csv` (and `.xlsx`). Characters that are not allowed in file names are removed.
