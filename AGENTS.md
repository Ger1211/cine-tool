# AGENTS.md

## Project overview

Client-side web app that automates HR payroll prep for a cinema complex. Processes clock-in data (Nelclock) and work schedules (Roster Excel) to generate two Excel files: Planilla de Ausencias and Trasnoches AVEL.

- **Zero-install**: single HTML file + vanilla JS + CSS, no build step, no server, no framework
- **All processing happens in the browser** — employee data never leaves the machine
- **No npm needed to run** — `npm`/`node_modules` exists only for parser testing via Node.js

## How to run

Open `index.html` directly in a browser, or serve the root directory with any static server:

```
npx serve .
```

## Architecture

```
index.html           — UI shell, loads all libs in order
css/style.css        — all styles
js/
  parser-nelclock.js — fixed-width TXT parser (Nelclock "Reporte municipal")
  parser-roster.js   — XLSX roster parser via SheetJS
  engines.js         — puntualidad, trasnoches, 4ta-jornada logic
  generator.js       — builds output Excel files (SheetJS)
  app.js             — IIFE controller, DOM wiring, localStorage persistence
```

### Load order matters

`index.html` loads scripts in this order at the bottom of `<body>`:
1. `parser-nelclock.js`
2. `parser-roster.js`
3. `engines.js`
4. `generator.js`
5. `app.js`

Each module attaches to the global `window` (e.g., `NelclockParser`, `RosterParser`, `Engines`, `Generator`). SheetJS (`XLSX`) is loaded from CDN before all of them.

### External dependency

SheetJS 0.20.3 from CDN:
```html
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
```

The app won't work offline unless you vendor this file.

- **Zero-install**: single HTML file + vanilla JS + CSS, no build step, no server, no framework
- **All processing happens in the browser** — employee data never leaves the machine
- **No npm needed to run** — `npm`/`node_modules` exists only for parser testing via Node.js

## How to run

Open `index.html` directly in a browser, or serve the root directory with any static server:

```
npx serve .
```

## Architecture

```
index.html           — UI shell, loads all libs in order
css/style.css        — all styles
js/
  parser-nelclock.js — fixed-width TXT parser (Nelclock "Reporte municipal")
  parser-roster.js   — XLSX roster parser via SheetJS
  engines.js         — puntualidad, trasnoches, 4ta-jornada logic
  generator.js       — builds output Excel files (SheetJS)
  app.js             — IIFE controller, DOM wiring, localStorage persistence
```

### Load order matters

`index.html` loads scripts in this order at the bottom of `<body>`:
1. `parser-nelclock.js`
2. `parser-roster.js`
3. `engines.js`
4. `generator.js`
5. `app.js`

Each module attaches to the global `window` (e.g., `NelclockParser`, `RosterParser`, `Engines`, `Generator`). SheetJS (`XLSX`) is loaded from CDN before all of them.

### External dependency

ExcelJS 4.4.0 from CDN:
```html
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
```

The app won't work offline unless you vendor this file. ExcelJS replaced SheetJS — it is used for both reading (roster) and writing (output Excel with full formatting: fonts, colors, borders, merges, number formats).

## Key data flow

1. User drags/drops two files: Nelclock TXT export + Roster XLSX
2. `NelclockParser.parse()` → `{ fichadas: [{legajo, nombre, apellido, entrada, salida}, ...] }`
3. `RosterParser.parse()` → array of `{ nombre, tipo (FT/PT), dias: [{fecha, entrada, salida, ...}] }`
4. `Engines.puntualidad()` / `trasnoches()` / `cuartaJornada()` cross-reference both
5. User fills manual forms (novedades, horas extras) — persisted to localStorage
6. `Generator` produces two Excel Blobs for download

## Roster parser fragility

The roster is an Excel template reused since 2023. Key facts:

- **Multiple sheets per sector** (Candy, Piso, Proyección). Sheets named `"vacaciones"` and `"pedidos"` are skipped by default
- **Week blocks are detected by scanning for `STAFF` / `FT/PT` marker cells** across rows 3-5. Each block spans 2 ID columns + 14 data columns (7 days × 2: entry + exit)
- **Dates come from Excel formulas** (e.g., `=C4+1`) — `getCellValue()` resolves them via the `result` property
- **Time values are stored as Excel epoch Dates** (1899-12-30) — `cellToString()` detects and converts to `HH:MM:SS`
- **Employee rows repeat the employee name at the start of each week block** within the same sheet row (horizontal repetition)
- If the roster template changes layout, `findWeekBlocks()` in `parser-roster.js:125` is the first place to debug
- Sheets with legacy dates (2023, 2024, 2025) are filtered by `targetMonth` param (YYYY-MM format)

## Nelclock TXT format

Fixed-width text export from Nelclock's "Reporte municipal de empleados". Sample (`ejemplos/Marcelo V (1).txt`):

```
              00011046         DNI     39744545               20397445454            Vandecaveye            Marcelo Alejandro      01/11/25 15:30:00             01/11/25 23:08:00
```

Parsed via regex in `parser-nelclock.js:15`. Format: `dd/mm/yy hh:mm:ss`. Dates after 2000 assumed.

Critical: file is read with **ISO-8859-1** encoding (`app.js:493`) because the TXT contains Spanish characters (e.g. "Página") in legacy encoding.

Note: the export may include employees with missing exit time (still clocked in). Parser handles `salida: null`.

## Name matching

Roster stores names as `"Nombre Apellido"`, Nelclock stores them as `apellido` + `nombre` separately. `engines.js:220-261` (`indexarFichadas` + `buscarFichadas`) does fuzzy matching:

1. Indexes fichadas by both `"apellido nombre"` and `"nombre apellido"`
2. Falls back to word-level containment matching

If an employee shows "Sin fichada" incorrectly, check name format in both sources.

## Test data

`ejemplos/` contains real (anonymized) sample files for development testing:
- `Marcelo V (1).txt` — single-employee Nelclock export (November 2025)
- `ROSTER JULIO 26.xlsx` — full roster with 11 sheets, legacy data mixed in
- `Planilla de Ausencias Julio 2026.xlsx` — expected output format (4 sheets)
- `Trasnoches AVEL JUL 26.xlsx` — expected trasnoches output format
- `andy ayuda rrhh.xlsx` — consolidated helper sheet (not input, reference only)
- `Fichadas de Marcelo (1).xlsx` — same as TXT but opened in Excel (misleading extension, content is fixed-width)

To test with Node.js: `node -e "const XLSX = require('xlsx'); ..."` (SheetJS is in `node_modules`).

## Output files

Two Excel files are generated:
1. `Planilla de Ausencias [Mes] [Año].xlsx` — 4 sheets: Novedades, 4ta jornada, Horas Extras, Puntualidad
2. `Trasnoches AVEL [MES_ABBR] [YY].xlsx` — 1 sheet: TRASNOCHES

Both attempt to match the exact layout of the existing manual templates.
