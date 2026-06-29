# Excel Workbench PWA

A browser-based workbench for understanding, filtering, and lightly transforming Excel files — locally, offline, and without touching a server.

Built for the moments when Excel feels too heavy, too awkward, or simply unavailable — especially on tablets and PWA-style workflows where macros are not a realistic option.

## Screenshots

![Excel Workbench PWA screenshot 1](./docs/images/excel-workbench-pwa-screenshot-1.png)
Main view with the sidebar collapsed.

![Excel Workbench PWA screenshot 2](./docs/images/excel-workbench-pwa-screenshot-2.png)
Main view with the sidebar expanded.

## Why This Exists

This project came out of a real workflow need, not a generic idea for "Excel in the browser."

In practice, a lot of Excel work is not really about spreadsheet authoring. It is about:

- making sense of messy, deeply nested workbooks quickly
- filtering and comparing data faster than standard Excel flows allow
- doing analytical work safely on a tablet or in a browser
- replacing some macro-heavy steps with simpler, browser-native tools

Classic Excel is powerful, but it is also awkward for inspection workflows, unreliable on tablet, and out of reach once you need anything beyond basic filtering without VBA.

Excel Workbench PWA fills that gap — without claiming to be something it is not.

**The goal is not to clone Excel.**  
The goal is to build a workbench *around* Excel files: local-first, safe for source files, genuinely useful on a tablet, and better at inspection and structure discovery than Excel's own tooling.

## What It Does

### File handling
- Open `.xlsx` and `.xlsm` files in the browser — no upload, no backend
- Drag and drop or file picker
- Auto-detect header rows, choose any sheet, configure starting column
- Save modified workbooks back to `.xlsx` or export to CSV

### Filtering and search
- Two independent text filters with multiple match modes (contains, starts with, equals)
- Date filter with presets (last N days, custom range) and column selection
- Quick search with highlight or filter mode and configurable column scope
- Empty / non-empty filtering per column
- Negate any filter

### Sorting and working views
- Multi-column sort with priority order
- Saved sort presets for quick switching between common working states
- Frozen header row
- Adjustable column widths and zoom level

### Workbook inspection
- Sheet-level structural summary: column count, row count, data layout, header guess
- Column profiler: data type, fill rate, numeric stats, date ranges per column
- Section navigator for visually segmented sheets
- Repeated block detector — identifies cycling column patterns (`od`, `do`, `od 2`, `do 2`, etc.)
- KPI extractor for dashboard-style or summary-heavy sheets

### Aggregation workbench
- Pivot-table-inspired grouping by up to three levels
- Multiple simultaneous measures: count, sum, average, median, min, max, distinct count, earliest, latest
- Source mode control — classic, Wide-to-Long, or auto
- HAVING filter and result search
- Scoped to current filtered view or the full sheet

### Wide-to-Long transformation
- Detect wide operational sheets with repeating column cycles
- Transform into a long-format view on the fly for cleaner aggregation and filtering

### Duration analysis
- Auto-detect date-pair columns (`od` / `do`, start / end, from / to)
- Compute durations, averages, and status breakdowns per entity
- Built for operational and process-tracking sheets

### Formula workbench
- Browse all formulas across the sheet in one panel
- Filter by function type, error type, or free-text search
- Navigate directly to any formula location in the table

### UI and experience
- Full Polish / English UI with live language switching
- Light and dark mode
- Responsive layout — comfortable on both desktop and tablet
- Smart cursor hints system
- Glassmorphic visual design with depth and motion

## Privacy and Data Safety

Workbook files are processed entirely inside the browser.

No file data is sent to a server. No accounts. No cloud storage. The app is designed so your Excel files stay on your device at every step — the only dependencies are the browser and a trusted build of the app.

## Offline

The app ships with a service worker and works fully offline after the first successful load. A background update mechanism checks for new versions and notifies you when one is available.

## How It Is Built

Excel Workbench PWA is built with vanilla JavaScript — no framework, no build step, no bundler. The core XLSX parsing library (`xlsx-js-style`) is bundled locally so the app can work fully offline without relying on a CDN.

The codebase is split into focused modules (file IO, analysis engine, UI controls, language, table rendering, formula workbench) and shares a single global state with clearly defined boundaries. Playwright smoke tests cover core UI flows.

## Start Locally

```bash
python3 -m http.server 8001
```

Then open `http://127.0.0.1:8001/`.

## Deploy

Dev: serwuj repo root (`npm run serve` lub `python3 -m http.server`).

Produkcja z minifikacją: `npm run build` → serwuj katalog `dist/` (Vercel: build command `npm run build`, output `dist`).

Bez buildu nadal działa serwowanie źródeł z root (jak dotąd).

## Install as PWA

On iPad / iPhone: open in Safari → Share → Add to Home Screen.

On desktop: look for the install icon in the browser address bar.

## Public Roadmap

Planned work and longer-term ideas live in [ROADMAP.md](./ROADMAP.md).

## Contributing

Contributions, bug reports, UX suggestions, and workbook-based edge cases are welcome.

Please read [CONTRIBUTING.md](./CONTRIBUTING.md) before contributing.

## License

Source code is available under the `MIT` license.

Exception: the project logo, branding elements, screenshots, visual identity, and other presentation assets remain the property of Mateusz Zmuda and are not intended for reuse as ready-made branding or product presentation materials without explicit permission.
