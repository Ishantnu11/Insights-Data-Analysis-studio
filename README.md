# InsightForge — Data Analysis Studio (local edition)

A single-file React + Framer Motion web app for data analysis and dashboards
(Power BI / Tableau style). No build step, no install, no backend.

## 1. Run it

**Easiest:** double-click `index.html` (Chrome, Edge, Firefox, Safari).

**Recommended** (avoids some browser restrictions on `file://`):

```bash
cd insightforge-local
python3 -m http.server 8080      # or: npx serve .
# open http://localhost:8080
```

You need an internet connection the first time: React, Framer Motion, Babel,
SheetJS, jsPDF, JSZip and the Plus Jakarta Sans font load from CDNs. For fully
offline use, download these files, put them in a `libs/` folder and change the
`<script src>` tags in `index.html` to point to them:

| Library | Version | URL |
|---|---|---|
| React | 18.2.0 | https://cdnjs.cloudflare.com/ajax/libs/react/18.2.0/umd/react.production.min.js |
| ReactDOM | 18.2.0 | https://cdnjs.cloudflare.com/ajax/libs/react-dom/18.2.0/umd/react-dom.production.min.js |
| Framer Motion | 10.16.4 | https://cdn.jsdelivr.net/npm/framer-motion@10.16.4/dist/framer-motion.js |
| Babel standalone | 7.23.9 | https://cdnjs.cloudflare.com/ajax/libs/babel-standalone/7.23.9/babel.min.js |
| SheetJS (xlsx) | 0.18.5 | https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js |
| jsPDF | 2.5.1 | https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js |
| JSZip | 3.10.1 | https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js |

## 2. Features

| Area | What you get |
|---|---|
| **Home** | Mosaic overview: row/column counts, live top-category bars, quick "ask" box |
| **Import** | CSV, TSV (paste from Google Sheets/Excel), JSON, Excel `.xlsx/.xls` (extra sheets become join/stack sources), free text |
| **Clean** | Remove duplicates, drop blank rows, trim text, fill blanks, to number, rename/delete column, keep rows where…, Undo (10 steps) |
| **Prepare** | Calculated columns (e.g. `Profit / Sales * 100`), joins (left/inner), stack/append datasets |
| **Profile** | Types, unique, missing, min / mean / median / std / max, outlier count, correlation matrix, auto-insights, searchable + sortable data grid |
| **Visuals** | Bar, line, area, pie, scatter, histogram, stacked bar, heatmap, box plot, funnel, KPI (with target), table |
| **Time analysis** | Running total, 3-point moving average, period-over-period %, year-over-year % |
| **Stats aggregates** | sum, avg, count, max, min, median, std, p90 |
| **Interactivity** | Slicer, date range, click a bar/slice to cross-filter, move / widen / drag-resize visuals |
| **Pivot** | Row × column cross-tab with totals and heat map |
| **Text analysis** | Keyword frequency and rule-based sentiment (positive / neutral / negative) |
| **AI chat** | Rule-based parser works everywhere (see §4 for the LLM mode) |
| **Output** | Download all (ZIP: updated CSV + Excel, filtered CSV, KPI CSV, chart PNGs, PDF, HTML report), Share report (standalone HTML), per-chart PNG |
| **Save** | "Save layout" stores dashboard + data (up to 20,000 rows) in the browser's localStorage |

## 3. Daily workflow

1. **Data & Profile** → upload or paste your data. Check the profile (missing values, outliers).
2. **Clean** (duplicates, blanks, types) → add **calculated columns** → **join** lookup tables.
3. **Dashboard** → add visuals with the builder, set a KPI target, add slicer/date range.
4. **Pivot** for ad-hoc cross-tabs.
5. **Download all (ZIP)** or **Share report** to hand results to others.
6. **Save layout** to resume later (same browser).

## 4. AI chat in local mode

The hosted version asks Claude (via the Claude artifact runtime) and gives it
tools to query your full dataset. That runtime does not exist locally, so:

* The **✦ Claude AI** checkbox starts **off**; the built-in rule-based parser
  handles requests like `average profit by region pie chart`, `top 5 categories
  by sales`, `total sales kpi`, `scatter units vs sales`.
* To use an LLM locally, define `window.claude` **before** the app script runs.
  The app calls `window.claude.use('sample')` and then
  `sample.json(prompt, { tools })`. A minimal shim backed by **your own
  server** (never put an API key in front-end code):

```html
<script>
window.claude = { use: async (name) => {
  if (name !== 'sample') return null;
  const sample = async () => ({ text: '' });
  sample.json = async (prompt /*, opts */) => {
    const r = await fetch('/api/ask', { method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ prompt }) });      // your server calls the Anthropic API
    return await r.json();                       // must return the JSON the prompt asks for
  };
  return sample;
}};
</script>
```

  Without `opts.tools` support in your shim, the prompt still contains
  pre-computed totals per category, so most "which X is highest" questions
  work; arbitrary queries need you to implement the `aggregate` / `find_rows`
  tool loop on your server.

## 5. Limits & notes

* Runs entirely in the browser — best for up to tens of thousands of rows.
* Charts show the top 24 groups; scatter shows 300 points.
* Calculated columns and joins are not covered by Undo; deleting a column can
  break visuals that use it. "Load sample" keeps existing visuals.
* No live database / Google Sheets connection — paste or upload exports.
* Year-over-year needs 13+ periods of data.
* Sentiment uses a small English word list — a quick signal, not a model.
* Not included: maps, combo charts, multiple dashboard pages, multi-filter slicers, drill-down hierarchies.

## 6. File layout

```
insightforge-local/
├── index.html   # the whole app (HTML + CSS + React/JSX)
└── README.md    # this file
```

The code lives in one `<script type="text/babel">` block. Main parts: data
helpers (`toRows`, `profile`, `group`, `AG`), `Chart` (SVG renderer),
`Grid`, `Pivot`, `Home`, and `App` (state, tabs, handlers).
