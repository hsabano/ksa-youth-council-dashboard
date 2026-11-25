# SAP KSA Youth Council · Impact & Execution Dashboard

Static dashboard (HTML, CSS, JavaScript; no build step) showing the council's initiatives, their impact against the three strategic pillars, and delivery progress from Microsoft Planner.

**This repository is public. It holds anonymised sample data only.** Team members appear as "Member A–J". Do not commit real names, Planner exports or the Council Data workbook.

## Files
| Path | Purpose |
|---|---|
| `site/` | The website Netlify publishes: `index.html`, `style.css`, `app.js`, logo and icons |
| `site/data-sample.js` | Fully filled sample data set (placeholder names), built by `tools/make_sample.py` |
| `site/data.js` | Optional static live data; `null` here. Live data normally comes from uploads (below) |
| `site/upload.js`, `site/convert.js` | Upload Data page: reads and checks the two Excel files in the browser |
| `site/vendor/xlsx.full.min.js` | SheetJS 0.18.5 (Apache-2.0) for reading Excel files |
| `site/Council_Data_Template.xlsx` | Blank Council Data template offered for download |
| `netlify/functions/` | `data.mts` serves live data (`GET /api/data`); `upload.mts` publishes it (`POST /api/upload`, passcode) |
| `tools/` | Offline converter (`build_data.py`) and sample builder (`make_sample.py`) |

## Uploading live data
Sidebar footer › **Upload data**. Choose the Planner export and the Council Data workbook; the page checks both and lists anything that must be fixed. When all checks pass, enter the upload passcode and publish. Live data is stored in Netlify Blobs, never in this repository. The passcode is the `UPLOAD_PASSCODE` environment variable in Netlify (scope: Functions).

## Sample / live toggle
The switch at the top of every page chooses the data set. The choice is remembered per browser. **Live** is disabled when `data.js` holds no data.

## Deploy
Netlify deploys automatically from `main`. To preview locally, open `index.html` in a browser or run `python3 -m http.server 8000`.

## Refreshing with real data
1. Export the Planner plan to Excel (Planner › … › Export plan to Excel).
2. Update the Council Data workbook (Metrics month column, Participants, Survey Results).
3. Run `python3 tools/build_data.py PLANNER_EXPORT.xlsx COUNCIL_DATA.xlsx -o data.js` (needs `pip install openpyxl`). To rebuild the sample, run `python3 tools/make_sample.py BLANK_TEMPLATE.xlsx`.
4. Fix anything listed under "Data health", then rerun.

**A data.js built from real data contains real names and results. Do not commit it while this repository is public.**

## Change process
Changes are previewed and approved before they are pushed to `main`.
