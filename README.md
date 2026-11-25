# SAP KSA Youth Council · Impact & Execution Dashboard

Static dashboard (HTML, CSS, JavaScript; no build step) showing the council's initiatives, their impact against the three strategic pillars, and delivery progress from Microsoft Planner.

**This repository is public. It holds anonymised sample data only.** Team members appear as "Member A–J". Do not commit real names, Planner exports or the Council Data workbook.

## Files
| File | Purpose |
|---|---|
| `index.html` | Page shell: sidebar, header, dialog |
| `style.css` | Layout, colours, responsive rules |
| `app.js` | Routing, filters, calculations, metric explanation cards |
| `data.js` | Snapshot data (sample, 25 Nov 2026) and metric definitions |
| `logo.jpg` | Council logo |
| `netlify.toml` | Publishes the repository root as-is, no build command |
| `robots.txt` | Asks search engines not to index the site |

## Deploy
Netlify deploys automatically from `main`. To preview locally, open `index.html` in a browser or run `python3 -m http.server 8000`.

## Change process
Changes are previewed and approved before they are pushed to `main`.
