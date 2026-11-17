# Juhayna Fleet Control Tower — v2

Master-data-driven dashboard: owned-truck KPIs, driver incentives, vendor analytics, cost estimates.
All Excel parsing happens in the browser — shipment files are never uploaded to any server.

## Deploy / update (GitHub Pages)
1. Upload `index.html` to the repo root (overwrite the old one) — Settings → Pages → Deploy from branch `main` /root if first time.
2. URL: `https://<account>.github.io/<repo>/` — hard-refresh after updates (add `?v=N`).

Sign-in: the password is checked by decrypting the shared `data.json` (AES-256-GCM, PBKDF2-SHA-256 600,000 iterations, random salt per save).
No password or password hash is stored in this public repository — never write the password in any file here.
To change it: sign in, then press **🔑 Change password** (needs the writer token on that browser). It re-encrypts `data.json` with the new password in one step.

## Daily use (planner)
Open URL → login → choose files:
1 · Owned raw (KPI tracker export — Shipment Update layout, headers row 4)
2 · Spot/Rent raw (SAP extract — headers row 2)
3 · Master data (`Juhayna-Master-Data-v2.3.xlsx`, kept privately by the team — not in this public repo; "download current" exports the master that is in the shared data)
→ RUN. With a writer token configured (⚙ Sync), data commits encrypted to `data.json`; every device sees it with the last-update time.

## Master data maintenance
- Owned Route Master: route → weight (days), single/double incentive.
- Truck Driver Master: truck ↔ driver, Temp D/C, Active flag.
- Spot-Rent Route Master: route → weight, RT km, spot rates dry/chilled (amber = estimates; confirm & set Source=Confirmed).
- Vendor Master: monthly rent dry/chilled, contracted MR trucks (drives MR utilization %).
Add rows freely — the dashboard reads whole columns. Definitions: Short trip <1 day, Medium =1, Long >1; performance = points ÷ (trucks × working days excl. Fridays); truck mix High >95% / Med 90–95% / Low <90%.

## Sync (shared data across devices)
Writer: ⚙ Sync → repo `account/repo` → fine-grained token (only this repo, Contents: Read & write). 🔧 Test explains any failure in plain English.
Viewers: nothing to configure when data.json is in this same repo.
Every push = a git commit → free versioned history.

## Export
⬇ Export Excel downloads a workbook of all tables exactly as filtered on screen.
