# Juhayna Fleet Control Tower

Static, single-file dashboard for owned-truck KPIs, driver incentives and vendor analytics.
**All Excel parsing happens in the user's browser — no shipment data is ever uploaded to GitHub or any server.**

## Publish on GitHub Pages (5 minutes)
1. Create a repository (e.g. `juhayna-dashboard`) — **Private repo works with Pages on paid plans; public repo is fine because the file contains no shipment data if you remove the preloaded snapshot (see below).**
2. Upload `index.html` to the repository root.
3. Settings → Pages → Source: `Deploy from a branch` → Branch `main` / root → Save.
4. URL: `https://<account>.github.io/juhayna-dashboard/` — share with planners.

Login: user `admin` — password set at build time (SHA-256 hash embedded; to change it, hash `user|password` and replace AUTH_HASH in index.html).

## Daily use
Open URL → login → choose the two Excel files (KPI tracker + SAP vendor extract) → RUN.

## Data & history — how it works
- GitHub only hosts the app. Excel files are **not** uploaded; they are read locally by the browser.
- The file ships with a baked demo/snapshot dataset (`const REAL=`). To publish without any real data, replace it with empty arrays: `const REAL={owned:[],vendor:[],names:{}}`.
- **History options** (no paid storage needed):
  a) Commit a monthly JSON snapshot to the repo (free, versioned by git).
  b) Browser localStorage per device (works on GitHub Pages).
  c) Free-tier backend (Google Sheets API / Supabase) if multi-user history is required later.

## Config
Thresholds, fuel price, km table: edit the `COSTMODEL` block at the top of index.html.

## Shared data across devices (data.json in the repo)
- Planner (writer): open dashboard → **⚙ Sync setup** → enter `account/repo` and a fine-grained token (repo access, permission **Contents: Read & write**). After every RUN, parsed data is committed to `data.json` automatically.
- Everyone else (viewers): press ⚙ Sync setup, enter the repo, leave token empty — or nothing at all if `data.json` sits next to index.html on Pages. On open they see the shared data and "last update" timestamp.
- Every upload is a git commit → free versioned history of all past uploads.
- Note: on a public repo, `data.json` is publicly readable. Use a private repo + Pages (paid) or keep the repo public only if shipment data is acceptable to expose; token itself is never stored in the repo.

## Data-leakage protection (free)
`data.json` is now **AES-256-GCM encrypted in the browser before upload**, key derived (PBKDF2, 100k iterations) from the login password. A public repo therefore exposes only ciphertext. Decryption happens after login on any device. Changing the password = changing the encryption key (update AUTH_HASH and re-push data once). Recommended hygiene: keep the data repo private anyway (free), rotate the password quarterly, and scope the writer token to Contents-only on that single repo.
