# Building index.html

`index.html` is generated. Edit `src/app.html`, then rebuild:

```
python3 tools/build.py src/app.html src/snapshot-2026-08.json src/master-v2.json /tmp/preview.html index.html
```

- `src/app.html` — the dashboard (markup, styles, script). `<!--BODY-->` splits head from body.
- `src/snapshot-2026-08.json` — the August 2026 data preloaded for first open (old format; shared data from Sync replaces it after login).
- `src/master-v2.json` — Master Data v2 as parsed by the dashboard's own reader (`readMaster`). Regenerate it whenever `Juhayna-Master-Data-v2.xlsx` changes.
- `tools/build_master_v2.py` — one-off migration that produced `Juhayna-Master-Data-v2.xlsx` from the v1 master.
