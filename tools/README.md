# Building index.html

`index.html` is generated. Edit `src/app.html`, then rebuild:

```
python3 tools/build.py src/app.html - src/master-v2.json /tmp/preview.html index.html
```

- `src/app.html` — the dashboard (markup, styles, script). `<!--BODY-->` splits head from body.
- No shipment data is baked in: this repository is public. Data arrives through the encrypted Sync (data.json) after login.
- `src/master-v2.json` — Master Data v2 as parsed by the dashboard's own reader (`readMaster`). Regenerate it whenever `Juhayna-Master-Data-v2.3.xlsx` changes.
- `tools/build_master_v2.py` — one-off migration that produced `Juhayna-Master-Data-v2.xlsx` from the v1 master.
- `tools/update_master_v2_1.py` — adds dated rates (price list +10% from 10 Mar 2026), invoice-based rates, new routes/vendors/SpPI, SAP equipment numbers and the Stoppage Reasons tab from the Jan–Sep invoice extract and the truck stoppage log.
- `tools/update_master_v2_2.py` — adds owned routes found in new owned shipment files, with estimated weight, RT km and incentives.
- `tools/update_master_v2_3.py` — adds the Extra Off Days tab (target = month days − Fridays − extra off days), makes Calendar a reference list, derives vendor rents from the MR report.
