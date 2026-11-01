"""Master v2.1 -> v2.2: add owned routes found in the Jan–Sep owned shipments, with estimated weight, RT km and incentives.

usage: python3 -I update_master_v2_2.py MASTER_V2_1.xlsx MISSING.json OUT.xlsx
MISSING.json: {route name: trips} (from the dashboard's M4 issue)
"""
import sys, re, json, statistics as st
from datetime import date
import openpyxl
from openpyxl.styles import PatternFill, Font

SRC, MISS, OUT = sys.argv[1:4]
EST = PatternFill("solid", fgColor="DDEBF7"); EST_FONT = Font(color="1F4E79", italic=True); NEW = PatternFill("solid", fgColor="FCE4D6")
rnorm = lambda s: re.sub(r"\s*([-+])\s*", r"\1", re.sub(r"\s+", " ", str(s or "")).strip().lower())
missing = json.load(open(MISS, encoding="utf-8"))
wb = openpyxl.load_workbook(SRC)

# spot-route km (current period rows)
sp = wb["Spot-Rent Route Master"]; SH = {str(c.value).strip(): c.column for c in sp[3] if c.value}
spot_km = {}
for r in range(4, sp.max_row + 1):
    n = sp.cell(r, SH["Route"]).value
    if n and sp.cell(r, SH["Valid to"]).value is None and sp.cell(r, SH["RT km"]).value:
        spot_km[rnorm(n)] = float(sp.cell(r, SH["RT km"]).value)
SYN = [(r"mid ?alex", "mid alex"), (r"alex ?east", "alex east"), (r"\balex\b", "alex"), (r"beh[ei]?ra|bahera|behera", "beheira"), (r"men[i]?a|mania", "menia"),
       (r"ass[iu]+[iu]?t", "assiut"), (r"b(ani|eni|\.)? ?s[uw]?[ie]+[fv]|bani swef", "bani swef"), (r"so?hag|suhag", "sohag"), (r"kena|qena|luxor", "kena (luxor)"),
       (r"as+wan", "aswan"), (r"matrouh", "marsa matrouh"), (r"tour|eltor", "el-tour"), (r"hurghada", "hurghada"), (r"isma[i]?l[i]?a", "ismailia"), (r"port ?said", "port said"),
       (r"damiat", "mansoura-islamic dist. (موزع دمياط)"), (r"zagazig|shar[k]?[iy]a", "zagazig"), (r"monof", "monofya"), (r"tanta", "tanta"), (r"mans[o]?ura", "mansoura"),
       (r"kal[iy]o?bia|qalyobia", "kalyobia"), (r"khos", "khosos"), (r"sawah|swaah", "elswaah"), (r"obour", "el-obour"), (r"warr?ak", "alwarrak"), (r"\bauc\b", "auc"),
       (r"wadi ?ho+f", "wadi hofe"), (r"suez", "suez"), (r"zayed", "zayed"), (r"dab[a]+", "eldaba"), (r"fast trade", "factory to dist. fast trade"),
       (r"trinity south", "factory to dis trinity south"), (r"sedeek", "factory to dis el sedeek"), (r"al batal", "factory to dist. al batal"), (r"hani adel", "factory to dis hani adel")]

# existing owned routes
ow = wb["Owned Route Master"]; OH = {str(c.value).strip(): c.column for c in ow[3] if c.value}
rows = []
for r in range(4, ow.max_row + 1):
    n = ow.cell(r, 1).value
    if not n: continue
    g = lambda h: ow.cell(r, OH[h]).value
    core = re.sub(r"^(juh(ayna)?)\s*-?\s*", "", str(n).strip(), flags=re.I)
    rows.append(dict(name=n, w=g("Weight (days)"), km=g("RT km"), i1=g("Incentive single (EGP)"), i2=g("Incentive double (EGP)"),
                     stops=max(1, len([p for p in re.split(r"\s*-\s*", core) if p.strip()]))))
ratio = st.median(x["i2"] / x["i1"] for x in rows if x["i1"] and x["i2"])
last = ow.max_row
added = []
for name, trips in sorted(missing.items(), key=lambda x: -x[1]):
    core = re.sub(r"^(juh(ayna)?)\s*-?\s*", "", name.strip(), flags=re.I)
    stops = [p for p in re.split(r"\s*-\s*", core) if p.strip()]
    kms = []
    for stp in stops:
        sn = stp.lower()
        for pat, tgt in SYN:
            if re.search(pat, sn) and rnorm(tgt) in spot_km:
                kms.append(spot_km[rnorm(tgt)]); break
    if kms:
        km = max(kms) * (1 + 0.1 * (len(kms) - 1)); kb = f"Estimated: max of {len(kms)} matched stop(s) from Spot-Rent km, +10% per extra stop"
    else:
        km = st.median(x["km"] for x in rows if x["w"] == 1 and x["km"]); kb = "Estimated: no stop matched — median km of weight-1 routes"
    km = int(round(km / 5) * 5)
    peers = [x["w"] for x in rows if x["km"] and x["w"] and abs(x["km"] / km - 1) <= .2]
    w = round(st.median(peers) * 2) / 2 if len(peers) >= 3 else 1
    ns = max(1, len(stops))
    bucket = [x["i1"] for x in rows if x["w"] == w and x["stops"] == ns and x["i1"]]
    if len(bucket) < 3: bucket = [x["i1"] for x in rows if x["w"] == w and x["i1"]]
    i1 = int(round(st.median(bucket) / 5) * 5); i2 = int(round(i1 * ratio / 5) * 5)
    last += 1
    vals = {"Route": name.strip(), "Weight (days)": w, "RT km": km, "Incentive single (EGP)": i1, "Incentive double (EGP)": i2, "Plant shuttle (Y/N)": "N",
            "Incentive status": "Estimated", "km status": "Estimated", "km basis": kb,
            "Comment": f"NEW from Jan–Sep owned shipments ({trips} trips): weight = median of routes within ±20% km; incentive = median of weight-{w} routes with {ns} stop(s), double = single × {ratio:.2f}",
            "Last updated": date(2026, 11, 1)}
    for h, col in OH.items():
        c = ow.cell(last, col, vals.get(h)); c.fill = NEW
        if h in ("Weight (days)", "RT km", "Incentive single (EGP)", "Incentive double (EGP)"): c.fill = EST; c.font = EST_FONT
        if h == "Last updated": c.number_format = "yyyy-mm-dd"
    added.append((name.strip(), trips, w, km, i1, i2))
rm = wb["READ ME"]; rm.cell(rm.max_row + 1, 1, f"UPDATE 1 Nov 2026 (v2.2): {len(added)} owned routes added from the Jan–Sep owned shipments with estimated weight, RT km and incentives (orange rows, blue values).")
wb.save(OUT)
for a in added: print(f"  {a[1]:3} trips  {a[0][:44]:44} w={a[2]} km={a[3]} single={a[4]} double={a[5]}")
