"""Master v2 -> v2.1: dated rates (price change 10 Mar 2026), invoice-based rates, new routes/vendors/SpPI,
SAP equipment numbers, Stoppage Reasons tab, invoice parameters.

usage: python3 -I update_master_v2_1.py MASTER_V2.xlsx INVOICES.xlsx STOPPAGES.xlsx OUT.xlsx
"""
import sys, re, collections, statistics as st
from datetime import date, datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from copy import copy

MV2, INV, STP, OUT = sys.argv[1:5]
CUT = date(2026, 3, 10)               # first day of the current price list (from invoices)
OLD_TO = date(2026, 3, 9)
UPLIFT = 1.10
TODAY = date(2026, 11, 1)
EST = PatternFill("solid", fgColor="DDEBF7"); EST_FONT = Font(color="1F4E79", italic=True)
NEW = PatternFill("solid", fgColor="FCE4D6"); NAVY = PatternFill("solid", fgColor="13294B")
norm = lambda s: re.sub(r"\s+", " ", str(s or "")).strip().lower()
D = lambda v: v.date() if isinstance(v, datetime) else v

wb = openpyxl.load_workbook(MV2)
report = collections.defaultdict(list)

def headers(ws):
    return {str(c.value).strip(): c.column for c in ws[3] if c.value}

# ---------------- invoices -> shipment totals ----------------
iws = openpyxl.load_workbook(INV, read_only=True, data_only=True)["Sheet1"]
it = iws.iter_rows(values_only=True); hdr = list(next(it)); ix = hdr.index
ship = {}
for r in it:
    if not r[0]: continue
    s = ship.setdefault(r[0], {"net": 0.0, "n": 0})
    s.update(route=str(r[ix("Shipment route")]).strip(), spi=r[ix("SpPI")], d=D(r[ix("PlDat.ChIn")]),
             vendor=re.sub(r"\s+", " ", str(r[ix("Number of forwarding agent")]).strip()), shty=r[ix("ShTy")])
    s["net"] += r[ix("Net value")] or 0; s["n"] += 1
SP_TEMP = {}
sws_sp = wb["SpPI Master"]
for r in sws_sp.iter_rows(min_row=4, values_only=True):
    if r[0]: SP_TEMP[str(r[0]).strip().upper()] = str(r[1]).strip().upper()
temp_of = lambda spi: SP_TEMP.get(str(spi).upper(), "C" if str(spi).upper().startswith("C") else "D")
act = collections.defaultdict(list)   # (route_norm, temp, period) -> nets
names = {}
for s in ship.values():
    k = norm(s["route"]); names.setdefault(k, s["route"])
    act[(k, temp_of(s["spi"]), "old" if s["d"] < CUT else "new")].append(s["net"])
med = lambda k: (round(st.median(act[k]), 2), len(act[k])) if len(act.get(k, [])) >= 3 else (None, len(act.get(k, [])))

# ---------------- SpPI: codes seen in invoices ----------------
have = set(SP_TEMP)
seen = collections.Counter(str(s["spi"]).upper() for s in ship.values())
r = sws_sp.max_row
for code, n in sorted(seen.items()):
    if code in have: continue
    r += 1; m = re.match(r"([A-Z])(\d+)", code)
    vals = [code, "C" if code.startswith("C") else "D", "Double" if m and int(m.group(2)) >= 40 else "Single", int(m.group(2)) if m else None,
            "Estimated", f"Added from Jan–Sep invoices ({n} shipments): temperature from first letter, pallets from digits — confirm"]
    for j, v in enumerate(vals, 1):
        c = sws_sp.cell(r, j, v); c.fill = NEW
    report["SpPI added"].append(f"{code} ({n} shipments)")

# ---------------- Spot-Rent Route Master: dated rates ----------------
ws = wb["Spot-Rent Route Master"]; H = headers(ws)
cols = ["Route", "Weight (days)", "RT km", "Rate dry (EGP/trip)", "Rate chilled (EGP/trip)", "Valid from", "Valid to",
        "Rate status dry", "Rate status chilled", "km status", "Notes", "Last updated"]
old_rows = []
for rr in range(4, ws.max_row + 1):
    if ws.cell(rr, 1).value: old_rows.append({h: ws.cell(rr, H[h]).value for h in cols} | {"_fill_km": True})
existing = {norm(x["Route"]) for x in old_rows}

def rate_pair(k, cur, status, temp):
    """returns (old_value, old_status, old_note, new_value, new_status, new_note)"""
    if status == "Confirmed" and cur:
        return (round(cur / UPLIFT, 2), "Confirmed", f"price list ÷ {UPLIFT}", cur, "Confirmed", "price list")
    mo, no_ = med((k, temp, "old")); mn, nn_ = med((k, temp, "new"))
    if mn is None and mo is not None: mn, nn_src = round(mo * UPLIFT, 2), f"old-period median × {UPLIFT}"
    else: nn_src = f"median of {nn_} invoices" if mn is not None else None
    if mo is None and mn is not None: mo, no_src = round(mn / UPLIFT, 2), f"current ÷ {UPLIFT}"
    else: no_src = f"median of {no_} invoices" if mo is not None else None
    if mn is None:            # no invoices at all: keep engine estimate
        if not cur: return (None, "Estimated", "", None, "Estimated", "")
        return (round(cur / UPLIFT, 2), "Estimated", f"engine estimate ÷ {UPLIFT}", cur, "Estimated", "engine estimate (no invoices)")
    return (mo, "Estimated", no_src, mn, "Estimated", nn_src)

out = []
for x in old_rows:
    k = norm(x["Route"])
    d = rate_pair(k, x["Rate dry (EGP/trip)"], x["Rate status dry"], "D")
    c = rate_pair(k, x["Rate chilled (EGP/trip)"], x["Rate status chilled"], "C")
    if x["Rate status dry"] != "Confirmed" and d[3] and x["Rate dry (EGP/trip)"] and d[3] != x["Rate dry (EGP/trip)"]: report["estimated rate replaced by invoice median"].append(f"{x['Route']} dry {x['Rate dry (EGP/trip)']:,.0f} → {d[3]:,.0f}")
    if x["Rate status chilled"] != "Confirmed" and c[3] and x["Rate chilled (EGP/trip)"] and c[3] != x["Rate chilled (EGP/trip)"]: report["estimated rate replaced by invoice median"].append(f"{x['Route']} chilled {x['Rate chilled (EGP/trip)']:,.0f} → {c[3]:,.0f}")
    base = dict(x)
    out.append(base | {"Rate dry (EGP/trip)": d[0], "Rate chilled (EGP/trip)": c[0], "Valid from": date(2026, 1, 1), "Valid to": OLD_TO,
                       "Rate status dry": d[1], "Rate status chilled": c[1], "Notes": f"before 10 Mar 2026 · dry: {d[2] or '—'} · chilled: {c[2] or '—'}", "_period": "old"})
    out.append(base | {"Rate dry (EGP/trip)": d[3], "Rate chilled (EGP/trip)": c[3], "Valid from": CUT, "Valid to": None,
                       "Rate status dry": d[4], "Rate status chilled": c[4], "Notes": f"from 10 Mar 2026 · dry: {d[5] or '—'} · chilled: {c[5] or '—'}", "_period": "new"})
# new routes from invoices
newk = sorted({norm(s["route"]) for s in ship.values()} - existing)
for k in newk:
    n_all = sum(len(act.get((k, t, p), [])) for t in "DC" for p in ("old", "new"))
    d = rate_pair(k, None, "Estimated", "D"); c = rate_pair(k, None, "Estimated", "C")
    base = {"Route": names[k], "Weight (days)": 1, "RT km": None, "km status": "Estimated", "Last updated": TODAY, "_new": True, "_fill_km": True}
    out.append(base | {"Rate dry (EGP/trip)": d[0], "Rate chilled (EGP/trip)": c[0], "Valid from": date(2026, 1, 1), "Valid to": OLD_TO, "Rate status dry": "Estimated", "Rate status chilled": "Estimated",
                       "Notes": f"NEW from invoices ({n_all} shipments) · before 10 Mar · dry: {d[2] or '—'} · chilled: {c[2] or '—'}", "_period": "old"})
    out.append(base | {"Rate dry (EGP/trip)": d[3], "Rate chilled (EGP/trip)": c[3], "Valid from": CUT, "Valid to": None, "Rate status dry": "Estimated", "Rate status chilled": "Estimated",
                       "Notes": f"NEW from invoices ({n_all} shipments) · from 10 Mar · dry: {d[5] or '—'} · chilled: {c[5] or '—'}", "_period": "new"})
    report["routes added from invoices"].append(f"{names[k]} ({n_all})")
# km for new routes: long-haul model from existing km vs current dry/chilled rate, else nearest-priced median
pts = [(x["Rate dry (EGP/trip)"], x["RT km"]) for x in out if x.get("_period") == "new" and not x.get("_new") and x["RT km"] and x["Rate dry (EGP/trip)"] and x["Rate dry (EGP/trip)"] >= 9000]
n = len(pts); sx = sum(p[0] for p in pts); sy = sum(p[1] for p in pts); sxx = sum(p[0] ** 2 for p in pts); sxy = sum(p[0] * p[1] for p in pts)
b = (n * sxy - sx * sy) / (n * sxx - sx * sx); a = (sy - b * sx) / n
short = sorted((x["Rate dry (EGP/trip)"], x["RT km"]) for x in out if x.get("_period") == "new" and not x.get("_new") and x["RT km"] and x["Rate dry (EGP/trip)"] and x["Rate dry (EGP/trip)"] < 9000)
newkm = {}
for x in out:
    if x.get("_new") and x["_period"] == "new":
        rate = x["Rate dry (EGP/trip)"] or ((x["Rate chilled (EGP/trip)"] or 0) / 1.12)
        if rate >= 9000: km, basis = a + b * rate, "price model (long-haul fit)"
        elif rate > 0: km, basis = st.median([p[1] for p in sorted(short, key=lambda p: abs(p[0] - rate))[:3]]), "nearest-priced routes"
        else: km, basis = 250, "default"
        newkm[norm(x["Route"])] = (max(30, int(round(km / 5) * 5)), basis)
for x in out:
    if x.get("_new"):
        x["RT km"], basis = newkm[norm(x["Route"])]; x["Notes"] += f" · km {basis}"

# rewrite the tab
ws.delete_rows(4, ws.max_row)
for i, x in enumerate(out, 4):
    for j, h in enumerate(cols, 1):
        c = ws.cell(i, j, x.get(h))
        if h in ("Valid from", "Valid to", "Last updated") and c.value: c.number_format = "yyyy-mm-dd"
        if x.get("_new"): c.fill = NEW
        if h == "RT km" or (h == "Rate dry (EGP/trip)" and x["Rate status dry"] != "Confirmed") or (h == "Rate chilled (EGP/trip)" and x["Rate status chilled"] != "Confirmed"):
            c.fill = EST; c.font = EST_FONT
ws.cell(2, 1).value = "Two validity periods per route: before and from 10 Mar 2026 (price list increase of 10%, found in the Jan–Sep invoices). Add a new row with new dates when the price list changes."

# ---------------- Vendor Contracts: vendors seen in invoices ----------------
vc = wb["Vendor Contracts"]; VH = headers(vc)
known = {re.sub(r"\s+", " ", str(vc.cell(r, 1).value).strip()) for r in range(4, vc.max_row + 1) if vc.cell(r, 1).value}
cnt = collections.Counter(s["vendor"] for s in ship.values())
r = vc.max_row
for v, nn in cnt.most_common():
    if v in known: continue
    r += 1
    vals = {"Vendor": v, "Valid from": date(2026, 1, 1), "Valid to": None, "Rent dry (EGP/truck/month)": None, "Rent chilled (EGP/truck/month)": None,
            "Contracted trucks": 0, "VAT applies (Y/N)": "Y", "Fuel paid by (Juhayna/Vendor)": "Juhayna", "Status": "Spot only", "Notes": f"added from Jan–Sep invoices ({nn} shipments); add rent if they ever run MR"}
    for h, col in VH.items():
        c = vc.cell(r, col, vals.get(h)); c.fill = NEW
    vc.cell(r, VH["Valid from"]).number_format = "yyyy-mm-dd"
    report["vendors added (spot only)"].append(f"{v} ({nn})")

# ---------------- Truck Master: SAP equipment number ----------------
sw = openpyxl.load_workbook(STP, read_only=True, data_only=True).active
eq = {}
for row in sw.iter_rows(min_row=2, values_only=True):
    if row[0] is None: continue
    m = re.findall(r"\d{3,}", str(row[1]))
    if m: eq.setdefault(m[0], (row[0], str(row[1]).strip()))
tm = wb["Truck Master"]; TH = headers(tm)
col = tm.max_column + 1
h = tm.cell(3, col, "SAP equipment"); h.fill = NAVY; h.font = Font(bold=True, color="FFFFFF")
tm.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 14
plates = set()
for rr in range(4, tm.max_row + 1):
    p = str(tm.cell(rr, 1).value or "").strip()
    if not p: continue
    plates.add(p)
    if p in eq: tm.cell(rr, col, eq[p][0])
    else: report["trucks without SAP equipment in stoppage log"].append(p)
for p, (e, desc) in eq.items():
    if p not in plates: report["stoppage-log trucks not in Truck Master"].append(f"{desc} (equipment {e})")

# ---------------- Stoppage Reasons ----------------
MAP = [  # code text, category, counts as stoppage, note
    ("نقلات سفر", "On trip", "N", "Truck working (travel trips) — not a stoppage"),
    ("مبيت بالفرع", "Waiting at branch", "Y", "Overnight at branch / drop point"),
    ("صيانة", "Maintenance", "Y", ""),
    ("عمرة", "Maintenance", "Y", "Overhaul"),
    ("حادثة", "Maintenance", "Y", "Accident repair"),
    ("ترخيص", "License", "Y", "Licence renewal"),
    ("رخصه مسحوبه", "License", "Y", "Licence withdrawn"),
    ("عدم وجود شهادة مؤسسه", "License", "Y", "Missing company certificate"),
    ("عدم إكتمال هيكل السائقين", "Missing driver", "Y", "Driver structure incomplete"),
    ("إنتظار تحميل أمام المصنع", "Waiting at plant", "Y", "Waiting to load at the plant"),
    ("عدم وجود خطه", "Other", "Y", "No plan / no route assigned"),
    ("تخزين فرشة رمضان", "Other", "Y", "Used as Ramadan stall storage"),
    ("ظروف أمنيه", "Other", "Y", "Security conditions"),
    ("دواعى أمنية", "Other", "Y", "Security reasons"),
    ("أسباب تشغيلية", "Other", "Y", "Operational reasons"),
]
sr = wb.create_sheet("Stoppage Reasons", index=wb.sheetnames.index("Driver Master") + 1)
sr.cell(1, 1, "STOPPAGE REASONS").font = Font(bold=True, size=13, color="13294B")
sr.cell(2, 1, "SAP 'Code text' from the truck stoppage log → root-cause category. Categories: On trip, Waiting at branch, Waiting at plant, Maintenance, License, Missing driver, Other.").font = Font(italic=True, color="5A6373")
for j, hh in enumerate(["Code text", "Category", "Counts as stoppage (Y/N)", "Status", "Notes"], 1):
    c = sr.cell(3, j, hh); c.fill = NAVY; c.font = Font(bold=True, color="FFFFFF")
for i, (code, cat, stp, note) in enumerate(MAP, 4):
    for j, v in enumerate([code, cat, stp, "Estimated", note], 1): sr.cell(i, j, v)
for ccol, w in zip("ABCDE", (34, 20, 14, 12, 44)): sr.column_dimensions[ccol].width = w
sr.freeze_panes = "A4"

# ---------------- Parameters ----------------
pm = wb["Parameters"]; last = pm.max_row
for i, row in enumerate([("Spot cost source", "Invoice when available", "Invoice when available / Rate card", "Use the SAP invoice (Net value + Tax amount) when the extract has it; otherwise the rate card"),
                         ("Invoice variance tolerance", 0.05, "fraction", "Rule X7: invoice more than 5% off the rate card valid on the trip date")], last + 1):
    for j, v in enumerate(row, 1): pm.cell(i, j, v).fill = NEW

# READ ME note
rm = wb["READ ME"]; rr = rm.max_row + 2
for line in ["UPDATE 1 Nov 2026 (v2.1): dated spot rates (price list +10% from 10 Mar 2026, found in Jan–Sep invoices); invoice-based rates for routes without a price;",
             "72 routes, 3 vendors and SpPI C04/C10 added from the invoices; SAP equipment numbers on Truck Master; new Stoppage Reasons tab; invoice parameters."]:
    rm.cell(rr, 1, line); rr += 1

wb.save(OUT)
for k, v in report.items():
    print(f"## {k}: {len(v)}")
    for x in v[:8]: print("    ", x)
    if len(v) > 8: print("     …")
print(f"km model for new routes: km = {a:.1f} + {b:.5f} × rate (n={n})")
