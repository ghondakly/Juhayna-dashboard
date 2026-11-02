"""Master v2.2 -> v2.3: Extra Off Days tab (target = month days - Fridays - extra off days), Calendar becomes reference,
vendor rents derived from the MR report, missing MR route added.

usage: python3 -I update_master_v2_3.py MASTER_V2_2.xlsx MR_REPORT.xlsx OUT.xlsx
"""
import sys, re, collections
from datetime import date
import openpyxl
from openpyxl.styles import PatternFill, Font

SRC, MR, OUT = sys.argv[1:4]
NAVY = PatternFill("solid", fgColor="13294B"); NEW = PatternFill("solid", fgColor="FCE4D6")
EST = PatternFill("solid", fgColor="DDEBF7"); EST_FONT = Font(color="1F4E79", italic=True)
rn = lambda s: re.sub(r"\s*([-+])\s*", r"\1", re.sub(r"\s+", " ", str(s or "")).strip().lower())
wb = openpyxl.load_workbook(SRC)
hdr = lambda ws: {str(c.value).strip(): c.column for c in ws[3] if c.value}

# ---- Extra Off Days (empty) + Calendar as reference ----
xo = wb.create_sheet("Extra Off Days", index=wb.sheetnames.index("Calendar") + 1)
xo.cell(1, 1, "EXTRA OFF DAYS").font = Font(bold=True, size=13, color="13294B")
xo.cell(2, 1, "Target working days = calendar days − Fridays − the dates listed here. Public holidays are NOT deducted unless listed. You can also manage this list on the dashboard's Working days screen.").font = Font(italic=True, color="5A6373")
for j, h in enumerate(["Date", "Reason"], 1):
    c = xo.cell(3, j, h); c.fill = NAVY; c.font = Font(bold=True, color="FFFFFF")
xo.column_dimensions["A"].width = 14; xo.column_dimensions["B"].width = 44; xo.freeze_panes = "A4"
cal = wb["Calendar"]; cal.cell(1, 1).value = "PUBLIC HOLIDAYS — REFERENCE ONLY"
cal.cell(2, 1).value = "Official 2026 days off (presidency.eg). NOT deducted from the target; add a date to Extra Off Days to deduct it."

# ---- MR report: per-point rent by vendor/temperature ----
M_SP = {}
for r in wb["SpPI Master"].iter_rows(min_row=4, values_only=True):
    if r[0]: M_SP[str(r[0]).upper()] = str(r[1]).upper()
sw = wb["Spot-Rent Route Master"]; SH = hdr(sw)
w_of = {}
for r in range(4, sw.max_row + 1):
    n = sw.cell(r, SH["Route"]).value
    if n: w_of[rn(n)] = sw.cell(r, SH["Weight (days)"]).value
ws = openpyxl.load_workbook(MR, read_only=True, data_only=True).active
it = ws.iter_rows(values_only=True); h = [str(c).strip() if c else "" for c in next(it)]; ix = h.index
modes = collections.defaultdict(collections.Counter); missing = collections.Counter(); vendors = collections.Counter()
for r in it:
    if not r[ix("Shipment")]: continue
    v = re.sub(r"\s+", " ", str(r[ix("Name of forwarding a")]).strip()); vendors[v] += 1
    route = r[ix("Name")]; w = w_of.get(rn(route))
    if w is None: missing[route] += 1; continue
    a = r[ix("Actul Shipment cost")]
    if r[ix("ShTy")] == "MR" and isinstance(a, (int, float)) and w:
        t = M_SP.get(str(r[ix("SpPI")]).upper(), "D")
        modes[(v, t)][round(a / w * 26 / 1.14 / 100) * 100] += 1

# ---- Vendor Contracts: rents from the MR report ----
vc = wb["Vendor Contracts"]; VH = hdr(vc)
rows = {re.sub(r"\s+", " ", str(vc.cell(r, 1).value).strip()): r for r in range(4, vc.max_row + 1) if vc.cell(r, 1).value}
changed = []
for v in sorted({k[0] for k in modes}):
    if v not in rows:
        r = vc.max_row + 1; rows[v] = r
        for hh, col in VH.items():
            vc.cell(r, col).fill = NEW
        vc.cell(r, VH["Vendor"], v); vc.cell(r, VH["Valid from"], date(2026, 1, 1)).number_format = "yyyy-mm-dd"
        vc.cell(r, VH["VAT applies (Y/N)"], "Y"); vc.cell(r, VH["Fuel paid by (Juhayna/Vendor)"], "Juhayna"); vc.cell(r, VH["Contracted trucks"], 0)
        changed.append(f"{v}: ADDED (in MR report, not in master)")
    r = rows[v]; notes = []
    for t, col in (("D", "Rent dry (EGP/truck/month)"), ("C", "Rent chilled (EGP/truck/month)")):
        c = modes.get((v, t))
        if not c: continue
        rent, n = c.most_common(1)[0]; tot = sum(c.values())
        old = vc.cell(r, VH[col]).value
        cell = vc.cell(r, VH[col], rent); cell.fill = EST; cell.font = EST_FONT
        notes.append(f"{'dry' if t == 'D' else 'chilled'} {rent:,.0f} = most common in MR report ({n}/{tot} trips; other rates used in some months)")
        if old != rent: changed.append(f"{v} {('dry' if t == 'D' else 'chilled')}: {old} → {rent:,.0f}")
    vc.cell(r, VH["Status"], "Derived (MR report)")
    vc.cell(r, VH["Notes"], " · ".join(notes) + " — net of VAT; the dashboard uses the MR report's actual cost when loaded")

# ---- missing MR routes ----
r = sw.max_row
for route, n in missing.items():
    for period, vf, vt in (("old", date(2026, 1, 1), date(2026, 3, 9)), ("new", date(2026, 3, 10), None)):
        r += 1
        vals = {"Route": route, "Weight (days)": 1, "RT km": 925, "Valid from": vf, "Valid to": vt, "Rate status dry": "Estimated", "Rate status chilled": "Estimated",
                "km status": "Estimated", "Notes": f"NEW from MR report ({n} shipments): weight 1, km ≈ Menia + Assiut route +10%; no spot rate (MR only)", "Last updated": date(2026, 11, 1)}
        for hh, col in SH.items():
            c = sw.cell(r, col, vals.get(hh)); c.fill = NEW
            if hh in ("Valid from", "Valid to", "Last updated") and c.value: c.number_format = "yyyy-mm-dd"
    changed.append(f"route added: {route} ({n})")

rm = wb["READ ME"]; rm.cell(rm.max_row + 1, 1, "UPDATE 2 Nov 2026 (v2.3): target = calendar days − Fridays − Extra Off Days (new tab; public holidays are reference only); vendor rents derived from the Jan–Sep MR report; MR route added.")
wb.save(OUT)
print("vendors in MR report:", len(vendors))
for c in changed: print("  ", c)
