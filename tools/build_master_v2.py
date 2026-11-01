"""Build Master Data v2 from the current (v1) master + the user's decisions.

usage: python3 -I build_master_v2.py MASTER_V1.xlsx OUT.xlsx
"""
import sys, re, statistics as st
from datetime import date
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

SRC, OUT = sys.argv[1:3]
TODAY = date(2026, 11, 1)
src = openpyxl.load_workbook(SRC)

NAVY = PatternFill("solid", fgColor="13294B")
EST = PatternFill("solid", fgColor="DDEBF7")
EST_FONT = Font(color="1F4E79", italic=True)
NEW = PatternFill("solid", fgColor="FCE4D6")
HDR_FONT = Font(bold=True, color="FFFFFF")

def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()

wb = openpyxl.Workbook()
wb.remove(wb.active)

def sheet(name, title, note, headers, widths=None):
    ws = wb.create_sheet(name)
    ws.cell(1, 1, title).font = Font(bold=True, size=13, color="13294B")
    ws.cell(2, 1, note).font = Font(italic=True, color="5A6373")
    for i, h in enumerate(headers, 1):
        c = ws.cell(3, i, h); c.fill = NAVY; c.font = HDR_FONT; c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A4"
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(h, max(12, min(40, len(h) + 4)))
    return ws

def put(ws, row, vals, est_cols=(), new=False):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row, i, v)
        if i in est_cols:
            c.fill = EST; c.font = EST_FONT
        elif new:
            c.fill = NEW

# ---------------- READ ME ----------------
rm = wb.create_sheet("READ ME")
for i, line in enumerate([
    "JUHAYNA FLEET CONTROL TOWER — MASTER DATA v2",
    "",
    "The dashboard validates every upload against these tabs before calculating anything. Headers on row 3 are read by NAME — keep them exactly; add rows freely.",
    "Every rate/contract/route row has a Status: Confirmed or Estimated. The Cost screen shows how much spend rests on Estimated values.",
    "",
    "Parameters ............ fuel price, VAT, working-day target, RE free trips and your policy decisions",
    "Calendar .............. public holidays (official presidency.eg 2026 list). Fridays are excluded automatically.",
    "SpPI Master ........... SpPI code → temperature (D/C), Single/Double, pallets",
    "Shipment Type Master .. ShTy → cost bucket: SPOT, MR (monthly rent), RE (returned pallets), NONE (ignored)",
    "Truck Master .......... owned truck plate (digits only, as in SAP 'Car No.2') → temperature, active",
    "Driver Master ......... driver code → name, usual truck",
    "Owned Route Master .... route → weight (days), RT km, incentive single/double, plant shuttle flag",
    "Spot-Rent Route Master  route → weight, RT km, spot rate dry/chilled, valid from/to",
    "Vendor Contracts ...... vendor → monthly rent dry/chilled per truck, contracted trucks, VAT, fuel paid by",
    "Route Aliases ......... spelling variants in SAP → the route name used in the masters",
    "",
    "Colours: blue italic = Estimated value (replace with actual and set Status to Confirmed). Orange = row added during the v2 migration.",
], 1):
    rm.cell(i, 1, line)
rm.cell(1, 1).font = Font(bold=True, size=14, color="13294B")
rm.column_dimensions["A"].width = 130

# ---------------- Parameters ----------------
ws = sheet("Parameters", "PARAMETERS", "Change values here — no code change needed. Policy rows record the decisions taken on 1 Nov 2026.",
           ["Parameter", "Value", "Unit", "Notes"], {"Parameter": 38, "Value": 18, "Unit": 16, "Notes": 70})
params = [
    ("Master version", 2, "", "Do not change"),
    ("Fuel price", 20.5, "EGP / litre", ""),
    ("Fuel consumption", 3, "km / litre", "Heavy truck average"),
    ("Reefer fuel factor", 1.5, "× fuel", "Applied to chilled trips"),
    ("VAT rate", 0.14, "fraction", "14%"),
    ("Spot price list excludes VAT", "Y", "Y/N", "Decision 4: Y → 14% VAT added to spot and billable RE freight"),
    ("Rented rent VAT", "Y", "Y/N", "VAT on monthly rent"),
    ("Working days per month target", 26, "days", "Used for projections (points per truck per month)"),
    ("Weekly day off", "Friday", "", "Decision 8: Fridays + Calendar holidays are not working days"),
    ("RE free trips per vendor per month", 3, "trips", "Decision 3: first N RE trips per vendor per calendar month are free, then route rate"),
    ("Rented truck count rule", "Computed", "Computed / Contracted / Higher", "Decision 2: Computed = round(MR points ÷ working days)"),
    ("Rented fuel paid by Juhayna", "Y", "Y/N", "Decision 6"),
    ("Unmaintained trips", "Count with flag", "Count with flag / Exclude", "Decision 1"),
    ("Owned fixed costs included", "N", "Y/N", "Decision 5: variable costs only (fuel + incentive)"),
    ("Same-day points warning", 3, "points", "Rule X3: one truck starting more than this many points (route-weight days) on one day"),
    ("Rate outlier tolerance", 0.4, "fraction", "Rule X5: rate more than 40% off routes of similar distance"),
    ("Master stale after", 60, "days", "Rule X6"),
]
for r, p in enumerate(params, 4):
    put(ws, r, p)

# ---------------- Calendar ----------------
ws = sheet("Calendar", "CALENDAR — PUBLIC HOLIDAYS", "Official 2026 days off (presidency.eg). Fridays are excluded automatically; list only extra non-working days.",
           ["Date", "Holiday", "Type", "Status", "Source"], {"Date": 14, "Holiday": 40, "Type": 16, "Status": 12, "Source": 30})
hol = [("2026-01-07", "Coptic Christmas"), ("2026-01-29", "January 25 Revolution & Police Day (observed)"),
       ("2026-03-19", "Eid al-Fitr"), ("2026-03-20", "Eid al-Fitr"), ("2026-03-21", "Eid al-Fitr"), ("2026-03-22", "Eid al-Fitr"), ("2026-03-23", "Eid al-Fitr"),
       ("2026-04-13", "Sham El-Nessim"), ("2026-04-25", "Sinai Liberation Day"), ("2026-05-07", "Labour Day (observed)"),
       ("2026-05-26", "Arafat Day"), ("2026-05-27", "Eid al-Adha"), ("2026-05-28", "Eid al-Adha"), ("2026-05-29", "Eid al-Adha"), ("2026-05-30", "Eid al-Adha"), ("2026-05-31", "Eid al-Adha"),
       ("2026-06-18", "Islamic New Year (observed)"), ("2026-07-02", "June 30 Revolution (observed)"), ("2026-07-23", "July 23 Revolution Day"),
       ("2026-08-27", "Prophet's Birthday (observed)"), ("2026-10-06", "Armed Forces Day")]
for r, (d, n) in enumerate(hol, 4):
    c = ws.cell(r, 1, date.fromisoformat(d)); c.number_format = "yyyy-mm-dd"
    put_vals = [None, n, "Public holiday", "Confirmed", "presidency.eg 2026 list"]
    for i, v in enumerate(put_vals[1:], 2):
        ws.cell(r, i, v)

# ---------------- SpPI ----------------
ws = sheet("SpPI Master", "SpPI MASTER", "Every SpPI code that appears in SAP must be here (rule M6).",
           ["SpPI", "Temperature (D/C)", "Load (Single/Double)", "Pallets", "Status", "Notes"])
for r, v in enumerate([("D24", "D", "Single", 24, "Confirmed", ""), ("C24", "C", "Single", 24, "Confirmed", ""),
                       ("C40", "C", "Double", 40, "Confirmed", "Double → incentive double"),
                       ("J24", "D", "Single", 24, "Confirmed", "Decision 7 (1 Nov 2026)")], 4):
    put(ws, r, v, new=(v[0] == "J24"))

# ---------------- Shipment types ----------------
ws = sheet("Shipment Type Master", "SHIPMENT TYPE MASTER", "Every ShTy in the spot/rent extract must be here (rule M7). Cost bucket: SPOT, MR, RE or NONE.",
           ["ShTy", "Cost bucket", "Description"], {"Description": 50})
for r, v in enumerate([("MR", "MR", "Monthly rent"), ("TR", "SPOT", "Spot factory → branch"), ("DTR", "SPOT", "Factory → distributor"),
                       ("DD", "SPOT", "Special customers"), ("ZTRB", "SPOT", "Branch → customer"), ("RE", "RE", "Returned pallets")], 4):
    put(ws, r, v)

# ---------------- Trucks & drivers ----------------
tws = src["Truck Driver Master"]
trucks = [r for r in tws.iter_rows(min_row=4, values_only=True) if r[0] is not None]
ws = sheet("Truck Master", "TRUCK MASTER", "Plate = digits of SAP 'Car No.2'. Temperature D/C. Unknown plates in uploads appear on the fix list (rule M1).",
           ["Plate", "Temperature (D/C)", "Plant", "Active (Y/N)", "Status", "Last updated"])
for r, t in enumerate(trucks, 4):
    put(ws, r, (str(t[0]).strip(), t[3] or "D", "", t[4] or "Y", "Confirmed", TODAY))
ws = sheet("Driver Master", "DRIVER MASTER", "Driver code = SAP 'Juhayna Driver No'. Unknown codes in uploads appear on the fix list (rule M2).",
           ["Driver code", "Driver name", "Usual truck", "Active (Y/N)", "Status", "Last updated"], {"Driver name": 34})
for r, t in enumerate(trucks, 4):
    put(ws, r, (str(t[2]).strip(), t[1], str(t[0]).strip(), "Y", "Confirmed", TODAY))

# ---------------- Spot-Rent routes ----------------
sws = src["Spot-Rent Route Master"]
spot = []
for row in sws.iter_rows(min_row=4):
    v = [c.value for c in row]
    if not v[0]:
        continue
    fills = [c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None for c in row]
    srcnote = str(v[6] or "")
    pl_dry = fills[4] and str(fills[4]).endswith("C6EFCE")
    pl_ch = fills[5] and str(fills[5]).endswith("C6EFCE")
    spot.append(dict(name=str(v[0]).strip(), w=v[2], km=v[3], rd=v[4], rc=v[5], sd="Confirmed" if pl_dry else "Estimated",
                     sc="Confirmed" if pl_ch else "Estimated", note=(srcnote + (" · " + str(v[7]) if len(v) > 7 and v[7] else ""))))
ws = sheet("Spot-Rent Route Master", "SPOT-RENT ROUTE MASTER", "Rates apply to trip dates between Valid from and Valid to (blank = open). Add a new row with new dates when the price list changes.",
           ["Route", "Weight (days)", "RT km", "Rate dry (EGP/trip)", "Rate chilled (EGP/trip)", "Valid from", "Valid to",
            "Rate status dry", "Rate status chilled", "km status", "Notes", "Last updated"], {"Route": 38, "Notes": 60})
for r, s in enumerate(spot, 4):
    put(ws, r, (s["name"], s["w"], s["km"], s["rd"], s["rc"], date(2026, 1, 1), None, s["sd"], s["sc"], "Estimated", s["note"], TODAY),
        est_cols=(3,) + ((4,) if s["sd"] == "Estimated" else ()) + ((5,) if s["sc"] == "Estimated" else ()))
    ws.cell(r, 6).number_format = "yyyy-mm-dd"; ws.cell(r, 12).number_format = "yyyy-mm-dd"
spot_km = {norm(s["name"]): float(s["km"]) for s in spot if s["km"]}

# ---------------- Owned routes (with RT km estimate) ----------------
SYN = [  # token pattern -> spot route name supplying the distance
    (r"mid ?alex", "mid alex"), (r"alex ?east", "alex east"), (r"alex ?west|alexwest", "alex"), (r"\balex\b", "alex"),
    (r"beh[ei]?ra|bahera|behera", "beheira"), (r"men[i]?a|mania", "menia"), (r"ass?[iu]+[iu]?t|assuit|assiut", "assiut"),
    (r"b(ani|eni|\.)? ?s[uw]?[ie]+[fv]|benisu[ie]f|banisweif|benisweif|bani swef", "bani swef"),
    (r"so?hag|suhag|asohag", "sohag"), (r"kena|qena|luxor", "kena (luxor)"), (r"as+wan", "aswan"),
    (r"matrouh", "marsa matrouh"), (r"tour", "el- tour"), (r"hurghada", "hurghada"), (r"isma[i]?l[i]?a", "ismailia"),
    (r"port ?said|portsaid", "port said"), (r"damiat", "mansoura-islamic dist. (موزع دمياط)"), (r"zagazig", "zagazig"),
    (r"shar[k]?[iy]a|sharkya|sharkia", "zagazig"), (r"monof", "monofya"), (r"tanta", "tanta"), (r"mans[o]?ura|mansora", "mansoura"),
    (r"kal[iy]o?bia|qalyobia|kaliobia|alkaliobia", "kalyobia"), (r"khos", "khosos"), (r"sawah|swaah", "elswaah"),
    (r"obour", "el-obour"), (r"warr?ak", "alwarrak"), (r"\bauc\b", "auc"), (r"wadi ?ho+f", "wadi hofe"), (r"suez", "suez"),
    (r"shab[r]?a[mn]ant", "shabranant(zayed)"), (r"zayed", "zayed"), (r"dab[a]+", "eldaba"), (r"fast trade", "factory to dist. fast trade"),
    (r"trinity south", "factory to dis trinity south"), (r"sedeek", "factory to dis el sedeek"), (r"islamic|markaz", "factory to dist. markaz islamy"),
    (r"ehab aram", "factory to dis ehab aram"), (r"el ?nada", "factory to dis el nada"), (r"mcdonald", "mcdonalds"), (r"amazon", "amazon ( suez )"),
    (r"logistic", "factory to dist. logistic"),
]
ows = src["Owned Route Master"]
orows, seen, dups = [], set(), []
for row in ows.iter_rows(min_row=5):
    v = [c.value for c in row]
    name = v[0]
    if not name:
        continue
    k = norm(name)
    if k in seen:
        dups.append(name); continue
    seen.add(k)
    est_inc = bool(v[4] and str(v[4]).startswith("EXTRAPOLATED"))
    stops = [p for p in re.split(r"\s*-\s*", re.sub(r"^(juh(ayna)?)\s*-?\s*", "", str(name).strip(), flags=re.I)) if p.strip()]
    plant = "plant" in k and v[1] == 0
    upper_origin = bool(re.search(r"(ass[iu]+[iu]?t|kena|qena) ?plant", k))
    kms = []
    for stp in stops:
        sn = norm(stp)
        if "plant" in sn:
            continue
        for pat, tgt in SYN:
            if re.search(pat, sn) and tgt in spot_km:
                kms.append(spot_km[tgt]); break
    orows.append(dict(name=str(name).strip(), w=v[1], i1=v[2], i2=v[3], est_inc=est_inc, plant=plant, upper=upper_origin, kms=kms, n=len(stops),
                      comment=(str(v[4]) if v[4] else "")))
by_w = {}
for o in orows:
    if o["kms"] and not o["plant"] and not o["upper"]:
        o["km"] = max(o["kms"]) * (1 + 0.1 * (len(o["kms"]) - 1)); o["kb"] = f"max of {len(o['kms'])} matched stop(s) from Spot-Rent km, +10% per extra stop"
        by_w.setdefault(o["w"], []).append(o["km"])
for o in orows:
    if "km" not in o:
        if o["upper"] or o["plant"]:
            o["km"] = 200; o["kb"] = "Upper-Egypt plant origin — no distance basis, placeholder 200 km: PLEASE FILL"
        else:
            ref = by_w.get(o["w"]) or by_w.get(1)
            o["km"] = st.median(ref); o["kb"] = f"median km of matched weight-{o['w']} routes"
    o["km"] = int(round(o["km"] / 5.0) * 5)

ws = sheet("Owned Route Master", "OWNED ROUTE MASTER", "Weight 0 is allowed only for plant shuttles (Plant shuttle = Y). RT km drives owned fuel cost.",
           ["Route", "Weight (days)", "RT km", "Incentive single (EGP)", "Incentive double (EGP)", "Plant shuttle (Y/N)",
            "Incentive status", "km status", "km basis", "Comment", "Last updated"], {"Route": 42, "km basis": 50, "Comment": 50})
for r, o in enumerate(orows, 4):
    put(ws, r, (o["name"], o["w"], o["km"], o["i1"], o["i2"], "Y" if o["plant"] else "N",
                "Estimated" if o["est_inc"] else "Confirmed", "Estimated", o["kb"], o["comment"], TODAY),
        est_cols=(3,) + ((4, 5) if o["est_inc"] else ()))

# ---------------- Vendor contracts ----------------
vws = src["Vendor Master"]
ws = sheet("Vendor Contracts", "VENDOR CONTRACTS", "One row per contract period. MR trips need a contract valid on the trip date (rule M3).",
           ["Vendor", "Valid from", "Valid to", "Rent dry (EGP/truck/month)", "Rent chilled (EGP/truck/month)", "Contracted trucks",
            "VAT applies (Y/N)", "Fuel paid by (Juhayna/Vendor)", "Status", "Notes"], {"Vendor": 36, "Notes": 40})
r = 4
for v in vws.iter_rows(min_row=4, values_only=True):
    if not v[0]:
        continue
    put(ws, r, (str(v[0]).strip(), date(2026, 1, 1), None, v[3], v[4], v[5], "Y", "Juhayna", "Estimated", "rent = engine default — replace with contract"),
        est_cols=(4, 5))
    ws.cell(r, 2).number_format = "yyyy-mm-dd"; r += 1

# ---------------- Aliases ----------------
ws = sheet("Route Aliases", "ROUTE ALIASES", "Map a route spelling found in SAP to the route name in Owned or Spot-Rent Route Master.",
           ["Alias (as in SAP)", "Maps to route", "Table (Owned/Spot)", "Notes"], {"Alias (as in SAP)": 40, "Maps to route": 40, "Notes": 40})

wb.move_sheet("READ ME", offset=-wb.index(wb["READ ME"]))
wb.save(OUT)
print("owned routes", len(orows), "dropped duplicates", dups)
print("owned km basis:", {k: sum(1 for o in orows if o["kb"].startswith(k)) for k in ("max of", "median", "plant")})
print("spot routes", len(spot), "confirmed dry", sum(s["sd"] == "Confirmed" for s in spot), "confirmed chilled", sum(s["sc"] == "Confirmed" for s in spot))
for o in orows[:12] + [o for o in orows if o["plant"]][:3]:
    print(f"  {o['name'][:40]:40} w={o['w']} km={o['km']:5} | {o['kb'][:60]}")
