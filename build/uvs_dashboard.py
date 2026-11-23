#!/usr/bin/env python3
"""UVS Games customer shipment dashboard builder (HD Shipping Solutions).

Usage:
  python3 uvs_dashboard.py --nexus nexus.json [--itracker itracker.json] [--carrier carrier_status.json]
                           [--state state.json] --out-state new_state.json --out-xlsx out.xlsx

nexus.json     raw nexus_query result (or a list of load objects) for client 2843, fields:
               id status freight_transportation_mode{abbreviation} proNumber bolNumber concatenated_references
               references{name pivot{value}} carrier{name}
               origin{place region_code appointment_at appointment_timezone arrived_at departed_at completed_at}
               destination{place region_code appointment_at appointment_timezone arrived_at completed_at}
itracker.json  {"<nexus load id>": <raw itracker LoadLiveTracking result>} for FTL loads (optional)
carrier_status.json  the carrier-website tracking file ({"rows":[...]}) - pass ONLY on the daily LTL run
state.json     previous state (keeps LTL carrier data between daily runs)
"""
import argparse, json, re, sys, os, base64
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

EXCLUDE = {"OPEN", "QUOTED", "RESERVED", "TRUCK_ORDER_NOT_USED", "RELEASED", "DRAFT"}
STATE_TZ = {
 'AL':'America/Chicago','AK':'America/Anchorage','AZ':'America/Phoenix','AR':'America/Chicago','CA':'America/Los_Angeles',
 'CO':'America/Denver','CT':'America/New_York','DE':'America/New_York','DC':'America/New_York','FL':'America/New_York',
 'GA':'America/New_York','HI':'Pacific/Honolulu','ID':'America/Boise','IL':'America/Chicago','IN':'America/Indiana/Indianapolis',
 'IA':'America/Chicago','KS':'America/Chicago','KY':'America/New_York','LA':'America/Chicago','ME':'America/New_York',
 'MD':'America/New_York','MA':'America/New_York','MI':'America/Detroit','MN':'America/Chicago','MS':'America/Chicago',
 'MO':'America/Chicago','MT':'America/Denver','NE':'America/Chicago','NV':'America/Los_Angeles','NH':'America/New_York',
 'NJ':'America/New_York','NM':'America/Denver','NY':'America/New_York','NC':'America/New_York','ND':'America/Chicago',
 'OH':'America/New_York','OK':'America/Chicago','OR':'America/Los_Angeles','PA':'America/New_York','RI':'America/New_York',
 'SC':'America/New_York','SD':'America/Chicago','TN':'America/Chicago','TX':'America/Chicago','UT':'America/Denver',
 'VT':'America/New_York','VA':'America/New_York','WA':'America/Los_Angeles','WV':'America/New_York','WI':'America/Chicago',
 'WY':'America/Denver','ON':'America/Toronto','QC':'America/Toronto','BC':'America/Vancouver','AB':'America/Edmonton'}
FTL_STATUS = {"COVERED":"Booked","DISPATCHED":"Dispatched","AT_PICKUP":"At Pickup","LOADED":"In Transit",
 "IN_TRANSIT":"In Transit","AT_CONSIGNEE":"At Delivery","DELIVERED":"Delivered","CLAIM":"Delivered",
 "BLOCKED":"On Hold","ON_SITE":"At Delivery","DROPPED":"Delivered"}
LTL_NEXUS_STATUS = {"COVERED":"Awaiting Pickup","DISPATCHED":"Awaiting Pickup","AT_PICKUP":"Awaiting Pickup",
 "LOADED":"In Transit","IN_TRANSIT":"In Transit","AT_CONSIGNEE":"Out for Delivery","DELIVERED":"Delivered","CLAIM":"Delivered"}

def tzname(state, fallback=None):
    return STATE_TZ.get((state or '').upper()) or fallback or 'America/Los_Angeles'

def abbr(tz, when):
    a = when.astimezone(ZoneInfo(tz)).tzname() or ''
    return {'EST':'ET','EDT':'ET','CST':'CT','CDT':'CT','MST':'MT','MDT':'MT','PST':'PT','PDT':'PT'}.get(a, a)

def p_dt(s):
    if not s: return None
    try: return datetime.strptime(s[:19], '%Y-%m-%d %H:%M:%S')
    except ValueError:
        try: return datetime.strptime(s[:10], '%Y-%m-%d')
        except ValueError: return None

def utc_to_local(s, tz):
    d = p_dt(s)
    if not d: return None
    return d.replace(tzinfo=timezone.utc).astimezone(ZoneInfo(tz))

def stamp(kind, when, tz=None, date_only=False, window=''):
    """A dashboard date cell: kind = ETA | Picked up | Delivered."""
    return {"kind": kind, "date": when.strftime('%Y-%m-%d %H:%M'), "tz": (abbr(tz, when) if tz and not date_only else ''),
            "date_only": date_only, "window": window}

def text(t): return {"kind": "text", "text": t}

def load_list(x):
    if isinstance(x, list): return x
    return x["data"]["loads"]["data"]

def po_of(l):
    refs = [r["pivot"]["value"] for r in (l.get("references") or []) if r.get("pivot") and (r["name"] or '').upper() in ("PO", "PURCHASE ORDER", "PO#", "PO NUMBER")]
    if refs: return ", ".join(refs)
    return l.get("concatenated_references") or ''

def city(s): return ", ".join(x for x in [s.get("place"), s.get("region_code")] if x) if s else ''

def appt_local(stop):
    """Nexus appointment is local wall time in the stop's appointment_timezone."""
    d = p_dt(stop.get("appointment_at"))
    if not d: return None, None
    tz = stop.get("appointment_timezone") or tzname(stop.get("region_code"))
    return d.replace(tzinfo=ZoneInfo(tz)), tz

def appt_stamp(stop):
    d, tz = appt_local(stop)
    if not d: return text("Pending")
    if d.hour == 0 and d.minute == 0: return stamp("ETA", d, date_only=True)
    return stamp("ETA", d, tz)

def pick_tracker(raw):
    try: arr = raw["data"]["loadTrackers"]["data"]
    except (KeyError, TypeError): return None
    if not arr: return None
    for t in arr:
        if any((d.get("pivot") or {}).get("is_tracking") for d in (t.get("drivers") or [])): return t
    return arr[0]

def ftl_row(l, trk):
    o, d = l.get("origin") or {}, l.get("destination") or {}
    otz, dtz = tzname(o.get("region_code"), o.get("appointment_timezone")), tzname(d.get("region_code"), d.get("appointment_timezone"))
    st = l["status"]
    # pickup
    pu_done = o.get("completed_at") or o.get("departed_at")
    if trk and not pu_done:
        for s in trk.get("stops") or []:
            if s.get("type") == "PICKUP" and (s.get("completed_at") or s.get("departed_at")):
                pu_done = s.get("completed_at") or s.get("departed_at")
    pu = stamp("Picked up", utc_to_local(pu_done, otz), otz) if pu_done else appt_stamp(o)
    # delivery
    de_done = d.get("completed_at") or (d.get("arrived_at") if st in ("DELIVERED", "CLAIM") else None)
    if de_done:
        de = stamp("Delivered", utc_to_local(de_done, dtz), dtz)
    elif st in ("DELIVERED", "CLAIM"):
        de = text("Delivered")
    else:
        de = None
        if trk:
            if trk.get("eta_at"):
                de = stamp("ETA", utc_to_local(trk["eta_at"], dtz), dtz)
            else:
                drop = [s for s in (trk.get("stops") or []) if s.get("type") == "DROP_OFF"]
                s = drop[-1] if drop else None
                if s and s.get("delayInMinutes") is not None and s.get("appointment_at"):
                    a = p_dt(s["appointment_at"]).replace(tzinfo=ZoneInfo(s.get("appointment_timezone") or dtz))
                    eta = a + timedelta(minutes=int(s["delayInMinutes"]))
                    de = stamp("ETA", eta.astimezone(ZoneInfo(dtz)), dtz)
        if not de: de = appt_stamp(d)
    status = FTL_STATUS.get(st, st.replace('_', ' ').title())
    if pu_done and status in ("Booked", "Dispatched", "At Pickup"): status = "In Transit"
    if trk and status == "In Transit" and (trk.get("delay_status") or {}).get("is_delayed"): status = "In Transit - Delayed"
    return {"status": status, "pickup": pu, "delivery": de, "pro": l.get("proNumber") or ''}

def parse_carrier_eta(eta):
    """'2026-11-24 10:30 AM-2:30 PM' -> (date, window)."""
    m = re.match(r'(\d{4}-\d{2}-\d{2})\s*(.*)', eta or '')
    if not m: return None, ''
    return datetime.strptime(m.group(1), '%Y-%m-%d'), m.group(2).strip()

def ltl_from_carrier(l, c, prev):
    o = l.get("origin") or {}
    cat, cs, det = c.get("category", ""), (c.get("carrier_status") or ''), (c.get("detail") or '')
    pro = c.get("pro") or l.get("proNumber") or ''
    if cat == "Delivered": status = "Delivered"
    elif cat == "At destination terminal":
        status = "Out for Delivery" if re.search(r'out for delivery', cs + ' ' + det, re.I) else "At Destination Terminal"
    elif cat == "In transit": status = "In Transit"
    elif cat == "Exception": status = "Delivery Exception"
    else:
        # needs manual check / not on carrier site yet
        status = (prev or {}).get("status") or LTL_NEXUS_STATUS.get(l["status"], "In Transit")
    picked = status not in ("Awaiting Pickup",)
    # pickup
    pu = (prev or {}).get("pickup") if (prev or {}).get("pickup", {}).get("kind") == "Picked up" else None
    if not pu:
        if picked:
            m = re.search(r'[Pp]icked up (\d{1,2})/(\d{1,2})', det)
            ad, _ = appt_local(o)
            if m and ad:
                pdate = datetime(ad.year, int(m.group(1)), int(m.group(2)))
            else:
                pdate = ad.replace(tzinfo=None) if ad else None
            pu = stamp("Picked up", pdate, date_only=True) if pdate else text("Picked up")
        else:
            pu = appt_stamp(o)
    # delivery
    if status == "Delivered":
        dv = c.get("delivered") or ''
        dd, win = parse_carrier_eta(dv)
        de = stamp("Delivered", dd, date_only=True, window=win) if dd else text("Delivered")
    else:
        dd, win = parse_carrier_eta(c.get("eta"))
        if dd: de = stamp("ETA", dd, date_only=True, window=win)
        elif re.search(r'appointment', cs + ' ' + det, re.I): de = text("ETA pending delivery appointment")
        else: de = text("ETA pending from carrier")
    return {"status": status, "pro": pro, "pickup": pu, "delivery": de, "checked": c.get("checked_at", "")}

def ltl_from_nexus(l):
    st = LTL_NEXUS_STATUS.get(l["status"], "In Transit")
    o = l.get("origin") or {}
    pu = appt_stamp(o)
    if st != "Awaiting Pickup":
        ad, _ = appt_local(o)
        pu = stamp("Picked up", ad.replace(tzinfo=None), date_only=True) if ad else text("Picked up")
    return {"status": st, "pro": l.get("proNumber") or '', "pickup": pu,
            "delivery": text("ETA pending from carrier"), "checked": ""}

STATUS_ORDER = ["Delivery Exception", "In Transit - Delayed", "Out for Delivery", "At Delivery", "At Destination Terminal",
                "In Transit", "At Pickup", "Dispatched", "Booked", "Awaiting Pickup", "On Hold", "Delivered"]
FILL = {"Delivered": ("E3F5EC", "0B6B48"), "Delivery Exception": ("FFE1E6", "B0123F"), "In Transit - Delayed": ("FFF1D6", "8A5200"),
        "Out for Delivery": ("DCEEFF", "0A4A86"), "At Delivery": ("DCEEFF", "0A4A86"), "At Destination Terminal": ("E8F1FA", "1D4E7A"),
        "In Transit": ("EEF3F8", "0A2540"), "On Hold": ("FFF1D6", "8A5200")}

def ltl_label(carrier, now):
    """Carrier check time as 'YYYY-MM-DDTHH:MM:SSZ' from the rows' checked_at ('YYYY-MM-DD HH:MM PT')."""
    for r in carrier.values():
        m = re.match(r'(\d{4}-\d{2}-\d{2} \d{1,2}:\d{2})', r.get("checked_at") or '')
        if m:
            d = datetime.strptime(m.group(1), '%Y-%m-%d %H:%M').replace(tzinfo=ZoneInfo('America/Los_Angeles'))
            return d.astimezone(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    return now.strftime('%Y-%m-%dT%H:%M:%SZ')

ITER = 600000
def _key(pw, salt, iters=None):
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iters or ITER).derive(pw.encode())

def enc_dump(obj, path, pw):
    """AES-256-GCM, key from PBKDF2-SHA256 (same format the web page decrypts)."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt, iv = os.urandom(16), os.urandom(12)
    ct = AESGCM(_key(pw, salt)).encrypt(iv, json.dumps(obj, separators=(",", ":")).encode(), None)
    b = lambda x: base64.b64encode(x).decode()
    json.dump({"v": 1, "iter": ITER, "salt": b(salt), "iv": b(iv), "ct": b(ct)}, open(path, "w"))

def enc_load(path, pw):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    e = json.load(open(path)); d = lambda k: base64.b64decode(e[k])
    return json.loads(AESGCM(_key(pw, d("salt"), e.get("iter"))).decrypt(d("iv"), d("ct"), None))

def build(args):
    pw = os.environ.get(args.key_env) if args.key_env else None
    if args.key_env and not pw: sys.exit(f"Set the {args.key_env} environment variable to the dashboard password")
    loads = [l for l in load_list(json.load(open(args.nexus))) if l["status"] not in EXCLUDE]
    trk = json.load(open(args.itracker)) if args.itracker else {}
    carrier = {}
    if args.carrier:
        for r in json.load(open(args.carrier)).get("rows", []): carrier[str(r["load_id"])] = r
    prev = (enc_load(args.state, pw) if pw else json.load(open(args.state))) if args.state and os.path.exists(args.state) else {}
    prev_ltl = prev.get("ltl", {})
    now = datetime.now(timezone.utc)
    rows, new_ltl = [], {}
    for l in loads:
        lid = str(l["id"])
        mode = ((l.get("freight_transportation_mode") or {}).get("abbreviation") or '').upper()
        if mode != "FTL": mode = "LTL"
        if mode == "FTL":
            r = ftl_row(l, pick_tracker(trk.get(lid)) if trk.get(lid) else None)
        else:
            if lid in carrier and (carrier[lid].get("category") != "Needs manual check" or lid not in prev_ltl):
                r = ltl_from_carrier(l, carrier[lid], prev_ltl.get(lid))
            elif lid in prev_ltl:
                r = prev_ltl[lid]
            else:
                r = ltl_from_nexus(l)
            new_ltl[lid] = r
        rows.append({"id": lid, "mode": mode, "status": r["status"], "pro": r.get("pro") or l.get("proNumber") or '',
                     "bol": l.get("bolNumber") or '', "po": po_of(l), "origin": city(l.get("origin")),
                     "dest": city(l.get("destination")), "pickup": r["pickup"], "delivery": r["delivery"]})
    rows.sort(key=lambda r: (STATUS_ORDER.index(r["status"]) if r["status"] in STATUS_ORDER else 5,
                             (r["delivery"].get("date") or '9999') if r["status"] != "Delivered" else '',
                             -(datetime.strptime(r["delivery"]["date"], '%Y-%m-%d %H:%M').timestamp()) if r["status"] == "Delivered" and r["delivery"].get("date") else 0))
    state = {"updated": now.strftime('%Y-%m-%dT%H:%M:%SZ'),
             "ltl_checked": (ltl_label(carrier, now) if args.carrier else prev.get("ltl_checked", "")),
             "ftl_checked": now.strftime('%Y-%m-%dT%H:%M:%SZ'), "ltl": new_ltl, "rows": rows}
    if args.out_xlsx: write_xlsx(state, args.out_xlsx)
    if args.out_json: write_json(state, args.out_json, pw)
    state.pop("rows")
    if pw: enc_dump(state, args.out_state, pw)
    else: json.dump(state, open(args.out_state, "w"), separators=(",", ":"))
    print(f"{len(rows)} shipments ({sum(r['mode']=='LTL' for r in rows)} LTL, {sum(r['mode']=='FTL' for r in rows)} FTL) -> {args.out_xlsx}")

def cell_value(c):
    """Return (value, number_format, is_actual)."""
    if c.get("kind") == "text": return c["text"], None, c["text"].startswith(("Delivered", "Picked"))
    d = datetime.strptime(c["date"], '%Y-%m-%d %H:%M')
    lab = c["kind"]
    tail = (' ' + c["window"]) if c.get("window") else ''
    tail = tail.replace('"', '')
    if c.get("date_only"): fmt = f'"{lab} "m/d/yyyy"{tail}"' if tail else f'"{lab} "m/d/yyyy'
    else: fmt = f'"{lab} "m/d/yyyy h:mm AM/PM" {c.get("tz", "")}"'
    return d, fmt, lab != "ETA"

def label(c):
    """Display text + sort key for a date cell (used by the web dashboard)."""
    if c.get("kind") == "text": return {"t": c["text"], "k": "Delivered" if c["text"].startswith("Delivered") else ("Picked up" if c["text"].startswith("Picked") else "Pending"), "d": ""}
    d = datetime.strptime(c["date"], '%Y-%m-%d %H:%M')
    t = d.strftime('%a %b %-d') if c.get("date_only") else d.strftime('%a %b %-d, %-I:%M %p') + (' ' + c["tz"] if c.get("tz") else '')
    if c.get("window"): t += ', ' + c["window"]
    return {"t": t, "k": c["kind"], "d": d.strftime('%Y-%m-%d')}

def write_json(state, path, pw=None):
    out = {"ftl_checked": state.get("ftl_checked", ""), "ltl_checked": state.get("ltl_checked", ""), "rows": []}
    for r in state["rows"]:
        out["rows"].append({"status": r["status"], "mode": r["mode"], "pro": r["pro"], "bol": r["bol"], "po": r["po"],
                            "origin": r["origin"], "dest": r["dest"], "pu": label(r["pickup"]), "del": label(r["delivery"])})
    if pw: enc_dump(out, path, pw)
    else: json.dump(out, open(path, "w"), separators=(",", ":"))

def write_xlsx(state, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
    NAVY, SKY = "0A2540", "8FD3FF"
    wb = Workbook(); ws = wb.active; ws.title = "Shipments"
    ws.sheet_view.showGridLines = False
    rows = state["rows"]
    pt = ZoneInfo('America/Los_Angeles')
    def fmt_pt(iso):
        if not iso: return "not yet"
        d = datetime.strptime(iso, '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc).astimezone(pt)
        return d.strftime('%b %-d, %Y %-I:%M %p') + ' PT'
    hdr = ["Status", "Mode", "PRO #", "BOL #", "PO", "Origin City", "Destination City", "Pickup (ETA / Actual)", "Delivery (ETA / Actual)"]
    widths = [24, 8, 15, 14, 12, 22, 24, 34, 38]
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    band = PatternFill("solid", fgColor=NAVY)
    for r in range(1, 5):
        for c in range(1, len(hdr) + 1): ws.cell(r, c).fill = band
    ws["A1"] = "UVS Games  |  Shipment Tracking"; ws["A1"].font = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
    ws["A2"] = "Provided by HD Shipping Solutions  ·  hdships.com"; ws["A2"].font = Font(size=11, color=SKY)
    n_open = sum(r["status"] != "Delivered" for r in rows)
    ws["A3"] = (f"{len(rows)} shipments  ·  {n_open} active  ·  {len(rows) - n_open} delivered  ·  "
                f"{sum(r['mode']=='LTL' for r in rows)} LTL, {sum(r['mode']=='FTL' for r in rows)} FTL")
    ws["A3"].font = Font(size=12, bold=True, color="FFFFFF")
    ws["A4"] = (f"Truckload (FTL) updated {fmt_pt(state.get('ftl_checked'))}, hourly from live truck tracking.   "
                f"LTL updated {fmt_pt(state.get('ltl_checked'))}, each morning from the carrier's tracking site.")
    ws["A4"].font = Font(size=10, color="DCEAF7")
    ws.row_dimensions[1].height = 34
    ws["A5"] = "Tip: use the arrow on any column heading to filter (for example Mode = LTL or FTL) or sort. Filters you set only change your own view."
    ws["A5"].font = Font(size=10, italic=True, color="53657A")
    H = 6
    for i, h in enumerate(hdr, 1):
        c = ws.cell(H, i, h); c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="14365A")
        c.alignment = Alignment(vertical="center")
    ws.row_dimensions[H].height = 22
    thin = Border(bottom=Side(style="thin", color="DFE6EE"))
    for k, r in enumerate(rows, H + 1):
        vals = [r["status"], r["mode"], r["pro"], r["bol"], r["po"], r["origin"], r["dest"]]
        for i, v in enumerate(vals, 1):
            c = ws.cell(k, i, v); c.border = thin; c.alignment = Alignment(vertical="center")
            if i in (3, 4, 5): c.number_format = "@"
        fg, fc = FILL.get(r["status"], ("FFFFFF", "0A2540"))
        sc = ws.cell(k, 1); sc.fill = PatternFill("solid", fgColor=fg); sc.font = Font(bold=True, color=fc)
        for i, key in ((8, "pickup"), (9, "delivery")):
            v, f, actual = cell_value(r[key])
            c = ws.cell(k, i, v); c.border = thin; c.alignment = Alignment(vertical="center", horizontal="left")
            if f: c.number_format = f
            c.font = Font(color="0B6B48", bold=True) if actual else Font(color="1D4E7A", italic=True)
    last = H + max(len(rows), 1)
    if rows:
        t = Table(displayName="Shipments", ref=f"A{H}:{get_column_letter(len(hdr))}{last}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.add_table(t)
    else:
        ws.cell(H + 1, 1, "No active shipments right now.")
    ws.freeze_panes = f"A{H + 1}"
    k = last + 2
    for line in ["How to read the dates: “ETA” is the current estimate; once it happens it changes to “Picked up” or “Delivered” with the actual date (green).",
                 "Times are local to the pickup or delivery city. LTL delivery windows are shown when the carrier provides one.",
                 "Questions on a shipment: HD Shipping Solutions LTL team · LTL@hdships.com · (775) 300-7970"]:
        ws.cell(k, 1, line).font = Font(size=10, color="53657A"); k += 1
    ws.print_title_rows = f"{H}:{H}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.page_setup.fitToHeight = 0
    wb.save(path)

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--nexus", required=True); ap.add_argument("--itracker"); ap.add_argument("--carrier")
    ap.add_argument("--state"); ap.add_argument("--out-state", required=True); ap.add_argument("--out-xlsx"); ap.add_argument("--out-json")
    ap.add_argument("--key-env", help="name of the env var holding the dashboard password; encrypts --out-json and the state file")
    build(ap.parse_args())
