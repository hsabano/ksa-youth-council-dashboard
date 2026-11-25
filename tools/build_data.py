#!/usr/bin/env python3
"""Build the dashboard's data.js from the Planner export and the Council Data workbook.

Usage:
    python3 tools/build_data.py PLANNER_EXPORT.xlsx COUNCIL_DATA.xlsx [-o data.js]

Requires: openpyxl (pip install openpyxl).
Prints a data-health report; nothing is uploaded anywhere.
Do not commit a data.js built from real data to a public repository.
"""
import argparse, datetime as dt, json, math, re, sys
from openpyxl import load_workbook

PILLARS = [  # order, colours and icons match the approved design
    {"name": "Culture", "color": "#073BD1", "icon": "sun",
     "def": "Strengthening SAP KSA’s culture: engagement, belonging, collaboration and a motivating workplace."},
    {"name": "Community", "color": "#248373", "icon": "network",
     "def": "An inclusive professional network where young staff share ideas, learn from each other and build relationships across SAP KSA."},
    {"name": "Business Priorities", "color": "#7561ac", "icon": "layers",
     "def": "Young staff bringing fresh ideas and practical solutions to SAP’s business goals in KSA."},
]
UNITS = {"%": "%", "pp": " pp", "score / 5": " / 5", "count": "", "per month": "", "steps": "", "M": "M"}
TIPS = {"Total tasks": "All Planner tasks in the selected initiatives. A shared task counts once.",
        "Open": "Tasks not yet marked complete in Planner.", "Completed": "Tasks marked complete in Planner.",
        "Late / overdue": "Completed after the due date, or still open past it.",
        "On time": "Completed on or before the due date. Tasks without a due date are excluded."}

warnings = []
def warn(msg): warnings.append(msg)
def clean(s):
    """Text from spreadsheets is untrusted: strip markup characters and invisible spaces."""
    if s is None: return ""
    return str(s).replace("​", "").replace("<", "‹").replace(">", "›").strip()
def is_example(id_, label):
    """Rows marked EXAMPLE in the workbook are guidance for whoever fills it in, never data."""
    return clean(label).upper() == "EXAMPLE" or clean(id_).upper().startswith("EX")
def key(s): return re.sub(r"\s+", " ", clean(s)).lower()
def as_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}", v.strip()): return dt.date.fromisoformat(v.strip()[:10])
    return None
def person(name):  # Planner writes "Last, First"
    name = clean(name)
    if "," in name:
        last, first = [x.strip() for x in name.split(",", 1)]
        return f"{first} {last.title() if last.isupper() else last}"
    return name
def num(v):
    if v is None or v == "": return None
    try: return float(v)
    except (TypeError, ValueError): return None
def tidy(x):
    if x is None: return None
    x = round(x, 2)
    return int(x) if x == int(x) else x

def sheet_rows(ws, header_row):
    head = [clean(c.value) for c in ws[header_row]]
    for r in ws.iter_rows(min_row=header_row + 1):
        vals = {head[i]: c for i, c in enumerate(r) if i < len(head) and head[i]}
        if any(c.value not in (None, "") for c in vals.values()): yield vals

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("planner"); ap.add_argument("council"); ap.add_argument("-o", "--out", default="data.js")
    ap.add_argument("--sample", action="store_true", help="build the sample data set (data-sample.js) instead of live data")
    args = ap.parse_args()
    pw = load_workbook(args.planner, data_only=True)
    cw = load_workbook(args.council, data_only=True)

    # ---- snapshot date
    as_of = None
    if "Plan" in pw.sheetnames:
        for r in sheet_rows(pw["Plan"], 1): as_of = as_date(r.get("Date of export").value if r.get("Date of export") else None)
    as_of = as_of or dt.date.today()

    # ---- settings
    st = {clean(r["Setting"].value): r["Value"].value for r in sheet_rows(cw["Settings"], 4)}
    headcount = num(st.get("SAP KSA headcount"))
    penalty = num(st.get("Delay penalty per week")) or 0.1
    low_n = num(st.get("Low-sample warning below")) or 10
    if not headcount: warn("Settings: KSA headcount is empty.")

    # ---- initiatives (Planner Setup), optionally overridden by Planner Goals
    inits = []
    for r in sheet_rows(cw["Planner Setup"], 4):
        name = clean(r["Initiative (dashboard name)"].value)
        if not name or not clean(r["Current Planner bucket"].value): continue
        pillar = clean(r["Pillar"].value)
        if not pillar: continue
        if pillar not in [p["name"] for p in PILLARS]: warn(f"Planner Setup: '{name}' has no valid pillar."); continue
        inits.append({"name": name, "bucket": key(r["Current Planner bucket"].value), "pillar": [p["name"] for p in PILLARS].index(pillar),
                      "focal": ", ".join(x.strip() for x in clean(r["Focal point(s)"].value).split(";") if x.strip()) or "Unassigned",
                      "start": as_date(r["Planned start"].value) if r.get("Planned start") else None,
                      "end": as_date(r["Planned finish"].value) if r.get("Planned finish") else None,
                      "status": clean(r["Goal status"].value) if r.get("Goal status") else "", "goal": False})
    if "Goals" in pw.sheetnames:
        for g in sheet_rows(pw["Goals"], 1):
            gname = key(g["Goal name"].value)
            for it in inits:
                if gname in (it["bucket"], key(it["name"])):
                    it["goal"] = True
                    it["start"] = as_date(g["Start date"].value) or it["start"]
                    it["end"] = as_date(g["Finish date"].value) or it["end"]
                    it["status"] = clean(g["Status"].value) or it["status"]
    for it in inits:
        if not it["goal"] and not it["end"]: warn(f"{it['name']}: no Planner Goal found. Create one named like the bucket, with start date, finish date and status.")
        elif not it["end"]: warn(f"{it['name']}: no planned finish date.")
        if it["focal"] == "Unassigned": warn(f"{it['name']}: no focal point.")
        if it["goal"] and not it["status"]: warn(f"{it['name']}: no goal status (shows as 'No status').")

    # ---- people and tasks from Planner
    people = []
    if "Users" in pw.sheetnames:
        people = [person(r["User Name"].value) for r in sheet_rows(pw["Users"], 1) if r.get("User Name") and r["User Name"].value]
    raw = []
    for r in sheet_rows(pw["Consolidated Data"], 1):
        b = key(r["Bucket"].value)
        idx = next((i for i, it in enumerate(inits) if it["bucket"] == b), None)
        title = clean(r["Task Name"].value)
        if idx is None: warn(f"Task '{title}': bucket '{clean(r['Bucket'].value)}' is not in Planner Setup; skipped."); continue
        who = [person(x) for x in clean(r["Assigned To"].value).split(";") if x.strip()]
        for w in who:
            if w not in people: people.append(w)
        items = [clean(x) for x in clean(r["Checklist Items"].value).split(";") if x.strip()]
        done_items = clean(r["Completed Checklist Items"].value).split("/")[0] if r.get("Completed Checklist Items") else "0"
        raw.append({"i": idx, "title": title, "who": [people.index(w) for w in who],
                    "start": as_date(r["Start date"].value), "due": as_date(r["Due date"].value),
                    "done": as_date(r["Completed Date"].value) if clean(r["Status"].value).lower() == "completed" else None,
                    "created": as_date(r["Created Date"].value),
                    "check": [items, int(done_items) if done_items.isdigit() else 0] if items else None})
        t = raw[-1]
        if not t["who"]: warn(f"Task '{title}': unassigned.")
        if not t["start"]: warn(f"Task '{title}': no start date.")
        if not t["due"]: warn(f"Task '{title}': no due date.")
        if t["check"]: warn(f"Task '{title}': has a {len(items)}-item checklist; split it into tasks so the steps are counted.")
        if t["due"] and t["created"] and t["due"] < t["created"]: warn(f"Task '{title}': due date is before it was created.")

    # ---- day numbering for the timeline
    dates = [d for t in raw for d in (t["start"], t["due"], t["done"], t["created"]) if d] + \
            [d for it in inits for d in (it["start"], it["end"]) if d] + [as_of]
    base = min(dates).replace(day=1)
    last = max(dates) + dt.timedelta(days=7)
    day = lambda d: None if d is None else (d - base).days + 1
    span = max(31, (last - base).days + 1)
    months = sorted({(d.year, d.month) for d in (base, last)})
    label = lambda y, m: dt.date(y, m, 1).strftime("%B")
    range_short = label(*months[0]) if len(months) == 1 else f"{label(*months[0])} – {label(*months[-1])}"
    range_label = f"{range_short} {last.year}"

    # ---- metrics (one row per metric, one column per month)
    mws = cw["Metrics"]
    head = {}
    for c in mws[4]:
        if c.value and not isinstance(c.value, (dt.date, dt.datetime)): head.setdefault(clean(c.value), c.column)
    month_cols = [(c.column, as_date(c.value)) for c in mws[4] if isinstance(c.value, (dt.date, dt.datetime))]
    month_cols = [(col, d) for col, d in month_cols if d <= as_of]
    metrics, by_id, no_target, no_actual = [[] for _ in inits], {}, [], []
    for r in range(5, mws.max_row + 1):
        mid = clean(mws.cell(r, head["Metric ID"]).value)
        if not mid or is_example(mid, mws.cell(r, head["Initiative"]).value): continue
        g = lambda h: mws.cell(r, head[h]).value if h in head else None
        ini = clean(g("Initiative"))
        unit = clean(g("Unit"))
        def val(cell):
            v = num(cell.value)
            if v is not None and unit == "%" and abs(v) <= 1 and "%" in (cell.number_format or ""): v *= 100
            return v
        series = [(d, val(mws.cell(r, col))) for col, d in month_cols]
        series = [(d, v) for d, v in series if v is not None]
        cmt = mws.cell(r, head["Metric"]).comment
        src = clean(g("Data source")) or (re.sub(r"^Data source:\s*", "", cmt.text.split("\n")[0]) if cmt else "")
        rec = {"id": mid, "init": ini, "level": clean(g("Level")), "name": clean(g("Metric")), "unit": unit,
               "target": num(g("Target")), "series": series, "note": clean(g("Latest note")),
               "def": clean(g("Definition (shown on hover)")), "how": clean(g("How it’s calculated")), "src": clean(src),
               "n": num(g("Responses")), "overview": clean(g("Feeds overview measure")), "survey": "survey" in clean(src).lower()}
        rec["actual"] = series[-1][1] if series else None
        if clean(g("Status")) == "Dropped": continue
        idx = next((i for i, it in enumerate(inits) if it["name"] == ini), None)
        if rec["level"] == "Time" and idx is not None:  # delivered on time, from Planner
            ts = [t for t in raw if t["i"] == idx]
            end = inits[idx]["end"]
            if ts and end:
                closed = all(t["done"] for t in ts)
                ref = max(t["done"] for t in ts) if closed else as_of  # open: judged against today
                late = max(0, (ref - end).days)
                rec["actual"] = max(0.0, 100 - 100 * penalty * math.ceil(late / 7))
                rec["note"] = rec["note"] if rec["note"] and rec["note"] != "Calculated from Planner" else \
                    ("Closed " if closed else "Open, ") + (f"{late} day(s) after the planned finish" if late else "on schedule")
            rec["unit"] = "%"
        if not rec["def"]: warn(f"Metric {mid}: no definition.")
        if rec["actual"] is not None and unit == "%" and 0 < rec["actual"] < 1: warn(f"Metric {mid}: {rec['actual']} looks like a fraction; type 30 for 30%.")
        if rec["actual"] is not None and rec["survey"] and (rec["n"] or 0) < low_n:
            warn(f"Metric {mid}: only {int(rec['n'] or 0)} survey responses (below {int(low_n)}); treat with care.")
        if rec["target"] is None and rec["level"] != "Time": no_target.append(mid)
        if rec["actual"] is None and rec["level"] != "Time": no_actual.append(mid)
        by_id[mid] = rec
        if idx is not None:
            rec["ref"] = f"M:{idx}:{len(metrics[idx])}"; metrics[idx].append(rec)

    if no_target: warn(f"{len(no_target)} metric(s) have no target, so they show as Tracking: {', '.join(no_target)}.")
    if no_actual: warn(f"{len(no_actual)} metric(s) have no actual up to {as_of:%b %Y}: {', '.join(no_actual)}.")
    # ---- pillar KPIs
    kws = cw["Pillar KPIs"]
    kh = {}
    for c in kws[4]:
        if c.value: kh.setdefault(clean(c.value), c.column)  # first column with a given header wins
    kpis, kpi_no_target = [[] for _ in PILLARS], []
    for r in range(5, kws.max_row + 1):
        kid = clean(kws.cell(r, kh["KPI ID"]).value)
        if not kid or is_example(kid, kws.cell(r, kh["Pillar"]).value): continue
        g = lambda h: kws.cell(r, kh[h]).value if h in kh else None
        pname = clean(g("Pillar"))
        if pname not in [p["name"] for p in PILLARS]: warn(f"KPI {kid}: unknown pillar."); continue
        links = [by_id.get(clean(g(h))) for h in ("Fed by metric 1", "Fed by metric 2") if clean(g(h))]
        if any(l is None for l in links) or not links: warn(f"KPI {kid}: linked metric not found."); links = [l for l in links if l]
        agg = clean(g("Aggregation")) or "Single"
        def combine(vals):
            vals = [v for v in vals if v is not None]
            if not vals: return None
            return sum(vals) if agg == "Sum" else sum(vals) / len(vals) if agg == "Average" else vals[0]
        unit = links[0]["unit"] if links else ""
        target = num(g("Target override")) if "Target override" in kh else num(g("Target"))
        if target is None and links:
            lt = [l["target"] for l in links if l["target"] is not None]
            if lt and len(lt) == len(links):
                target = sum(lt) if agg == "Sum" else sum(lt) / len(lt) if agg == "Average" else lt[0]
        if target is None: kpi_no_target.append(kid)
        chips = [l.get("ref") for l in links]
        if links and all(chips):
            joined = " and ".join("{" + c + "}" for c in chips)
            how = ("Equals " if agg == "Single" or len(chips) == 1 else "Sum of " if agg == "Sum" else "Average of ") + joined + "."
        elif links:
            how = links[0]["how"] or "Council-wide measure."
        else:
            how = "Not linked to a metric yet."
        kpis[[p["name"] for p in PILLARS].index(pname)].append({
            "id": kid, "name": clean(g("KPI")), "unit": unit, "target": target, "links": links, "combine": combine,
            "def": clean(g("Definition (shown on hover)")), "how": how,
            "src": links[0]["src"] if links else "", "from": [c for c in chips if c],
            "inits": sorted({int(c.split(":")[1]) for c in chips if c})})

    if kpi_no_target: warn(f"{len(kpi_no_target)} pillar KPI(s) have no target, so they are left out of the pillar score: {', '.join(kpi_no_target)}.")
    # ---- pillar scores now and per month (for sparkline and trend)
    def kpi_value(k, upto=None):
        vals = []
        for l in k["links"]:
            s = [v for d, v in l["series"] if upto is None or d <= upto]
            vals.append(s[-1] if s else (l["actual"] if upto is None else None))
        return k["combine"](vals)
    def pillar_score(p, upto=None):
        att = [min(kpi_value(k, upto) / k["target"], 1) * 100 for k in kpis[p] if k["target"] and kpi_value(k, upto) is not None]
        return round(sum(att) / len(att)) if att else None
    pillars_out = []
    for p, meta in enumerate(PILLARS):
        hist = [s for s in (pillar_score(p, d) for _, d in month_cols[:-1]) if s is not None]
        pillars_out.append({"name": meta["name"], "color": meta["color"], "icon": meta["icon"],
                            "prior": hist[-1] if hist else None, "history": hist,
                            "kpis": [[k["name"], tidy(kpi_value(k)), tidy(k["target"]), UNITS.get(k["unit"], ""), k["inits"]] for k in kpis[p]]})

    # ---- overview measures come from metrics marked in “Feeds overview measure”
    reach, survey = [[] for _ in inits], [[0, 0] for _ in inits]
    feeds = lambda name: [m for m in by_id.values() if m["overview"] == name]
    students = [0] * len(inits)
    for m in feeds("Students reached"):
        if m.get("ref") and m["actual"] is not None: students[int(m["ref"].split(":")[1])] += tidy(m["actual"])
    sat = [[0.0, 0] for _ in inits]
    for m in feeds("Average satisfaction"):
        if m.get("ref") and m["actual"] is not None and m["n"]:
            i = int(m["ref"].split(":")[1]); sat[i][0] += m["actual"] * m["n"]; sat[i][1] += int(m["n"])
        elif m.get("ref") and m["actual"] is not None: warn(f"Metric {m['id']}: satisfaction has no Responses, so it is left out of the weighted average.")
    if any(n for _, n in sat): survey = [[round(a / n, 2), n] if n else [0, 0] for a, n in sat]
    emp = feeds("Employees reached")
    employees = tidy(emp[0]["actual"]) if emp and emp[0]["actual"] is not None else None
    sm = (feeds("Students reached") or [None])[0]
    sat_refs = [m["ref"] for m in feeds("Average satisfaction") if m.get("ref")]

    # ---- assemble
    status_map = {"on track": True, "completed": True, "done": True, "at risk": False, "behind": False, "off track": False}
    stud_ref = sm.get("ref") if sm else None
    data = {
        "asOf": as_of.isoformat(), "sample": args.sample, "headcount": headcount, "people": people,
        "base": base.isoformat(), "today": day(as_of), "span": span, "rangeLabel": range_label, "rangeShort": range_short,
        "pillars": pillars_out,
        "initiatives": [[it["name"], it["pillar"], it["focal"],
                         it["start"].isoformat() if it["start"] else None, it["end"].isoformat() if it["end"] else None,
                         min([t["start"] or t["created"] for t in raw if t["i"] == i and (t["start"] or t["created"])], default=None) and
                         min([t["start"] or t["created"] for t in raw if t["i"] == i and (t["start"] or t["created"])]).isoformat(),
                         status_map.get(it["status"].lower()),
                         [[m["level"], m["name"], tidy(m["actual"]), tidy(m["target"]), UNITS.get(m["unit"], ""), m["note"] or None, tidy(m["n"])] for m in metrics[i]]]
                        for i, it in enumerate(inits)],
        "reach": reach, "survey": survey, "students": students, "employees": employees,
        "tasks": [[t["i"], t["title"], t["who"], day(t["start"]), day(t["due"]), day(t["done"])] + ([t["check"]] if t["check"] else []) for t in raw],
    }
    info = {
        "pillars": [{"def": p["def"]} for p in PILLARS],
        "kpis": [[{"def": k["def"], "how": k["how"], "src": k["src"] or "See linked metric", "from": k["from"]} for k in row] for row in kpis],
        "metrics": [[{"def": m["def"], "how": m["how"], "src": m["src"]} for m in row] for row in metrics],
        "stats": [
            {"name": "Employees reached", "def": "Distinct SAP KSA employees who took part in at least one council initiative.",
             "how": (by_id["CW-02"]["how"] if "CW-02" in by_id else "Unique employees across every initiative’s participant list.") + " Entered once a month by the dashboard owner.",
             "src": (emp[0]["src"] if emp else "Participant lists")},
            {"name": "Students reached", "def": "Students engaged through council workshops.",
             "how": ("Equals {" + stud_ref + "}.") if stud_ref else "Students who attended council workshops.",
             "src": sm["src"] if sm else "Workshop attendance"},
            {"name": "Average satisfaction", "def": "Average participant rating across council surveys.",
             "how": "Satisfaction ratings weighted by their number of responses: " + (", ".join("{" + r + "}" for r in sat_refs) or "no satisfaction metrics yet") + ".",
             "src": "Feedback surveys (Microsoft Forms)"},
            {"name": "Initiatives on track", "def": "Share of initiatives whose focal points report them on track.",
             "how": "Initiatives with goal status On track ÷ initiatives shown.", "src": "Microsoft Planner goal status"}],
        "tips": TIPS,
    }
    var = "DATA_SAMPLE" if args.sample else "DATA_LIVE"
    data["info"] = info
    js = ("// Generated by tools/build_data.py on " + dt.datetime.now().strftime("%Y-%m-%d %H:%M") +
          ". Do not edit by hand.\nwindow." + var + " = " + json.dumps(data, ensure_ascii=False, indent=1) + ";\n")
    open(args.out, "w", encoding="utf-8").write(js)

    done = sum(1 for t in raw if t["done"])
    print(f"Wrote {args.out}: snapshot {as_of}, {len(inits)} initiatives, {len(raw)} tasks ({done} completed, {len(raw) - done} open), "
          f"{sum(len(m) for m in metrics)} metrics, {sum(len(k) for k in kpis)} pillar KPIs, {len(people)} people.")
    print(f"\nData health: {len(warnings)} item(s) to fix" + (":" if warnings else "."))
    for w in warnings: print("  -", w)

if __name__ == "__main__":
    main()
