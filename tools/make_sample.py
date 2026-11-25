#!/usr/bin/env python3
"""Create the fully filled SAMPLE inputs (placeholder names only) and build data-sample.js from them.

Usage:
    python3 tools/make_sample.py BLANK_COUNCIL_TEMPLATE.xlsx

Writes tools/sample/Planner_Sample.xlsx and tools/sample/Council_Data_Sample.xlsx,
then runs build_data.py --sample to produce data-sample.js. All values are illustrative.
"""
import datetime as dt, os, subprocess, sys
from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "sample")
D = lambda m, d: dt.date(2026 if m >= 10 else 2027, m, d)
AS_OF = dt.date(2027, 1, 15)
MEMBERS = [f"Member {c}" for c in "ABCDEFGH"]
BUCKETS = {"Intro Video": "SAP KSA Youth Council Introduction Video", "Youth Workshops": "University Engagement",
           "Steps Challenge": "Steps challenge", "Interview Series": "Internal Interview Series"}
GOALS = {  # planned start, planned finish, status
    "Intro Video": (D(11, 3), D(11, 29), "Completed"), "Youth Workshops": (D(11, 2), D(2, 28), "On track"),
    "Steps Challenge": (D(11, 2), D(12, 20), "Completed"), "Interview Series": (D(11, 16), D(3, 31), "At risk")}
FOCAL = {"Intro Video": "Member D; Member I", "Youth Workshops": "Member G",
         "Steps Challenge": "Member H; Member E; Member J", "Interview Series": "Member A"}
# initiative, title, assignees, start, due, done, checklist (items, completed)
TASKS = [
 ("Intro Video", "Agree the video brief", "DA", D(11, 3), D(11, 5), D(11, 5), None),
 ("Intro Video", "Draft the introduction script", "D", D(11, 5), D(11, 8), D(11, 9), None),
 ("Intro Video", "Review the storyboard", "B", D(11, 8), D(11, 10), D(11, 12), None),
 ("Intro Video", "Record council introductions", "CD", D(11, 12), D(11, 15), D(11, 15), None),
 ("Intro Video", "Edit the first cut", "B", D(11, 16), D(11, 19), D(11, 20), None),
 ("Intro Video", "Review captions and branding", "A", D(11, 20), D(11, 22), D(11, 23), None),
 ("Intro Video", "Publish the launch video", "D", D(11, 24), D(11, 29), D(11, 30), None),
 ("Youth Workshops", "Confirm university contacts", "G", D(11, 2), D(11, 4), D(11, 4), None),
 ("Youth Workshops", "Scope the student workshop", "GF", D(11, 4), D(11, 7), D(11, 9), None),
 ("Youth Workshops", "Draft the invitation", "F", D(11, 9), D(11, 12), D(11, 11), None),
 ("Youth Workshops", "Prepare the facilitator guide", "C", D(11, 10), D(11, 13), D(11, 16), None),
 ("Youth Workshops", "Run the pilot workshop", "GC", D(11, 30), D(12, 3), D(12, 3), None),
 ("Youth Workshops", "Collect the post-event survey", "F", D(12, 3), D(12, 10), D(12, 9), None),
 ("Youth Workshops", "Confirm second university", "G", D(12, 10), D(1, 10), None, None),
 ("Youth Workshops", "Follow up with third university", "F", D(12, 15), D(1, 25), None, None),
 ("Youth Workshops", "Book January session space", "CG", D(1, 5), D(1, 20), None, (["Shortlist rooms", "Confirm catering", "Send calendar invite"], 1)),
 ("Steps Challenge", "Agree collaboration with Health, Safety & Wellbeing", "EH", D(11, 2), D(11, 8), D(11, 4), None),
 ("Steps Challenge", "Set up the challenge in Fit@SAP", "H", D(11, 5), D(11, 12), D(11, 14), None),
 ("Steps Challenge", "Launch registration", "HE", D(11, 15), D(11, 20), D(11, 20), None),
 ("Steps Challenge", "Send mid-challenge reminder", "E", D(12, 1), D(12, 3), D(12, 2), None),
 ("Steps Challenge", "Review final-week retention", "E", D(12, 15), D(12, 18), D(12, 19), None),
 ("Steps Challenge", "Collect participant feedback", "H", D(12, 19), D(12, 22), D(12, 27), None),
 ("Interview Series", "Select the first guests", "A", D(11, 16), D(11, 23), D(11, 25), None),
 ("Interview Series", "Confirm guest availability", "B", D(11, 24), D(11, 30), D(11, 30), None),
 ("Interview Series", "Prepare interview questions", "AB", D(12, 1), D(12, 5), D(12, 6), None),
 ("Interview Series", "Record the first interview", "C", D(12, 8), D(12, 12), D(12, 12), None),
 ("Interview Series", "Edit the footage", "B", D(12, 13), D(12, 20), D(12, 22), None),
 ("Interview Series", "Publish episode 1", "A", D(12, 22), D(12, 24), D(12, 24), None),
 ("Interview Series", "Record the second interview", "C", D(1, 5), D(1, 12), D(1, 13), None),
 ("Interview Series", "Publish episode 2", "A", D(1, 14), D(1, 21), None, None),
]
# metric id: target, owner, {month: value}, responses (latest), latest note
NOV, DEC, JAN = D(11, 1), D(12, 1), D(1, 1)
METRICS = {
 "I1-01": (50, "Member D", {DEC: 42, JAN: 51}, None, "204 of 400 employees"),
 "I1-02": (30, "Member D", {DEC: 16, JAN: 24}, 135, "30% baseline → 54%"),
 "I1-03": (20, "Member D", {DEC: 11, JAN: 17}, None, "+6 since December"),
 "I1-05": (60, "Member B", {DEC: 58, JAN: 63}, None, "Average of 2:14 watched"),
 "I1-06": (4.2, "Member D", {DEC: 4.3, JAN: 4.4}, 96, "96 ratings"),
 "I2-01": (60, "Member G", {DEC: 60, JAN: 92}, None, "Cumulative registrations"),
 "I2-02": (50, "Member G", {DEC: 48, JAN: 48}, None, "Pilot workshop attendance"),
 "I2-03": (4.2, "Member F", {DEC: 4.4}, 48, "48 of 48 attendees answered"),
 "I2-04": (6, "Member G", {DEC: 5, JAN: 6}, None, "Facilitators to date"),
 "I2-05": (70, "Member F", {DEC: 68}, 48, "Post-event survey"),
 "I2-06": (10, "Member G", {DEC: 6, JAN: 9}, None, "CVs and applications"),
 "I3-01": (30, "Member H", {NOV: 15, DEC: 18}, None, "72 of 400 employees"),
 "I3-02": (10, "Member H", {NOV: 7, DEC: 9}, None, "Final team count"),
 "I3-03": (60, "Member E", {DEC: 61}, None, "Week 1 86% · week 2 64% · week 3 61%"),
 "I3-04": (5, "Member H", {NOV: 1.8, DEC: 3.6}, None, "3.6M of 5M steps"),
 "I3-05": (70, "Member E", {DEC: 61}, 50, "Post-challenge survey"),
 "I3-06": (60, "Member E", {DEC: 52}, 50, "Post-challenge survey"),
 "I3-08": (4.2, "Member H", {DEC: 4.1}, 50, "Post-challenge survey"),
 "I4-01": (1, "Member A", {NOV: 0, DEC: 1, JAN: 1}, None, "Episode 2 due 21 Jan"),
 "I4-02": (100, "Member A", {DEC: 84, JAN: 96}, None, "Average across episodes"),
 "I4-03": (60, "Member B", {DEC: 53, JAN: 57}, None, "Average across episodes"),
 "I4-04": (10, "Member A", {DEC: 8, JAN: 11}, None, "Questions and nominations"),
 "I4-05": (70, "Member A", {DEC: 72, JAN: 74}, 40, "Pulse survey"),
 "I4-06": (20, "Member A", {DEC: 12, JAN: 15}, None, "Coffee chats and mentoring"),
 "I4-08": (4.2, "Member A", {DEC: 4.5, JAN: 4.4}, 40, "Episode feedback"),
 "I1-04": (100, "Member D", {}, None, None), "I2-07": (100, "Member G", {}, None, None),
 "I3-07": (100, "Member H", {}, None, None), "I4-07": (100, "Member A", {}, None, None),
 "CW-01": (4.2, "Member I", {NOV: 3.7, JAN: 3.9}, 210, "Quarterly pulse survey"),
 "CW-02": (250, "Member I", {NOV: 96, DEC: 180, JAN: 212}, None, "Deduplicated across all initiatives"),
}

def planner_export(path):
    wb = Workbook(); ws = wb.active; ws.title = "Plan"
    ws.append(["Plan ID", "Plan name", "Date of export"]); ws.append(["sample", "Initiatives-Execution Tracker", AS_OF.isoformat()])
    cd = wb.create_sheet("Consolidated Data")
    cd.append(["Task ID", "Task Name", "Bucket", "Goal", "Status", "Priority", "Assigned To", "Created By", "Created Date", "Due date",
               "Start date", "Finish date", "Is Recurring", "Late", "Completed Date", "Completed By", "Completed Checklist Items", "Checklist Items", "Labels", "Notes"])
    for n, (ini, title, who, start, due, done, check) in enumerate(TASKS, 1):
        names = ";".join(f"{m}" for m in (f"Member {c}" for c in who))
        cd.append([f"T{n}", title, BUCKETS[ini], ini, "Completed" if done else ("In progress" if start <= AS_OF else "Not started"), "Medium",
                   names, "Member D", start.isoformat(), due.isoformat(), start.isoformat(), None, "false", "false",
                   done.isoformat() if done else None, names.split(";")[0] if done else None,
                   f"{check[1]}/{len(check[0])}" if check else None, ";".join(check[0]) if check else None, None, None])
    g = wb.create_sheet("Goals")
    g.append(["Goal ID", "Goal name", "Status", "Priority", "Start date", "Finish date", "Connected tasks", "Description"])
    for n, (ini, (s, f, st)) in enumerate(GOALS.items(), 1):
        g.append([f"G{n}", BUCKETS[ini], st, "Medium", s.isoformat(), f.isoformat(), sum(1 for t in TASKS if t[0] == ini), ""])
    b = wb.create_sheet("Buckets"); b.append(["Bucket ID", "Bucket Name"])
    for n, name in enumerate(BUCKETS.values(), 1): b.append([f"B{n}", name])
    u = wb.create_sheet("Users"); u.append(["User ID", "User Name", "Email"])
    for n, m in enumerate(MEMBERS, 1): u.append([f"U{n}", m, f"member.{m[-1].lower()}@example.com"])
    wb.save(path)

def council_workbook(template, path):
    wb = load_workbook(template)
    wb["Settings"]["B5"] = 400
    ps = wb["Planner Setup"]
    for r in range(5, 13):
        ini = ps.cell(r, 1).value
        if ini in FOCAL: ps.cell(r, 4).value = FOCAL[ini]
    ms = wb["Metrics"]
    head, months = {}, {}
    for c in ms[4]:
        if isinstance(c.value, str): head.setdefault(c.value, c.column)
        elif isinstance(c.value, (dt.date, dt.datetime)): months[c.value.date() if isinstance(c.value, dt.datetime) else c.value] = c.column
    for r in range(5, ms.max_row + 1):
        mid = ms.cell(r, 1).value
        if mid not in METRICS: continue
        tgt, owner, vals, n, note = METRICS[mid]
        ms.cell(r, head["Target"]).value = tgt; ms.cell(r, head["Owner"]).value = owner
        ms.cell(r, head["Responses"]).value = n; ms.cell(r, head["Latest note"]).value = note
        ms.cell(r, head["Status"]).value = "Adopt"
        for m, v in vals.items(): ms.cell(r, months[m]).value = v
    wb.save(path)

if __name__ == "__main__":
    if len(sys.argv) != 2: sys.exit(__doc__)
    os.makedirs(OUT, exist_ok=True)
    pl, cw = os.path.join(OUT, "Planner_Sample.xlsx"), os.path.join(OUT, "Council_Data_Sample.xlsx")
    planner_export(pl); council_workbook(sys.argv[1], cw)
    subprocess.run([sys.executable, os.path.join(HERE, "build_data.py"), pl, cw, "--sample",
                    "-o", os.path.join(os.path.dirname(HERE), "data-sample.js")], check=True)
