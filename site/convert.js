// Reads the Planner export and the Council Data workbook (via SheetJS) and builds the dashboard's live data.
// Returns { errors, warnings, data, summary }. Any error blocks publishing; warnings are data-health notes.
// Mirrors tools/build_data.py.
(function () {
  const PILLARS = [
    { name: 'Culture', color: '#073BD1', icon: 'sun', def: 'Strengthening SAP KSA’s culture: engagement, belonging, collaboration and a motivating workplace.' },
    { name: 'Community', color: '#248373', icon: 'network', def: 'An inclusive professional network where young staff share ideas, learn from each other and build relationships across SAP KSA.' },
    { name: 'Business Priorities', color: '#7561ac', icon: 'layers', def: 'Young staff bringing fresh ideas and practical solutions to SAP’s business goals in KSA.' }];
  const PNAMES = PILLARS.map(p => p.name);
  const UNITS = { '%': '%', pp: ' pp', 'score / 5': ' / 5', count: '', 'per month': '', steps: '', M: 'M' };
  const TIPS = { 'Total tasks': 'All Planner tasks in the selected initiatives. A shared task counts once.', Open: 'Tasks not yet marked complete in Planner.', Completed: 'Tasks marked complete in Planner.', 'Late / overdue': 'Completed after the due date, or still open past it.', 'On time': 'Completed on or before the due date. Tasks without a due date are excluded.' };
  const STATUS = { 'on track': true, completed: true, done: true, 'at risk': false, behind: false, 'off track': false };

  const clean = v => v == null ? '' : String(v).replace(/​/g, '').replace(/</g, '‹').replace(/>/g, '›').trim();
  const key = v => clean(v).replace(/\s+/g, ' ').toLowerCase();
  const isExample = (id, label) => clean(label).toUpperCase() === 'EXAMPLE' || clean(id).toUpperCase().startsWith('EX');
  const pad = n => String(n).padStart(2, '0');
  const iso = d => d ? `${d.y}-${pad(d.m)}-${pad(d.d)}` : null;
  const ord = d => Date.UTC(d.y, d.m - 1, d.d) / 864e5;
  function asDate(v) {
    if (v instanceof Date && !isNaN(v)) { const x = new Date(v.getTime() + 12 * 3600e3); return { y: x.getFullYear(), m: x.getMonth() + 1, d: x.getDate() } }  // round: Excel dates can arrive a few seconds before midnight
    const m = typeof v === 'string' && v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? { y: +m[1], m: +m[2], d: +m[3] } : null;
  }
  const fromOrd = n => { const x = new Date(n * 864e5); return { y: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate() } };
  const fmt = d => d ? new Date(Date.UTC(d.y, d.m - 1, d.d)).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' }) : '';
  const fmtMY = d => d ? new Date(Date.UTC(d.y, d.m - 1, 1)).toLocaleDateString('en-GB', { month: 'short', year: 'numeric', timeZone: 'UTC' }) : '';
  const tidy = x => x == null ? null : Math.round(x * 100) / 100;
  function person(n) {
    n = clean(n); if (!n.includes(',')) return n;
    const [last, first] = n.split(',').map(s => s.trim());
    return `${first} ${last === last.toUpperCase() ? last[0] + last.slice(1).toLowerCase() : last}`;
  }

  // Sheet helpers. Rows and columns are 1-based like Excel.
  function sheetAPI(ws) {
    const ref = ws && ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { e: { r: -1, c: -1 } };
    const cell = (r, c) => ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })];
    const val = (r, c) => { const x = cell(r, c); return x ? x.v : null };
    const head = row => { const h = {}; for (let c = 1; c <= ref.e.c + 1; c++) { const v = val(row, c); if (typeof v === 'string' && v.trim() && !(clean(v) in h)) h[clean(v)] = c } return h };
    return { cell, val, head, rows: ref.e.r + 1, cols: ref.e.c + 1 };
  }
  const colName = c => XLSX.utils.encode_col(c - 1);

  function convert(plannerBook, councilBook, files) {
    const errors = [], warnings = [];
    const err = (where, problem, fix) => errors.push({ where, problem, fix });
    const warn = m => warnings.push(m);

    // ---------- structure checks ----------
    const pNeed = ['Consolidated Data'], cNeed = ['Settings', 'Planner Setup', 'Metrics', 'Pillar KPIs'];
    const missP = pNeed.filter(s => !plannerBook.SheetNames.includes(s));
    const missC = cNeed.filter(s => !councilBook.SheetNames.includes(s));
    if (missP.length) err(files[0], `This doesn't look like a Planner export: the sheet “${missP.join('”, “')}” is missing.`, 'In Planner, use … › Export plan to Excel and upload that file without editing it.');
    if (missC.length) err(files[1], `This doesn't look like the Council Data workbook: sheet(s) “${missC.join('”, “')}” missing.`, 'Download the template from this page, copy your data into it, and keep the sheet names unchanged.');
    if (errors.length) return { errors, warnings };

    const P = sheetAPI(plannerBook.Sheets['Consolidated Data']);
    const ph = P.head(1);
    const pCols = ['Task Name', 'Bucket', 'Status', 'Assigned To', 'Created Date', 'Due date', 'Start date', 'Completed Date'];
    const missCol = pCols.filter(c => !(c in ph));
    if (missCol.length) err(`${files[0]} › Consolidated Data`, `Column(s) missing: ${missCol.join(', ')}.`, 'Re-export the plan from Planner; do not delete or rename columns.');

    const S = sheetAPI(councilBook.Sheets.Settings), PS = sheetAPI(councilBook.Sheets['Planner Setup']);
    const M = sheetAPI(councilBook.Sheets.Metrics), K = sheetAPI(councilBook.Sheets['Pillar KPIs']);
    const mh = M.head(4), kh = K.head(4), psh = PS.head(4), sh = S.head(4);
    const need = (h, cols, where) => { const miss = cols.filter(c => !(c in h)); if (miss.length) err(where, `Column(s) missing or renamed: ${miss.join(', ')}.`, 'Use the latest template from this page and do not rename the header row (row 4).') };
    const oldLinks = 'Fed by metric 1' in kh;  // earlier template versions linked metrics on the Pillar KPIs tab
    need(mh, ['Metric ID', 'Initiative', 'Level', 'Metric', 'Unit', 'Target', 'Status', 'Data source', 'Responses', 'Feeds overview measure', 'Latest note', 'Definition (shown on hover)', 'How it’s calculated'].concat(oldLinks ? [] : ['Feeds pillar KPI']), `${files[1]} › Metrics`);
    need(kh, ['KPI ID', 'Pillar', 'KPI', 'Aggregation', 'Target override', 'Definition (shown on hover)'].concat(oldLinks ? ['Fed by metric 2'] : []), `${files[1]} › Pillar KPIs`);
    need(psh, ['Initiative (dashboard name)', 'Current Planner bucket', 'Pillar', 'Focal point(s)'], `${files[1]} › Planner Setup`);
    need(sh, ['Setting', 'Value'], `${files[1]} › Settings`);
    if (errors.length) return { errors, warnings };

    // ---------- settings ----------
    const settings = {};
    for (let r = 5; r <= S.rows; r++) { const k = clean(S.val(r, sh.Setting)); if (k) settings[k] = S.val(r, sh.Value) }
    const num = v => { if (v == null || v === '') return null; const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, '')); return isFinite(n) ? n : NaN };
    const headcount = num(settings['SAP KSA headcount']);
    if (headcount === null) warn('Settings: KSA headcount is empty.');
    else if (isNaN(headcount) || headcount <= 0) err(`${files[1]} › Settings`, `KSA headcount “${settings['SAP KSA headcount']}” is not a positive number.`, 'Type the headcount as a whole number, e.g. 400.');
    const penalty = num(settings['Delay penalty per week']) || 0.1, lowN = num(settings['Low-sample warning below']) || 10;
    const onT = num(settings['On-target threshold']), nearT = num(settings['Near-target threshold']);
    if ([onT, nearT].some(x => x !== null && (isNaN(x) || x <= 0 || x > 1)) || (onT && nearT && nearT >= onT))
      err(`${files[1]} › Settings`, 'Status thresholds are invalid.', 'Use percentages between 1% and 100%, with the near-target threshold below the on-target threshold (defaults 90% and 70%).');

    // ---------- initiatives ----------
    const inits = [];
    for (let r = 5; r <= PS.rows; r++) {
      const name = clean(PS.val(r, psh['Initiative (dashboard name)'])), bucket = clean(PS.val(r, psh['Current Planner bucket']));
      const pillar = clean(PS.val(r, psh.Pillar));
      if (!name || !bucket || name === '☐') continue;
      if (!pillar) { err(`${files[1]} › Planner Setup, row ${r}`, `Initiative “${name}” has no pillar.`, 'Pick Culture, Community or Business Priorities in the Pillar column.'); continue }
      if (!PNAMES.includes(pillar)) { err(`${files[1]} › Planner Setup, row ${r}`, `Pillar “${pillar}” is not one of the three strategic pillars.`, 'Pick Culture, Community or Business Priorities from the drop-down.'); continue }
      if (inits.some(i => i.name === name)) { err(`${files[1]} › Planner Setup, row ${r}`, `Initiative “${name}” appears twice.`, 'Each initiative needs exactly one row.'); continue }
      if (inits.some(i => i.bucket === key(bucket))) { err(`${files[1]} › Planner Setup, row ${r}`, `Planner bucket “${bucket}” is mapped to two initiatives.`, 'Each Planner bucket belongs to exactly one initiative.'); continue }
      const focal = clean(PS.val(r, psh['Focal point(s)'])).split(';').map(s => s.trim()).filter(Boolean).join(', ');
      inits.push({ name, bucket: key(bucket), bucketLabel: bucket, pillar: PNAMES.indexOf(pillar), focal: focal || 'Unassigned', start: null, end: null, status: '', goal: false });
    }
    if (!inits.length) err(`${files[1]} › Planner Setup`, 'No initiatives are listed.', 'Add one row per initiative with its Planner bucket name and pillar.');
    if (errors.length) return { errors, warnings };

    // Planner Goals: planned dates and status
    if (plannerBook.SheetNames.includes('Goals')) {
      const G = sheetAPI(plannerBook.Sheets.Goals), gh = G.head(1);
      for (let r = 2; r <= G.rows; r++) {
        const gname = key(G.val(r, gh['Goal name'])); if (!gname) continue;
        const it = inits.find(i => i.bucket === gname || key(i.name) === gname); if (!it) continue;
        it.goal = true; it.start = asDate(G.val(r, gh['Start date'])) || it.start; it.end = asDate(G.val(r, gh['Finish date'])) || it.end; it.status = clean(G.val(r, gh.Status)) || it.status;
      }
    }
    // snapshot date
    let asOf = null;
    if (plannerBook.SheetNames.includes('Plan')) { const PL = sheetAPI(plannerBook.Sheets.Plan), plh = PL.head(1); if (plh['Date of export']) asOf = asDate(PL.val(2, plh['Date of export'])) }
    if (!asOf) { const t = new Date(); asOf = { y: t.getFullYear(), m: t.getMonth() + 1, d: t.getDate() }; warn('Planner export has no export date; today is used as the snapshot date.') }

    // ---------- people and tasks ----------
    const people = [];
    if (plannerBook.SheetNames.includes('Users')) { const U = sheetAPI(plannerBook.Sheets.Users), uh = U.head(1); for (let r = 2; r <= U.rows; r++) { const n = clean(U.val(r, uh['User Name'])); if (n) people.push(person(n)) } }
    const raw = [], unknownBuckets = new Set();
    for (let r = 2; r <= P.rows; r++) {
      const title = clean(P.val(r, ph['Task Name'])); if (!title) continue;
      const b = key(P.val(r, ph.Bucket)), idx = inits.findIndex(i => i.bucket === b);
      if (idx < 0) { unknownBuckets.add(clean(P.val(r, ph.Bucket))); continue }
      const who = clean(P.val(r, ph['Assigned To'])).split(';').map(person).filter(Boolean);
      who.forEach(w => { if (!people.includes(w)) people.push(w) });
      const items = ph['Checklist Items'] ? clean(P.val(r, ph['Checklist Items'])).split(';').map(clean).filter(Boolean) : [];
      const doneItems = ph['Completed Checklist Items'] ? parseInt(clean(P.val(r, ph['Completed Checklist Items'])).split('/')[0]) || 0 : 0;
      const completed = clean(P.val(r, ph.Status)).toLowerCase() === 'completed';
      const t = { i: idx, title, who: who.map(w => people.indexOf(w)), start: asDate(P.val(r, ph['Start date'])), due: asDate(P.val(r, ph['Due date'])), done: completed ? asDate(P.val(r, ph['Completed Date'])) : null, created: asDate(P.val(r, ph['Created Date'])), check: items.length ? [items, doneItems] : null };
      if (completed && !t.done) err(`${files[0]} › row ${r}`, `Task “${title}” is marked Completed but has no completed date.`, 'Re-export from Planner; do not edit the Completed Date column.');
      if (!t.who.length) warn(`Task “${title}”: unassigned.`);
      if (!t.start) warn(`Task “${title}”: no start date.`);
      if (!t.due) warn(`Task “${title}”: no due date.`);
      if (t.check) warn(`Task “${title}”: has a ${items.length}-item checklist; split it into tasks so the steps are counted.`);
      if (t.due && t.created && ord(t.due) < ord(t.created)) warn(`Task “${title}”: due date is before it was created.`);
      raw.push(t);
    }
    if (unknownBuckets.size) err(`${files[1]} › Planner Setup`, `Planner bucket(s) not listed: “${[...unknownBuckets].join('”, “')}”. Their tasks would be missing from the dashboard.`, 'Add each bucket to the Current Planner bucket column (exact name), or move those tasks into an existing bucket in Planner.');
    if (!raw.length) err(files[0], 'The Planner export contains no tasks for the initiatives in Planner Setup.', 'Check that bucket names in Planner Setup match Planner exactly.');
    inits.forEach(it => {
      if (!it.goal && !it.end) warn(`${it.name}: no Planner Goal found. Create one named like the bucket, with start date, finish date and status.`);
      else if (!it.end) warn(`${it.name}: no planned finish date.`);
      if (it.focal === 'Unassigned') warn(`${it.name}: no focal point.`);
      if (it.goal && !it.status) warn(`${it.name}: no goal status (shows as “No status”).`);
    });

    // ---------- metrics ----------
    const monthCols = [];
    for (let c = 1; c <= M.cols; c++) { const v = M.val(4, c); const d = v instanceof Date ? asDate(v) : null; if (d && ord(d) <= ord(asOf)) monthCols.push([c, d]) }
    const metrics = inits.map(() => []), byId = {}, noTarget = [], noActual = [];
    const g = (r, h) => M.val(r, mh[h]);
    for (let r = 5; r <= M.rows; r++) {
      const id = clean(g(r, 'Metric ID')); if (!id || isExample(id, g(r, 'Initiative'))) continue;
      const where = `${files[1]} › Metrics, row ${r} (${id})`;
      if (byId[id]) { err(where, `Metric ID ${id} is used twice.`, 'Give every metric a unique ID.'); continue }
      const ini = clean(g(r, 'Initiative')), unit = clean(g(r, 'Unit')), level = clean(g(r, 'Level'));
      if (clean(g(r, 'Status')) === 'Dropped') continue;
      const idx = inits.findIndex(i => i.name === ini);
      if (idx < 0 && ini !== 'Council-wide') { err(where, `Initiative “${ini}” is not listed in Planner Setup.`, 'Pick an initiative from the drop-down, or add it to Planner Setup first.'); continue }
      if (!(unit in UNITS)) err(where, `Unit “${unit}” is not recognised.`, `Pick one of: ${Object.keys(UNITS).join(', ')}.`);
      const numeric = (c, label) => { const raw = M.val(r, c); if (raw == null || raw === '') return null; const n = num(raw); if (isNaN(n)) { err(where, `${label} “${raw}” is not a number.`, 'Type numbers only, e.g. 30 for 30%; no % signs, text or “n/a”.'); return null } const cl = M.cell(r, c); if (unit === '%' && n > 0 && n < 1 && cl && /%/.test(cl.z || '')) return n * 100; return n };
      const target = numeric(mh.Target, 'Target'), n = numeric(mh.Responses, 'Responses');
      const series = [];
      for (const [c, d] of monthCols) { const v = numeric(c, `${fmtMY(d)} value`); if (v != null) series.push([d, v]) }
      for (const [d, v] of series.concat(target != null ? [[null, target]] : [])) {
        const lbl = d ? `${fmtMY(d)} value` : 'Target';
        if (unit === '%' && v > 0 && v < 1) err(where, `${lbl} ${v} looks like a fraction.`, 'Type 30 for 30%, not 0.3.');
        if (unit === '%' && (v < 0 || v > 1000)) err(where, `${lbl} ${v} is not a valid percentage.`, 'Type a number between 0 and 100 (above 100 only for over-achievement).');
        if (unit === 'score / 5' && (v < 1 || v > 5)) err(where, `${lbl} ${v} is outside the 1–5 rating scale.`, 'Type the average rating between 1 and 5, e.g. 4.3.');
        if (['count', 'per month', 'M'].includes(unit) && v < 0) err(where, `${lbl} ${v} is negative.`, 'Counts cannot be negative.');
      }
      if (n != null && (n < 0 || Math.round(n) !== n)) err(where, `Responses “${n}” must be a whole number.`, 'Type how many people answered, e.g. 48.');
      const src = clean(g(r, 'Data source')), survey = /survey/i.test(src);
      const rec = { id, ini, level, unit, target, series, n, src, survey, feeds: oldLinks ? '' : clean(g(r, 'Feeds pillar KPI')), row: r, name: clean(g(r, 'Metric')), note: clean(g(r, 'Latest note')), def: clean(g(r, 'Definition (shown on hover)')), how: clean(g(r, 'How it’s calculated')), overview: clean(g(r, 'Feeds overview measure')) };
      rec.actual = series.length ? series[series.length - 1][1] : null;
      if (level === 'Time' && idx >= 0) {
        const ts = raw.filter(t => t.i === idx), end = inits[idx].end;
        if (ts.length && end) { const closed = ts.every(t => t.done); const ref = closed ? Math.max(...ts.map(t => ord(t.done))) : ord(asOf); const late = Math.max(0, ref - ord(end)); rec.actual = Math.max(0, 100 - 100 * penalty * Math.ceil(late / 7)); if (!rec.note || rec.note === 'Calculated from Planner') rec.note = (closed ? 'Closed ' : 'Open, ') + (late ? `${late} day(s) after the planned finish` : 'on schedule') }
        rec.unit = '%';
      }
      if (!rec.def) warn(`Metric ${id}: no definition.`);
      if (rec.target == null && level !== 'Time') noTarget.push(id);
      if (rec.actual == null && level !== 'Time') noActual.push(id);
      if (rec.actual != null && survey && (n || 0) < lowN) warn(`Metric ${id}: only ${n || 0} survey responses (below ${lowN}); treat with care.`);
      if (rec.actual != null && rec.overview === 'Average satisfaction' && !n) err(where, 'This satisfaction rating has no Responses.', 'Type how many people answered, so the overview can weight it correctly.');
      byId[id] = rec;
      if (idx >= 0) { rec.ref = `M:${idx}:${metrics[idx].length}`; metrics[idx].push(rec) }
    }
    if (noTarget.length) warn(`${noTarget.length} metric(s) have no target, so they show as Tracking: ${noTarget.join(', ')}.`);
    if (noActual.length) warn(`${noActual.length} metric(s) have no actual up to ${fmtMY(asOf)}: ${noActual.join(', ')}.`);

    // ---------- pillar KPIs ----------
    const kpis = PILLARS.map(() => []), kNoTarget = [], kpiIds = new Set();
    const kg = (r, h) => K.val(r, kh[h]);
    for (let r = 5; r <= K.rows; r++) {
      const id = clean(kg(r, 'KPI ID')); if (!id || isExample(id, kg(r, 'Pillar'))) continue;
      const where = `${files[1]} › Pillar KPIs, row ${r} (${id})`, pname = clean(kg(r, 'Pillar'));
      if (!PNAMES.includes(pname)) { err(where, `Pillar “${pname}” is not one of the three strategic pillars.`, 'Pick a pillar from the drop-down.'); continue }
      if (kpiIds.has(id)) { err(where, `KPI ID ${id} is used twice.`, 'Give every pillar KPI a unique ID.'); continue }
      kpiIds.add(id);
      const agg = clean(kg(r, 'Aggregation')) || 'Single';
      if (!['Single', 'Sum', 'Average'].includes(agg)) { err(where, `Aggregation “${agg}” is not recognised.`, 'Pick Single, Sum or Average.'); continue }
      let links;
      if (oldLinks) {
        const ids = ['Fed by metric 1', 'Fed by metric 2'].map(h => clean(kg(r, h))).filter(Boolean);
        const bad = ids.filter(x => !byId[x]); if (bad.length) { err(where, `Linked metric ${bad.join(', ')} does not exist (or is Dropped).`, 'Pick a Metric ID from the drop-down.'); continue }
        links = ids.map(x => byId[x]);
      } else links = Object.values(byId).filter(m => m.feeds === id);
      if (!links.length) { warn(`KPI ${id} (${clean(kg(r, 'KPI'))}): no metric feeds it yet, so it shows no data. Pick it in “Feeds pillar KPI” on the Metrics tab.`) }
      if (agg === 'Single' && links.length > 1) { err(where, `Aggregation is Single but ${links.length} metrics feed this KPI (${links.map(l => l.id).join(', ')}).`, 'On the Metrics tab keep “Feeds pillar KPI” = ' + id + ' on one metric only, or set Aggregation to Sum or Average here.'); continue }
      const kpPillar = PNAMES.indexOf(pname);
      links.forEach(l => { const ii = inits.findIndex(i => i.name === l.ini); if (ii >= 0 && inits[ii].pillar !== kpPillar) warn(`Metric ${l.id} (${l.ini}, ${PNAMES[inits[ii].pillar]}) feeds ${id}, a ${pname} KPI: a metric normally feeds its own initiative’s pillar.`) });
      const combine = vals => { vals = vals.filter(v => v != null); if (!vals.length) return null; return agg === 'Sum' ? vals.reduce((a, b) => a + b, 0) : agg === 'Average' ? vals.reduce((a, b) => a + b, 0) / vals.length : vals[0] };
      let target = num(kg(r, 'Target override'));
      if (!links.length) { kpis[kpPillar].push({ id, name: clean(kg(r, 'KPI')), unit: '', target: isNaN(target) ? null : target, links: [], combine: () => null, def: clean(kg(r, 'Definition (shown on hover)')), how: 'No metric feeds this KPI yet.', src: '', from: [], inits: [] }); continue }
      if (isNaN(target)) { err(where, `Target override “${kg(r, 'Target override')}” is not a number.`, 'Type a number or leave it empty.'); continue }
      if (target == null) { const lt = links.map(l => l.target); if (lt.length && lt.every(x => x != null)) target = agg === 'Sum' ? lt.reduce((a, b) => a + b, 0) : agg === 'Average' ? lt.reduce((a, b) => a + b, 0) / lt.length : lt[0] }
      if (target == null) kNoTarget.push(id);
      const chips = links.map(l => l.ref).filter(Boolean);
      const how = links.length && chips.length === links.length ? (agg === 'Single' || chips.length === 1 ? 'Equals ' : agg === 'Sum' ? 'Sum of ' : 'Average of ') + chips.map(c => `{${c}}`).join(' and ') + '.' : (links[0].how || 'Council-wide measure.');
      kpis[PNAMES.indexOf(pname)].push({ id, name: clean(kg(r, 'KPI')), unit: links[0].unit, target, links, combine, def: clean(kg(r, 'Definition (shown on hover)')), how, src: links[0].src, from: chips, inits: [...new Set(chips.map(c => +c.split(':')[1]))].sort() });
    }
    if (!oldLinks) Object.values(byId).forEach(m => { if (m.feeds && !kpiIds.has(m.feeds)) err(`${files[1]} › Metrics, row ${m.row} (${m.id})`, `“Feeds pillar KPI” is ${m.feeds}, but there is no such KPI on the Pillar KPIs tab.`, 'Pick a KPI ID from the drop-down, or add that KPI on the Pillar KPIs tab first.') });
    if (kNoTarget.length) warn(`${kNoTarget.length} pillar KPI(s) have no target, so they are left out of the pillar score: ${kNoTarget.join(', ')}.`);
    if (errors.length) return { errors, warnings };

    const kpiValue = (k, upto) => k.combine(k.links.map(l => { const s = l.series.filter(([d]) => !upto || ord(d) <= ord(upto)); return s.length ? s[s.length - 1][1] : (upto ? null : l.actual) }));
    const pillarScore = (p, upto) => { const att = kpis[p].filter(k => k.target && kpiValue(k, upto) != null).map(k => Math.min(kpiValue(k, upto) / k.target, 1) * 100); return att.length ? Math.round(att.reduce((a, b) => a + b, 0) / att.length) : null };

    // ---------- overview measures ----------
    const all = Object.values(byId), feeds = name => all.filter(m => m.overview === name);
    const students = inits.map(() => 0), sat = inits.map(() => [0, 0]);
    feeds('Students reached').forEach(m => { if (m.ref && m.actual != null) students[+m.ref.split(':')[1]] += tidy(m.actual) });
    feeds('Average satisfaction').forEach(m => { if (m.ref && m.actual != null && m.n) { const i = +m.ref.split(':')[1]; sat[i][0] += m.actual * m.n; sat[i][1] += m.n } });
    const emp = feeds('Employees reached')[0];

    // ---------- timeline frame ----------
    const dates = raw.flatMap(t => [t.start, t.due, t.done, t.created]).concat(inits.flatMap(i => [i.start, i.end]), [asOf]).filter(Boolean);
    const first = dates.reduce((a, b) => ord(a) < ord(b) ? a : b), base = { y: first.y, m: first.m, d: 1 };
    const last = fromOrd(Math.max(...dates.map(ord)) + 7), day = d => d ? ord(d) - ord(base) + 1 : null;
    const mname = (y, m) => new Date(Date.UTC(y, m - 1, 1)).toLocaleDateString('en-GB', { month: 'long', timeZone: 'UTC' });
    const rangeShort = base.y === last.y && base.m === last.m ? mname(base.y, base.m) : `${mname(base.y, base.m)} – ${mname(last.y, last.m)}`;

    const data = {
      asOf: iso(asOf), sample: false, headcount, people, base: iso(base), today: day(asOf), span: Math.max(31, ord(last) - ord(base) + 1),
      rangeLabel: `${rangeShort} ${last.y}`, rangeShort,
      pillars: PILLARS.map((pm, p) => { const hist = monthCols.slice(0, -1).map(([, d]) => pillarScore(p, d)).filter(x => x != null); return { name: pm.name, color: pm.color, icon: pm.icon, prior: hist.length ? hist[hist.length - 1] : null, history: hist, kpis: kpis[p].map(k => [k.name, tidy(kpiValue(k)), tidy(k.target), UNITS[k.unit] || '', k.inits]) } }),
      initiatives: inits.map((it, i) => { const ts = raw.filter(t => t.i === i), st = ts.map(t => t.start || t.created).filter(Boolean); return [it.name, it.pillar, it.focal, iso(it.start), iso(it.end), st.length ? iso(st.reduce((a, b) => ord(a) < ord(b) ? a : b)) : null, STATUS[it.status.toLowerCase()] ?? null, metrics[i].map(m => [m.level, m.name, tidy(m.actual), tidy(m.target), UNITS[m.unit] || '', m.note || null, tidy(m.n)])] }),
      reach: inits.map(() => []), survey: sat.map(([a, n]) => n ? [Math.round(a / n * 100) / 100, n] : [0, 0]), students, employees: emp && emp.actual != null ? tidy(emp.actual) : null,
      tasks: raw.map(t => [t.i, t.title, t.who, day(t.start), day(t.due), day(t.done)].concat(t.check ? [t.check] : [])),
      info: {
        pillars: PILLARS.map(p => ({ def: p.def })),
        kpis: kpis.map(row => row.map(k => ({ def: k.def, how: k.how, src: k.src || 'See linked metric', from: k.from }))),
        metrics: metrics.map(row => row.map(m => ({ def: m.def, how: m.how, src: m.src }))),
        stats: [
          { name: 'Employees reached', def: 'Distinct SAP KSA employees who took part in at least one council initiative.', how: (byId['CW-02'] ? byId['CW-02'].how : 'Unique employees across every initiative’s participant list.') + ' Entered once a month by the dashboard owner.', src: emp ? emp.src : 'Participant lists' },
          { name: 'Students reached', def: 'Students engaged through council workshops.', how: feeds('Students reached').filter(m => m.ref).map(m => `Equals {${m.ref}}.`)[0] || 'Students who attended council workshops.', src: (feeds('Students reached')[0] || {}).src || 'Workshop attendance' },
          { name: 'Average satisfaction', def: 'Average participant rating across council surveys.', how: 'Satisfaction ratings weighted by their number of responses: ' + (feeds('Average satisfaction').filter(m => m.ref).map(m => `{${m.ref}}`).join(', ') || 'no satisfaction metrics yet') + '.', src: 'Feedback surveys (Microsoft Forms)' },
          { name: 'Initiatives on track', def: 'Share of initiatives whose focal points report them on track.', how: 'Initiatives with goal status On track ÷ initiatives shown.', src: 'Microsoft Planner goal status' }],
        tips: TIPS }
    };
    const done = raw.filter(t => t.done).length;
    const summary = { asOf: fmt(asOf), initiatives: inits.length, tasks: raw.length, done, open: raw.length - done, metrics: all.length, withActual: all.filter(m => m.actual != null).length, kpis: kpis.flat().length, people: people.length, scores: PILLARS.map((p, i) => [p.name, pillarScore(i)]) };
    return { errors, warnings, data, summary };
  }
  window.CouncilConvert = { convert };
})();
