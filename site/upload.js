// Upload Data page: download the template, upload the two workbooks, check them, publish to the live dashboard.
(function () {
  const files = { planner: null, council: null };
  let result = null;
  const h = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const when = s => { try { return new Date(s).toLocaleString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) } catch (e) { return s } };

  function loadScript(src) {
    return new Promise((ok, fail) => { if (document.querySelector(`script[src="${src}"]`)) return ok(); const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => fail(new Error('Could not load ' + src)); document.head.appendChild(s) });
  }

  function statusCard() {
    const m = window.LIVE_META;
    if (m) return `<div class="upstat ok"><b>Live data is published.</b> Last upload ${h(when(m.uploadedAt))}${m.uploadedBy ? ' by ' + h(m.uploadedBy) : ''}${m.files && m.files.length ? ' · ' + m.files.map(h).join(' + ') : ''}. Snapshot ${h(window.DATA_LIVE ? isoDate(window.DATA_LIVE.asOf) + ' ' + window.DATA_LIVE.asOf.slice(0, 4) : '')}.</div>`;
    if (window.DATA_LIVE) return `<div class="upstat ok"><b>Live data is loaded from this site’s files</b> (snapshot ${h(isoDate(window.DATA_LIVE.asOf) + ' ' + window.DATA_LIVE.asOf.slice(0, 4))}).</div>`;
    return '<div class="upstat"><b>No live data has been published yet.</b> The dashboard is showing sample data. Upload both files below to publish live data for everyone.</div>';
  }

  function slot(id, n, title, desc, accept) {
    const f = files[id];
    return `<div class="panel step"><div class="stephead"><span class="stepn">${n}</span><h3>${title}</h3></div><p class="sub">${desc}</p>
      <label class="drop${f ? ' has' : ''}" for="up-${id}" data-slot="${id}"><input id="up-${id}" type="file" accept="${accept}" onchange="UploadPage.pick('${id}',this.files[0])">
      <span class="dropicon" aria-hidden="true">${f ? '✓' : '↑'}</span><span>${f ? '<b>' + h(f.name) + '</b><small>' + Math.round(f.size / 1024) + ' KB · click to replace</small>' : '<b>Choose file</b><small>or drag it here · .xlsx</small>'}</span></label></div>`;
  }

  function resultPanel() {
    if (!result) return '';
    if (result.busy) return '<div class="panel upres"><p>Checking the files…</p></div>';
    if (result.fatal) return `<div class="panel upres bad"><h3>The files could not be read</h3><p>${h(result.fatal)}</p><p class="sub">Make sure both are Excel (.xlsx) files: the Planner export as downloaded, and the Council Data workbook from the template.</p></div>`;
    const w = result.warnings.length ? `<details class="upwarn"><summary>${result.warnings.length} data-health note${result.warnings.length === 1 ? '' : 's'} (won’t block the upload)</summary><ul>${result.warnings.map(x => '<li>' + h(x) + '</li>').join('')}</ul></details>` : '';
    if (result.errors.length) return `<div class="panel upres bad" role="alert"><h3>Upload can’t be accepted: ${result.errors.length} issue${result.errors.length === 1 ? '' : 's'} to fix</h3><p class="sub">Nothing has been published. Fix these in the files and upload them again.</p>
      <ol class="uperrs">${result.errors.map(e => `<li><span class="upwhere">${h(e.where)}</span><b>${h(e.problem)}</b><span class="upfix">How to fix: ${h(e.fix)}</span></li>`).join('')}</ol>${w}</div>`;
    const s = result.summary;
    return `<div class="panel upres good"><h3>Checks passed: ready to publish</h3>
      <div class="upsum"><div><small>Snapshot</small><b>${h(s.asOf)}</b></div><div><small>Tasks</small><b>${s.tasks}</b><small>${s.done} done · ${s.open} open</small></div><div><small>Metrics with a reading</small><b>${s.withActual} / ${s.metrics}</b></div>${s.scores.map(([n, v]) => `<div><small>${h(n)}</small><b>${v == null ? '—' : v + '%'}</b></div>`).join('')}</div>${w}
      <div class="uppub"><label>Upload passcode<input id="up-pass" type="password" autocomplete="off"></label><label>Your name (optional)<input id="up-name" type="text" maxlength="80"></label>
      <button class="control primary" id="up-go" onclick="UploadPage.publish()">Publish to live dashboard</button></div><p id="up-msg" class="sub" role="status"></p></div>`;
  }

  function page() {
    return `<p class="sub">Publish new live data for everyone who opens the dashboard. Files are checked in your browser first; nothing is published unless every check passes.</p>
      ${statusCard()}
      <div class="upgrid">
        <div class="panel step"><div class="stephead"><span class="stepn">1</span><h3>Get the data template</h3></div><p class="sub">The Council Data workbook: metrics, targets and monthly results. Each tab explains how to fill it.</p><p><a class="control primary dl" href="Council_Data_Template.xlsx" download>Download template (.xlsx)</a></p><p class="sub">Already filling it in? Keep using your copy.</p></div>
        ${slot('planner', 2, 'Planner export', 'In Planner: … › Export plan to Excel. Upload the file as downloaded.', '.xlsx')}
        ${slot('council', 3, 'Council Data workbook', 'The filled-in template for this month.', '.xlsx')}
      </div><div id="up-result">${resultPanel()}</div>`;
  }

  async function check() {
    if (!files.planner || !files.council) { result = null; return redraw() }
    result = { busy: true }; redraw();
    try {
      await loadScript('vendor/xlsx.full.min.js').catch(() => loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js')); await loadScript('convert.js');
      const read = async f => XLSX.read(await f.arrayBuffer(), { cellDates: true, cellNF: true });
      let pb, cb;
      try { pb = await read(files.planner) } catch (e) { throw new Error(`“${files.planner.name}” is not a readable Excel file.`) }
      try { cb = await read(files.council) } catch (e) { throw new Error(`“${files.council.name}” is not a readable Excel file.`) }
      result = CouncilConvert.convert(pb, cb, [files.planner.name, files.council.name]);
    } catch (e) { result = { fatal: e.message || String(e), errors: [], warnings: [] } }
    redraw();
  }

  function redraw() {
    const el = document.getElementById('page'); if (!el || !el.querySelector('.upgrid')) return;
    const grid = el.querySelector('.upgrid');
    grid.children[1].outerHTML = slot('planner', 2, 'Planner export', 'In Planner: … › Export plan to Excel. Upload the file as downloaded.', '.xlsx');
    grid.children[2].outerHTML = slot('council', 3, 'Council Data workbook', 'The filled-in template for this month.', '.xlsx');
    document.getElementById('up-result').innerHTML = resultPanel(); wireDrops();
  }

  function wireDrops() {
    document.querySelectorAll('.drop').forEach(d => {
      d.ondragover = e => { e.preventDefault(); d.classList.add('over') };
      d.ondragleave = () => d.classList.remove('over');
      d.ondrop = e => { e.preventDefault(); d.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) pick(d.dataset.slot, f) };
    });
  }

  function pick(id, f) { if (!f) return; files[id] = f; check() }

  async function publish() {
    const msg = document.getElementById('up-msg'), btn = document.getElementById('up-go');
    const pass = document.getElementById('up-pass').value;
    if (!pass) { msg.textContent = 'Enter the upload passcode.'; return }
    btn.disabled = true; msg.textContent = 'Publishing…';
    try {
      const res = await fetch('/api/upload', { method: 'POST', headers: { 'content-type': 'application/json', 'x-upload-passcode': pass },
        body: JSON.stringify({ data: result.data, meta: { files: [files.planner.name, files.council.name], uploadedBy: document.getElementById('up-name').value.trim(), warnings: result.warnings.length } }) });
      let body = {}; try { body = await res.json() } catch (e) { }
      if (res.status === 404 || res.status === 405) throw new Error('Publishing works only on the hosted (Netlify) dashboard, not in this preview.');
      if (res.status === 429) throw new Error('Too many attempts. Wait a minute and try again.');
      if (!res.ok) throw new Error((body.errors || ['The server rejected the upload.']).join(' '));
      window.DATA_LIVE = result.data; window.LIVE_META = { uploadedAt: body.uploadedAt, files: [files.planner.name, files.council.name], uploadedBy: document.getElementById('up-name').value.trim() };
      try { localStorage.setItem('ycds', 'live') } catch (e) { }
      window.name = 'ycds:live';
      msg.innerHTML = '<b>Published.</b> Everyone now sees this data under Live data. <button class="control" onclick="location.hash=\'overview\';location.reload()">View live dashboard</button>';
    } catch (e) { msg.textContent = e.message || String(e); btn.disabled = false }
  }

  window.UploadPage = { page, init: wireDrops, pick, publish };
})();
