(() => {
  if (document.getElementById('hi-toggle')) return;

  // ── Field definitions ──────────────────────────────────────────────────────
  const FIELDS = [
    { key: 'so_number',         label: 'SO #',                 formLabel: 'SO #',                  type: 'text'     },
    { key: 'job_name',          label: 'Job Name',             formLabel: 'Job Name',              type: 'text'     },
    { key: 'customer_name',     label: 'Customer Name',        formLabel: 'Customer Name',         type: 'dropdown' },
    { key: 'color',             label: 'Color',                formLabel: 'Color',                 type: 'text'     },
    { key: 'series',            label: 'Series',               formLabel: 'Series',                type: 'dropdown' },
    { key: 'edging',            label: 'Edging',               formLabel: 'Edging',                type: 'radio'    },
    { key: 'net_usd_value',     label: 'Net USD Value',        formLabel: 'Net USD Value',         type: 'text'     },
    { key: 'customer_req_date', label: 'Customer Request Date',formLabel: 'Customer Request Date', type: 'text'     },
    { key: 'special_materials', label: 'Special Materials?',   formLabel: 'Special Materials?',    type: 'checkbox' },
  ];

  const SKIP_FIELDS = [
    { key: 'unit_stalls',   label: 'Unit Stalls'            },
    { key: 'units_urinal',  label: 'Units Urinal'           },
    { key: 'floor_gap',     label: 'Floor Gap'              },
    { key: 'handover_date', label: 'Handover Document Date' },
  ];

  // ── Inject panel HTML ──────────────────────────────────────────────────────
  const toggle = document.createElement('button');
  toggle.id = 'hi-toggle';
  toggle.title = 'Handover Importer';
  toggle.innerHTML = '📋';
  document.body.appendChild(toggle);

  const panel = document.createElement('div');
  panel.id = 'hi-panel';
  panel.innerHTML = `
    <div class="hi-header">
      <div class="hi-header-dot"></div>
      <div>
        <div class="hi-header-title">Handover Importer</div>
        <div class="hi-header-sub">Drop Excel \u2192 fields fill automatically</div>
      </div>
    </div>
    <div class="hi-body">
      <div class="hi-drop" id="hi-drop">
        <input type="file" accept=".xlsx,.xls" id="hi-file" />
        <div class="hi-drop-icon">📄</div>
        <div class="hi-drop-hint"><b>Click to browse</b> or drag &amp; drop<br>your handover .xlsx file</div>
      </div>
      <div class="hi-pill" id="hi-pill">
        <div class="hi-pill-icon">📗</div>
        <div>
          <div class="hi-pill-name" id="hi-pill-name"></div>
          <div class="hi-pill-sub" id="hi-pill-sub"></div>
        </div>
      </div>
      <div class="hi-fields" id="hi-fields" style="display:none"></div>
      <div class="hi-status" id="hi-status"></div>
      <div class="hi-success" id="hi-success">
        <div class="hi-success-icon">✅</div>
        <div class="hi-success-msg">Fields filled!</div>
        <div class="hi-success-sub">Review the form above<br>and click <b>Submit</b> when ready.</div>
      </div>
      <button class="hi-btn" id="hi-btn">⚡ Fill Form Fields</button>
      <button class="hi-reset" id="hi-reset">↩ Start over</button>
    </div>`;
  document.body.appendChild(panel);

  document.getElementById('hi-btn').addEventListener('click', hiFill);
  document.getElementById('hi-reset').addEventListener('click', hiReset);

  // ── Toggle panel ───────────────────────────────────────────────────────────
  toggle.addEventListener('click', () => {
    const open = panel.classList.toggle('visible');
    toggle.classList.toggle('open', open);
    toggle.innerHTML = open ? '✕' : '📋';
  });

  // ── Drag & drop ────────────────────────────────────────────────────────────
  const dz = document.getElementById('hi-drop');
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('over');
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  });
  document.getElementById('hi-file').addEventListener('change', e => {
    if (e.target.files[0]) processFile(e.target.files[0]);
  });

  // ── State ──────────────────────────────────────────────────────────────────
  let parsedData = null;

  // ── Parse Excel ────────────────────────────────────────────────────────────
  function processFile(file) {
    setStatus('Parsing…');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        parsedData = parseHandover(ws);
        const count = Object.values(parsedData).filter(v => v).length;
        document.getElementById('hi-pill-name').textContent = file.name;
        document.getElementById('hi-pill-sub').textContent = `${count} of ${FIELDS.length} fields extracted`;
        document.getElementById('hi-pill').classList.add('show');
        dz.style.display = 'none';
        renderFields(parsedData);
        setStatus('');
        document.getElementById('hi-btn').classList.add('show');
        document.getElementById('hi-reset').classList.add('show');
      } catch (err) {
        setStatus('⚠ Parse error: ' + err.message, 'err');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function parseHandover(ws) {
    const g = (addr) => {
      const c = ws[addr];
      if (!c) return null;
      if (c.t === 'd' || c.v instanceof Date) return c.v;
      return c.v;
    };

    const soNum = String(g('B2') || '').replace(/^SO\s*/i, '').trim();
    const jobName = String(g('B3') || '').trim();
    const addrRaw = String(g('B35') || g('B36') || '');
    const customerName = addrRaw.split(',')[0].trim();

    const colorRaw = String(g('B34') || '').replace(/^Formica\s*/i, '').trim();
    const codeMatch = colorRaw.match(/(\d{4}-\d{2,3})\s+(.+)/);
    const color = codeMatch ? `${codeMatch[1]} ${codeMatch[2].trim()}` : colorRaw;

    const rangeRaw = String(g('B22') || '');
    const series = /Flow HPL/i.test(rangeRaw) ? 'Flow HPL' : rangeRaw.split(' ').slice(0, 2).join(' ');
    const edging = /HPL/i.test(rangeRaw) ? 'HPL' : '';

    const netUSD = ((parseFloat(g('F22')) || 0) + (parseFloat(g('F23')) || 0) + (parseFloat(g('F24')) || 0)).toFixed(2);

    let reqDate = g('C32');
    if (reqDate instanceof Date) {
      reqDate = `${String(reqDate.getMonth()+1).padStart(2,'0')}/${String(reqDate.getDate()).padStart(2,'0')}/${reqDate.getFullYear()}`;
    } else {
      reqDate = reqDate ? String(reqDate) : '';
    }

    const specialMat = String(g('B40') || '').trim();

    return { so_number: soNum, job_name: jobName, customer_name: customerName, color, series, edging, net_usd_value: netUSD, customer_req_date: reqDate, special_materials: specialMat };
  }

  // ── Render preview list ────────────────────────────────────────────────────
  function renderFields(data) {
    const container = document.getElementById('hi-fields');
    container.style.display = 'block';
    container.innerHTML = '';
    for (const f of FIELDS) {
      const val = data[f.key];
      const row = document.createElement('div');
      row.className = 'hi-field';
      row.innerHTML = `
        <div class="hi-field-status ${val ? 'ok' : 'skip'}">${val ? '✓' : '—'}</div>
        <div class="hi-field-name">${f.label}</div>
        <div class="hi-field-val ${val ? '' : 'dim'}">${val || 'not found'}</div>`;
      container.appendChild(row);
    }
    const div = document.createElement('div');
    div.style.cssText = 'height:1px;background:#1e2026;margin:4px 0;';
    container.appendChild(div);
    for (const f of SKIP_FIELDS) {
      const row = document.createElement('div');
      row.className = 'hi-field';
      row.innerHTML = `
        <div class="hi-field-status skip">–</div>
        <div class="hi-field-name">${f.label}</div>
        <div class="hi-field-val dim">fill manually</div>`;
      container.appendChild(row);
    }
  }

  // ── Fill all fields ────────────────────────────────────────────────────────
  async function hiFill() {
    if (!parsedData) return;
    const btn = document.getElementById('hi-btn');
    btn.disabled = true;
    btn.textContent = 'Filling…';

    let filled = 0;
    const missed = [];

    for (const f of FIELDS) {
      const val = parsedData[f.key];
      if (!val) continue;
      let ok = false;
      if      (f.type === 'radio')    ok = fillRadio(f.formLabel, val);
      else if (f.type === 'checkbox') ok = fillCheckbox(f.formLabel, val);
      else if (f.type === 'dropdown') ok = await fillDropdown(f.formLabel, val);
      else                            ok = fillText(f.formLabel, val);
      if (ok) filled++; else missed.push(f.label);
      await sleep(150);
    }

    btn.classList.remove('show');
    document.getElementById('hi-success').classList.add('show');
    setStatus(
      missed.length ? `⚠ Not matched: ${missed.join(', ')}` : `✓ ${filled} fields filled`,
      missed.length ? 'err' : 'ok'
    );
  }

  // ── Helpers ────────────────────────────────────────────────────────────────

  function findContainer(labelText) {
    const t = normalise(labelText);
    for (const c of document.querySelectorAll('.css-a718ui')) {
      if (normalise(c.innerText || '').startsWith(t)) return c;
    }
    return null;
  }

  function fillText(labelText, value) {
    const c = findContainer(labelText);
    if (!c) return false;
    const input = c.querySelector('input[type=text], input[type=number], textarea');
    if (!input) return false;
    setReactValue(input, value);
    return true;
  }

  function fillRadio(labelText, value) {
    const c = findContainer(labelText);
    if (!c) return false;
    const radios = c.querySelectorAll('input[type=radio]');
    for (const r of radios) {
      if (normalise(r.value) === normalise(value)) { r.click(); return true; }
    }
    for (const r of radios) {
      if (normalise(r.value).includes(normalise(value)) || normalise(value).includes(normalise(r.value))) {
        r.click(); return true;
      }
    }
    return false;
  }

  function fillCheckbox(labelText, value) {
    const c = findContainer(labelText);
    if (!c) return false;
    const boxes = c.querySelectorAll('input[type=checkbox]');
    const words = normalise(value).split(/\s+/).filter(w => w.length > 2);
    let any = false;
    for (const cb of boxes) {
      const lbl = normalise(cb.closest('label')?.innerText || cb.parentElement?.innerText || '');
      if (words.some(w => lbl.includes(w))) {
        if (!cb.checked) cb.click();
        any = true;
      }
    }
    return any;
  }

  // Smartsheet lodestar combobox:
  // Click [role="combobox"] to open, wait, then click the matching [role="option"]
  async function fillDropdown(labelText, value) {
    const c = findContainer(labelText);
    if (!c) return false;

    const combobox = c.querySelector('[role="combobox"]');
    if (!combobox) return false;

    const menuId = combobox.getAttribute('aria-controls');

    // Open it
    combobox.click();
    await sleep(450);

    // Find options — prefer the linked listbox, fall back to whole document
    const menu = menuId ? document.getElementById(menuId) : null;
    const options = (menu || document).querySelectorAll('[role="option"]');

    for (const opt of options) {
      const optText = normalise(opt.innerText || opt.textContent || '');
      if (optText.includes(normalise(value)) || normalise(value).includes(optText)) {
        opt.click();
        return true;
      }
    }

    // Close if no match
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    return false;
  }

  // Bypass React's synthetic event system by using the native prototype setter
  function setReactValue(el, value) {
    const proto = el.tagName === 'TEXTAREA'
      ? window.HTMLTextAreaElement.prototype
      : window.HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    if (setter) setter.call(el, value); else el.value = value;
    el.dispatchEvent(new Event('input',  { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function normalise(str) {
    return String(str).trim().replace(/\s+/g, ' ').replace(/[*:]/g, '').toLowerCase();
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  // ── Reset ──────────────────────────────────────────────────────────────────
  function hiReset() {
    parsedData = null;
    dz.style.display = 'block';
    document.getElementById('hi-pill').classList.remove('show');
    document.getElementById('hi-fields').style.display = 'none';
    document.getElementById('hi-fields').innerHTML = '';
    document.getElementById('hi-success').classList.remove('show');
    const btn = document.getElementById('hi-btn');
    btn.classList.remove('show');
    btn.disabled = false;
    btn.textContent = '⚡ Fill Form Fields';
    document.getElementById('hi-reset').classList.remove('show');
    document.getElementById('hi-file').value = '';
    setStatus('');
  }

  function setStatus(msg, type = '') {
    const el = document.getElementById('hi-status');
    el.textContent = msg;
    el.className = 'hi-status' + (type ? ' ' + type : '');
  }

})();
