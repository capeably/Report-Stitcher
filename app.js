'use strict';

/* ============================================================
   0. ACTIVE STITCH CONFIG  (config-driven engine bootstrap)
   ============================================================
   The active StitchConfig (Emory-specific schema, xlsx builders, dashboard
   layout, etc.) lives in configs/<id>.js. configs/index.js's resolver picks
   one based on the localStorage client code + selected config id. If neither
   is set, the resolver returns null — app.js then shows the client-code gate
   dialog instead of running the normal init.

   CONFIG is captured as `let` rather than `const` to allow future live-switch
   between configs without a page reload; Phase 2 still reloads on switch for
   simplicity but the engine no longer depends on CONFIG being immutable.
*/

let CONFIG = (typeof window !== 'undefined' && typeof window.getActiveStitchConfig === 'function')
  ? window.getActiveStitchConfig()
  : null;

/* ============================================================
   1. CONSTANTS
   ============================================================ */

const NAVY_ARGB        = 'FF1F3864';
const WHITE_ARGB       = 'FFFFFFFF';
const LIGHT_BLUE_ARGB  = 'FFD9E1F2';
const GRAY_FILL_ARGB   = 'FFE7E6E6';
const RED_TITLE_ARGB   = 'FFC00000';

const NAVY_HEX  = '#1F3864';
const SOFT_BLUE = '#8FAADC';   // for "Enrolled" series in stacked charts (lighter than navy)
const NAVY_DARK = '#15264a';

// Column-picker preferences are namespaced per config id so choices don't
// leak across stitches. Returns null if CONFIG hasn't resolved yet — callers
// short-circuit in that case (the column picker is hidden until stitch runs).
function getColumnsStorageKey() {
  return CONFIG ? `stitcher.columns.${CONFIG.id}.v1` : null;
}

/* ============================================================
   2. STATE
   ============================================================ */

const STATE = {
  primary:   { rows: null, fileName: null },
  secondary: { rows: null, fileName: null },
  stitched:  null,    // [{ primary, secondary, method, score, subtypeBucket }]
  unmatched: null,    // [secondary rows]
  // Candidate matches rejected by CONFIG.matchValidator — kept around so the
  // config can surface them on a separate sheet for investigation rather than
  // silently dropping them.
  rejectedMatches: null, // [{ primary, secondary, method, reason }]
  methodCounts: null, // keyed by CONFIG.matchStrategy[i].method
  testRemovedCount: 0,
  // The actual rows CONFIG.testRowFilter pulled out — surfaced under Match
  // Summary so the user can validate the filter's calls.
  testRemovedRows: { primary: [], secondary: [] },
  columns: null,      // [ { ...DEFAULT col, enabled: bool } ]
};

/* ============================================================
   3. MATCH-KEY NORMALIZATION
   ============================================================ */
// Engine-side normalizers shared across configs. matchStrategy.normalize names
// one of these ('email' | 'phone' | 'trim') or supplies an inline function.

function normEmail(s) {
  return (s || '').toString().trim().toLowerCase();
}

function normPhone(s) {
  if (s == null) return '';
  const digits = String(s).replace(/\D/g, '');
  if (digits.length === 11 && digits.startsWith('1')) return digits.slice(1);
  return digits;
}

/* ============================================================
   4. CSV LAYER
   ============================================================ */

async function readCsv(file) {
  // Auto-detect encoding: UTF-8 (with or without BOM) is preferred; otherwise
  // assume Salesforce default (Windows-1252). Strict UTF-8 decode throws on
  // invalid sequences, which is how cp1252 high-bytes get caught.
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);

  let encoding;
  if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
    encoding = 'utf-8';
  } else {
    try {
      new TextDecoder('utf-8', { fatal: true }).decode(buf);
      encoding = 'utf-8';
    } catch (e) {
      encoding = 'windows-1252';
    }
  }

  const text = new TextDecoder(encoding).decode(buf);

  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true,
      skipEmptyLines: 'greedy',
      dynamicTyping: false,
      transformHeader: h => h.trim(),
      complete: (res) => {
        if (res.errors && res.errors.length) {
          const fatal = res.errors.find(e => e.type === 'Quotes' || e.type === 'Delimiter');
          if (fatal) console.warn('CSV parse warnings:', res.errors);
        }
        resolve(res.data);
      },
      error: (err) => reject(err),
    });
  });
}

function validateHeaders(rows, required, label) {
  if (!rows || rows.length === 0) {
    throw new Error(`${label} appears to be empty.`);
  }
  const headers = Object.keys(rows[0]);
  const missing = required.filter(r => !headers.includes(r));
  if (missing.length) {
    throw new Error(`${label} is missing required columns: ${missing.join(', ')}`);
  }
}

/* ============================================================
   CACHE LAYER (IndexedDB) — persist uploaded CSVs across sessions
   ============================================================
   Stores parsed CSV rows so users can come back to view the dashboard
   without re-uploading. Cache writes happen on every successful upload;
   reads happen once at init() to restore last-known-good state. The
   restore path re-runs validateHeaders + stitch, so a code change that
   alters schema requirements falls back gracefully (cache is wiped
   instead of producing stale results). All data stays on-device. */

const CACHE_DB = 'stitcher-cache';
const CACHE_STORE = 'csvs';
const CACHE_VERSION = 1;

function cacheOpenDb() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) return reject(new Error('IndexedDB unavailable'));
    const req = indexedDB.open(CACHE_DB, CACHE_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(CACHE_STORE)) {
        db.createObjectStore(CACHE_STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror   = () => reject(req.error);
  });
}

// Cache keys are namespaced as `{clientCode}:{configId}:{target}` so each
// (client, config) pair keeps its own restore state. Records created before
// Phase 2 (with bare 'cm'/'pa'/'primary'/'secondary' keys) become orphaned —
// they're harmless and get cleared on the next Reset.
function cacheKey(target) {
  const code = window.getClientCode ? window.getClientCode() : '';
  const id   = CONFIG ? CONFIG.id : '';
  return `${code}:${id}:${target}`;
}

async function cachePutCsv(target, fileName, rows) {
  const db = await cacheOpenDb();
  const id = cacheKey(target);
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CACHE_STORE, 'readwrite');
    tx.objectStore(CACHE_STORE).put({ id, target, fileName, rows, savedAt: Date.now() });
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
  });
}

// Returns records for the active (client, config) keyed by target ('primary'
// / 'secondary'). Records from other clients/configs are skipped.
async function cacheLoadAll() {
  const db = await cacheOpenDb();
  const prefix = cacheKey('');   // `${code}:${id}:`
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CACHE_STORE, 'readonly');
    const out = {};
    tx.objectStore(CACHE_STORE).openCursor().onsuccess = (e) => {
      const cur = e.target.result;
      if (cur) {
        const rec = cur.value;
        if (typeof rec.id === 'string' && rec.id.startsWith(prefix) && rec.target) {
          out[rec.target] = rec;
        }
        cur.continue();
      } else {
        resolve(out);
      }
    };
    tx.onerror = () => reject(tx.error);
  });
}

async function cacheDeleteCsv(target) {
  const db = await cacheOpenDb();
  const id = cacheKey(target);
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CACHE_STORE, 'readwrite');
    tx.objectStore(CACHE_STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
  });
}

// Reset clears only the active (client, config) records — other clients'
// caches survive a Reset on this machine.
async function cacheClearAll() {
  const db = await cacheOpenDb();
  const prefix = cacheKey('');
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CACHE_STORE, 'readwrite');
    const store = tx.objectStore(CACHE_STORE);
    store.openCursor().onsuccess = (e) => {
      const cur = e.target.result;
      if (cur) {
        if (typeof cur.value.id === 'string' && cur.value.id.startsWith(prefix)) {
          cur.delete();
        }
        cur.continue();
      }
    };
    tx.oncomplete = () => resolve();
    tx.onerror    = () => reject(tx.error);
  });
}

/* ============================================================
   5. STITCH
   ============================================================ */

// Named normalizers referenced by string from CONFIG.matchStrategy[i].normalize.
// A rule may also supply an inline (row → string) function instead of a name.
const NORMALIZERS = {
  trim:  (s) => (s == null ? '' : String(s).trim()),
  email: normEmail,
  phone: normPhone,
};

function resolveNormalizer(spec) {
  if (typeof spec === 'function') return spec;
  const fn = NORMALIZERS[spec];
  if (!fn) throw new Error(`Unknown matchStrategy normalizer: ${spec}`);
  return fn;
}

// Builds one normalized index per matchStrategy rule, keyed by rule.method.
// Empty/blank keys are skipped so they don't match each other.
function indexPrimary(primaryRows, matchStrategy) {
  const indexes = {};
  for (const rule of matchStrategy) {
    const norm = resolveNormalizer(rule.normalize);
    const idx  = new Map();
    for (const row of primaryRows) {
      const key = norm(row[rule.primaryCol]);
      if (!key) continue;
      if (!idx.has(key)) idx.set(key, []);
      idx.get(key).push(row);
    }
    indexes[rule.method] = idx;
  }
  return indexes;
}

// Tries rules in order; first hit wins. Returns null if no rule matched.
function findMatches(secondaryRow, indexes, matchStrategy) {
  for (const rule of matchStrategy) {
    const norm = resolveNormalizer(rule.normalize);
    const key  = norm(secondaryRow[rule.secondaryCol]);
    if (!key) continue;
    const idx  = indexes[rule.method];
    if (idx && idx.has(key)) {
      return { method: rule.method, candidates: idx.get(key) };
    }
  }
  return null;
}

function stitch(primaryRows, secondaryRows) {
  // Split each side into kept rows + filtered-out test rows. The test rows
  // surface in Match Summary so the user can spot misclassifications.
  const primaryClean = [], primaryTest = [];
  for (const r of primaryRows) (CONFIG.testRowFilter(r) ? primaryTest : primaryClean).push(r);
  const secondaryClean = [], secondaryTest = [];
  for (const r of secondaryRows) (CONFIG.testRowFilter(r) ? secondaryTest : secondaryClean).push(r);
  const testRemovedCount = primaryTest.length + secondaryTest.length;
  const testRemovedRows  = { primary: primaryTest, secondary: secondaryTest };

  const matchStrategy = CONFIG.matchStrategy;
  const primaryIdx    = indexPrimary(primaryClean, matchStrategy);

  // Optional validator: returns true (or a truthy value) if a candidate
  // (primary, secondary) pair makes logical sense. Reject reasons surface in
  // rejectedMatches (string or true → no reason). When a secondary has at
  // least one valid candidate, that valid set goes to tiebreaker as usual;
  // when ALL its candidates are rejected, the secondary moves to `unmatched`
  // and each rejected candidate is recorded for the config to surface (e.g.
  // a separate xlsx sheet for investigation).
  const validator = (typeof CONFIG.matchValidator === 'function') ? CONFIG.matchValidator : null;

  const stitched        = [];
  const unmatched       = [];
  const rejectedMatches = [];
  const methodCounts    = {};
  for (const rule of matchStrategy) methodCounts[rule.method] = 0;

  for (const secondary of secondaryClean) {
    const match = findMatches(secondary, primaryIdx, matchStrategy);
    if (!match) { unmatched.push(secondary); continue; }

    let validCandidates = match.candidates;
    let rejected = [];
    if (validator) {
      validCandidates = [];
      for (const cand of match.candidates) {
        const verdict = validator(cand, secondary, { method: match.method });
        if (verdict === true || (verdict && verdict.valid !== false)) {
          validCandidates.push(cand);
        } else {
          rejected.push({
            cand,
            reason: (verdict && typeof verdict === 'object' && verdict.reason) || null,
          });
        }
      }
    }

    if (validCandidates.length > 0) {
      const { row: primary, score } = CONFIG.tiebreaker(validCandidates, secondary);
      methodCounts[match.method]++;
      const matchInfo = { method: match.method, score };
      stitched.push({
        primary, secondary,
        method: match.method,
        score: Math.round(score * 1000) / 1000,
        subtypeBucket: CONFIG.derivedFields.subtype_bucket(primary, secondary, matchInfo),
      });
    } else {
      // No valid candidate — treat secondary as unmatched and log every
      // rejected candidate so the config can surface them.
      unmatched.push(secondary);
      for (const r of rejected) {
        rejectedMatches.push({
          primary:   r.cand,
          secondary,
          method:    match.method,
          reason:    r.reason,
        });
      }
    }
  }

  return { stitched, unmatched, methodCounts, testRemovedCount, testRemovedRows, rejectedMatches };
}

/* ============================================================
   6. CELL VALUE LOOKUP  (unifies primary/secondary/derived sources)
   ============================================================ */

function getCellValue(row, col) {
  const { primary, secondary, method, score, subtypeBucket } = row;
  if (col.source === 'primary') {
    let v = primary[col.sourceField];
    if ((v == null || v === '') && col.fallbackField) v = secondary[col.fallbackField];
    return v == null ? '' : v;
  }
  if (col.source === 'secondary') {
    let v = secondary[col.sourceField];
    if ((v == null || v === '') && col.fallbackField) v = primary[col.fallbackField];
    return v == null ? '' : v;
  }
  if (col.source === 'derived') {
    // Fast paths for fields the engine pre-computes during stitch (cached on
    // the stitched row directly so we don't re-invoke the config function for
    // every cell render):
    if (col.sourceField === 'subtype_bucket') return subtypeBucket;
    if (col.sourceField === 'match_method')   return method;
    if (col.sourceField === 'course_score')   return score;
    // Anything else: invoke the config's derivedField on demand. Returns
    // null/undefined → empty string in the cell.
    const fn = CONFIG && CONFIG.derivedFields && CONFIG.derivedFields[col.sourceField];
    if (typeof fn === 'function') {
      const v = fn(primary, secondary, { method, score });
      return v == null ? '' : v;
    }
  }
  return '';
}

/* ============================================================
   7. COLUMN CONFIG  (load/save/render)
   ============================================================ */

function loadSavedConfig() {
  const key = getColumnsStorageKey();
  if (!key) return null;
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const saved = JSON.parse(raw);
    return Array.isArray(saved) ? saved : null;
  } catch (e) {
    console.warn('Failed to load column config.', e);
    return null;
  }
}

function sanitizeKey(s) {
  return String(s).toLowerCase().replace(/\W+/g, '_').replace(/^_+|_+$/g, '');
}

// Canonical default column list given the actual file headers. The 19 documented
// defaults come first (enabled). Any additional CM/PA header in the uploaded
// files is appended (disabled), so the user can opt-in to columns like
// "Lead Source Details" or "Contact Owner" without touching code.
function buildDefaultColumnList(primaryHeaders, secondaryHeaders) {
  const cols = CONFIG.defaultColumns.map(d => ({ ...d, enabled: true }));
  const coveredPrimary = new Set();
  const coveredSecondary = new Set();
  for (const c of cols) {
    if (c.source === 'primary') coveredPrimary.add(c.sourceField);
    if (c.source === 'secondary') coveredSecondary.add(c.sourceField);
  }
  for (const h of (primaryHeaders || [])) {
    if (!coveredPrimary.has(h)) {
      cols.push({ key: 'primary__' + sanitizeKey(h), label: h, source: 'primary', sourceField: h, enabled: false });
      coveredPrimary.add(h);
    }
  }
  for (const h of (secondaryHeaders || [])) {
    if (!coveredSecondary.has(h)) {
      cols.push({ key: 'secondary__' + sanitizeKey(h), label: h, source: 'secondary', sourceField: h, enabled: false });
      coveredSecondary.add(h);
    }
  }
  return cols;
}

// Build the column list applying any saved user preferences (order, label, enabled).
// New columns the saved config didn't know about are appended at the end.
function buildColumnList(primaryHeaders, secondaryHeaders) {
  const fullDefaults = buildDefaultColumnList(primaryHeaders, secondaryHeaders);
  const saved = loadSavedConfig();
  if (!saved) return fullDefaults;

  const knownByKey = new Map(fullDefaults.map(c => [c.key, c]));
  const reordered = [];
  const seenKeys = new Set();
  for (const item of saved) {
    if (!item || !knownByKey.has(item.key)) continue;
    const def = knownByKey.get(item.key);
    reordered.push({
      ...def,
      label: typeof item.label === 'string' && item.label.trim() ? item.label : def.label,
      enabled: typeof item.enabled === 'boolean' ? item.enabled : def.enabled,
    });
    seenKeys.add(item.key);
  }
  for (const c of fullDefaults) {
    if (!seenKeys.has(c.key)) reordered.push(c);
  }
  return reordered;
}

function saveColumnConfig(cols) {
  const key = getColumnsStorageKey();
  if (!key) return;
  try {
    const data = cols.map(c => ({ key: c.key, label: c.label, enabled: c.enabled }));
    localStorage.setItem(key, JSON.stringify(data));
  } catch (e) {
    console.warn('Failed to save column config.', e);
  }
}

function renderColumnPicker(container, cols, onChange) {
  container.innerHTML = '';
  cols.forEach((col, idx) => {
    const row = document.createElement('div');
    row.className = 'col-row';
    row.dataset.key = col.key;

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = col.enabled;
    cb.addEventListener('change', () => { col.enabled = cb.checked; onChange(); });

    const upBtn = document.createElement('button');
    upBtn.className = 'reorder-btn';
    upBtn.textContent = '↑';
    upBtn.disabled = idx === 0;
    upBtn.title = 'Move up';
    upBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (idx > 0) {
        [cols[idx-1], cols[idx]] = [cols[idx], cols[idx-1]];
        renderColumnPicker(container, cols, onChange);
        onChange();
      }
    });

    const downBtn = document.createElement('button');
    downBtn.className = 'reorder-btn';
    downBtn.textContent = '↓';
    downBtn.disabled = idx === cols.length - 1;
    downBtn.title = 'Move down';
    downBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (idx < cols.length - 1) {
        [cols[idx], cols[idx+1]] = [cols[idx+1], cols[idx]];
        renderColumnPicker(container, cols, onChange);
        onChange();
      }
    });

    const labelInput = document.createElement('input');
    labelInput.type = 'text';
    labelInput.value = col.label;
    labelInput.className = 'rename-input';
    labelInput.addEventListener('input', () => { col.label = labelInput.value; onChange(); });

    const sourceSpan = document.createElement('span');
    sourceSpan.className = 'col-source ' + col.source;
    sourceSpan.textContent = col.source === 'derived' ? 'CALC' : col.source.toUpperCase();

    row.append(cb, upBtn, downBtn, labelInput, sourceSpan);
    container.append(row);
  });
}

/* ============================================================
   8. PREVIEW TABLE + KPIs
   ============================================================ */

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderPreviewTable() {
  const enabled = STATE.columns.filter(c => c.enabled);
  const rows = STATE.stitched;
  const limit = Math.min(100, rows.length);
  const tbl = document.getElementById('preview-table');

  let html = '<thead><tr>';
  for (const col of enabled) html += `<th>${escapeHtml(col.label)}</th>`;
  html += '</tr></thead><tbody>';
  for (let i = 0; i < limit; i++) {
    const r = rows[i];
    html += '<tr>';
    for (const col of enabled) {
      let val = getCellValue(r, col);
      if (col.key === 'course_score') {
        const cls = val === 1 ? 'score-1' : (val === 0 ? 'score-0' : 'score-fuzzy');
        html += `<td class="${cls}">${val.toFixed(3)}</td>`;
      } else {
        html += `<td title="${escapeHtml(val)}">${escapeHtml(val)}</td>`;
      }
    }
    html += '</tr>';
  }
  html += '</tbody>';
  tbl.innerHTML = html;

  document.getElementById('preview-meta').textContent =
    `Showing ${limit.toLocaleString()} of ${rows.length.toLocaleString()} stitched rows · ${enabled.length} columns selected`;
}

function renderKpis() {
  // Primary-input funnel — headline metrics. Scope: primary rows whose date
  // dimension (CM Update Date for Emory) falls in FILTERS.focus. This matches
  // the Campaign Summary xlsx tab's framing — both reflect Focus-period
  // campaign activity. The match-info comes from STATE.stitched, which has
  // already been bounded to PA ∈ Focus × CM ∈ Influence.
  const dateOf = CONFIG.dashboard && CONFIG.dashboard.dateOf;
  const primaryInFocus = STATE.primary.rows.filter(r =>
    !CONFIG.testRowFilter(r) && (!dateOf || dateInRange(dateOf.primary(r), FILTERS.focus))
  );
  const primaryToStitched = new Map();
  for (const s of STATE.stitched) {
    if (!primaryToStitched.has(s.primary)) primaryToStitched.set(s.primary, []);
    primaryToStitched.get(s.primary).push(s);
  }
  let countNc = 0, countCx = 0, countRg = 0, countEn = 0;
  for (const primary of primaryInFocus) {
    const status = CONFIG.deriveCourseStatus(primaryToStitched.get(primary) || []);
    if (status === 'Not Converted')                 countNc++;
    else if (status === 'Cancelled/Withdrawn/Etc')  countCx++;
    else if (status === 'Registered')               countRg++;
    else if (status === 'Enrolled')                 countEn++;
  }
  const converted = countRg + countEn;
  const conversionRate = primaryInFocus.length === 0 ? 0 : (converted / primaryInFocus.length) * 100;

  document.getElementById('kpi-primary-total').textContent        = primaryInFocus.length.toLocaleString();
  document.getElementById('kpi-primary-notconverted').textContent = countNc.toLocaleString();
  document.getElementById('kpi-primary-cancelled').textContent    = countCx.toLocaleString();
  document.getElementById('kpi-primary-registered').textContent   = countRg.toLocaleString();
  document.getElementById('kpi-primary-enrolled').textContent     = countEn.toLocaleString();
  document.getElementById('kpi-conversion-rate').textContent =
    conversionRate < 10 ? conversionRate.toFixed(1) + '%' : Math.round(conversionRate) + '%';

  // Secondary-input matching — secondary metrics. STATE.stitched and
  // STATE.unmatched are already Focus-bounded by runStitch, so these counts
  // reflect PAs ∈ Focus.
  const stitched = STATE.stitched.length;
  const unmatchedSecondary = STATE.unmatched.length;
  document.getElementById('kpi-secondary-total').textContent     = (stitched + unmatchedSecondary).toLocaleString();
  document.getElementById('kpi-stitched').textContent            = stitched.toLocaleString();
  document.getElementById('kpi-secondary-unmatched').textContent = unmatchedSecondary.toLocaleString();

  // Match details — counts in the same order as CONFIG.matchStrategy.
  const m = STATE.methodCounts;
  document.getElementById('match-methods').textContent =
    CONFIG.matchStrategy.map(rule => (m[rule.method] || 0).toLocaleString()).join(' / ');

  let s1 = 0, sFuzzy = 0, s0 = 0;
  for (const r of STATE.stitched) {
    if (r.score === 1) s1++; else if (r.score === 0) s0++; else sFuzzy++;
  }
  document.getElementById('score-dist').textContent   = `${s1} / ${sFuzzy} / ${s0}`;
  document.getElementById('test-removed').textContent = STATE.testRemovedCount.toLocaleString();
  renderTestRowsRemoved();
}

// List the rows CONFIG.testRowFilter pulled out, grouped by source side, so
// the user can validate the filter's calls. Container hides when count is 0;
// each side's column hides when that side contributed nothing.
function renderTestRowsRemoved() {
  const container = document.getElementById('test-rows-removed');
  if (!container) return;

  const removed = STATE.testRemovedRows || { primary: [], secondary: [] };
  if (removed.primary.length + removed.secondary.length === 0) {
    container.hidden = true;
    return;
  }
  container.hidden = false;

  const renderItem = (r) => {
    const first = ((r['First Name'] || '') + '').trim();
    const last  = ((r['Last Name']  || '') + '').trim();
    const email = ((r['Email']      || '') + '').trim();
    const name  = [first, last].filter(Boolean).join(' ') || '(no name)';
    const emailHtml = email ? ` · <span class="test-row-email">${escapeHtml(email)}</span>` : '';
    return `<li><span class="test-row-name">${escapeHtml(name)}</span>${emailHtml}</li>`;
  };

  const fillSide = (kind, rows) => {
    const col      = document.getElementById(`test-rows-${kind}-col`);
    const list     = document.getElementById(`test-rows-${kind}-list`);
    const heading  = col.querySelector('.test-rows-source-heading');
    if (!col || !list || !heading) return;
    if (rows.length === 0) { col.hidden = true; list.innerHTML = ''; return; }
    const inputLabel = (CONFIG.inputs && CONFIG.inputs[kind] && CONFIG.inputs[kind].label) || kind;
    heading.textContent = `${inputLabel} (${rows.length.toLocaleString()})`;
    list.innerHTML = rows.map(renderItem).join('');
    col.hidden = false;
  };

  fillSide('primary',   removed.primary);
  fillSide('secondary', removed.secondary);
}

/* ============================================================
   9. AGGREGATIONS  (used for in-page charts AND the xlsx Summary tables)
   ============================================================ */

function aggregateAll(stitched) {
  // Returns three structures with consistent shape:
  //   subtype:  { buckets: { Website: [{name,reg,enr},...], Social:[...], Unknown:[...] } }
  //   parent:   [{ name, reg, enr }] (sorted desc by total)
  //   course:   [{ name, reg, enr }] (sorted desc by total)
  const subtypeMap = new Map();   // bucket → Map(subtypeName → { reg, enr })
  const parentMap  = new Map();
  const courseMap  = new Map();

  // Counts include only Registered + Enrolled per R1 Feedback.
  for (const r of stitched) {
    const status = r.secondary['Status'];
    const isReg = status === 'Registered';
    const isEnr = status === 'Enrolled';
    if (!isReg && !isEnr) continue;

    const subtype = (r.primary['Sub-Type'] || '').trim() || '(blank)';
    const bkt = r.subtypeBucket;
    if (!subtypeMap.has(bkt)) subtypeMap.set(bkt, new Map());
    const sub = subtypeMap.get(bkt);
    if (!sub.has(subtype)) sub.set(subtype, { reg:0, enr:0 });
    if (isReg) sub.get(subtype).reg++;
    if (isEnr) sub.get(subtype).enr++;

    const parent = (r.primary['Parent Campaign Name'] || '').trim() || '(blank)';
    if (!parentMap.has(parent)) parentMap.set(parent, { reg:0, enr:0 });
    if (isReg) parentMap.get(parent).reg++;
    if (isEnr) parentMap.get(parent).enr++;

    const course = (r.secondary['Course Name'] || '').trim() || '(blank)';
    if (!courseMap.has(course)) courseMap.set(course, { reg:0, enr:0 });
    if (isReg) courseMap.get(course).reg++;
    if (isEnr) courseMap.get(course).enr++;
  }

  // Materialize subtype in locked bucket order, with Unknown last.
  const subtypeBuckets = {};
  const orderedBuckets = [...CONFIG.bucketOrder];
  if (subtypeMap.has('Unknown')) orderedBuckets.push('Unknown');
  for (const bkt of orderedBuckets) {
    if (!subtypeMap.has(bkt)) continue;
    const arr = [];
    for (const [name, c] of subtypeMap.get(bkt)) arr.push({ name, reg:c.reg, enr:c.enr });
    arr.sort((a,b) => (b.reg+b.enr) - (a.reg+a.enr));
    subtypeBuckets[bkt] = arr;
  }

  const sortByTotal = arr => arr.sort((a,b) => (b.reg+b.enr) - (a.reg+a.enr));
  const parentArr = sortByTotal([...parentMap.entries()].map(([name,c]) => ({ name, reg:c.reg, enr:c.enr })));
  const courseArr = sortByTotal([...courseMap.entries()].map(([name,c]) => ({ name, reg:c.reg, enr:c.enr })));

  return { subtypeBuckets, parent: parentArr, course: courseArr };
}

/* ============================================================
   10. CHART RENDER  (in-page + offscreen for PNG embed)
   ============================================================ */

const PAGE_CHARTS = {};   // key → Chart instance
const OFF_CHARTS  = {};   // key → Chart instance for offscreen high-res render

function buildHorizontalStackedConfig(labels, regSeries, enrSeries, opts = {}) {
  const isPng = !!opts.forPng;
  // For PNG render we draw at 2x the embed size, so font sizes scale up to stay
  // crisp when displayed in Excel at the embed dimensions.
  const scale = isPng ? (opts.scale || 2) : 1;
  const fontSize   = (isPng ? 13 : 12) * scale;
  const totalsSize = (isPng ? 14 : 11) * scale;
  const padRight   = (isPng ? 44 : 30) * scale;

  const cfg = {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Registered', data: regSeries, backgroundColor: NAVY_HEX,  borderWidth: 0 },
        { label: 'Enrolled',   data: enrSeries, backgroundColor: SOFT_BLUE, borderWidth: 0 },
      ],
    },
    options: {
      indexAxis: 'y',
      responsive: !isPng,
      maintainAspectRatio: false,
      animation: false,
      devicePixelRatio: isPng ? 1 : (window.devicePixelRatio || 1),
      layout: { padding: { right: padRight, top: 6 * scale, bottom: 6 * scale } },
      scales: {
        x: {
          stacked: true,
          beginAtZero: true,
          ticks: { font: { size: fontSize, family: 'Arial' }, color: '#000', precision: 0 },
          grid:  { color: 'rgba(0,0,0,.12)' },
        },
        y: {
          stacked: true,
          ticks: {
            // autoSkip: false ensures every label renders. With long labels (e.g.
            // courses) Chart.js's default would silently hide labels it deemed
            // crowded — leaving bars with no label. The host card's height calc
            // is responsible for giving each row enough vertical space (see the
            // per-row multipliers in renderInPageCharts).
            autoSkip: false,
            font: { size: fontSize, family: 'Arial' },
            color: '#000',
            // Wrap long category labels across two lines so they don't clip on the
            // canvas's left edge. Returning an array → Chart.js renders multi-line.
            callback: function(value) {
              const label = this.getLabelForValue(value);
              if (!label || label.length <= 34) return label;
              const mid = Math.floor(label.length / 2);
              let breakAt = label.lastIndexOf(' ', mid + 8);
              if (breakAt < mid - 12) breakAt = label.indexOf(' ', mid);
              if (breakAt < 0) return label.length > 60 ? label.slice(0, 57) + '…' : label;
              return [label.slice(0, breakAt), label.slice(breakAt + 1)];
            },
          },
          grid:  { display: false },
        },
      },
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            font: { size: fontSize, family: 'Arial', weight: '600' },
            color: '#000',
            boxWidth: 14 * scale,
            padding: 8 * scale,
          },
        },
        // Tooltip body sits ABOVE the cursor (yAlign: 'bottom' anchors the
        // caret at the body's bottom). caretPadding gives breathing room so
        // the tooltip doesn't hug the cursor.
        tooltip: { enabled: !isPng, yAlign: 'bottom', caretPadding: 8 },
      },
    },
    plugins: [totalsLabelPlugin(totalsSize)],
  };

  // onClick / onHover plumbing for click-to-drill in distribution charts.
  // opts.onClick is called with (categoryIdx, datasetIdx) when a bar segment
  // is clicked. opts.onHover is optional — defaulted to a cursor:pointer flip
  // if onClick is set, so users see the chart's interactivity affordance.
  if (typeof opts.onClick === 'function') {
    cfg.options.onClick = (evt, elements) => {
      if (!elements.length) return;
      const el = elements[0];
      opts.onClick(el.index, el.datasetIndex);
    };
    cfg.options.onHover = (evt, elements) => {
      const target = evt && evt.native && evt.native.target;
      if (target && target.style) target.style.cursor = elements.length ? 'pointer' : 'default';
    };
  }

  return cfg;
}

function totalsLabelPlugin(fontSize) {
  return {
    id: 'totalsLabel',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const meta1 = chart.getDatasetMeta(1);
      if (!meta1) return;
      ctx.save();
      ctx.fillStyle = '#000';
      ctx.font = `bold ${fontSize}px Arial, sans-serif`;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      const data0 = chart.data.datasets[0].data;
      const data1 = chart.data.datasets[1].data;
      meta1.data.forEach((bar, i) => {
        const total = (Number(data0[i]) || 0) + (Number(data1[i]) || 0);
        if (total === 0) return;
        ctx.fillText(String(total), bar.x + Math.max(4, fontSize * 0.3), bar.y);
      });
      ctx.restore();
    },
  };
}

function destroyChart(map, key) {
  if (map[key]) { try { map[key].destroy(); } catch(e) {} delete map[key]; }
}

function renderInPageCharts(agg, filteredPrimarySet) {
  // Status order in the datasets: [Registered, Enrolled]. datasetIdx 0 → Registered, 1 → Enrolled.
  const STATUS_BY_DATASET = ['Registered', 'Enrolled'];
  const onSegmentClick = (dimension, labels) => (idx, datasetIdx) => {
    const value  = labels[idx];
    const status = STATUS_BY_DATASET[datasetIdx];
    if (!value || !status) return;
    openDistributionDrilldown(dimension, value, status, filteredPrimarySet);
  };

  // Sub-type chart: detail rows only (NO bucket parents — matches stitch.py chart source).
  const sLabels = [], sReg = [], sEnr = [];
  for (const bkt of Object.keys(agg.subtypeBuckets)) {
    for (const r of agg.subtypeBuckets[bkt]) {
      sLabels.push(r.name);
      sReg.push(r.reg);
      sEnr.push(r.enr);
    }
  }
  destroyChart(PAGE_CHARTS, 'subtype');
  const subCanvas = document.getElementById('chart-subtype');
  subCanvas.parentElement.style.height = Math.max(220, sLabels.length * 38 + 80) + 'px';
  PAGE_CHARTS.subtype = new Chart(subCanvas, buildHorizontalStackedConfig(sLabels, sReg, sEnr, {
    onClick: onSegmentClick('subType', sLabels),
  }));

  destroyChart(PAGE_CHARTS, 'parent');
  const pLabels = agg.parent.map(r => r.name);
  const pReg = agg.parent.map(r => r.reg);
  const pEnr = agg.parent.map(r => r.enr);
  const pCanvas = document.getElementById('chart-parent');
  pCanvas.parentElement.style.height = Math.max(260, pLabels.length * 30 + 80) + 'px';
  PAGE_CHARTS.parent = new Chart(pCanvas, buildHorizontalStackedConfig(pLabels, pReg, pEnr, {
    onClick: onSegmentClick('parentCampaign', pLabels),
  }));

  destroyChart(PAGE_CHARTS, 'course');
  const cLabels = agg.course.map(r => r.name);
  const cReg = agg.course.map(r => r.reg);
  const cEnr = agg.course.map(r => r.enr);
  const cCanvas = document.getElementById('chart-course');
  // Course names tend to be the longest and often wrap to 2 lines via the tick
  // callback — 44px per row gives wrapped labels enough room (was 28). The
  // sub-type (38) and parent (30) multipliers stay; their labels rarely wrap
  // and autoSkip:false catches the rest.
  cCanvas.parentElement.style.height = Math.max(280, cLabels.length * 44 + 90) + 'px';
  PAGE_CHARTS.course = new Chart(cCanvas, buildHorizontalStackedConfig(cLabels, cReg, cEnr, {
    onClick: onSegmentClick('course', cLabels),
  }));
}

// Open the drilldown for a Distribution chart bar segment. Iterates STATE.stitched
// (NOT the dashboard's primaryDataset) because a single primary row with two
// matched secondary rows of different statuses contributes 1 to two different
// bar segments — so attribution must walk stitched rows directly. The dashboard
// filter is reapplied via filteredPrimarySet (each primary's CM passed the
// date / parent / sub-type / bucket filter).
function openDistributionDrilldown(dimension, dimensionValue, status, filteredPrimarySet) {
  const matchesDimension = (s) => {
    if (dimension === 'course')         return (s.secondary['Course Name']       || '').trim() === dimensionValue;
    if (dimension === 'parentCampaign') return ((s.primary['Parent Campaign Name'] || '').trim() || '(blank)') === dimensionValue;
    if (dimension === 'subType')        return ((s.primary['Sub-Type']             || '').trim() || '(blank)') === dimensionValue;
    return false;
  };

  const matching = STATE.stitched
    .filter(s => filteredPrimarySet.has(s.primary)
              && (s.secondary['Status'] || '').trim() === status
              && matchesDimension(s))
    .map(s => ({
      // Reshape the stitched row into the same fields renderDrilldownPage expects.
      primary:        s.primary,
      secondary:      s.secondary,
      bestMatch:      s,
      courseStatus:   status,
      parentCampaign: (s.primary['Parent Campaign Name'] || '').trim() || '(blank)',
      campaignName:   (s.primary['Campaign Name']        || '').trim() || '(blank)',
      subType:        (s.primary['Sub-Type']             || '').trim() || '(blank)',
      subTypeBucket:  CONFIG.derivedFields.subtype_bucket(s.primary),
      primaryActivity:  s.primary._memberStatusUpdate,
      secondaryCreated: s.secondary._secondaryCreated,
      courseStart:      s.secondary._courseStart,
    }));

  openDrilldown(`${dimensionValue} — ${status} — ${matching.length.toLocaleString()} record${matching.length === 1 ? '' : 's'}`, matching);
}

async function renderOffscreenChartPng(canvasId, labels, reg, enr) {
  // Compute embed dimensions FIRST, then render the canvas at the same aspect ratio
  // (scaled up 2x for crisp display in Excel). Matching aspect ratios prevents
  // Excel from stretching the bitmap when the user resizes columns/rows.
  const embedW = 720;
  const embedH = Math.max(280, labels.length * 28 + 100);
  const scale  = 2;
  const canvas = document.getElementById(canvasId);
  canvas.width  = embedW * scale;
  canvas.height = embedH * scale;
  destroyChart(OFF_CHARTS, canvasId);
  const cfg = buildHorizontalStackedConfig(labels, reg, enr, { forPng: true, scale });
  OFF_CHARTS[canvasId] = new Chart(canvas, cfg);
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
  return { dataUrl: canvas.toDataURL('image/png'), embedW, embedH };
}

/* ============================================================
   11. xlsx — ENGINE PRIMITIVES & CONFIG BRIDGE
   ============================================================
   Style constants, generic xlsx helpers, and the window.RS bridge that the
   active config's outputSheet builders pull their bindings from. Sheet
   builders themselves live in configs/<id>.js. The orchestration that
   iterates CONFIG.outputSheets is in section 12 below. */

const FONT_HEADER = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE_ARGB } };
const FONT_BODY   = { name: 'Arial', size: 10 };
const FILL_NAVY   = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY_ARGB } };
const FILL_LBLUE  = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT_BLUE_ARGB } };
const FILL_GRAY   = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRAY_FILL_ARGB } };
const BORDER_THIN = {
  top:    { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  left:   { style: 'thin', color: { argb: 'FF000000' } },
  right:  { style: 'thin', color: { argb: 'FF000000' } },
};
const BORDER_MED_BOTTOM = { ...BORDER_THIN, bottom: { style: 'medium', color: { argb: 'FF000000' } } };
const BORDER_MED_TOP    = { ...BORDER_THIN, top:    { style: 'medium', color: { argb: 'FF000000' } } };

function colLetter(n) {
  // 1 -> A, 27 -> AA
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function letterToCol(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function applyBorderRange(ws, ref, border) {
  // Apply the same border to every cell in a range string like 'A1:E1'.
  const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return;
  const c1 = letterToCol(m[1]), r1 = +m[2];
  const c2 = letterToCol(m[3]), r2 = +m[4];
  for (let rr = r1; rr <= r2; rr++) {
    for (let cc = c1; cc <= c2; cc++) {
      ws.getCell(`${colLetter(cc)}${rr}`).border = border;
    }
  }
}

function escapeFormula(s) {
  return String(s).replace(/"/g, '""');
}

function findEnabledColIndex(enabled, key) {
  // Returns 1-based column index in the Stitched Data sheet, or null if not enabled.
  const i = enabled.findIndex(c => c.key === key);
  return i === -1 ? null : i + 1;
}

// Bindings the config's outputSheet builders pull from. Function declarations
// here are hoisted, so this object literal is safe to construct mid-file. STATE
// is a live mutable reference — builders always read fresh state.
window.RS = {
  STATE,
  FONT_HEADER, FONT_BODY,
  FILL_NAVY, FILL_LBLUE, FILL_GRAY,
  BORDER_THIN, BORDER_MED_BOTTOM, BORDER_MED_TOP,
  NAVY_ARGB, WHITE_ARGB, LIGHT_BLUE_ARGB, GRAY_FILL_ARGB, RED_TITLE_ARGB,
  colLetter, letterToCol, applyBorderRange, escapeFormula, findEnabledColIndex,
  getCellValue, aggregateAll, dateInRange, parseSfDate,
  renderChartPng: renderOffscreenChartPng,
};

/* ============================================================
   12. xlsx — BUILD ORCHESTRATION
   ============================================================
   Generic loop over CONFIG.outputSheets. Each builder gets (wb, ctx); the
   engine stores ctx.sheets[sheet.name] = builder return value so downstream
   builders can read their handoffs. Builders are async-aware (await each). */

async function generateXlsx(opts = {}) {
  if (!CONFIG || !CONFIG.outputSheets || !CONFIG.outputSheets.length) {
    throw new Error('Active config has no outputSheets defined.');
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Report Stitcher';
  wb.created = new Date();

  // Snapshot active Focus + Influence windows for the builders. We pass the
  // refs (mutable) rather than copying so a builder can apply tightening
  // logic without re-reading globals. dateOf hooks travel along so builders
  // don't need to know the row-cache field names.
  const ctx = {
    sheets:    {},
    focus:     FILTERS.focus,
    influence: FILTERS.influence,
    dateOf:    CONFIG.dashboard && CONFIG.dashboard.dateOf,
  };
  for (const sheet of CONFIG.outputSheets) {
    const result = await sheet.builder(wb, ctx);
    if (result) ctx.sheets[sheet.name] = result;
  }

  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const stamp  = new Date().toISOString().slice(0, 10);
  const suffix = opts.filenameSuffix || '';
  saveAs(blob, `Stitched_Report${suffix}_${stamp}.xlsx`);
}

// Filter-bar variant: generate an xlsx whose builders see only the dashboard-
// filtered cohort. Implemented as a STATE-swap so the existing builders work
// unchanged: they read STATE.stitched + STATE.primary.rows, we temporarily
// replace those with the filtered subsets and restore in `finally`. The
// generic engine's future contract should pass cohort data through ctx so
// this swap-and-restore can go away.
async function generateFilteredXlsx() {
  if (!DASH.initialized || !DASH.primaryDataset) {
    throw new Error('Dashboard is not initialized.');
  }
  const filtered = applyDashboardFilters(DASH.primaryDataset);
  if (filtered.length === 0) {
    throw new Error('No rows match the current filter.');
  }
  const filteredPrimarySet  = new Set(filtered.map(d => d.primary));
  const filteredStitched    = STATE.stitched.filter(s => filteredPrimarySet.has(s.primary));
  const filteredPrimaryRows = STATE.primary.rows.filter(p => filteredPrimarySet.has(p));

  const origStitched     = STATE.stitched;
  const origPrimaryRows  = STATE.primary.rows;
  STATE.stitched      = filteredStitched;
  STATE.primary.rows  = filteredPrimaryRows;
  try {
    await generateXlsx({ filenameSuffix: '_filtered' });
  } finally {
    STATE.stitched     = origStitched;
    STATE.primary.rows = origPrimaryRows;
  }
}

/* ============================================================
   13. DASHBOARD — interactive exploration tab
   ============================================================
   Single-tab interactive dashboard with date-range slider, multi-select
   filters, live KPIs, conversion funnel chart, dual-line acquisition vs.
   conversion time-series, and click-to-drill-down modal. */

const PALETTE = {
  navy:     '#1F3864',
  blue:     '#8FAADC',
  green:    '#1f7a3a',
  amber:    '#d97706',
  slate:    '#6b7280',
  burgundy: '#9b1c1c',
};
// Emory-specific status canonical strings, status palette, and Sub-Type bucket
// chip options moved to CONFIG.dashboard (see configs/emory-cm-pa.js). Engine
// reads them via CONFIG.dashboard.{statusOrder,statusColor,bucketOptions}.

const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const ONE_DAY = 86400000;

// Two date windows scope the stitch (see runStitch):
//   Focus     — Participant Created Date range; bounds the secondary input.
//   Influence — Campaign Member Updated Date range; bounds the primary input
//               (typically wider than Focus to capture earlier campaign touches).
// The other filters (parent, sub-type, status) apply at render time only.
const FILTERS = {
  focus:           { dateMin: null, dateMax: null },
  influence:       { dateMin: null, dateMax: null },
  parentCampaigns: null,   // null = all, else Set<string>
  subTypes:        null,
  subTypeBuckets:  null,
  courseStatuses:  null,
};
const DASH = {
  initialized:    false,
  primaryDataset:      null,    // [{ primary, secondary, courseStatus, parentCampaign, subType, ... }]
  parentList:     [],
  parentSelected: null,    // Set<string>
  subTypeList:    [],
  subTypeSelected: null,   // Set<string>
  // Available bounds (from raw data) — set at first stitch, used to clamp
  // the sliders. *Default is the auto-fill value used by Reset Filters.
  focusMinAvail:      null, focusMaxAvail:      null, focusDefault:      null,
  influenceMinAvail:  null, influenceMaxAvail:  null, influenceDefault:  null,
  funnelChart:    null,
  timeseriesChart:null,
  drillRows:      [],
  drillPage:      0,
  drillPageSize:  12,
};

/* --- Date helpers -------------------------------------------------------- */

function parseSfDate(s) {
  if (s == null || s === '') return null;
  const str = String(s).trim();
  if (!str) return null;
  // ISO: YYYY-MM-DD[ THH:MM[:SS]]
  const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (iso) {
    const d = new Date(+iso[1], +iso[2]-1, +iso[3], +(iso[4]||0), +(iso[5]||0), +(iso[6]||0));
    return isFinite(d.getTime()) ? d : null;
  }
  // US: M/D/YYYY [HH:MM[:SS]] [AM|PM]
  const us = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?/i);
  if (us) {
    let year = +us[3];
    if (year < 100) year += year < 50 ? 2000 : 1900;
    let h = +(us[4] || 0);
    const m = +(us[5] || 0);
    const sec = +(us[6] || 0);
    const ampm = us[7];
    if (ampm) {
      if (ampm.toUpperCase() === 'PM' && h < 12) h += 12;
      if (ampm.toUpperCase() === 'AM' && h === 12) h = 0;
    }
    const d = new Date(year, +us[1]-1, +us[2], h, m, sec);
    return isFinite(d.getTime()) ? d : null;
  }
  return null;
}

function fmtDate(d) {
  if (!(d instanceof Date) || !isFinite(d.getTime())) return '—';
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
}

function startOfDay(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function startOfWeek(d) {
  const day = d.getDay(); // 0 = Sun
  const offset = day === 0 ? 6 : day - 1;  // ISO week starts Mon
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() - offset);
}
function startOfMonth(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}
function endOfMonth(d) {
  // Last calendar day of the month at start-of-day. Day 0 of next month = last day of this month.
  return new Date(d.getFullYear(), d.getMonth() + 1, 0);
}

// Range-check helper. `range` is { dateMin, dateMax } (both optional). A null
// date never matches; the bound is inclusive on both sides at start-of-day.
function dateInRange(date, range) {
  if (!date || !range) return false;
  if (range.dateMin && date < range.dateMin) return false;
  if (range.dateMax && date > new Date(range.dateMax.getTime() + 86400000 - 1)) return false;
  return true;
}

// Wipe the Focus + Influence windows so autoFillDateRanges re-seeds them
// from data on the next stitch. Called after a fresh upload / remove / Reset
// — the previous window may no longer make sense with new data.
function clearActiveDateFilters() {
  FILTERS.focus.dateMin     = null;
  FILTERS.focus.dateMax     = null;
  FILTERS.influence.dateMin = null;
  FILTERS.influence.dateMax = null;
}

function cacheParsedDatesOnRows() {
  // Called once at stitch time. Each row gets a parsed Date attached so we
  // never re-parse during filter changes.
  for (const r of STATE.primary.rows) {
    if (r._datesCached) continue;
    r._memberStatusUpdate = parseSfDate(r['Member Status Update Date']);
    r._memberFirstResponded = parseSfDate(r['Member First Responded Date']);
    r._datesCached = true;
  }
  for (const r of STATE.secondary.rows) {
    if (r._datesCached) continue;
    r._secondaryCreated   = parseSfDate(r['Program Participant: Created Date']);
    r._courseStart = parseSfDate(r['Course Instance: Start Date']);
    r._datesCached = true;
  }
}

/* --- Primary-centric dataset for the dashboard -------------------------- */

function buildPrimaryDataset() {
  const primaryClean = STATE.primary.rows.filter(r => !CONFIG.testRowFilter(r));
  const primaryToStitched = new Map();
  for (const s of STATE.stitched) {
    if (!primaryToStitched.has(s.primary)) primaryToStitched.set(s.primary, []);
    primaryToStitched.get(s.primary).push(s);
  }
  const rankSecondary = (s) => {
    const st = (s.secondary['Status'] || '').trim();
    if (st === 'Enrolled')   return 3;
    if (st === 'Registered') return 2;
    return 1;
  };
  return primaryClean.map(primary => {
    const matches = primaryToStitched.get(primary) || [];
    const courseStatus = CONFIG.deriveCourseStatus(matches);
    const best = matches.length === 0 ? null : matches.reduce((a, b) => rankSecondary(b) > rankSecondary(a) ? b : a);
    return {
      primary,
      secondary:      best ? best.secondary : null,
      bestMatch:      best,
      courseStatus,
      parentCampaign: (primary['Parent Campaign Name'] || '').trim() || '(blank)',
      campaignName:   (primary['Campaign Name']        || '').trim() || '(blank)',
      subType:        (primary['Sub-Type']             || '').trim() || '(blank)',
      subTypeBucket:  CONFIG.derivedFields.subtype_bucket(primary),
      primaryActivity:    primary._memberStatusUpdate,
      secondaryCreated:   best ? best.secondary._secondaryCreated : null,
      courseStart:        best ? best.secondary._courseStart      : null,
    };
  });
}

function applyDashboardFilters(dataset) {
  // Date filtering is no longer needed here — runStitch bounds the inputs
  // by FILTERS.focus / FILTERS.influence at stitch time, so DASH.primaryDataset
  // (built from STATE.stitched) is already date-scoped.
  const { parentCampaigns, subTypes, subTypeBuckets, courseStatuses } = FILTERS;
  return dataset.filter(d => {
    if (parentCampaigns && !parentCampaigns.has(d.parentCampaign)) return false;
    if (subTypes        && !subTypes.has(d.subType)) return false;
    if (subTypeBuckets  && !subTypeBuckets.has(d.subTypeBucket)) return false;
    if (courseStatuses  && !courseStatuses.has(d.courseStatus)) return false;
    return true;
  });
}

/* --- Tab routing -------------------------------------------------------- */

function setActiveTab(name) {
  if (name !== 'configure' && name !== 'dashboard') name = 'configure';
  if (name === 'dashboard' && !DASH.initialized) {
    name = 'configure';
    if (location.hash === '#dashboard') location.hash = '#configure';
  }
  document.querySelectorAll('.tab-link').forEach(link => {
    link.classList.toggle('active', link.dataset.tab === name);
  });
  document.querySelectorAll('.tab-content').forEach(el => {
    el.hidden = el.id !== `tab-${name}`;
  });
  // Tab swaps preserve scroll position by default, which surfaces the wrong
  // chunk of the new tab when the user clicked from mid-page (e.g. the
  // "Slice it on the Dashboard" callout near the bottom of Step 2). Reset.
  window.scrollTo({ top: 0, behavior: 'smooth' });
  if (name === 'dashboard') {
    // Charts were created while the panel was display:none, so canvas
    // dimensions may be stale. Fire resize after layout settles so Chart.js
    // recomputes widths.
    requestAnimationFrame(() => {
      window.dispatchEvent(new Event('resize'));
      refreshDashboard();
    });
  }
}

function enableDashboardTab(showCue = true) {
  const link = document.getElementById('tab-link-dashboard');
  // Diff first so a re-stitch in the same session doesn't re-fire the cue.
  // Cache restores also skip the cue (showCue=false) — it's a quiet welcome-back.
  const wasDisabled = link.classList.contains('disabled');
  link.classList.remove('disabled');
  link.removeAttribute('title');
  if (wasDisabled && showCue) {
    link.classList.add('fresh');
    link.addEventListener('click', () => link.classList.remove('fresh'), { once: true });
  }
}

/* --- Filter UI ---------------------------------------------------------- */

// Re-stitch with current Focus/Influence windows (which the date controls
// have just mutated in place on FILTERS.focus / FILTERS.influence). Skips
// the "fresh data" cue and scroll-into-view since this is a refinement of
// an existing result, not a brand-new stitch.
function onDateRangeChange() {
  runStitch({ fromCache: true, skipScroll: true });
}

// Helpers used by both the Configure-tab Date Ranges step and the
// Dashboard's filter-bar date inputs.
function _dateToIsoInput(d) {
  if (!d) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// Wire a single date input against a mutable FILTERS.*.{dateMin|dateMax}.
// Clamps to [availMin, availMax] and to the other end of the range, then
// fires onChange (which typically re-stitches).
function wireDateInput(inputId, availMin, availMax, range, kind, onChange) {
  const inp = document.getElementById(inputId);
  if (!inp) return;
  inp.min = _dateToIsoInput(availMin);
  inp.max = _dateToIsoInput(availMax);
  inp.value = _dateToIsoInput(kind === 'min' ? range.dateMin : range.dateMax);
  inp.onchange = () => {
    if (!inp.value) return;
    const parsed = startOfDay(new Date(inp.value + 'T00:00:00'));
    if (isNaN(parsed.getTime())) return;
    const lo = availMin.getTime(), hi = availMax.getTime();
    const ts = Math.max(lo, Math.min(hi, parsed.getTime()));
    const next = startOfDay(new Date(ts));
    if (kind === 'min') {
      range.dateMin = (range.dateMax && next > range.dateMax) ? range.dateMax : next;
    } else {
      range.dateMax = (range.dateMin && next < range.dateMin) ? range.dateMin : next;
    }
    inp.value = _dateToIsoInput(range[kind === 'min' ? 'dateMin' : 'dateMax']);
    onChange();
  };
}

// Wire the (min, max) pair for one date range. Used twice in initDashboard
// (Focus + Influence) and reused by renderDateRangesStep for the Configure
// tab inputs.
function wireDateRangeControls(opts) {
  const { minInputId, maxInputId, availMin, availMax, filterRange, onChange } = opts;
  wireDateInput(minInputId, availMin, availMax, filterRange, 'min', onChange);
  wireDateInput(maxInputId, availMin, availMax, filterRange, 'max', onChange);
}

// Configure-tab Date Ranges step — same control surface as the dashboard's
// filter-bar pair, just on the static Configure flow before xlsx generation.
function renderDateRangesStep() {
  if (!CONFIG.dashboard || !CONFIG.dashboard.dateOf) return;

  wireDateRangeControls({
    minInputId: 'cfg-focus-min',     maxInputId: 'cfg-focus-max',
    availMin:   DASH.focusMinAvail,  availMax:   DASH.focusMaxAvail,
    filterRange: FILTERS.focus,      onChange:   onDateRangeChange,
  });
  wireDateRangeControls({
    minInputId: 'cfg-influence-min', maxInputId: 'cfg-influence-max',
    availMin:   DASH.influenceMinAvail, availMax: DASH.influenceMaxAvail,
    filterRange: FILTERS.influence,  onChange:   onDateRangeChange,
  });

  const btn = document.getElementById('btn-dates-reset');
  if (btn) {
    btn.onclick = () => {
      if (DASH.focusDefault) {
        FILTERS.focus.dateMin = DASH.focusDefault.dateMin;
        FILTERS.focus.dateMax = DASH.focusDefault.dateMax;
      }
      if (DASH.influenceDefault) {
        FILTERS.influence.dateMin = DASH.influenceDefault.dateMin;
        FILTERS.influence.dateMax = DASH.influenceDefault.dateMax;
      }
      onDateRangeChange();
    };
  }
}

function initDashboard(opts = {}) {
  // Date caches + Focus/Influence windows are already set by runStitch (which
  // gated everything in this function). Build the per-CM dashboard dataset
  // from the date-bounded STATE.stitched.
  DASH.primaryDataset = buildPrimaryDataset();

  // Initial filter state for non-date filters. Note: date ranges are owned
  // by runStitch (FILTERS.focus / FILTERS.influence). All on except Not
  // Converted (per consultant pitfall §5.3).
  FILTERS.parentCampaigns = null;
  FILTERS.subTypes        = null;
  FILTERS.subTypeBuckets  = new Set(CONFIG.dashboard.bucketOptions);
  FILTERS.courseStatuses  = new Set(['Cancelled/Withdrawn/Etc', 'Registered', 'Enrolled']);

  // Both date ranges wired through the same helper. Range changes trigger
  // re-stitch (Focus bounds secondary, Influence bounds primary), which in
  // turn re-renders the whole pipeline (KPIs, dashboard, etc).
  wireDateRangeControls({
    minInputId: 'filter-focus-min',
    maxInputId: 'filter-focus-max',
    availMin:   DASH.focusMinAvail,
    availMax:   DASH.focusMaxAvail,
    filterRange: FILTERS.focus,
    onChange:   onDateRangeChange,
  });
  wireDateRangeControls({
    minInputId: 'filter-influence-min',
    maxInputId: 'filter-influence-max',
    availMin:   DASH.influenceMinAvail,
    availMax:   DASH.influenceMaxAvail,
    filterRange: FILTERS.influence,
    onChange:   onDateRangeChange,
  });

  // Parent campaign multi-select
  DASH.parentList = [...new Set(DASH.primaryDataset.map(d => d.parentCampaign))].sort((a,b) => a.localeCompare(b));
  DASH.parentSelected = new Set(DASH.parentList);
  renderParentMultiselect();

  // Sub-Type multi-select (raw Sub-Type values, distinct from the bucket chip group below)
  DASH.subTypeList = [...new Set(DASH.primaryDataset.map(d => d.subType))].sort((a,b) => a.localeCompare(b));
  DASH.subTypeSelected = new Set(DASH.subTypeList);
  renderSubTypeMultiselect();

  // Sub-Type Bucket chips
  renderChipGroup('filter-bucket-chips', CONFIG.dashboard.bucketOptions, FILTERS.subTypeBuckets, (set) => {
    FILTERS.subTypeBuckets = set.size === 0 ? null : set;
    refreshDashboard();
  });

  // Course Status chips
  renderChipGroup('filter-status-chips', CONFIG.dashboard.statusOrder, FILTERS.courseStatuses, (set) => {
    FILTERS.courseStatuses = set.size === 0 ? null : set;
    refreshDashboard();
  });

  // Reset filters — both date ranges back to the auto-filled defaults plus
  // the non-date filters back to their initial state. Triggers a re-stitch
  // since date ranges changed.
  document.getElementById('btn-filter-reset').onclick = () => {
    FILTERS.subTypeBuckets  = new Set(CONFIG.dashboard.bucketOptions);
    FILTERS.courseStatuses  = new Set(['Cancelled/Withdrawn/Etc', 'Registered', 'Enrolled']);
    FILTERS.parentCampaigns = null;
    FILTERS.subTypes        = null;
    DASH.parentSelected     = new Set(DASH.parentList);
    DASH.subTypeSelected    = new Set(DASH.subTypeList);
    if (DASH.focusDefault) {
      FILTERS.focus.dateMin = DASH.focusDefault.dateMin;
      FILTERS.focus.dateMax = DASH.focusDefault.dateMax;
    }
    if (DASH.influenceDefault) {
      FILTERS.influence.dateMin = DASH.influenceDefault.dateMin;
      FILTERS.influence.dateMax = DASH.influenceDefault.dateMax;
    }
    renderChipGroup('filter-bucket-chips', CONFIG.dashboard.bucketOptions, FILTERS.subTypeBuckets, (set) => {
      FILTERS.subTypeBuckets = set.size === 0 ? null : set;
      refreshDashboard();
    });
    renderChipGroup('filter-status-chips', CONFIG.dashboard.statusOrder, FILTERS.courseStatuses, (set) => {
      FILTERS.courseStatuses = set.size === 0 ? null : set;
      refreshDashboard();
    });
    renderParentMultiselect();
    renderSubTypeMultiselect();
    onDateRangeChange();
  };

  DASH.initialized = true;
  enableDashboardTab(opts.showCue !== false);
}

function renderChipGroup(containerId, options, activeSet, onChange) {
  const container = document.getElementById(containerId);
  container.innerHTML = '';
  for (const opt of options) {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'chip' + (activeSet.has(opt) ? ' active' : '');
    chip.textContent = opt;
    chip.dataset.status = opt;
    chip.addEventListener('click', () => {
      if (activeSet.has(opt)) activeSet.delete(opt);
      else activeSet.add(opt);
      chip.classList.toggle('active', activeSet.has(opt));
      onChange(activeSet);
    });
    container.appendChild(chip);
  }
}

// Generic multi-select dropdown shared by Parent Campaign and Sub-Type filters.
// opts: {
//   toggleId, dropdownId,
//   listGetter:     () => string[],     // current source list (re-evaluated each rebuild)
//   selectedGetter: () => Set<string>,  // mutable selection set
//   onCommit:       (sel) => void,      // fires on every checkbox change (incl. all-row)
//   nounPlural,                         // 'campaigns', 'sub-types', etc.
//   nounAllLabel,                       // 'All campaigns', 'All sub-types' (no caret)
//   searchPlaceholder,                  // optional, defaults to 'Search…'
// }
//
// The toggle's onclick does NOT stopPropagation — letting the click bubble means
// each multi-select's own outside-click handler closes any sibling multi-select
// that's open, so only one dropdown is ever open at a time without explicit
// cross-coordination. The `_docHandlerInstalled` guard lives on the toggle DOM
// element (not the function) so multiple toggles each install their own handler
// exactly once.
function renderMultiselect(opts) {
  const {
    toggleId, dropdownId,
    listGetter, selectedGetter, onCommit,
    nounPlural, nounAllLabel,
    searchPlaceholder,
  } = opts;

  const toggle   = document.getElementById(toggleId);
  const dropdown = document.getElementById(dropdownId);
  if (!toggle || !dropdown) return;

  function updateLabel() {
    const list  = listGetter();
    const sel   = selectedGetter();
    const total = list.length;
    const n     = sel.size;
    if (n === total)      toggle.innerHTML = `${escapeHtml(nounAllLabel)} (${total}) &#x25BE;`;
    else if (n === 0)     toggle.innerHTML = `No ${escapeHtml(nounPlural)} selected &#x25BE;`;
    else if (n === 1)     toggle.innerHTML = `${escapeHtml([...sel][0]).slice(0, 32)} &#x25BE;`;
    else                  toggle.innerHTML = `${n} of ${total} ${escapeHtml(nounPlural)} &#x25BE;`;
  }

  function rebuild(filterText) {
    dropdown.innerHTML = '';
    const search = document.createElement('input');
    search.className   = 'filter-multi-search';
    search.placeholder = searchPlaceholder || 'Search…';
    search.value       = filterText || '';
    search.addEventListener('input', e => rebuild(e.target.value));
    dropdown.appendChild(search);
    setTimeout(() => search.focus(), 0);

    const list    = listGetter();
    const sel     = selectedGetter();
    const ft      = (filterText || '').toLowerCase();
    const visible = list.filter(p => !ft || p.toLowerCase().includes(ft));

    const allRow = document.createElement('label');
    allRow.className = 'filter-multi-option all-row';
    const allCb = document.createElement('input');
    allCb.type  = 'checkbox';
    allCb.checked = visible.length > 0 && visible.every(p => sel.has(p));
    allCb.indeterminate = !allCb.checked && visible.some(p => sel.has(p));
    allCb.addEventListener('change', () => {
      if (allCb.checked) for (const p of visible) sel.add(p);
      else               for (const p of visible) sel.delete(p);
      rebuild(filterText);
      updateLabel();
      onCommit(sel);
    });
    const allLbl = document.createElement('span');
    allLbl.textContent = ft ? `Select all visible (${visible.length})` : `Select all (${list.length})`;
    allRow.appendChild(allCb);
    allRow.appendChild(allLbl);
    dropdown.appendChild(allRow);

    for (const p of visible) {
      const row = document.createElement('label');
      row.className = 'filter-multi-option';
      const cb = document.createElement('input');
      cb.type    = 'checkbox';
      cb.checked = sel.has(p);
      cb.addEventListener('change', () => {
        if (cb.checked) sel.add(p);
        else            sel.delete(p);
        updateLabel();
        onCommit(sel);
      });
      const lbl = document.createElement('span');
      lbl.textContent = p;
      row.appendChild(cb);
      row.appendChild(lbl);
      dropdown.appendChild(row);
    }
  }

  toggle.onclick = () => {
    if (dropdown.hidden) { rebuild(''); dropdown.hidden = false; }
    else                 { dropdown.hidden = true; }
  };

  if (!toggle._docHandlerInstalled) {
    document.addEventListener('click', (e) => {
      if (!dropdown.contains(e.target) && e.target !== toggle && !toggle.contains(e.target)) {
        dropdown.hidden = true;
      }
    });
    toggle._docHandlerInstalled = true;
  }

  updateLabel();
}

// Strip the trailing caret from a "[noun] ▾" config label so the helper can
// re-insert the count + caret in the right spots.
function _stripCaret(s) {
  return String(s || '').replace(/\s*[▾▼]\s*$/, '');
}

function renderParentMultiselect() {
  const allLbl = (CONFIG.dashboard && CONFIG.dashboard.labels && CONFIG.dashboard.labels.filterParentAll) || 'All campaigns ▾';
  renderMultiselect({
    toggleId:        'filter-parent-toggle',
    dropdownId:      'filter-parent-dropdown',
    listGetter:      () => DASH.parentList,
    selectedGetter:  () => DASH.parentSelected,
    onCommit:        (sel) => {
      FILTERS.parentCampaigns = sel.size === DASH.parentList.length ? null : sel;
      refreshDashboard();
    },
    nounPlural:      'campaigns',
    nounAllLabel:    _stripCaret(allLbl),
    searchPlaceholder: 'Search…',
  });
}

function renderSubTypeMultiselect() {
  const allLbl = (CONFIG.dashboard && CONFIG.dashboard.labels && CONFIG.dashboard.labels.filterSubTypeAll) || 'All sub-types ▾';
  renderMultiselect({
    toggleId:        'filter-subtype-toggle',
    dropdownId:      'filter-subtype-dropdown',
    listGetter:      () => DASH.subTypeList,
    selectedGetter:  () => DASH.subTypeSelected,
    onCommit:        (sel) => {
      FILTERS.subTypes = sel.size === DASH.subTypeList.length ? null : sel;
      refreshDashboard();
    },
    nounPlural:      'sub-types',
    nounAllLabel:    _stripCaret(allLbl),
    searchPlaceholder: 'Search sub-types…',
  });
}

/* --- Refresh orchestration --------------------------------------------- */

function refreshDashboard() {
  if (!DASH.initialized) return;
  const filtered = applyDashboardFilters(DASH.primaryDataset);
  renderDashKpis(filtered, DASH.primaryDataset);
  renderFunnelChart(filtered);
  renderTimeSeriesChart(filtered);
  // Distribution charts: feed the subset of stitched rows whose CM passed the
  // dashboard filter. aggregateAll only counts Registered + Enrolled so the
  // status chips don't matter to these charts — date/parent/bucket filters do.
  // The set is also threaded into the chart's onClick handler so the
  // drilldown reuses the same cohort.
  const filteredPrimarySet = new Set(filtered.map(d => d.primary));
  const filteredStitched = STATE.stitched.filter(s => filteredPrimarySet.has(s.primary));
  renderInPageCharts(aggregateAll(filteredStitched), filteredPrimarySet);
  renderCampaignSummaryTable(filtered);
  renderFilterSummary(filtered, DASH.primaryDataset);
}

function renderCampaignSummaryTable(filtered) {
  const tbl = document.getElementById('campaign-summary-table');
  if (filtered.length === 0) {
    tbl.innerHTML = '<tbody><tr><td class="campaign-summary-empty">No Campaign Members match the current filter.</td></tr></tbody>';
    return;
  }

  // Group by parent → campaign → status counts
  const byParent = new Map();
  for (const d of filtered) {
    const p = d.parentCampaign;
    const c = d.campaignName;
    if (!byParent.has(p)) byParent.set(p, new Map());
    const inner = byParent.get(p);
    if (!inner.has(c)) inner.set(c, { 'Not Converted':0, 'Cancelled/Withdrawn/Etc':0, 'Registered':0, 'Enrolled':0 });
    inner.get(c)[d.courseStatus]++;
  }

  const sortedParents = [...byParent.keys()].sort((a, b) => a.localeCompare(b));
  let html = '<thead><tr>'
    + '<th class="col-text">Parent Campaign</th>'
    + '<th class="col-text">Campaign Name</th>'
    + '<th>Not Converted</th>'
    + '<th>Cancelled / Withdrawn / Etc</th>'
    + '<th>Enrolled</th>'
    + '<th>Registered</th>'
    + '<th>Total</th>'
    + '</tr></thead><tbody>';

  let grNc = 0, grCx = 0, grEn = 0, grRg = 0;
  for (const parent of sortedParents) {
    const inner = byParent.get(parent);
    const sortedCampaigns = [...inner.keys()].sort((a, b) => a.localeCompare(b));
    let pNc = 0, pCx = 0, pEn = 0, pRg = 0;

    sortedCampaigns.forEach((campaign, idx) => {
      const c = inner.get(campaign);
      const total = c['Not Converted'] + c['Cancelled/Withdrawn/Etc'] + c['Enrolled'] + c['Registered'];
      html += '<tr>'
        + `<td class="col-text" title="${escapeHtml(parent)}">${idx === 0 ? `<span class="parent-name">${escapeHtml(parent)}</span>` : ''}</td>`
        + `<td class="col-text indent" title="${escapeHtml(campaign)}">${escapeHtml(campaign)}</td>`
        + `<td>${c['Not Converted'].toLocaleString()}</td>`
        + `<td>${c['Cancelled/Withdrawn/Etc'].toLocaleString()}</td>`
        + `<td>${c['Enrolled'].toLocaleString()}</td>`
        + `<td>${c['Registered'].toLocaleString()}</td>`
        + `<td>${total.toLocaleString()}</td>`
        + '</tr>';
      pNc += c['Not Converted']; pCx += c['Cancelled/Withdrawn/Etc'];
      pEn += c['Enrolled']; pRg += c['Registered'];
    });

    html += '<tr class="subtotal">'
      + `<td class="col-text"></td>`
      + `<td class="col-text">Subtotal</td>`
      + `<td>${pNc.toLocaleString()}</td>`
      + `<td>${pCx.toLocaleString()}</td>`
      + `<td>${pEn.toLocaleString()}</td>`
      + `<td>${pRg.toLocaleString()}</td>`
      + `<td>${(pNc + pCx + pEn + pRg).toLocaleString()}</td>`
      + '</tr>';
    grNc += pNc; grCx += pCx; grEn += pEn; grRg += pRg;
  }

  html += '<tr class="grand-total">'
    + `<td class="col-text">Grand Total</td>`
    + `<td class="col-text"></td>`
    + `<td>${grNc.toLocaleString()}</td>`
    + `<td>${grCx.toLocaleString()}</td>`
    + `<td>${grEn.toLocaleString()}</td>`
    + `<td>${grRg.toLocaleString()}</td>`
    + `<td>${(grNc + grCx + grEn + grRg).toLocaleString()}</td>`
    + '</tr>';

  html += '</tbody>';
  tbl.innerHTML = html;
}

function renderFilterSummary(filtered, all) {
  const summary = document.getElementById('filter-summary');
  const f = FILTERS.focus, i = FILTERS.influence;
  summary.innerHTML =
    `Showing <strong>${filtered.length.toLocaleString()}</strong> of ${all.length.toLocaleString()} Campaign Members ` +
    `· Focus <strong>${fmtDate(f.dateMin)} – ${fmtDate(f.dateMax)}</strong>` +
    ` · Influence <strong>${fmtDate(i.dateMin)} – ${fmtDate(i.dateMax)}</strong>`;

  // Disable the filtered-xlsx button when nothing's selected — generating an
  // empty xlsx is useless and the builders would throw on the empty cohort.
  const btn = document.getElementById('btn-filtered-xlsx');
  if (btn) btn.disabled = filtered.length === 0;
}

function renderDashKpis(filtered) {
  let nc = 0, cx = 0, rg = 0, en = 0;
  for (const d of filtered) {
    if (d.courseStatus === 'Not Converted')                 nc++;
    else if (d.courseStatus === 'Cancelled/Withdrawn/Etc')  cx++;
    else if (d.courseStatus === 'Registered')               rg++;
    else if (d.courseStatus === 'Enrolled')                 en++;
  }
  document.getElementById('dash-kpi-total').textContent        = filtered.length.toLocaleString();
  document.getElementById('dash-kpi-notconverted').textContent = nc.toLocaleString();
  document.getElementById('dash-kpi-cancelled').textContent    = cx.toLocaleString();
  document.getElementById('dash-kpi-registered').textContent   = rg.toLocaleString();
  document.getElementById('dash-kpi-enrolled').textContent     = en.toLocaleString();
}

/* --- Funnel chart ------------------------------------------------------- */

function renderFunnelChart(filtered) {
  const counts = { 'Not Converted': 0, 'Cancelled/Withdrawn/Etc': 0, 'Registered': 0, 'Enrolled': 0 };
  for (const d of filtered) counts[d.courseStatus]++;
  const total = filtered.length || 1;
  const pct = (n) => ((n / total) * 100).toFixed(1) + '%';

  const canvas = document.getElementById('chart-funnel');
  if (DASH.funnelChart) { try { DASH.funnelChart.destroy(); } catch(e) {} DASH.funnelChart = null; }

  // Build aria-label dynamically for screen readers
  const ariaParts = CONFIG.dashboard.statusOrder.map(s => `${s} ${counts[s]}`);
  canvas.setAttribute('aria-label', `Conversion funnel: total ${filtered.length}, ${ariaParts.join(', ')}`);

  DASH.funnelChart = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: ['Funnel'],
      datasets: CONFIG.dashboard.statusOrder.map(status => ({
        label: `${status} — ${counts[status]} (${pct(counts[status])})`,
        data: [counts[status]],
        backgroundColor: CONFIG.dashboard.statusColor[status],
        borderWidth: 0,
        _statusKey: status,
      })),
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      layout: { padding: { top: 4, bottom: 4 } },
      scales: {
        x: { stacked: true, beginAtZero: true, ticks: { font: { size: 12 }, color: '#000', precision: 0 }, grid: { color: 'rgba(0,0,0,.06)' } },
        y: { stacked: true, display: false },
      },
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 12 }, color: '#000', boxWidth: 14, padding: 10 } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const s = CONFIG.dashboard.statusOrder[ctx.datasetIndex];
              return `${s}: ${ctx.raw.toLocaleString()} (${pct(ctx.raw)})`;
            },
          },
        },
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const status = CONFIG.dashboard.statusOrder[elements[0].datasetIndex];
        const matching = filtered.filter(d => d.courseStatus === status);
        openDrilldown(`${status} — ${matching.length.toLocaleString()} Campaign Member${matching.length === 1 ? '' : 's'}`, matching);
      },
    },
    plugins: [funnelLabelPlugin()],
  });
}

function funnelLabelPlugin() {
  return {
    id: 'funnelLabel',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      ctx.save();
      ctx.font = 'bold 12px Arial, sans-serif';
      ctx.textBaseline = 'middle';
      ctx.textAlign = 'center';
      ctx.fillStyle = '#fff';
      for (let i = 0; i < chart.data.datasets.length; i++) {
        const meta = chart.getDatasetMeta(i);
        const value = chart.data.datasets[i].data[0];
        if (value === 0) continue;
        const bar = meta.data[0];
        const segWidth = bar.x - bar.base;
        if (segWidth >= 28) {
          ctx.fillText(String(value), (bar.base + bar.x) / 2, bar.y);
        }
      }
      ctx.restore();
    },
  };
}

/* --- Time-series chart -------------------------------------------------- */

function renderTimeSeriesChart(filtered) {
  // CM acquisition (primaryActivity) and PA conversion (secondaryCreated), filtered cohort only.
  // X-axis spans Influence (the wider window) so CM lookback is visible; the
  // PA-conversion line naturally falls inside Focus since stitch is Focus-bounded.
  const primaryDates = filtered.filter(d => d.primaryActivity).map(d => d.primaryActivity);
  const secondaryDates = filtered.filter(d => d.secondaryCreated).map(d => d.secondaryCreated);

  const infMin = FILTERS.influence.dateMin;
  const infMax = FILTERS.influence.dateMax;
  const rangeMs = (infMax && infMin) ? infMax - infMin : 0;
  const days = rangeMs / ONE_DAY;
  const bin = days < 60 ? 'day' : days < 540 ? 'week' : 'month';
  const binStart = bin === 'day' ? startOfDay : bin === 'week' ? startOfWeek : startOfMonth;
  const binAdvance = (start) => {
    if (bin === 'day')  return new Date(start.getTime() + ONE_DAY);
    if (bin === 'week') return new Date(start.getTime() + 7 * ONE_DAY);
    return new Date(start.getFullYear(), start.getMonth() + 1, 1);
  };

  // Build bin sequence covering Influence (or fall back to whatever dates we have)
  const startBin = infMin ? binStart(infMin) : (primaryDates.length || secondaryDates.length ? binStart(new Date(Math.min(...[...primaryDates, ...secondaryDates].map(d => d.getTime())))) : new Date());
  const endBin   = infMax ? binStart(infMax) : startBin;
  const bins = [];
  for (let cur = new Date(startBin.getTime()); cur <= endBin; cur = binAdvance(cur)) {
    bins.push(cur.getTime());
  }

  const cmCounts = new Map(bins.map(t => [t, 0]));
  const paCounts = new Map(bins.map(t => [t, 0]));
  for (const d of primaryDates) {
    const k = binStart(d).getTime();
    if (cmCounts.has(k)) cmCounts.set(k, cmCounts.get(k) + 1);
  }
  for (const d of secondaryDates) {
    const k = binStart(d).getTime();
    if (paCounts.has(k)) paCounts.set(k, paCounts.get(k) + 1);
  }

  const labels   = bins.map(t => fmtBinLabel(new Date(t), bin));
  const cmSeries = bins.map(t => cmCounts.get(t));
  const paSeries = bins.map(t => paCounts.get(t));

  const canvas = document.getElementById('chart-timeseries');
  canvas.setAttribute('aria-label', `Time series: ${primaryDates.length} CMs acquired, ${secondaryDates.length} PAs created in the filtered range, binned by ${bin}.`);

  if (DASH.timeseriesChart) { try { DASH.timeseriesChart.destroy(); } catch(e) {} DASH.timeseriesChart = null; }
  DASH.timeseriesChart = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: `CMs acquired (${primaryDates.length})`,
          data: cmSeries,
          borderColor: PALETTE.navy,
          backgroundColor: PALETTE.navy,
          borderWidth: 2.5,
          tension: 0.25,
          yAxisID: 'y',
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false,
        },
        {
          label: `PAs created (${secondaryDates.length})`,
          data: paSeries,
          borderColor: PALETTE.amber,
          backgroundColor: PALETTE.amber,
          borderWidth: 2.5,
          tension: 0.25,
          yAxisID: 'y1',
          pointRadius: 3,
          pointHoverRadius: 5,
          fill: false,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { ticks: { font: { size: 11 }, color: '#000', maxRotation: 45, minRotation: 0 } },
        y:  { position: 'left',  beginAtZero: true, ticks: { font: { size: 11 }, color: PALETTE.navy,  precision: 0 }, title: { display: true, text: 'CMs acquired',  color: PALETTE.navy  } },
        y1: { position: 'right', beginAtZero: true, ticks: { font: { size: 11 }, color: PALETTE.amber, precision: 0 }, grid: { drawOnChartArea: false }, title: { display: true, text: 'PAs created', color: PALETTE.amber } },
      },
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 12 }, color: '#000', padding: 10 } },
      },
      onClick: (evt, elements) => {
        if (!elements.length) return;
        const el = elements[0];
        const idx = el.index;
        const dsIdx = el.datasetIndex;
        const t0 = bins[idx];
        const t1 = idx + 1 < bins.length ? bins[idx + 1] : binAdvance(new Date(t0)).getTime();
        const labelTxt = labels[idx];
        if (dsIdx === 0) {
          const matching = filtered.filter(d => d.primaryActivity && d.primaryActivity.getTime() >= t0 && d.primaryActivity.getTime() < t1);
          openDrilldown(`CMs acquired ${labelTxt} — ${matching.length.toLocaleString()}`, matching);
        } else {
          const matching = filtered.filter(d => d.secondaryCreated && d.secondaryCreated.getTime() >= t0 && d.secondaryCreated.getTime() < t1);
          openDrilldown(`PAs created ${labelTxt} — ${matching.length.toLocaleString()}`, matching);
        }
      },
    },
  });
}

function fmtBinLabel(d, bin) {
  if (bin === 'day')  return `${d.getMonth()+1}/${d.getDate()}`;
  if (bin === 'week') return `Wk of ${d.getMonth()+1}/${d.getDate()}`;
  return `${MONTH_ABBR[d.getMonth()]} ${String(d.getFullYear()).slice(-2)}`;
}

/* --- Drill-down modal -------------------------------------------------- */

function openDrilldown(title, cmRows) {
  DASH.drillRows = cmRows;
  DASH.drillPage = 0;
  document.getElementById('drilldown-title').textContent = title;
  document.getElementById('drilldown-meta').innerHTML =
    `<strong>${cmRows.length.toLocaleString()}</strong> Campaign Member${cmRows.length === 1 ? '' : 's'} ` +
    `match the current filter and selection.`;
  renderDrilldownPage();
  document.getElementById('drilldown-dialog').showModal();
}

function renderDrilldownPage() {
  const { drillRows: rows, drillPage: page, drillPageSize: size } = DASH;
  const start = page * size;
  const end   = Math.min(rows.length, start + size);
  const slice = rows.slice(start, end);

  const headers = ['Parent Campaign', 'Campaign', 'Status', 'Last', 'First', 'Email', 'CM Update', 'PA Course', 'PA Status', 'PA Created'];
  let html = '<thead><tr>';
  for (const h of headers) html += `<th>${escapeHtml(h)}</th>`;
  html += '</tr></thead><tbody>';
  for (const d of slice) {
    html += '<tr>';
    html += `<td>${escapeHtml(d.parentCampaign)}</td>`;
    html += `<td>${escapeHtml(d.campaignName)}</td>`;
    html += `<td>${escapeHtml(d.courseStatus)}</td>`;
    html += `<td>${escapeHtml(d.primary['Last Name'] || '')}</td>`;
    html += `<td>${escapeHtml(d.primary['First Name'] || '')}</td>`;
    html += `<td>${escapeHtml(d.primary['Email'] || '')}</td>`;
    html += `<td>${d.primaryActivity ? escapeHtml(fmtDate(d.primaryActivity)) : ''}</td>`;
    html += `<td>${d.secondary ? escapeHtml(d.secondary['Course Name'] || '') : ''}</td>`;
    html += `<td>${d.secondary ? escapeHtml(d.secondary['Status'] || '') : ''}</td>`;
    html += `<td>${d.secondaryCreated ? escapeHtml(fmtDate(d.secondaryCreated)) : ''}</td>`;
    html += '</tr>';
  }
  if (slice.length === 0) html += `<tr><td colspan="${headers.length}" style="text-align:center;color:var(--text-muted);padding:24px;">(no matching rows)</td></tr>`;
  html += '</tbody>';
  document.getElementById('drilldown-table').innerHTML = html;

  document.getElementById('drilldown-prev').disabled = page === 0;
  document.getElementById('drilldown-next').disabled = end >= rows.length;
  document.getElementById('drilldown-pager-meta').textContent =
    rows.length === 0 ? '0 of 0' : `${start + 1}–${end} of ${rows.length.toLocaleString()}`;
}

function setupDrilldownHandlers() {
  document.getElementById('drilldown-close').addEventListener('click', () => {
    document.getElementById('drilldown-dialog').close();
  });
  document.getElementById('drilldown-prev').addEventListener('click', () => {
    if (DASH.drillPage > 0) { DASH.drillPage--; renderDrilldownPage(); }
  });
  document.getElementById('drilldown-next').addEventListener('click', () => {
    if ((DASH.drillPage + 1) * DASH.drillPageSize < DASH.drillRows.length) { DASH.drillPage++; renderDrilldownPage(); }
  });
}

/* ============================================================
   14. UI WIRE-UP
   ============================================================ */

function showToast(msg, isError) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.toggle('error', !!isError);
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3500);
}

function showError(containerId, msg) {
  const el = document.getElementById(containerId);
  if (!el) { showToast(msg, true); return; }
  el.innerHTML = `<div class="error-banner">${escapeHtml(msg)}</div>`;
  el.hidden = false;
}

function clearError(containerId) {
  const el = document.getElementById(containerId);
  if (el) { el.innerHTML = ''; el.hidden = true; }
}

function refreshStitchButton() {
  const btn = document.getElementById('btn-stitch');
  const hint = document.getElementById('stitch-hint');
  const ready = STATE.primary.rows && STATE.secondary.rows;
  btn.disabled = !ready;
  hint.textContent = ready
    ? `Ready: ${STATE.primary.rows.length.toLocaleString()} CM rows, ${STATE.secondary.rows.length.toLocaleString()} PA rows.`
    : 'Upload both reports to enable.';
}

function setupDropZone(zoneEl, target, required, label) {
  const input = zoneEl.querySelector('input[type="file"]');
  const status = zoneEl.querySelector('.file-status');

  const handleFile = async (file) => {
    if (!file) return;
    clearError('upload-error');
    status.innerHTML = `<span class="spinner"></span>Reading…`;
    try {
      const rows = await readCsv(file);
      validateHeaders(rows, required, label);
      STATE[target].rows = rows;
      STATE[target].fileName = file.name;
      // Fresh data invalidates the cached parsed-date flags and the user's
      // current Focus/Influence windows — autoFillDateRanges will re-seed
      // from the new data on the next stitch.
      clearActiveDateFilters();
      zoneEl.classList.add('loaded');
      // A fresh upload supersedes any "restored from last session" badge.
      zoneEl.classList.remove('from-cache');
      status.innerHTML = `<span class="filename">${escapeHtml(file.name)}</span><br><span class="row-count">${rows.length.toLocaleString()} rows</span> <span class="clear-link" data-clear="${target}">remove</span>`;
      refreshStitchButton();
      cachePutCsv(target, file.name, rows).catch(e => console.warn('Cache write failed', e));
      refreshResetButton();
    } catch (err) {
      console.error(err);
      zoneEl.classList.remove('loaded');
      status.innerHTML = '';
      STATE[target].rows = null;
      STATE[target].fileName = null;
      showError('upload-error', `${label}: ${err.message}`);
      refreshStitchButton();
    }
  };

  zoneEl.addEventListener('click', (e) => {
    // Don't fire when clicking the inline 'remove' link
    if (e.target.classList.contains('clear-link')) return;
    input.click();
  });
  input.addEventListener('change', () => handleFile(input.files[0]));
  zoneEl.addEventListener('dragover', (e) => { e.preventDefault(); zoneEl.classList.add('dragover'); });
  zoneEl.addEventListener('dragleave', () => zoneEl.classList.remove('dragover'));
  zoneEl.addEventListener('drop', (e) => {
    e.preventDefault();
    zoneEl.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });

  // Delegate clear-link clicks
  zoneEl.addEventListener('click', (e) => {
    if (e.target.dataset && e.target.dataset.clear) {
      e.stopPropagation();
      e.preventDefault();
      STATE[target].rows = null;
      STATE[target].fileName = null;
      clearActiveDateFilters();
      zoneEl.classList.remove('loaded', 'from-cache');
      status.innerHTML = '';
      input.value = '';
      refreshStitchButton();
      cacheDeleteCsv(target).catch(e => console.warn('Cache delete failed', e));
      refreshResetButton();
    }
  });
}

function rerenderDownstream() {
  // After any column-picker edit, refresh preview only (charts use derived data, not column choices)
  if (!STATE.stitched) return;
  saveColumnConfig(STATE.columns);
  renderPreviewTable();
}

// Walks the primary + secondary rows once to find the date bounds that drive
// the Focus + Influence default windows, and stashes them on DASH for the
// slider min/max. Called from runStitch every time, but only writes into
// FILTERS.* fields that are still null — so user-adjusted ranges survive a
// re-stitch (e.g. a Reset Filters click hits the *Default values).
function autoFillDateRanges() {
  const dateOf = CONFIG.dashboard && CONFIG.dashboard.dateOf;
  if (!dateOf || typeof dateOf.primary !== 'function' || typeof dateOf.secondary !== 'function') return;

  // Influence — full span of primary date dimension (CM Update Date for Emory)
  let infMin = null, infMax = null;
  for (const r of STATE.primary.rows) {
    const d = dateOf.primary(r);
    if (!d) continue;
    if (!infMin || d < infMin) infMin = d;
    if (!infMax || d > infMax) infMax = d;
  }
  if (!infMin) { infMin = new Date(); infMax = new Date(); }
  DASH.influenceMinAvail = startOfDay(infMin);
  DASH.influenceMaxAvail = startOfDay(infMax);
  DASH.influenceDefault  = { dateMin: DASH.influenceMinAvail, dateMax: DASH.influenceMaxAvail };

  // Focus — month-rounded span of secondary date dimension (PA Created Date for Emory)
  let focMin = null, focMax = null;
  for (const r of STATE.secondary.rows) {
    const d = dateOf.secondary(r);
    if (!d) continue;
    if (!focMin || d < focMin) focMin = d;
    if (!focMax || d > focMax) focMax = d;
  }
  if (!focMin) { focMin = infMin; focMax = infMax; }
  DASH.focusMinAvail = startOfMonth(focMin);
  DASH.focusMaxAvail = endOfMonth(focMax);
  DASH.focusDefault  = { dateMin: DASH.focusMinAvail, dateMax: DASH.focusMaxAvail };

  // Seed FILTERS only on first stitch (or after a user clears them, e.g.
  // re-upload). Subsequent re-stitches preserve the user's chosen window.
  if (!FILTERS.focus.dateMin)     FILTERS.focus.dateMin     = DASH.focusDefault.dateMin;
  if (!FILTERS.focus.dateMax)     FILTERS.focus.dateMax     = DASH.focusDefault.dateMax;
  if (!FILTERS.influence.dateMin) FILTERS.influence.dateMin = DASH.influenceDefault.dateMin;
  if (!FILTERS.influence.dateMax) FILTERS.influence.dateMax = DASH.influenceDefault.dateMax;
}

// Filter primary rows by Influence, secondary by Focus. Rows whose date
// dimension is null are excluded (we can't place them on a time axis).
function filterByActiveDateRanges() {
  const dateOf = CONFIG.dashboard.dateOf;
  const primaryInRange   = STATE.primary.rows.filter(r => dateInRange(dateOf.primary(r),   FILTERS.influence));
  const secondaryInRange = STATE.secondary.rows.filter(r => dateInRange(dateOf.secondary(r), FILTERS.focus));
  return { primaryInRange, secondaryInRange };
}

function runStitch(opts = {}) {
  if (!STATE.primary.rows || !STATE.secondary.rows) return;

  // Date caches must exist before any date-aware filtering (autoFill walks
  // rows via CONFIG.dashboard.dateOf, which reads the underscore-prefixed
  // caches this populates). Idempotent — short-circuits on r._datesCached.
  cacheParsedDatesOnRows();

  // Resolve current Focus + Influence windows (first call seeds defaults;
  // re-stitches after a date-range change preserve the user's choices).
  autoFillDateRanges();

  // Scope the stitch: secondary ∈ Focus, primary ∈ Influence.
  const { primaryInRange, secondaryInRange } = filterByActiveDateRanges();

  const t0 = performance.now();
  try {
    const result = stitch(primaryInRange, secondaryInRange);
    STATE.stitched         = result.stitched;
    STATE.unmatched        = result.unmatched;
    STATE.methodCounts     = result.methodCounts;
    STATE.testRemovedCount = result.testRemovedCount;
    STATE.testRemovedRows  = result.testRemovedRows || { primary: [], secondary: [] };
    STATE.rejectedMatches  = result.rejectedMatches || [];
  } catch (err) {
    console.error(err);
    showError('upload-error', `Stitch failed: ${err.message}`);
    return;
  }
  const t1 = performance.now();
  console.log(`Stitch complete in ${(t1-t0).toFixed(0)}ms (${primaryInRange.length} primary × ${secondaryInRange.length} secondary)`);

  // Build the column list now that we know which headers each CSV provides.
  const primaryHeaders = STATE.primary.rows.length ? Object.keys(STATE.primary.rows[0]) : [];
  const secondaryHeaders = STATE.secondary.rows.length ? Object.keys(STATE.secondary.rows[0]) : [];
  STATE.columns = buildColumnList(primaryHeaders, secondaryHeaders);

  // Reveal downstream sections
  for (const id of ['step-dates','step-stats','step-columns','step-preview','step-download']) {
    const el = document.getElementById(id);
    if (el) el.hidden = false;
  }
  renderKpis();
  renderColumnPicker(document.getElementById('column-list'), STATE.columns, rerenderDownstream);
  renderPreviewTable();
  if (typeof renderDateRangesStep === 'function') renderDateRangesStep();
  // The 3 distribution charts (sub-type / parent / course) now live in the Dashboard tab
  // and render on demand via refreshDashboard, so they pick up the active filters.

  // Initialize the Dashboard tab now that we have a stitched dataset.
  // Skip the "fresh data" cue on cache restore — it's a quiet welcome-back, not new arrival.
  try { initDashboard({ showCue: !opts.fromCache }); }
  catch (err) { console.error('Dashboard init failed:', err); }

  refreshResetButton();

  // Auto-scroll only when the user clicked Stitch — restoring from cache
  // shouldn't yank the page. opts.skipScroll lets re-stitch from a date
  // change suppress scroll too even though it's not a cache restore.
  if (!opts.fromCache && !opts.skipScroll) {
    document.getElementById('step-stats').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

/* --- Cache restore + reset orchestration -------------------------------- */

async function restoreFromCacheIfPresent() {
  let cached;
  try { cached = await cacheLoadAll(); }
  catch (e) { console.warn('IndexedDB unavailable; skipping restore.', e); return; }

  if (!cached.primary || !cached.secondary) {
    refreshResetButton();
    return;
  }

  // Re-validate against current schema. If app code has tightened requirements
  // since the cache was written, wipe rather than restore broken state.
  try {
    validateHeaders(cached.primary.rows, CONFIG.inputs.primary.requiredCols,   `Cached ${CONFIG.inputs.primary.label}`);
    validateHeaders(cached.secondary.rows, CONFIG.inputs.secondary.requiredCols, `Cached ${CONFIG.inputs.secondary.label}`);
  } catch (e) {
    console.warn('Cached files no longer meet schema; clearing.', e);
    await cacheClearAll().catch(() => {});
    refreshResetButton();
    return;
  }

  STATE.primary.rows     = cached.primary.rows;
  STATE.primary.fileName = cached.primary.fileName;
  STATE.secondary.rows     = cached.secondary.rows;
  STATE.secondary.fileName = cached.secondary.fileName;
  markZoneRestored('primary', cached.primary);
  markZoneRestored('secondary', cached.secondary);
  refreshStitchButton();
  refreshResetButton();

  // Silently re-stitch so the dashboard is ready the moment the user clicks the tab.
  runStitch({ fromCache: true });
}

function markZoneRestored(target, cached) {
  const zoneEl = document.getElementById('drop-' + target);
  const status = zoneEl.querySelector('.file-status');
  zoneEl.classList.add('loaded', 'from-cache');
  status.innerHTML =
    `<span class="filename">${escapeHtml(cached.fileName)}</span><br>` +
    `<span class="row-count">${cached.rows.length.toLocaleString()} rows</span> ` +
    `<span class="cache-tag">Restored</span> ` +
    `<span class="clear-link" data-clear="${target}">remove</span>`;
}

function refreshResetButton() {
  const btn = document.getElementById('btn-reset');
  if (!btn) return;
  const hasData = !!(STATE.primary.rows || STATE.secondary.rows || STATE.stitched);
  btn.hidden = !hasData;
}

async function resetApp() {
  const hasData = !!(STATE.primary.rows || STATE.secondary.rows || STATE.stitched);
  if (!hasData) return;
  if (!confirm('Reset everything? This clears the cached files, the dashboard, and any stitched data on this device. Column preferences and theme tweaks are kept.')) return;

  await cacheClearAll().catch(e => console.warn('Cache clear failed', e));

  // In-memory state
  STATE.primary.rows = null;     STATE.primary.fileName = null;
  STATE.secondary.rows = null;     STATE.secondary.fileName = null;
  STATE.stitched = null;
  STATE.unmatched = null;
  STATE.methodCounts = null;
  STATE.testRemovedCount = 0;
  STATE.testRemovedRows = { primary: [], secondary: [] };
  STATE.rejectedMatches = null;
  STATE.columns = null;
  clearActiveDateFilters();

  // Drop zones
  for (const target of ['primary', 'secondary']) {
    const zoneEl = document.getElementById('drop-' + target);
    zoneEl.classList.remove('loaded', 'from-cache', 'dragover');
    zoneEl.querySelector('.file-status').innerHTML = '';
    zoneEl.querySelector('input[type="file"]').value = '';
  }

  // Hide all downstream sections
  for (const id of ['step-stats', 'step-columns', 'step-preview', 'step-download']) {
    document.getElementById(id).hidden = true;
  }

  // Tear down dashboard charts and disable the tab
  if (DASH.funnelChart)     { try { DASH.funnelChart.destroy(); }     catch(e){} DASH.funnelChart = null; }
  if (DASH.timeseriesChart) { try { DASH.timeseriesChart.destroy(); } catch(e){} DASH.timeseriesChart = null; }
  for (const k of Object.keys(PAGE_CHARTS)) {
    try { PAGE_CHARTS[k].destroy(); } catch(e) {}
    delete PAGE_CHARTS[k];
  }
  DASH.initialized = false;
  DASH.primaryDataset = null;
  const dashLink = document.getElementById('tab-link-dashboard');
  dashLink.classList.add('disabled');
  dashLink.classList.remove('fresh');
  dashLink.setAttribute('title', 'Upload and stitch reports first');

  if (location.hash === '#dashboard') location.hash = '#configure';
  else setActiveTab('configure');

  clearError('upload-error');

  refreshStitchButton();
  refreshResetButton();
  showToast('Reset complete. Upload your reports to start over.');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Paint the active config's text into static DOM nodes. Called once at init
// before listeners are wired so labels are correct on first frame.
//
// Three painting modes:
//   1. document.title and the .intro-card subtree (eyebrow/h1/sub) are set
//      directly from CONFIG.label / CONFIG.intro.
//   2. Drop-zone strong texts read CONFIG.inputs.primary.label / secondary.label.
//   3. Any element with [data-label="<key>"] gets its textContent set from
//      CONFIG.dashboard.labels[key]; [data-label-title="<key>"] sets the title
//      attribute. The HTML keeps a fallback default value so the page is
//      readable if JS doesn't run.
function applyConfigToDom() {
  if (!CONFIG) return;

  document.title = `Report Stitcher — ${CONFIG.label}`;

  if (CONFIG.intro) {
    const eyebrow = document.querySelector('.intro-card .intro-eyebrow');
    const title   = document.querySelector('.intro-card h1');
    const sub     = document.querySelector('.intro-card .intro-sub');
    if (eyebrow && CONFIG.intro.eyebrow != null) eyebrow.textContent = CONFIG.intro.eyebrow;
    if (title   && CONFIG.intro.title   != null) title.innerHTML    = CONFIG.intro.title;
    if (sub     && CONFIG.intro.sub     != null) sub.textContent    = CONFIG.intro.sub;
  }

  if (CONFIG.inputs) {
    const p = document.querySelector('#drop-primary .drop-content strong');
    const s = document.querySelector('#drop-secondary .drop-content strong');
    if (p && CONFIG.inputs.primary   && CONFIG.inputs.primary.label)   p.textContent = CONFIG.inputs.primary.label;
    if (s && CONFIG.inputs.secondary && CONFIG.inputs.secondary.label) s.textContent = CONFIG.inputs.secondary.label;
  }

  const labels = CONFIG.dashboard && CONFIG.dashboard.labels;
  if (labels) {
    document.querySelectorAll('[data-label]').forEach(el => {
      const v = labels[el.dataset.label];
      if (v != null) el.textContent = v;
    });
    document.querySelectorAll('[data-label-title]').forEach(el => {
      const v = labels[el.dataset.labelTitle];
      if (v != null) el.title = v;
    });
  }
}

// Client-code gate. Shown when CONFIG is null (no code in localStorage, or
// the saved code doesn't match any registered config's clientCodes) AND when
// the user clicks "change" from the app-strip. Submit path validates, saves
// to localStorage, and reloads — the resolver in configs/index.js picks up
// the new code on the fresh page load.
//
// Re-entrant safe: the submit handler is assigned via form.onsubmit (single
// property) rather than addEventListener so re-invoking showClientGate from
// the "change" link doesn't stack handlers.
function showClientGate() {
  const dlg     = document.getElementById('client-gate-dialog');
  const form    = document.getElementById('client-gate-form');
  const input   = document.getElementById('client-gate-input');
  const errorEl = document.getElementById('client-gate-error');

  if (!dlg || !form || !input || !errorEl) {
    // No dialog in DOM — fall back to a native prompt loop so the gate still works.
    let code = '';
    while (true) {
      code = (prompt('Enter your client code:') || '').trim().toLowerCase();
      if (!code) return;
      if (window.getConfigsForClient(code).length > 0) break;
      alert(`No configurations available for "${code}".`);
    }
    localStorage.setItem(window.CLIENT_KEY, code);
    localStorage.removeItem(window.CONFIG_KEY);
    location.reload();
    return;
  }

  // Pre-fill with the saved code (if any) so the user sees what they had
  // before — e.g. they're returning after the registry filter changed, or
  // they clicked "change" to swap clients.
  const existing = localStorage.getItem(window.CLIENT_KEY);
  if (existing) input.value = existing;

  errorEl.hidden = true;
  errorEl.textContent = '';

  form.onsubmit = (e) => {
    e.preventDefault();
    const code = input.value.trim().toLowerCase();
    if (!code) {
      errorEl.textContent = 'Please enter a code.';
      errorEl.hidden = false;
      return;
    }
    const matches = window.getConfigsForClient(code);
    if (matches.length === 0) {
      errorEl.textContent = `No configurations available for "${code}". Check the code with whoever invited you.`;
      errorEl.hidden = false;
      input.select();
      return;
    }
    localStorage.setItem(window.CLIENT_KEY, code);
    localStorage.removeItem(window.CONFIG_KEY);   // reset to first available
    location.reload();
  };

  if (typeof dlg.showModal === 'function') dlg.showModal();
  else dlg.setAttribute('open', '');
  setTimeout(() => { input.focus(); input.select(); }, 0);
}

// Paint the app-strip's client display + config chooser + "change" link.
// Selecting a different config or clicking "change" reloads so the rest of
// the engine can re-init against the new config without partial-state risk.
function wireClientStrip() {
  const strip      = document.getElementById('client-strip');
  const display    = document.getElementById('client-code-display');
  const select     = document.getElementById('config-select');
  const sep        = document.getElementById('config-select-sep');
  const changeLink = document.getElementById('change-client-link');
  if (!strip || !display || !select || !changeLink) return;

  const code = window.getClientCode();
  display.textContent = code;

  const available = window.getConfigsForClient(code);
  select.innerHTML = '';
  for (const c of available) {
    const opt = document.createElement('option');
    opt.value = c.id;
    opt.textContent = c.label;
    if (CONFIG && c.id === CONFIG.id) opt.selected = true;
    select.appendChild(opt);
  }
  const hasChooser = available.length > 1;
  select.hidden = !hasChooser;
  if (sep) sep.hidden = !hasChooser;

  select.addEventListener('change', () => {
    localStorage.setItem(window.CONFIG_KEY, select.value);
    location.reload();
  });

  // The "change" link opens the gate dialog in place rather than the old
  // clear-storage-and-reload cycle. That cycle was a no-op while only one
  // config is registered (the resolver in configs/index.js auto-fills the
  // lone config's code on reload), so the user never actually saw the gate.
  // Going through showClientGate() lets the user type a different code at
  // any time, which matters during development with a single-config app.
  changeLink.addEventListener('click', (e) => {
    e.preventDefault();
    showClientGate();
  });

  strip.hidden = false;
}

async function init() {
  // STATE.columns is built after stitch (we need the actual CSV headers to
  // include all source columns, not just the documented defaults).

  // No active config → show the client-code gate and bail out. The submit
  // handler reloads, so the rest of init runs against a resolved CONFIG on
  // the next page load.
  if (!CONFIG) {
    showClientGate();
    return;
  }

  applyConfigToDom();
  wireClientStrip();

  setupDropZone(document.getElementById('drop-primary'), 'primary', CONFIG.inputs.primary.requiredCols,   CONFIG.inputs.primary.label);
  setupDropZone(document.getElementById('drop-secondary'), 'secondary', CONFIG.inputs.secondary.requiredCols, CONFIG.inputs.secondary.label);

  // Wrap so the click event isn't passed in as the opts arg.
  document.getElementById('btn-stitch').addEventListener('click', () => runStitch());

  document.getElementById('btn-reset').addEventListener('click', resetApp);

  document.getElementById('btn-reset-cols').addEventListener('click', () => {
    const primaryHeaders = STATE.primary.rows ? Object.keys(STATE.primary.rows[0]) : [];
    const secondaryHeaders = STATE.secondary.rows ? Object.keys(STATE.secondary.rows[0]) : [];
    STATE.columns = buildDefaultColumnList(primaryHeaders, secondaryHeaders);
    saveColumnConfig(STATE.columns);
    renderColumnPicker(document.getElementById('column-list'), STATE.columns, rerenderDownstream);
    renderPreviewTable();
  });
  document.getElementById('btn-uncheck-cols').addEventListener('click', () => {
    STATE.columns.forEach(c => c.enabled = false);
    saveColumnConfig(STATE.columns);
    renderColumnPicker(document.getElementById('column-list'), STATE.columns, rerenderDownstream);
    renderPreviewTable();
  });
  document.getElementById('btn-check-cols').addEventListener('click', () => {
    STATE.columns.forEach(c => c.enabled = true);
    saveColumnConfig(STATE.columns);
    renderColumnPicker(document.getElementById('column-list'), STATE.columns, rerenderDownstream);
    renderPreviewTable();
  });

  document.getElementById('btn-xlsx').addEventListener('click', async () => {
    const btn = document.getElementById('btn-xlsx');
    const status = document.getElementById('xlsx-status');
    btn.disabled = true;
    status.innerHTML = '<span class="spinner" style="border-color:rgba(31,56,100,.2);border-top-color:#1F3864;"></span>Building xlsx…';
    try {
      await generateXlsx();
      status.textContent = 'Done.';
      setTimeout(() => { status.textContent = ''; }, 2500);
    } catch (err) {
      console.error(err);
      status.textContent = '';
      showToast('xlsx generation failed: ' + err.message, true);
    } finally {
      btn.disabled = false;
    }
  });

  // Filter-bar "Generate xlsx (current filter)" — same lifecycle as btn-xlsx,
  // but routes through generateFilteredXlsx for the cohort-narrowing path.
  // renderFilterSummary keeps the disabled state in sync with filter results.
  const btnFilteredXlsx = document.getElementById('btn-filtered-xlsx');
  if (btnFilteredXlsx) {
    btnFilteredXlsx.addEventListener('click', async () => {
      const status = document.getElementById('filtered-xlsx-status');
      const wasDisabled = btnFilteredXlsx.disabled;
      btnFilteredXlsx.disabled = true;
      if (status) status.innerHTML = '<span class="spinner" style="border-color:rgba(31,56,100,.2);border-top-color:#1F3864;"></span>Building xlsx…';
      try {
        await generateFilteredXlsx();
        if (status) status.textContent = 'Done.';
        setTimeout(() => { if (status) status.textContent = ''; }, 2500);
      } catch (err) {
        console.error(err);
        if (status) status.textContent = '';
        showToast('xlsx generation failed: ' + err.message, true);
      } finally {
        // Restore prior disabled state (if filter result is empty, renderFilterSummary
        // will re-disable it on the next refresh; otherwise leave it enabled).
        btnFilteredXlsx.disabled = wasDisabled;
      }
    });
  }

  // Tab routing
  document.querySelectorAll('.tab-link').forEach(link => {
    link.addEventListener('click', (e) => {
      if (link.classList.contains('disabled')) { e.preventDefault(); return; }
    });
  });
  window.addEventListener('hashchange', () => {
    setActiveTab((location.hash.replace('#','') || 'configure'));
  });
  setupDrilldownHandlers();

  // Restore cached CSVs before routing so a deep-link to #dashboard lands
  // correctly when there's a previous session waiting.
  await restoreFromCacheIfPresent();

  setActiveTab(location.hash.replace('#','') || 'configure');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
