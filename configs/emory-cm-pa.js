'use strict';

/* ============================================================
   Emory: Campaign Member ↔ Participant
   ============================================================
   First built-in StitchConfig — reproduces the original Emory-app behavior on
   the new config-driven engine. Sub-Type bucket data, course-similarity Jaccard
   helpers, test-row filter, and CM Course-Status derivation all live in this
   module's closure. The engine in app.js stays generic.
*/

(function() {

  // ===== Emory-specific data tables =====

  const SUBTYPE_BUCKET = {
    'Course Landing Page'    : 'Website',
    'Website RFI'            : 'Website',
    'Facebook Lead Form'     : 'Social',
    'Instagram'              : 'Social',
    'LinkedIn Lead Form'     : 'Social',
    'Instagram Sponsored Ad' : 'Social',
  };
  // Locked Sub-Type Bucket display order (Unknown is appended last when present).
  // Drives both the xlsx Summary parent rows and the dashboard summary table.
  const BUCKET_ORDER = ['Website', 'Social'];

  // Stopwords stripped before Jaccard-tokenizing course names. Tweak with care:
  // these directly affect course_score and the tiebreaker's choice.
  const COURSE_STOPWORDS = new Set([
    'the','a','an','of','for','and','to','with','in','on',
    '&','-','program','certificate','course'
  ]);

  // ===== Emory-specific helpers =====
  // (These were top-level functions in the original app.js — now closure-private
  // to this config module so the engine has no implicit Emory knowledge.)

  function bucketSubType(subtype) {
    const s = (subtype || '').toString().trim();
    if (!s) return 'Unknown';
    return SUBTYPE_BUCKET[s] || 'Unknown';
  }

  const TOKEN_RE = /[A-Za-z0-9]+/g;
  function courseTokens(s) {
    if (!s) return new Set();
    const matches = String(s).toLowerCase().match(TOKEN_RE) || [];
    return new Set(matches.filter(t => !COURSE_STOPWORDS.has(t)));
  }
  function courseSimilarity(a, b) {
    if (!a || !b) return 0.0;
    if (String(a).trim().toLowerCase() === String(b).trim().toLowerCase()) return 1.0;
    const ta = courseTokens(a);
    const tb = courseTokens(b);
    if (ta.size === 0 || tb.size === 0) return 0.0;
    let inter = 0;
    for (const t of ta) if (tb.has(t)) inter++;
    const union = ta.size + tb.size - inter;
    return union === 0 ? 0.0 : inter / union;
  }

  const TEST_RE = /(?:^|\s)test(?:\s|$)/i;
  function isTestRow(row) {
    const first = (row['First Name'] || '').toString().trim();
    const last  = (row['Last Name']  || '').toString().trim();
    if (first.toLowerCase() === 'test' || last.toLowerCase() === 'test') return true;
    if (TEST_RE.test(first) || TEST_RE.test(last)) return true;
    return false;
  }

  // Maps a CM's set of stitched matches to one of the four canonical Course-Status
  // strings. These literal strings are referenced by xlsx COUNTIFS formulas in the
  // Summary sheets — see plan locked decision #3. Do not rename without also
  // updating every COUNTIFS criterion that mentions them.
  function deriveCourseStatus(matches) {
    if (!matches || matches.length === 0) return 'Not Converted';
    const statuses = matches.map(s => (s.secondary['Status'] || '').trim());
    if (statuses.includes('Enrolled'))   return 'Enrolled';
    if (statuses.includes('Registered')) return 'Registered';
    return 'Cancelled/Withdrawn/Etc';
  }

  // ===== xlsx builders =====
  // The four sheet builders below are registered on the config's `outputSheets`
  // array (see bottom of file). The engine's generateXlsx() iterates that array
  // and calls each builder with (wb, ctx); ctx.sheets[sheetName] threads each
  // builder's return value to downstream builders (Participant Summary reads
  // Stitched Data's enabledCols/dataLastRow; Campaign Summary reads Campaign
  // Members' table name + last row).
  //
  // Builders pull engine-side primitives (style constants, generic xlsx helpers,
  // STATE, aggregateAll, renderChartPng) from `window.RS`, which app.js sets up
  // AFTER configs/*.js have registered. We can't destructure at IIFE time —
  // each builder calls rs() on entry to get the live bindings.

  function rs() {
    if (!window.RS) throw new Error('Report Stitcher: window.RS is not initialized — did app.js load?');
    return window.RS;
  }

  // Used only by buildCampaignSummarySheet — computes the (parent → campaign)
  // row spine for the pivot. Cell counts themselves are COUNTIFS formulas
  // against the CampaignMembers table, so this aggregation only needs to
  // identify *which* parent/campaign rows to emit.
  function aggregateCmStatusByCampaign(ctx) {
    const { STATE, dateInRange } = rs();
    const dateOf = ctx && ctx.dateOf;
    const focus  = ctx && ctx.focus;
    // Scope to CMs ∈ Focus so the Campaign Summary's row spine matches the
    // CampaignMembers (Focus) table its COUNTIFS formulas reference. Without
    // this filter, older CMs (in Influence but outside Focus) would appear as
    // (parent, campaign) rows with all-zero counts.
    const cmClean = STATE.primary.rows.filter(r =>
      !isTestRow(r) && (!dateOf || dateInRange(dateOf.primary(r), focus))
    );
    const primaryToSecondaries = new Map();
    for (const s of STATE.stitched) {
      if (!primaryToSecondaries.has(s.primary)) primaryToSecondaries.set(s.primary, []);
      primaryToSecondaries.get(s.primary).push(s.secondary);
    }
    const out = new Map();
    for (const cm of cmClean) {
      const matched = primaryToSecondaries.get(cm) || [];
      let bucket;
      if (matched.length === 0) {
        bucket = 'notConverted';
      } else {
        const statuses = matched.map(p => (p['Status'] || '').trim());
        if (statuses.includes('Enrolled'))        bucket = 'enrolled';
        else if (statuses.includes('Registered')) bucket = 'registered';
        else                                       bucket = 'cancelled';
      }
      const parent   = (cm['Parent Campaign Name'] || '').trim() || '(blank)';
      const campaign = (cm['Campaign Name']        || '').trim() || '(blank)';
      if (!out.has(parent)) out.set(parent, new Map());
      const inner = out.get(parent);
      if (!inner.has(campaign)) inner.set(campaign, { notConverted:0, cancelled:0, enrolled:0, registered:0 });
      inner.get(campaign)[bucket]++;
    }
    return out;
  }

  // ----- Sheet 1: Stitched Data ---------------------------------------------
  function buildStitchedSheet(wb, ctx) {
    const { STATE, FONT_HEADER, FONT_BODY, FILL_NAVY, BORDER_THIN, NAVY_ARGB, getCellValue, parseSfDate } = rs();
    const ws = wb.addWorksheet('Stitched Data', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.properties.tabColor = { argb: NAVY_ARGB };
    const enabled = STATE.columns.filter(c => c.enabled);
    if (enabled.length === 0) throw new Error('Column picker: at least one column must be selected.');

    // Columns flagged `type: 'date'` in CONFIG.defaultColumns get their raw
    // Salesforce date strings parsed into Date objects so Excel writes them
    // as real date cells (sortable, filterable) rather than text.
    const tableColumns = enabled.map(c => ({ name: c.label, filterButton: true }));
    const tableRows = STATE.stitched.map(r => enabled.map(col => {
      const v = getCellValue(r, col);
      if (col.key === 'course_score') return Number(v);
      if (col.type === 'date') {
        if (v == null || v === '') return null;
        if (v instanceof Date) return v;
        const d = parseSfDate(v);
        return d || v;  // fall back to raw string if unparseable
      }
      return v == null ? '' : v;
    }));
    // ExcelJS requires at least one row in addTable; pad with empty if no data.
    if (tableRows.length === 0) tableRows.push(enabled.map(() => ''));

    ws.addTable({
      name: 'StitchedData',
      ref: 'A1',
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: tableColumns,
      rows: tableRows,
    });

    // Override header styling to navy + white Arial 11 bold (the table style sets a different theme color).
    const headerRow = ws.getRow(1);
    headerRow.height = 22;
    headerRow.eachCell((cell, colNum) => {
      if (colNum > enabled.length) return;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: false };
      cell.border = BORDER_THIN;
    });

    // Body styling + column widths. Date columns get a fixed width (~12) and
    // an m/d/yyyy display format so Excel sorts them as dates; other columns
    // autosize against the longest stringified value.
    for (let c = 1; c <= enabled.length; c++) {
      const col = enabled[c-1];
      if (col.type === 'date') {
        ws.getColumn(c).width  = 12;
        ws.getColumn(c).numFmt = 'm/d/yyyy';
        continue;
      }
      let maxLen = col.label.length;
      for (let r = 0; r < tableRows.length; r++) {
        const v = tableRows[r][c-1];
        const s = v == null ? '' : String(v);
        if (s.length > maxLen) maxLen = s.length;
      }
      ws.getColumn(c).width = Math.min(Math.max(maxLen + 2, 10), 45);
    }
    for (let r = 2; r <= tableRows.length + 1; r++) {
      const row = ws.getRow(r);
      row.eachCell((cell, colNum) => {
        if (colNum > enabled.length) return;
        cell.font = FONT_BODY;
        cell.alignment = { vertical: 'middle' };
      });
    }

    return { ws, enabledCols: enabled, dataLastRow: tableRows.length + 1 };
  }

  // ----- Sheet 2: Participant Summary ---------------------------------------
  // Builds the Sub-Type Bucket / Sub-Type table (left), Parent Campaign +
  // Course tables (right), then embeds PNG renders of the three distribution
  // charts at anchor rows derived from the table sizes. Charts are embedded
  // inline (rather than as a separate post-pass) because the anchor rows are
  // local knowledge.
  async function buildParticipantSummarySheet(wb, ctx) {
    const {
      STATE,
      FONT_HEADER, FONT_BODY, FILL_NAVY, FILL_LBLUE, FILL_GRAY,
      BORDER_THIN, BORDER_MED_BOTTOM,
      NAVY_ARGB, WHITE_ARGB,
      colLetter, applyBorderRange, findEnabledColIndex,
      escapeFormula, aggregateAll, renderChartPng,
    } = rs();

    const stitchedCtx = ctx.sheets['Stitched Data'];
    if (!stitchedCtx) throw new Error('Participant Summary builder ran before Stitched Data.');
    const { enabledCols, dataLastRow } = stitchedCtx;

    const ws = wb.addWorksheet('Participant Summary', { views: [{ showGridLines: false }] });
    ws.properties.tabColor = { argb: NAVY_ARGB };

    const dataSheet = "'Stitched Data'";
    const lastRow = dataLastRow;

    const colStatus  = findEnabledColIndex(enabledCols, 'status');
    const colSubBkt  = findEnabledColIndex(enabledCols, 'subtype_bucket');
    const colSub     = findEnabledColIndex(enabledCols, 'subtype');
    const colParent  = findEnabledColIndex(enabledCols, 'parent_campaign');
    const colCourse  = findEnabledColIndex(enabledCols, 'course_registered');

    // Helper: sheet-qualified absolute column reference like 'Stitched Data'!$C$2:$C$N
    const colRef = (c) => c == null ? null : `${dataSheet}!$${colLetter(c)}$2:$${colLetter(c)}$${lastRow}`;

    const ref_status  = colRef(colStatus);
    const ref_subBkt  = colRef(colSubBkt);
    const ref_sub     = colRef(colSub);
    const ref_parent  = colRef(colParent);
    const ref_course  = colRef(colCourse);

    // Conditional "Cancelled/Other" column — shown only if the stitched
    // secondary set contains any Status values besides Registered or Enrolled.
    // Counted via COUNTIFS (<>Registered, <>Enrolled, <>blank) so the bucket
    // is computed in-formula rather than baked into the column source.
    const hasOther = STATE.stitched.some(s => {
      const st = (s.secondary['Status'] || '').toString().trim();
      return st && st !== 'Registered' && st !== 'Enrolled';
    });
    const otherClause = ref_status
      ? `${ref_status},"<>Registered",${ref_status},"<>Enrolled",${ref_status},"<>"`
      : null;

    // Default column widths. E + J hold Cancelled/Other (hidden below when
    // !hasOther); F + K hold Total. Gap column F was removed in favor of the
    // extra Total column — the wider G (38) provides enough visual separation
    // from the sub-Type Total at F.
    ws.getColumn(1).width = 28;
    ws.getColumn(2).width = 30;
    ws.getColumn(3).width = 13;
    ws.getColumn(4).width = 13;
    ws.getColumn(5).width = 15;
    ws.getColumn(6).width = 13;
    ws.getColumn(7).width = 38;
    ws.getColumn(8).width = 13;
    ws.getColumn(9).width = 13;
    ws.getColumn(10).width = 15;
    ws.getColumn(11).width = 13;

    // ===== Title (A1:E1) =====
    ws.mergeCells('A1:E1');
    const titleCell = ws.getCell('A1');
    titleCell.value = 'Stitched Report — Summary';
    titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: NAVY_ARGB } };
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 26;

    // ===== Note row (A2:E2) — italic gray scope explanation =====
    // Cap at column E to avoid overlapping the Parent Campaign table on the
    // right (its header sits in G2). Row 2 height gets inflated to fit the
    // wrapped note, which also stretches G2 — that's an accepted trade-off.
    const noteText = (CONFIG.dashboard && CONFIG.dashboard.labels && CONFIG.dashboard.labels.noteParticipantSummary) || '';
    if (noteText) {
      ws.mergeCells('A2:E2');
      const noteCell = ws.getCell('A2');
      noteCell.value = noteText;
      noteCell.font  = { name: 'Arial', size: 9.5, italic: true, color: { argb: 'FF6B7280' } };
      noteCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      ws.getRow(2).height = 56;
    }

    // ===== KPI block (A3:B5) =====
    const kpiNavyCell = (row, label, formulaOrValue, isFormula) => {
      const a = ws.getCell(`A${row}`);
      a.value = label;
      a.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE_ARGB } };
      a.fill = FILL_NAVY;
      a.alignment = { horizontal: 'right', vertical: 'middle' };
      const b = ws.getCell(`B${row}`);
      b.value = isFormula ? { formula: formulaOrValue } : formulaOrValue;
      b.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE_ARGB } };
      b.fill = FILL_NAVY;
      b.alignment = { horizontal: 'center', vertical: 'middle' };
      b.numFmt = '#,##0';
    };
    const kpiPlainCell = (row, label, formula) => {
      const a = ws.getCell(`A${row}`);
      a.value = label;
      a.font = { name: 'Arial', size: 11, color: { argb: NAVY_ARGB }, bold: true };
      a.alignment = { horizontal: 'right', vertical: 'middle' };
      const b = ws.getCell(`B${row}`);
      b.value = { formula };
      b.font = { name: 'Arial', size: 11, color: { argb: NAVY_ARGB }, bold: true };
      b.alignment = { horizontal: 'center', vertical: 'middle' };
      b.numFmt = '#,##0';
    };

    // Total Influenced = COUNTA over the Status column (counts every stitched row).
    if (ref_status) {
      kpiNavyCell(3, 'Total Influenced', `COUNTA(${ref_status})`, true);
      kpiPlainCell(4, 'Total Registered', `COUNTIF(${ref_status},"Registered")`);
      kpiPlainCell(5, 'Total Enrolled',   `COUNTIF(${ref_status},"Enrolled")`);
    } else {
      kpiNavyCell(3, 'Total Influenced', STATE.stitched.length, false);
      kpiPlainCell(4, 'Total Registered', STATE.stitched.filter(r => r.secondary['Status']==='Registered').length.toString());
      kpiPlainCell(5, 'Total Enrolled',   STATE.stitched.filter(r => r.secondary['Status']==='Enrolled').length.toString());
    }
    // Bottom medium border on row 5
    ['A5','B5'].forEach(addr => {
      const c = ws.getCell(addr);
      c.border = { ...c.border, bottom: { style: 'medium', color: { argb: 'FF000000' } } };
    });

    // ===== Sub-Type Bucket / Sub-Type table (A7 onwards) =====
    const agg = aggregateAll(STATE.stitched);
    let r = 7;
    ws.mergeCells(`A${r}:F${r}`);
    const sectionTitle = ws.getCell(`A${r}`);
    sectionTitle.value = 'Sub-Type Bucket / Sub-Type';
    sectionTitle.font = { name: 'Arial', size: 12, bold: true, color: { argb: NAVY_ARGB } };
    sectionTitle.fill = FILL_GRAY;
    sectionTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorderRange(ws, `A${r}:F${r}`, BORDER_THIN);
    r++;

    // Header row
    const subHeaders = ['Sub-Type Bucket', 'Sub-Type', 'Registered', 'Enrolled', 'Cancelled/Other', 'Total'];
    subHeaders.forEach((label, idx) => {
      const cell = ws.getCell(`${colLetter(idx+1)}${r}`);
      cell.value = label;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = BORDER_THIN;
    });
    ws.getRow(r).height = 20;
    r++;

    // Bucket parent rows + detail rows
    for (const bkt of Object.keys(agg.subtypeBuckets)) {
      const details = agg.subtypeBuckets[bkt];
      if (details.length === 0) continue;

      const aCell = ws.getCell(`A${r}`);
      aCell.value = bkt;
      aCell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      aCell.fill = FILL_LBLUE;
      aCell.alignment = { horizontal: 'left', vertical: 'middle' };
      aCell.border = BORDER_THIN;
      const bCell = ws.getCell(`B${r}`);
      bCell.value = '';
      bCell.fill = FILL_LBLUE;
      bCell.border = BORDER_THIN;
      if (ref_subBkt && ref_status) {
        ws.getCell(`C${r}`).value = { formula: `COUNTIFS(${ref_subBkt},$A${r},${ref_status},"Registered")` };
        ws.getCell(`D${r}`).value = { formula: `COUNTIFS(${ref_subBkt},$A${r},${ref_status},"Enrolled")` };
        ws.getCell(`E${r}`).value = { formula: `COUNTIFS(${ref_subBkt},$A${r},${otherClause})` };
      } else {
        ws.getCell(`C${r}`).value = details.reduce((a,b)=>a+b.reg,0);
        ws.getCell(`D${r}`).value = details.reduce((a,b)=>a+b.enr,0);
        ws.getCell(`E${r}`).value = 0;
      }
      ws.getCell(`F${r}`).value = { formula: `C${r}+D${r}+E${r}` };
      ['C','D','E','F'].forEach(c => {
        const cell = ws.getCell(`${c}${r}`);
        cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
        cell.fill = FILL_LBLUE;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER_THIN;
        cell.numFmt = '#,##0';
      });
      r++;

      for (const d of details) {
        ws.getCell(`A${r}`).value = '';
        ws.getCell(`A${r}`).border = BORDER_THIN;
        const bCell2 = ws.getCell(`B${r}`);
        bCell2.value = '  ' + d.name;     // two-space indent matches stitch.py style
        bCell2.font = FONT_BODY;
        bCell2.alignment = { horizontal: 'left', vertical: 'middle' };
        bCell2.border = BORDER_THIN;
        if (ref_sub && ref_status) {
          ws.getCell(`C${r}`).value = { formula: `COUNTIFS(${ref_sub},"${escapeFormula(d.name)}",${ref_status},"Registered")` };
          ws.getCell(`D${r}`).value = { formula: `COUNTIFS(${ref_sub},"${escapeFormula(d.name)}",${ref_status},"Enrolled")` };
          ws.getCell(`E${r}`).value = { formula: `COUNTIFS(${ref_sub},"${escapeFormula(d.name)}",${otherClause})` };
        } else {
          ws.getCell(`C${r}`).value = d.reg;
          ws.getCell(`D${r}`).value = d.enr;
          ws.getCell(`E${r}`).value = 0;
        }
        ws.getCell(`F${r}`).value = { formula: `C${r}+D${r}+E${r}` };
        ['C','D','E','F'].forEach(c => {
          const cell = ws.getCell(`${c}${r}`);
          cell.font = FONT_BODY;
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = BORDER_THIN;
          cell.numFmt = '#,##0';
        });
        r++;
      }
    }

    // Grand Total row for the sub-type table
    const aTotal = ws.getCell(`A${r}`);
    aTotal.value = 'Grand Total';
    aTotal.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
    aTotal.fill = FILL_GRAY;
    aTotal.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getCell(`B${r}`).value = '';
    ws.getCell(`B${r}`).fill = FILL_GRAY;
    ws.getCell(`C${r}`).value = { formula: 'B4' };
    ws.getCell(`D${r}`).value = { formula: 'B5' };
    ws.getCell(`E${r}`).value = ref_status
      ? { formula: `COUNTIFS(${otherClause})` }
      : 0;
    ws.getCell(`F${r}`).value = { formula: `C${r}+D${r}+E${r}` };
    ['A','B','C','D','E','F'].forEach(c => {
      const cell = ws.getCell(`${c}${r}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: c === 'A' || c === 'B' ? 'left' : 'center', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
      if (c === 'C' || c === 'D' || c === 'E' || c === 'F') cell.numFmt = '#,##0';
    });
    const subTableLastRow = r;
    r++;

    // ===== Parent Campaign table (G1:K?) =====
    let pr = 1;
    ws.mergeCells(`G${pr}:K${pr}`);
    const pcTitle = ws.getCell(`G${pr}`);
    pcTitle.value = 'Registrations by Parent Campaign';
    pcTitle.font = { name: 'Arial', size: 12, bold: true, color: { argb: NAVY_ARGB } };
    pcTitle.fill = FILL_GRAY;
    pcTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorderRange(ws, `G${pr}:K${pr}`, BORDER_THIN);
    pr++;

    ['Parent Campaign','Registered','Enrolled','Cancelled/Other','Total'].forEach((label, idx) => {
      const cell = ws.getCell(`${colLetter(7+idx)}${pr}`);
      cell.value = label;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = BORDER_THIN;
    });
    pr++;

    const parentDetailRows = [];
    for (const row of agg.parent) {
      parentDetailRows.push(pr);
      const gCell = ws.getCell(`G${pr}`);
      gCell.value = row.name;
      gCell.font = FONT_BODY;
      gCell.alignment = { horizontal: 'left', vertical: 'middle' };
      gCell.border = BORDER_THIN;
      if (ref_parent && ref_status) {
        ws.getCell(`H${pr}`).value = { formula: `COUNTIFS(${ref_parent},"${escapeFormula(row.name)}",${ref_status},"Registered")` };
        ws.getCell(`I${pr}`).value = { formula: `COUNTIFS(${ref_parent},"${escapeFormula(row.name)}",${ref_status},"Enrolled")` };
        ws.getCell(`J${pr}`).value = { formula: `COUNTIFS(${ref_parent},"${escapeFormula(row.name)}",${otherClause})` };
      } else {
        ws.getCell(`H${pr}`).value = row.reg;
        ws.getCell(`I${pr}`).value = row.enr;
        ws.getCell(`J${pr}`).value = 0;
      }
      ws.getCell(`K${pr}`).value = { formula: `H${pr}+I${pr}+J${pr}` };
      ['H','I','J','K'].forEach(c => {
        const cell = ws.getCell(`${c}${pr}`);
        cell.font = FONT_BODY;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER_THIN;
        cell.numFmt = '#,##0';
      });
      pr++;
    }

    // Grand Total row
    const pcTotalRow = pr;
    ws.getCell(`G${pr}`).value = 'Grand Total';
    ws.getCell(`H${pr}`).value = { formula: parentDetailRows.length ? `SUM(H${parentDetailRows[0]}:H${parentDetailRows[parentDetailRows.length-1]})` : '0' };
    ws.getCell(`I${pr}`).value = { formula: parentDetailRows.length ? `SUM(I${parentDetailRows[0]}:I${parentDetailRows[parentDetailRows.length-1]})` : '0' };
    ws.getCell(`J${pr}`).value = { formula: parentDetailRows.length ? `SUM(J${parentDetailRows[0]}:J${parentDetailRows[parentDetailRows.length-1]})` : '0' };
    ws.getCell(`K${pr}`).value = { formula: `H${pr}+I${pr}+J${pr}` };
    ['G','H','I','J','K'].forEach(c => {
      const cell = ws.getCell(`${c}${pr}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: c === 'G' ? 'left' : 'center', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
      if (c !== 'G') cell.numFmt = '#,##0';
    });
    pr++;
    pr++;   // blank gap before Course table

    // ===== Course table (G{pr}:K?) =====
    const courseTitleRow = pr;
    ws.mergeCells(`G${pr}:K${pr}`);
    const cTitle = ws.getCell(`G${pr}`);
    cTitle.value = 'Registrations by Course';
    cTitle.font = { name: 'Arial', size: 12, bold: true, color: { argb: NAVY_ARGB } };
    cTitle.fill = FILL_GRAY;
    cTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorderRange(ws, `G${pr}:K${pr}`, BORDER_THIN);
    pr++;

    ['Course','Registered','Enrolled','Cancelled/Other','Total'].forEach((label, idx) => {
      const cell = ws.getCell(`${colLetter(7+idx)}${pr}`);
      cell.value = label;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = BORDER_THIN;
    });
    pr++;

    const courseDetailRows = [];
    for (const row of agg.course) {
      courseDetailRows.push(pr);
      const gCell = ws.getCell(`G${pr}`);
      gCell.value = row.name;
      gCell.font = FONT_BODY;
      gCell.alignment = { horizontal: 'left', vertical: 'middle' };
      gCell.border = BORDER_THIN;
      if (ref_course && ref_status) {
        ws.getCell(`H${pr}`).value = { formula: `COUNTIFS(${ref_course},"${escapeFormula(row.name)}",${ref_status},"Registered")` };
        ws.getCell(`I${pr}`).value = { formula: `COUNTIFS(${ref_course},"${escapeFormula(row.name)}",${ref_status},"Enrolled")` };
        ws.getCell(`J${pr}`).value = { formula: `COUNTIFS(${ref_course},"${escapeFormula(row.name)}",${otherClause})` };
      } else {
        ws.getCell(`H${pr}`).value = row.reg;
        ws.getCell(`I${pr}`).value = row.enr;
        ws.getCell(`J${pr}`).value = 0;
      }
      ws.getCell(`K${pr}`).value = { formula: `H${pr}+I${pr}+J${pr}` };
      ['H','I','J','K'].forEach(c => {
        const cell = ws.getCell(`${c}${pr}`);
        cell.font = FONT_BODY;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = BORDER_THIN;
        cell.numFmt = '#,##0';
      });
      pr++;
    }

    // Course grand total row
    ws.getCell(`G${pr}`).value = 'Grand Total';
    ws.getCell(`H${pr}`).value = { formula: courseDetailRows.length ? `SUM(H${courseDetailRows[0]}:H${courseDetailRows[courseDetailRows.length-1]})` : '0' };
    ws.getCell(`I${pr}`).value = { formula: courseDetailRows.length ? `SUM(I${courseDetailRows[0]}:I${courseDetailRows[courseDetailRows.length-1]})` : '0' };
    ws.getCell(`J${pr}`).value = { formula: courseDetailRows.length ? `SUM(J${courseDetailRows[0]}:J${courseDetailRows[courseDetailRows.length-1]})` : '0' };
    ws.getCell(`K${pr}`).value = { formula: `H${pr}+I${pr}+J${pr}` };
    ['G','H','I','J','K'].forEach(c => {
      const cell = ws.getCell(`${c}${pr}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: c === 'G' ? 'left' : 'center', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
      if (c !== 'G') cell.numFmt = '#,##0';
    });

    // Hide the Cancelled/Other columns (E for sub-Type table, J for Parent
    // + Course tables) when the secondary data has no non-Reg/Enr statuses.
    // Data + formulas remain in the cells; only the visible column is hidden.
    if (!hasOther) {
      ws.getColumn(5).hidden  = true;
      ws.getColumn(10).hidden = true;
    }

    // Chart anchor rows. Sub-Type chart sits below its table (left side);
    // Parent + Course charts sit to the RIGHT of their tables (parent table
    // now spans G-K with the Cancelled/Other column, so the chart goes at
    // column M = 0-based 12, leaving column L as a one-column visual buffer).
    // Vertical anchors match each table's title row.
    const subtypeChartAnchorRow = subTableLastRow + 2;
    const parentChartAnchorRow  = 1;             // aligned with Parent table title (row 1)
    const courseChartAnchorRow  = courseTitleRow; // aligned with Course table title
    const RIGHT_CHART_COL       = 12;            // column M (0-based)

    // ===== Embed PNG charts at the anchor rows we just computed =====
    // Each PNG carries an in-canvas title matching its source-table heading
    // (Registrations by Sub-Type / Parent Campaign / Course), pulled from
    // CONFIG.dashboard.labels so they stay in sync with the dashboard cards.
    const labels = (CONFIG.dashboard && CONFIG.dashboard.labels) || {};
    // Sub-Type detail series (excludes bucket parents).
    const sLabels = [], sReg = [], sEnr = [];
    for (const bkt of Object.keys(agg.subtypeBuckets)) {
      for (const r of agg.subtypeBuckets[bkt]) {
        sLabels.push(r.name); sReg.push(r.reg); sEnr.push(r.enr);
      }
    }
    const sub    = await renderChartPng('off-subtype', sLabels, sReg, sEnr, { title: labels.chartSubtypeTitle || 'Registrations by Sub-Type' });
    const pLabels = agg.parent.map(r => r.name);
    const pReg    = agg.parent.map(r => r.reg);
    const pEnr    = agg.parent.map(r => r.enr);
    const parent  = await renderChartPng('off-parent', pLabels, pReg, pEnr, { title: labels.chartParentTitle || 'Registrations by Parent Campaign' });
    const cLabels = agg.course.map(r => r.name);
    const cReg    = agg.course.map(r => r.reg);
    const cEnr    = agg.course.map(r => r.enr);
    const course  = await renderChartPng('off-course', cLabels, cReg, cEnr, { title: labels.chartCourseTitle || 'Registrations by Course' });

    // Embed at the exact dimensions the canvas was rendered for so Excel
    // doesn't stretch the bitmap when columns/rows are resized.
    const subId = wb.addImage({ base64: sub.dataUrl, extension: 'png' });
    ws.addImage(subId, {
      tl:  { col: 0, row: subtypeChartAnchorRow - 1 },
      ext: { width: sub.embedW, height: sub.embedH },
      editAs: 'oneCell',
    });
    const parentId = wb.addImage({ base64: parent.dataUrl, extension: 'png' });
    ws.addImage(parentId, {
      tl:  { col: RIGHT_CHART_COL, row: parentChartAnchorRow - 1 },
      ext: { width: parent.embedW, height: parent.embedH },
      editAs: 'oneCell',
    });
    const courseId = wb.addImage({ base64: course.dataUrl, extension: 'png' });
    ws.addImage(courseId, {
      tl:  { col: RIGHT_CHART_COL, row: courseChartAnchorRow - 1 },
      ext: { width: course.embedW, height: course.embedH },
      editAs: 'oneCell',
    });

    return { ws, subtypeChartAnchorRow, parentChartAnchorRow, courseChartAnchorRow };
  }

  // ----- Sheet 3: Campaign Members ------------------------------------------
  // Every CM (post test-row removal) appears exactly once. If matched to a
  // Participant, the PA's key fields appear inline (best match wins when there
  // are several: Enrolled > Registered > anything else). Wrapped as an Excel
  // Table named "CampaignMembers" so the Course Status column drives the
  // Campaign Summary's COUNTIFS formulas.
  // Shared body for both Campaign Members sheets. `scope(row)` decides which
  // CMs land in this sheet — used to split into Focus / Older variants while
  // keeping the column layout, sort order, and styling identical.
  function buildCampaignMembersScopedSheet(wb, ctx, opts) {
    const { STATE, FONT_HEADER, FONT_BODY, FILL_NAVY, BORDER_THIN } = rs();
    const { sheetName, tableName, scope } = opts;

    const ws = wb.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 1 }] });

    const cmClean = STATE.primary.rows.filter(r => !isTestRow(r) && scope(r));

    // Reverse-index: CM row → list of stitched matches (each carries pa, method, score)
    const primaryToStitched = new Map();
    for (const s of STATE.stitched) {
      if (!primaryToStitched.has(s.primary)) primaryToStitched.set(s.primary, []);
      primaryToStitched.get(s.primary).push(s);
    }

    // When a CM has multiple PA matches, pick the one with the highest course status.
    const rankPa = (s) => {
      const status = (s.secondary['Status'] || '').trim();
      if (status === 'Enrolled')   return 3;
      if (status === 'Registered') return 2;
      return 1;
    };

    const headers = [
      'Parent Campaign Name', 'Campaign Name', 'Sub-Type Bucket', 'Sub-Type',
      'Last Name', 'First Name', 'Email', 'Phone', 'Contact ID',
      'Member Status', 'Member Status Update Date', 'Member First Responded Date',
      'CM Related Course', 'Course Status', 'Match Method', 'Course Match Score',
      'Days to Convert',
      'PA Course Name', 'PA Status', 'PA Created Date',
      'PA Course Instance: Start Date', 'PA Contact ID',
    ];

    // Sort by Parent Campaign → Campaign Name → Last Name for ease of scanning.
    const sortedCms = [...cmClean].sort((a, b) => {
      const pa = (a['Parent Campaign Name'] || '');
      const pb = (b['Parent Campaign Name'] || '');
      if (pa !== pb) return pa.localeCompare(pb);
      const ca = (a['Campaign Name'] || '');
      const cb = (b['Campaign Name'] || '');
      if (ca !== cb) return ca.localeCompare(cb);
      return (a['Last Name'] || '').localeCompare(b['Last Name'] || '');
    });

    const rows = [];
    for (const cm of sortedCms) {
      const matches = primaryToStitched.get(cm) || [];
      const courseStatus = deriveCourseStatus(matches);
      const best = matches.length === 0 ? null : matches.reduce((a, b) => rankPa(b) > rankPa(a) ? b : a);
      const pa = best ? best.secondary : null;
      // Engagement → registration lag (whole days). Null if no PA match or
      // either date is missing. Goes in the new "Days to Convert" column.
      let daysToConvert = '';
      if (pa) {
        const updated = cm._memberStatusUpdate;
        const created = pa._secondaryCreated;
        if (updated && created) daysToConvert = Math.round((created.getTime() - updated.getTime()) / 86400000);
      }
      // Date cells get the pre-parsed Date objects from the row caches so
      // Excel writes them as real date cells (sortable, filterable). Falls
      // back to null when unparsed/blank — Excel renders blank.
      rows.push([
        cm['Parent Campaign Name'] || '',
        cm['Campaign Name'] || '',
        bucketSubType(cm['Sub-Type']),
        cm['Sub-Type'] || '',
        cm['Last Name'] || '',
        cm['First Name'] || '',
        cm['Email'] || '',
        cm['Phone'] || '',
        cm['Contact ID'] || '',
        cm['Member Status'] || '',
        cm._memberStatusUpdate   || null,
        cm._memberFirstResponded || null,
        cm['Related Course'] || '',
        courseStatus,
        best ? best.method : '',
        best ? best.score : '',
        daysToConvert,
        pa ? (pa['Course Name'] || '') : '',
        pa ? (pa['Status'] || '') : '',
        pa ? (pa._secondaryCreated || null) : null,
        pa ? (pa._courseStart      || null) : null,
        pa ? (pa['Contact ID'] || '') : '',
      ]);
    }

    // ExcelJS requires at least one row in addTable.
    if (rows.length === 0) rows.push(headers.map(() => ''));

    ws.addTable({
      name: tableName,
      ref: 'A1',
      headerRow: true,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: headers.map(name => ({ name, filterButton: true })),
      rows,
    });

    // Override header styling (table style sets a different theme color).
    const headerRow = ws.getRow(1);
    headerRow.height = 22;
    headerRow.eachCell((cell, colNum) => {
      if (colNum > headers.length) return;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.border = BORDER_THIN;
    });

    // Date columns get fixed width + m/d/yyyy display so Excel sorts them as
    // real dates. Indices below are 1-based and pinned to the headers array
    // above (Member Status Update Date, Member First Responded Date, PA
    // Created Date, PA Course Instance: Start Date).
    const DATE_COL_INDICES = new Set([11, 12, 20, 21]);

    // Column widths. Date columns hardcoded; everything else autosizes against
    // the longest stringified value, capped at 45 chars.
    for (let c = 1; c <= headers.length; c++) {
      if (DATE_COL_INDICES.has(c)) {
        ws.getColumn(c).width  = 12;
        ws.getColumn(c).numFmt = 'm/d/yyyy';
        continue;
      }
      let maxLen = headers[c - 1].length;
      for (const row of rows) {
        const v = row[c - 1];
        const s = v == null ? '' : String(v);
        if (s.length > maxLen) maxLen = s.length;
      }
      ws.getColumn(c).width = Math.min(Math.max(maxLen + 2, 10), 45);
    }
    for (let r = 2; r <= rows.length + 1; r++) {
      const row = ws.getRow(r);
      row.eachCell((cell, colNum) => {
        if (colNum > headers.length) return;
        cell.font = FONT_BODY;
        cell.alignment = { vertical: 'middle' };
      });
    }

    return {
      ws,
      sheetName,
      tableName,
      dataLastRow: rows.length + 1,
    };
  }

  // ----- Sheet 3: Campaign Members (Focus) — CMs ∈ Focus -------------------
  // tableName 'CampaignMembers' preserved so the Campaign Summary's COUNTIFS
  // formulas keep referencing the Focus-scoped table without change.
  function buildCampaignMembersSheet(wb, ctx) {
    const { dateInRange } = rs();
    const dateOf = ctx.dateOf;
    return buildCampaignMembersScopedSheet(wb, ctx, {
      sheetName: 'Campaign Members (Focus)',
      tableName: 'CampaignMembers',
      scope: (cm) => {
        if (!dateOf) return true;   // no date dimension → include all
        const d = dateOf.primary(cm);
        return d && dateInRange(d, ctx.focus);
      },
    });
  }

  // ----- Sheet 4: Campaign Members (Older) — CMs ∈ Influence ∖ Focus -------
  // Older CMs (matched and unmatched) that fell inside the broader Influence
  // window but outside Focus. Their own table so Campaign Summary stays
  // Focus-only.
  function buildCampaignMembersOlderSheet(wb, ctx) {
    const { dateInRange } = rs();
    const dateOf = ctx.dateOf;
    return buildCampaignMembersScopedSheet(wb, ctx, {
      sheetName: 'Campaign Members (Older)',
      tableName: 'CampaignMembersOlder',
      scope: (cm) => {
        if (!dateOf) return false;
        const d = dateOf.primary(cm);
        if (!d) return false;
        return dateInRange(d, ctx.influence) && !dateInRange(d, ctx.focus);
      },
    });
  }

  // ----- Sheet 5: Post-Reg Engagement (conditional) ------------------------
  // Lists the candidate matches matchValidator rejected because the CM's
  // First Responded Date is after the PA's Created Date. Built only when
  // there's something to show — returns null otherwise so the engine skips
  // adding the worksheet.
  function buildPostRegEngagementSheet(wb, ctx) {
    const { STATE, FONT_HEADER, FONT_BODY, FILL_NAVY, BORDER_THIN, NAVY_ARGB, colLetter } = rs();

    const rejected = (STATE.rejectedMatches || []).filter(r =>
      !r.reason || r.reason === 'engagement-after-registration'
    );
    if (rejected.length === 0) return null;

    const ws = wb.addWorksheet('Post-Reg Engagement', { views: [{ state: 'frozen', ySplit: 3 }] });

    // Title (row 1)
    ws.mergeCells('A1:O1');
    const title = ws.getCell('A1');
    title.value = 'Post-Registration Engagement (excluded from matching)';
    title.font = { name: 'Arial', size: 14, bold: true, color: { argb: NAVY_ARGB } };
    title.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 26;

    // Italic-gray scope note (row 2)
    const noteText = (CONFIG.dashboard && CONFIG.dashboard.labels && CONFIG.dashboard.labels.notePostRegEngagement) || '';
    if (noteText) {
      ws.mergeCells('A2:O2');
      const noteCell = ws.getCell('A2');
      noteCell.value = noteText;
      noteCell.font  = { name: 'Arial', size: 9.5, italic: true, color: { argb: 'FF6B7280' } };
      noteCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      ws.getRow(2).height = 56;
    }

    // Header row 3
    const headers = [
      'Parent Campaign Name', 'Campaign Name', 'Sub-Type',
      'Last Name', 'First Name', 'Email', 'Contact ID',
      'Member First Responded Date', 'Member Status Update Date',
      'PA Created Date', 'Days Reg Preceded Engagement',
      'Match Method', 'PA Course Name', 'PA Status', 'PA Contact ID',
    ];
    headers.forEach((h, i) => {
      const cell = ws.getCell(`${colLetter(i+1)}3`);
      cell.value = h;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = BORDER_THIN;
    });
    ws.getRow(3).height = 32;

    // Compute days-preceded per row and sort by the worst offenders first.
    const enriched = rejected.map(rec => {
      const cm = rec.primary, pa = rec.secondary;
      const responded = cm._memberFirstResponded;
      const created   = pa._secondaryCreated;
      const daysPreceded = (responded && created)
        ? Math.round((responded.getTime() - created.getTime()) / 86400000)
        : '';
      return { rec, daysPreceded };
    }).sort((a, b) => {
      const ax = typeof a.daysPreceded === 'number' ? a.daysPreceded : 0;
      const bx = typeof b.daysPreceded === 'number' ? b.daysPreceded : 0;
      return bx - ax;
    });

    // Body
    let r = 4;
    for (const { rec, daysPreceded } of enriched) {
      const cm = rec.primary, pa = rec.secondary;
      const cells = [
        cm['Parent Campaign Name'] || '',
        cm['Campaign Name'] || '',
        cm['Sub-Type'] || '',
        cm['Last Name'] || '',
        cm['First Name'] || '',
        cm['Email'] || '',
        cm['Contact ID'] || '',
        cm._memberFirstResponded || null,
        cm._memberStatusUpdate   || null,
        pa._secondaryCreated     || null,
        daysPreceded,
        rec.method,
        pa['Course Name'] || '',
        pa['Status'] || '',
        pa['Contact ID'] || '',
      ];
      cells.forEach((v, i) => {
        const cell = ws.getCell(`${colLetter(i+1)}${r}`);
        cell.value = v;
        cell.font  = FONT_BODY;
        cell.alignment = { vertical: 'middle' };
        cell.border = BORDER_THIN;
      });
      r++;
    }

    // Column widths + date formatting. Indices below are 1-based against `headers`.
    const DATE_COL_INDICES = new Set([8, 9, 10]);   // First Responded, Status Update, PA Created
    const FIXED_WIDTHS = {
      1: 28, 2: 30, 3: 22,            // Parent / Campaign / Sub-Type
      4: 16, 5: 16, 6: 26, 7: 14,     // Last / First / Email / Contact ID
      11: 14,                          // Days Reg Preceded Engagement
      12: 14,                          // Match Method
      13: 38, 14: 14, 15: 14,         // PA Course Name / PA Status / PA Contact ID
    };
    for (let c = 1; c <= headers.length; c++) {
      if (DATE_COL_INDICES.has(c)) {
        ws.getColumn(c).width  = 12;
        ws.getColumn(c).numFmt = 'm/d/yyyy';
        continue;
      }
      ws.getColumn(c).width = FIXED_WIDTHS[c] || 14;
    }

    return {
      ws,
      sheetName: 'Post-Reg Engagement',
      dataLastRow: r - 1,
    };
  }

  // ----- Sheet 6: Campaign Summary ------------------------------------------
  // Pivot-style view of CM counts by Course Status bucket, grouped by Parent
  // Campaign → Campaign Name with subtotals and a grand total. Count cells are
  // COUNTIFS formulas against the CampaignMembers structured table, so editing
  // rows on the data sheet auto-updates this view.
  function buildCampaignSummarySheet(wb, ctx) {
    const {
      STATE,
      FONT_HEADER, FONT_BODY, FILL_NAVY, FILL_LBLUE, FILL_GRAY,
      BORDER_THIN, BORDER_MED_BOTTOM,
      NAVY_ARGB,
      colLetter, escapeFormula,
    } = rs();

    const ws = wb.addWorksheet('Campaign Summary', { views: [{ state: 'frozen', ySplit: 3 }] });
    ws.properties.tabColor = { argb: NAVY_ARGB };
    const data = aggregateCmStatusByCampaign(ctx);

    // Structured table references — auto-resize when rows are added/removed.
    // The CampaignMembers table is created by buildCampaignMembersSheet (the
    // Focus sheet). Formulas resolve by name at open time, so this sheet can
    // be ordered ahead of Campaign Members in the tab strip — we just need
    // both sheets in the final workbook.
    const tbl         = 'CampaignMembers';
    const refParent   = `${tbl}[Parent Campaign Name]`;
    const refCampaign = `${tbl}[Campaign Name]`;
    const refStatus   = `${tbl}[Course Status]`;

    // Title (A1:G1)
    ws.mergeCells('A1:G1');
    const title = ws.getCell('A1');
    title.value = 'Campaign Members — Course Status Distribution';
    title.font = { name: 'Arial', size: 14, bold: true, color: { argb: NAVY_ARGB } };
    title.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 26;

    // Note row 2 — italic gray explanation of the Focus-only scope.
    ws.mergeCells('A2:G2');
    const noteCell = ws.getCell('A2');
    const noteText = (CONFIG.dashboard && CONFIG.dashboard.labels && CONFIG.dashboard.labels.noteCampaignSummary) || '';
    if (noteText) {
      noteCell.value = noteText;
      noteCell.font  = { name: 'Arial', size: 10, italic: true, color: { argb: 'FF6B7280' } };
      noteCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      ws.getRow(2).height = 44;
    }

    // Header row 3 — column D's user-facing label is "Cancelled/Other"; the
    // underlying COUNTIFS criterion stays the canonical "Cancelled/Withdrawn/Etc"
    // string deriveCourseStatus emits (locked decision #3).
    const headers = ['Parent Campaign', 'Campaign Name', 'Not Converted', 'Cancelled/Other', 'Enrolled', 'Registered', 'Total'];
    headers.forEach((h, i) => {
      const cell = ws.getCell(`${colLetter(i+1)}3`);
      cell.value = h;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = BORDER_THIN;
    });
    ws.getRow(3).height = 32;

    // Hide column D (Cancelled/Other) when the secondary data has no
    // non-Reg/Enr statuses — formulas stay valid (will just resolve to 0).
    const hasOther = STATE.stitched.some(s => {
      const st = (s.secondary['Status'] || '').toString().trim();
      return st && st !== 'Registered' && st !== 'Enrolled';
    });
    if (!hasOther) ws.getColumn(4).hidden = true;

    const styleNumCells = (rowNum, opts = {}) => {
      ['C','D','E','F','G'].forEach(col => {
        const cell = ws.getCell(`${col}${rowNum}`);
        cell.font = opts.font || FONT_BODY;
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = opts.border || BORDER_THIN;
        cell.numFmt = '#,##0';
        if (opts.fill) cell.fill = opts.fill;
      });
    };

    // Build a COUNTIFS formula. Pass null campaign to mean "any campaign" (subtotals).
    const cellFormula = (parentVal, campaignVal, statusVal) => {
      return campaignVal == null
        ? `COUNTIFS(${refParent},"${escapeFormula(parentVal)}",${refStatus},"${statusVal}")`
        : `COUNTIFS(${refParent},"${escapeFormula(parentVal)}",${refCampaign},"${escapeFormula(campaignVal)}",${refStatus},"${statusVal}")`;
    };

    let r = 4;
    const sortedParents = [...data.keys()].sort((a, b) => a.localeCompare(b));

    for (const parent of sortedParents) {
      const inner = data.get(parent);
      const sortedCampaigns = [...inner.keys()].sort((a, b) => a.localeCompare(b));

      sortedCampaigns.forEach((campaign, idx) => {
        ws.getCell(`A${r}`).value = idx === 0 ? parent : '';
        ws.getCell(`B${r}`).value = campaign;
        ws.getCell(`C${r}`).value = { formula: cellFormula(parent, campaign, 'Not Converted') };
        ws.getCell(`D${r}`).value = { formula: cellFormula(parent, campaign, 'Cancelled/Withdrawn/Etc') };
        ws.getCell(`E${r}`).value = { formula: cellFormula(parent, campaign, 'Enrolled') };
        ws.getCell(`F${r}`).value = { formula: cellFormula(parent, campaign, 'Registered') };
        ws.getCell(`G${r}`).value = { formula: `C${r}+D${r}+E${r}+F${r}` };

        ['A', 'B'].forEach(col => {
          const cell = ws.getCell(`${col}${r}`);
          cell.font = (idx === 0 && col === 'A')
            ? { name: 'Arial', size: 10, bold: true, color: { argb: NAVY_ARGB } }
            : FONT_BODY;
          cell.alignment = { vertical: 'middle' };
          cell.border = BORDER_THIN;
        });
        styleNumCells(r);
        r++;
      });

      // Subtotal row (light blue) — formula counts CMs in this parent across all campaigns
      ws.getCell(`A${r}`).value = '';
      ws.getCell(`B${r}`).value = 'Subtotal';
      ws.getCell(`C${r}`).value = { formula: cellFormula(parent, null, 'Not Converted') };
      ws.getCell(`D${r}`).value = { formula: cellFormula(parent, null, 'Cancelled/Withdrawn/Etc') };
      ws.getCell(`E${r}`).value = { formula: cellFormula(parent, null, 'Enrolled') };
      ws.getCell(`F${r}`).value = { formula: cellFormula(parent, null, 'Registered') };
      ws.getCell(`G${r}`).value = { formula: `C${r}+D${r}+E${r}+F${r}` };
      ['A', 'B'].forEach(col => {
        const cell = ws.getCell(`${col}${r}`);
        cell.fill = FILL_LBLUE;
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: NAVY_ARGB } };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
        cell.border = BORDER_THIN;
      });
      styleNumCells(r, {
        fill: FILL_LBLUE,
        font: { name: 'Arial', size: 10, bold: true, color: { argb: NAVY_ARGB } },
      });
      r++;
    }

    // Grand Total — count across the entire CampaignMembers table
    ws.getCell(`A${r}`).value = 'Grand Total';
    ws.getCell(`B${r}`).value = '';
    ws.getCell(`C${r}`).value = { formula: `COUNTIF(${refStatus},"Not Converted")` };
    ws.getCell(`D${r}`).value = { formula: `COUNTIF(${refStatus},"Cancelled/Withdrawn/Etc")` };
    ws.getCell(`E${r}`).value = { formula: `COUNTIF(${refStatus},"Enrolled")` };
    ws.getCell(`F${r}`).value = { formula: `COUNTIF(${refStatus},"Registered")` };
    ws.getCell(`G${r}`).value = { formula: `C${r}+D${r}+E${r}+F${r}` };
    ['A', 'B'].forEach(col => {
      const cell = ws.getCell(`${col}${r}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
    });
    styleNumCells(r, {
      fill: FILL_GRAY,
      font: { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } },
      border: BORDER_MED_BOTTOM,
    });

    const widths = [38, 48, 14, 22, 11, 12, 10];
    widths.forEach((w, i) => ws.getColumn(i + 1).width = w);
  }

  // ===== Config registration =====

  window.registerStitchConfig({
    id: 'emory-cm-pa',
    label: 'Campaign Member ↔ Participant',
    description: 'Stitch Salesforce Campaign Members to Program Participants.',
    clientCodes: ['emory'],

    // ----- Inputs (Step 1 drop zones + header validation) -------------------
    inputs: {
      primary: {
        label: 'Campaign Member CSV',
        requiredCols: [
          'Last Name','First Name','Email','Contact ID','Sub-Type',
          'Related Course','Parent Campaign Name','Campaign Name',
          'Member Status Update Date','Member First Responded Date',
        ],
      },
      secondary: {
        label: 'Participant CSV',
        requiredCols: [
          'Last Name','First Name','Email','Contact ID','Course Name',
          'Status','Program Participant: Created Date',
          'Course Instance: Start Date',
        ],
      },
    },

    // ----- Hero intro card --------------------------------------------------
    // Renders into #intro-card on first paint. `title` accepts inline HTML
    // (<em>, <br>) so the italicized phrase keeps its serif treatment.
    intro: {
      eyebrow: 'Salesforce report join',
      title:   'Stitch <em>Campaign Members</em><br>to <em>Participants</em>.',
      sub:     'Drop your two CSV exports, review the matches, then download a styled, multi-sheet xlsx with charts and a live dashboard.',
    },

    // ----- Default column picker (the locked 19) ----------------------------
    // `source` values remain 'cm' | 'pa' | 'derived' for Slice B. The generic
    // 'primary' | 'secondary' rename is part of the Slice F sweep, alongside
    // STATE.primary / STATE.secondary → STATE.primary / STATE.secondary.
    defaultColumns: [
      { key:'pa_created',         label:'Program Participant: Created Date', source:'secondary',     sourceField:'Program Participant: Created Date', type:'date' },
      { key:'course_registered',  label:'Course (Registered)',                source:'secondary',     sourceField:'Course Name' },
      { key:'status',             label:'Status',                             source:'secondary',     sourceField:'Status' },
      { key:'last_name',          label:'Last Name',                          source:'secondary',     sourceField:'Last Name', fallbackField:'Last Name' },
      { key:'first_name',         label:'First Name',                         source:'secondary',     sourceField:'First Name', fallbackField:'First Name' },
      { key:'email',              label:'Email',                              source:'secondary',     sourceField:'Email' },
      { key:'parent_campaign',    label:'Parent Campaign Name',               source:'primary',     sourceField:'Parent Campaign Name' },
      { key:'campaign',           label:'Campaign Name',                      source:'primary',     sourceField:'Campaign Name' },
      { key:'subtype_bucket',     label:'Sub-Type Bucket',                    source:'derived',sourceField:'subtype_bucket' },
      { key:'subtype',            label:'Sub-Type',                           source:'primary',     sourceField:'Sub-Type' },
      { key:'course_start',       label:'Course Instance: Start Date',        source:'secondary',     sourceField:'Course Instance: Start Date', type:'date' },
      { key:'first_responded',    label:'Member First Responded Date',        source:'primary',     sourceField:'Member First Responded Date', type:'date' },
      { key:'status_update',      label:'Member Status Update Date',          source:'primary',     sourceField:'Member Status Update Date', type:'date' },
      { key:'cm_related_course',  label:'CM Related Course',                  source:'primary',     sourceField:'Related Course' },
      { key:'phone',              label:'Phone',                              source:'secondary',     sourceField:'Phone', fallbackField:'Phone' },
      { key:'pa_contact_id',      label:'Contact ID (PA)',                    source:'secondary',     sourceField:'Contact ID' },
      { key:'cm_contact_id',      label:'Contact ID (CM)',                    source:'primary',     sourceField:'Contact ID' },
      { key:'match_method',       label:'Match Method',                       source:'derived',sourceField:'match_method' },
      { key:'course_score',       label:'Course Match Score',                 source:'derived',sourceField:'course_score' },
      { key:'days_to_convert',    label:'Days to Convert',                    source:'derived',sourceField:'days_to_convert' },
    ],

    // ----- Match strategy (declarative; engine still uses indexCm + findCmMatches
    //       internally — Slice F rewires those off of matchStrategy) ----------
    matchStrategy: [
      { method: 'ContactID', primaryCol: 'Contact ID', secondaryCol: 'Contact ID', normalize: 'trim'  },
      { method: 'Email',     primaryCol: 'Email',      secondaryCol: 'Email',      normalize: 'email' },
      { method: 'Phone',     primaryCol: 'Phone',      secondaryCol: 'Phone',      normalize: 'phone' },
    ],

    // ----- Test-row filter (applied to BOTH primary and secondary inputs) ----
    testRowFilter: isTestRow,

    // ----- Multi-candidate tiebreaker --------------------------------------
    // Called for every match (single or multi-candidate). Returns { row, score };
    // the score also flows into derivedFields.course_score via matchInfo.
    tiebreaker: (candidates, secondaryRow) => {
      let best = candidates[0];
      let bestScore = courseSimilarity(best['Related Course'], secondaryRow['Course Name']);
      for (let i = 1; i < candidates.length; i++) {
        const s = courseSimilarity(candidates[i]['Related Course'], secondaryRow['Course Name']);
        if (s > bestScore) { best = candidates[i]; bestScore = s; }
      }
      return { row: best, score: bestScore };
    },

    // ----- Derived fields (computed once per stitched row) -----------------
    // Signature: (primary, secondary, matchInfo) → any. subtype_bucket tolerates
    // undefined secondary/matchInfo so the engine can also call it from a
    // CM-only context (xlsx Campaign Members sheet, dashboard CM-funnel)
    // without synthesizing a stitched row.
    derivedFields: {
      subtype_bucket: (primary)                        => bucketSubType(primary['Sub-Type']),
      match_method:   (primary, secondary, matchInfo)  => matchInfo && matchInfo.method,
      course_score:   (primary, secondary, matchInfo)  => matchInfo && matchInfo.score,
      // Whole-day count between Member Status Update Date (engagement) and
      // Program Participant: Created Date (registration). Null when either
      // date is missing or the row isn't a stitched match. Reads the parsed-
      // date caches the engine populates pre-stitch (_memberStatusUpdate /
      // _secondaryCreated).
      days_to_convert: (primary, secondary) => {
        if (!primary || !secondary) return null;
        const updated = primary._memberStatusUpdate;
        const created = secondary._secondaryCreated;
        if (!updated || !created) return null;
        return Math.round((created.getTime() - updated.getTime()) / 86400000);
      },
    },

    // ----- Match validator -------------------------------------------------
    // Engine calls this for every candidate (primary, secondary) pair after
    // matchStrategy finds them. Returning true keeps the candidate; returning
    // { valid:false, reason } drops it. For Emory: reject when the CM's
    // First Responded Date is AFTER the PA's Created Date — the form
    // submission happened after the registration, so it couldn't logically
    // have driven it. Rejections surface on the "Post-Reg Engagement" sheet
    // (built only when there's something to show).
    //
    // Note: First Responded (NOT Member Status Update) is the right anchor
    // here. Status Update can drift forward over time even for someone whose
    // initial engagement was long before registration.
    matchValidator: (primary, secondary) => {
      const responded = primary._memberFirstResponded;
      const created   = secondary._secondaryCreated;
      // Missing dates — can't validate; default to accepting.
      if (!responded || !created) return true;
      if (responded.getTime() <= created.getTime()) return true;
      return { valid: false, reason: 'engagement-after-registration' };
    },

    // ----- CM-level Course Status derivation (load-bearing, see helper) ----
    deriveCourseStatus,

    // ----- Locked Sub-Type Bucket display order ----------------------------
    bucketOrder: BUCKET_ORDER,

    // ----- xlsx output sheets ---------------------------------------------
    // Rendered in array order. Each builder receives (wb, ctx); the engine
    // threads each builder's return value into ctx.sheets[name] so downstream
    // builders can read their handoff (e.g. Participant Summary reads
    // 'Stitched Data', Campaign Summary reads 'Campaign Members'). Chart PNG
    // embeds happen inline inside buildParticipantSummarySheet so the anchor
    // rows can be computed from the table sizes it just laid out.
    outputSheets: [
      { name: 'Stitched Data',             builder: buildStitchedSheet },
      { name: 'Participant Summary',       builder: buildParticipantSummarySheet },
      { name: 'Campaign Summary',          builder: buildCampaignSummarySheet },
      { name: 'Campaign Members (Focus)',  builder: buildCampaignMembersSheet },
      { name: 'Campaign Members (Older)',  builder: buildCampaignMembersOlderSheet },
      // Conditional — builder returns null and no worksheet is added when
      // there are no rejected matches to surface.
      { name: 'Post-Reg Engagement',       builder: buildPostRegEngagementSheet },
    ],

    // ----- Dashboard layout ------------------------------------------------
    // Static data the engine's dashboard renderer reads. The renderer itself
    // (section 13 of app.js) handles layout + interaction generically;
    // anything Emory-specific lives here. statusOrder is load-bearing — its
    // strings must match what deriveCourseStatus returns (locked decision #3).
    dashboard: {
      statusOrder: ['Not Converted', 'Cancelled/Withdrawn/Etc', 'Registered', 'Enrolled'],
      statusColor: {
        'Not Converted':           '#6b7280',   // slate
        'Cancelled/Withdrawn/Etc': '#9b1c1c',   // burgundy
        'Registered':              '#1F3864',   // navy (matches xlsx NAVY)
        'Enrolled':                '#1f7a3a',   // green
      },
      bucketOptions: ['Website', 'Social', 'Unknown'],

      // Date dimension extractors — the engine reads these to scope the
      // stitch (primary by Influence, secondary by Focus) and to drive the
      // dashboard's two date sliders. Functions return Date | null so the
      // engine can decide what to do with rows lacking a parseable date.
      // cacheParsedDatesOnRows (engine-side) populates the underscore-prefixed
      // caches these point at, pre-stitch.
      dateOf: {
        primary:   (row) => row && row._memberStatusUpdate || null,
        secondary: (row) => row && row._secondaryCreated   || null,
      },

      // User-facing strings. The engine reads these via [data-label] /
      // [data-label-title] attributes in index.html (painted by
      // applyConfigToDom at init). Defaults in the HTML remain as a graceful
      // fallback if JS is slow to bind.
      labels: {
        // KPI block — Step 2 + Dashboard funnel share these
        kpiPrimaryTotal:        'Total CMs',
        kpiStatusNotConverted:  'Not Converted',
        kpiStatusCancelled:     'Cancelled / Withdrawn / Etc',
        kpiStatusRegistered:    'Registered',
        kpiStatusEnrolled:      'Enrolled',
        kpiSecondaryTotal:      'Total PAs',
        kpiStitched:            'Stitched (PA ↔ CM)',
        kpiSecondaryUnmatched:  'PAs without CM',
        kpiConversionRate:      'CM conversion rate',
        kpiDashPrimaryTotal:    'CMs (filtered)',

        // Section headers
        sectionPrimaryFunnel:     'Campaign Member funnel',
        sectionSecondaryMatching: 'Participant matching',
        sectionDashboardFunnel:   'Filtered Campaign Member funnel',
        sectionDistributions:     'Distributions (Registered + Enrolled)',
        sectionCampaignSummary:   'Campaign Summary',

        // Match-details breakdown line
        matchMethodLabel:    'Match method ContactID / Email / Phone:',
        scoreDistLabel:      'Course score 1.0 / fuzzy / 0.0:',
        testRemovedLabel:    'Test rows removed:',

        // Filter UI
        filterParentLabel:   'Parent Campaign',
        filterParentAll:     'All campaigns ▾',
        filterSubTypeLabel:  'Sub-Type',
        filterSubTypeAll:    'All sub-types ▾',
        filterBucketLabel:   'Sub-Type Bucket',
        filterStatusLabel:   'Course Status',

        // Date Ranges step (Configure tab) + dashboard dual sliders
        stepDateRangesHeading: 'Date ranges',
        stepDateRangesBlurb:   'Focus + Influence windows',
        stepDateRangesHelp:    'These windows scope the stitch. Focus narrows which Participant registrations are counted; Influence is the broader pool of Campaign Member activity that might have driven those registrations. Defaults auto-fill from your uploads.',
        focusRangeLabel:       'Focus Date Range',
        focusRangeSub:         'Participant Created Date',
        focusRangeHelp:        'The primary window of activity you’re analyzing. Drives both the Participant Summary (registrations created in this range) and the Campaign Summary (Campaign Member engagement within this range). Defaults to the first day of the earliest Participant month through the last day of the latest.',
        influenceRangeLabel:   'Influence Date Range',
        influenceRangeSub:     'Campaign Member Updated Date',
        influenceRangeHelp:    'The broader window of Campaign Member engagement eligible to match Participants in Focus. Should start at least 1–2 years before Focus so earlier campaign touches still count. Defaults to the full span of your CM data.',

        // Filter-bar "Generate xlsx (current filter)" button
        btnFilteredXlsxLabel: 'Generate xlsx (current filter)',

        // Chart titles + captions
        chartFunnelCaption:     'Click any segment to drill into the matching Campaign Members.',
        funnelIncludeNotConvertedLabel: 'Include "Not Converted"',
        chartTimeseriesCaption: 'Click any point to drill into the Campaign Members or Participants in that time bin.',
        chartSubtypeTitle:      'Registrations by Sub-Type',
        chartParentTitle:       'Registrations by Parent Campaign',
        chartCourseTitle:       'Registrations by Course',

        // "Slice it on the Dashboard" callout sub-line
        calloutSub: 'Filter by date, parent campaign, sub-type, or status — and drill into any chart segment.',

        // xlsx italic-gray notes rendered under each affected sheet title.
        // Explain to the recipient why Participant and Campaign Summary
        // counts may differ even though they come from the same dataset.
        noteParticipantSummary: 'Counts include all registrations in the Focus Date Range, matched against Campaign Members whose update date falls within the (wider) Influence Date Range. If this is higher than the Campaign Summary’s Grand Total, that’s expected — Campaign Summary is scoped to engagement within Focus Date Range only.',
        noteCampaignSummary:    'Counts include only Campaign Members whose updated date falls within the Focus Date Range. Older Campaign Members that contributed to registrations in Focus live in the “Campaign Members (Older)” sheet. Counts here will be lower than the Participant Summary’s because of this narrower scope.',
        notePostRegEngagement:  'Candidate matches excluded from the stitch because the Campaign Member’s First Responded Date is after the Participant’s Created Date — the form submission happened after the registration, so it couldn’t logically have influenced it. Review and clean the underlying data if any of these look like real attribution.',
      },
    },
  });

})();
