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
  function aggregateCmStatusByCampaign() {
    const { STATE } = rs();
    const cmClean = STATE.primary.rows.filter(r => !isTestRow(r));
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
    const { STATE, FONT_HEADER, FONT_BODY, FILL_NAVY, BORDER_THIN, getCellValue } = rs();
    const ws = wb.addWorksheet('Stitched Data', { views: [{ state: 'frozen', ySplit: 1 }] });
    const enabled = STATE.columns.filter(c => c.enabled);
    if (enabled.length === 0) throw new Error('Column picker: at least one column must be selected.');

    const tableColumns = enabled.map(c => ({ name: c.label, filterButton: true }));
    const tableRows = STATE.stitched.map(r => enabled.map(col => {
      const v = getCellValue(r, col);
      if (col.key === 'course_score') return Number(v);
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

    // Body styling + column widths (autosize, capped at 45 chars).
    for (let c = 1; c <= enabled.length; c++) {
      let maxLen = enabled[c-1].label.length;
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

    // Default column widths
    ws.getColumn(1).width = 28;
    ws.getColumn(2).width = 30;
    ws.getColumn(3).width = 13;
    ws.getColumn(4).width = 13;
    ws.getColumn(5).width = 13;
    ws.getColumn(6).width = 3;
    ws.getColumn(7).width = 38;
    ws.getColumn(8).width = 13;
    ws.getColumn(9).width = 13;
    ws.getColumn(10).width = 13;

    // ===== Title (A1:E1) =====
    ws.mergeCells('A1:E1');
    const titleCell = ws.getCell('A1');
    titleCell.value = 'Stitched Report — Summary';
    titleCell.font = { name: 'Arial', size: 14, bold: true, color: { argb: NAVY_ARGB } };
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 26;

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
    ws.mergeCells(`A${r}:E${r}`);
    const sectionTitle = ws.getCell(`A${r}`);
    sectionTitle.value = 'Sub-Type Bucket / Sub-Type';
    sectionTitle.font = { name: 'Arial', size: 12, bold: true, color: { argb: NAVY_ARGB } };
    sectionTitle.fill = FILL_GRAY;
    sectionTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorderRange(ws, `A${r}:E${r}`, BORDER_THIN);
    r++;

    // Header row
    const subHeaders = ['Sub-Type Bucket', 'Sub-Type', 'Registered', 'Enrolled', 'Total'];
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
      } else {
        ws.getCell(`C${r}`).value = details.reduce((a,b)=>a+b.reg,0);
        ws.getCell(`D${r}`).value = details.reduce((a,b)=>a+b.enr,0);
      }
      ws.getCell(`E${r}`).value = { formula: `C${r}+D${r}` };
      ['C','D','E'].forEach(c => {
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
        } else {
          ws.getCell(`C${r}`).value = d.reg;
          ws.getCell(`D${r}`).value = d.enr;
        }
        ws.getCell(`E${r}`).value = { formula: `C${r}+D${r}` };
        ['C','D','E'].forEach(c => {
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
    ws.getCell(`E${r}`).value = { formula: `C${r}+D${r}` };
    ['A','B','C','D','E'].forEach(c => {
      const cell = ws.getCell(`${c}${r}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: c === 'A' || c === 'B' ? 'left' : 'center', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
      if (c === 'C' || c === 'D' || c === 'E') cell.numFmt = '#,##0';
    });
    const subTableLastRow = r;
    r++;

    // ===== Parent Campaign table (G1:J?) =====
    let pr = 1;
    ws.mergeCells(`G${pr}:J${pr}`);
    const pcTitle = ws.getCell(`G${pr}`);
    pcTitle.value = 'Registrations by Parent Campaign';
    pcTitle.font = { name: 'Arial', size: 12, bold: true, color: { argb: NAVY_ARGB } };
    pcTitle.fill = FILL_GRAY;
    pcTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorderRange(ws, `G${pr}:J${pr}`, BORDER_THIN);
    pr++;

    ['Parent Campaign','Registered','Enrolled','Total'].forEach((label, idx) => {
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
      } else {
        ws.getCell(`H${pr}`).value = row.reg;
        ws.getCell(`I${pr}`).value = row.enr;
      }
      ws.getCell(`J${pr}`).value = { formula: `H${pr}+I${pr}` };
      ['H','I','J'].forEach(c => {
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
    ws.getCell(`J${pr}`).value = { formula: `H${pr}+I${pr}` };
    ['G','H','I','J'].forEach(c => {
      const cell = ws.getCell(`${c}${pr}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: c === 'G' ? 'left' : 'center', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
      if (c !== 'G') cell.numFmt = '#,##0';
    });
    pr++;
    pr++;   // blank gap before Course table

    // ===== Course table (G{pr}:J?) =====
    ws.mergeCells(`G${pr}:J${pr}`);
    const cTitle = ws.getCell(`G${pr}`);
    cTitle.value = 'Registrations by Course';
    cTitle.font = { name: 'Arial', size: 12, bold: true, color: { argb: NAVY_ARGB } };
    cTitle.fill = FILL_GRAY;
    cTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    applyBorderRange(ws, `G${pr}:J${pr}`, BORDER_THIN);
    pr++;

    ['Course','Registered','Enrolled','Total'].forEach((label, idx) => {
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
      } else {
        ws.getCell(`H${pr}`).value = row.reg;
        ws.getCell(`I${pr}`).value = row.enr;
      }
      ws.getCell(`J${pr}`).value = { formula: `H${pr}+I${pr}` };
      ['H','I','J'].forEach(c => {
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
    ws.getCell(`J${pr}`).value = { formula: `H${pr}+I${pr}` };
    ['G','H','I','J'].forEach(c => {
      const cell = ws.getCell(`${c}${pr}`);
      cell.fill = FILL_GRAY;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY_ARGB } };
      cell.alignment = { horizontal: c === 'G' ? 'left' : 'center', vertical: 'middle' };
      cell.border = BORDER_MED_BOTTOM;
      if (c !== 'G') cell.numFmt = '#,##0';
    });

    // Chart anchor rows are derived from the table sizes we just emitted.
    const subtypeChartAnchorRow = subTableLastRow + 2;
    const parentChartAnchorRow  = pcTotalRow + 2;
    const courseChartAnchorRow  = subTableLastRow + 2 + 13;   // matches stitch.py spacing

    // ===== Embed PNG charts at the anchor rows we just computed =====
    // Sub-Type detail series (excludes bucket parents).
    const sLabels = [], sReg = [], sEnr = [];
    for (const bkt of Object.keys(agg.subtypeBuckets)) {
      for (const r of agg.subtypeBuckets[bkt]) {
        sLabels.push(r.name); sReg.push(r.reg); sEnr.push(r.enr);
      }
    }
    const sub    = await renderChartPng('off-subtype', sLabels, sReg, sEnr);
    const pLabels = agg.parent.map(r => r.name);
    const pReg    = agg.parent.map(r => r.reg);
    const pEnr    = agg.parent.map(r => r.enr);
    const parent  = await renderChartPng('off-parent', pLabels, pReg, pEnr);
    const cLabels = agg.course.map(r => r.name);
    const cReg    = agg.course.map(r => r.reg);
    const cEnr    = agg.course.map(r => r.enr);
    const course  = await renderChartPng('off-course', cLabels, cReg, cEnr);

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
      tl:  { col: 6, row: parentChartAnchorRow - 1 },
      ext: { width: parent.embedW, height: parent.embedH },
      editAs: 'oneCell',
    });
    const courseId = wb.addImage({ base64: course.dataUrl, extension: 'png' });
    ws.addImage(courseId, {
      tl:  { col: 0, row: courseChartAnchorRow - 1 },
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
  function buildCampaignMembersSheet(wb, ctx) {
    const { STATE, FONT_HEADER, FONT_BODY, FILL_NAVY, BORDER_THIN } = rs();

    const ws = wb.addWorksheet('Campaign Members', { views: [{ state: 'frozen', ySplit: 1 }] });

    const cmClean = STATE.primary.rows.filter(r => !isTestRow(r));

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
        cm['Member Status Update Date'] || '',
        cm['Member First Responded Date'] || '',
        cm['Related Course'] || '',
        courseStatus,
        best ? best.method : '',
        best ? best.score : '',
        pa ? (pa['Course Name'] || '') : '',
        pa ? (pa['Status'] || '') : '',
        pa ? (pa['Program Participant: Created Date'] || '') : '',
        pa ? (pa['Course Instance: Start Date'] || '') : '',
        pa ? (pa['Contact ID'] || '') : '',
      ]);
    }

    // ExcelJS requires at least one row in addTable.
    if (rows.length === 0) rows.push(headers.map(() => ''));

    ws.addTable({
      name: 'CampaignMembers',
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

    // Column widths (autosize, capped at 45 chars).
    for (let c = 1; c <= headers.length; c++) {
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
      sheetName: 'Campaign Members',
      tableName: 'CampaignMembers',
      dataLastRow: rows.length + 1,
    };
  }

  // ----- Sheet 4: Campaign Summary ------------------------------------------
  // Pivot-style view of CM counts by Course Status bucket, grouped by Parent
  // Campaign → Campaign Name with subtotals and a grand total. Count cells are
  // COUNTIFS formulas against the CampaignMembers structured table, so editing
  // rows on the data sheet auto-updates this view.
  function buildCampaignSummarySheet(wb, ctx) {
    const {
      FONT_HEADER, FONT_BODY, FILL_NAVY, FILL_LBLUE, FILL_GRAY,
      BORDER_THIN, BORDER_MED_BOTTOM,
      NAVY_ARGB,
      colLetter, escapeFormula,
    } = rs();

    const cmCtx = ctx.sheets['Campaign Members'];
    if (!cmCtx) throw new Error('Campaign Summary builder ran before Campaign Members.');

    const ws = wb.addWorksheet('Campaign Summary', { views: [{ state: 'frozen', ySplit: 3 }] });
    const data = aggregateCmStatusByCampaign();

    // Structured table references — auto-resize when rows are added/removed.
    const tbl = cmCtx.tableName;
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

    // Header row 3
    const headers = ['Parent Campaign', 'Campaign Name', 'Not Converted', 'Cancelled/Withdrawn/Etc', 'Enrolled', 'Registered', 'Total'];
    headers.forEach((h, i) => {
      const cell = ws.getCell(`${colLetter(i+1)}3`);
      cell.value = h;
      cell.font = FONT_HEADER;
      cell.fill = FILL_NAVY;
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = BORDER_THIN;
    });
    ws.getRow(3).height = 32;

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
      { key:'pa_created',         label:'Program Participant: Created Date', source:'secondary',     sourceField:'Program Participant: Created Date' },
      { key:'course_registered',  label:'Course (Registered)',                source:'secondary',     sourceField:'Course Name' },
      { key:'status',             label:'Status',                             source:'secondary',     sourceField:'Status' },
      { key:'last_name',          label:'Last Name',                          source:'secondary',     sourceField:'Last Name', fallbackField:'Last Name' },
      { key:'first_name',         label:'First Name',                         source:'secondary',     sourceField:'First Name', fallbackField:'First Name' },
      { key:'email',              label:'Email',                              source:'secondary',     sourceField:'Email' },
      { key:'parent_campaign',    label:'Parent Campaign Name',               source:'primary',     sourceField:'Parent Campaign Name' },
      { key:'campaign',           label:'Campaign Name',                      source:'primary',     sourceField:'Campaign Name' },
      { key:'subtype_bucket',     label:'Sub-Type Bucket',                    source:'derived',sourceField:'subtype_bucket' },
      { key:'subtype',            label:'Sub-Type',                           source:'primary',     sourceField:'Sub-Type' },
      { key:'course_start',       label:'Course Instance: Start Date',        source:'secondary',     sourceField:'Course Instance: Start Date' },
      { key:'first_responded',    label:'Member First Responded Date',        source:'primary',     sourceField:'Member First Responded Date' },
      { key:'status_update',      label:'Member Status Update Date',          source:'primary',     sourceField:'Member Status Update Date' },
      { key:'cm_related_course',  label:'CM Related Course',                  source:'primary',     sourceField:'Related Course' },
      { key:'phone',              label:'Phone',                              source:'secondary',     sourceField:'Phone', fallbackField:'Phone' },
      { key:'pa_contact_id',      label:'Contact ID (PA)',                    source:'secondary',     sourceField:'Contact ID' },
      { key:'cm_contact_id',      label:'Contact ID (CM)',                    source:'primary',     sourceField:'Contact ID' },
      { key:'match_method',       label:'Match Method',                       source:'derived',sourceField:'match_method' },
      { key:'course_score',       label:'Course Match Score',                 source:'derived',sourceField:'course_score' },
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
      { name: 'Stitched Data',       builder: buildStitchedSheet },
      { name: 'Participant Summary', builder: buildParticipantSummarySheet },
      { name: 'Campaign Members',    builder: buildCampaignMembersSheet },
      { name: 'Campaign Summary',    builder: buildCampaignSummarySheet },
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
        dateModeToggleTitle: 'Toggle between Activity date (CM update / PA created) and Course start date',

        // Filter-bar "Generate xlsx (current filter)" button
        btnFilteredXlsxLabel: 'Generate xlsx (current filter)',

        // Chart titles + captions
        chartFunnelCaption:     'Click any segment to drill into the matching Campaign Members.',
        chartTimeseriesCaption: 'Click any point to drill into the Campaign Members or Participants in that time bin.',
        chartSubtypeTitle:      'Registrations by Sub-Type',
        chartParentTitle:       'Registrations by Parent Campaign',
        chartCourseTitle:       'Registrations by Course',

        // "Slice it on the Dashboard" callout sub-line
        calloutSub: 'Filter by date, parent campaign, sub-type, or status — and drill into any chart segment.',
      },
    },
  });

})();
