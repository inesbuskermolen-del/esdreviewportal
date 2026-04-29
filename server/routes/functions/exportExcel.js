import ExcelJS from 'exceljs';
import { supabase } from '../../config/supabase.js';
import { attachUser, requireGIW } from '../../middleware/auth.js';

// POST /api/functions/exportExcel
export default [attachUser, requireGIW, async (req, res) => {
  try {
    const { projectId } = req.body;
    if (!projectId) return res.status(400).json({ error: 'Missing projectId' });

    const [
      { data: projectsArr },
      { data: credits },
      { data: comments },
      { data: drawingReqs },
      { data: excellenceOpps },
    ] = await Promise.all([
      supabase.from('projects').select('*').eq('id', projectId),
      supabase.from('credits').select('*').eq('project_id', projectId),
      supabase.from('credit_comments').select('*').eq('project_id', projectId),
      supabase.from('drawing_requirements').select('*').eq('project_id', projectId),
      supabase.from('esd_excellence_opportunities').select('*').eq('project_id', projectId),
    ]);

    const project = projectsArr?.[0];
    if (!project) return res.status(404).json({ error: 'Project not found' });

    // Build comment map by credit_id
    const commentsByCreditId = {};
    for (const c of (comments || [])) {
      if (!commentsByCreditId[c.credit_id]) commentsByCreditId[c.credit_id] = [];
      commentsByCreditId[c.credit_id].push(c);
    }

    // Group credits by category
    const categoryOrder = [];
    const categoryMap = {};
    for (const credit of (credits || [])) {
      const cat = credit.category || 'Uncategorised';
      if (!categoryMap[cat]) { categoryMap[cat] = []; categoryOrder.push(cat); }
      categoryMap[cat].push(credit);
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'GIW ESD Review Portal';
    workbook.created = new Date();

    const TEAL = '1F6B75';
    const GREEN_FILL = 'C6EFCE';
    const GREEN_FONT = '276221';
    const ORANGE_FILL = 'FCE4D6';
    const ORANGE_FONT = '8B2500';
    const GREY_FILL = 'D9D9D9';
    const YELLOW_FILL = 'FFEB9C';
    const HEADER_BG = 'D9E1F2';

    const tealFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + TEAL } };
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + HEADER_BG } };
    const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GREEN_FILL } };
    const orangeFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + ORANGE_FILL } };
    const greyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GREY_FILL } };
    const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + YELLOW_FILL } };
    const boldFont = { bold: true };
    const whiteFont = { bold: true, color: { argb: 'FFFFFFFF' } };
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
      left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
      bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
      right: { style: 'thin', color: { argb: 'FFD0D0D0' } },
    };

    function applyHeaderStyle(row) {
      row.eachCell((cell) => {
        cell.fill = headerFill;
        cell.font = boldFont;
        cell.border = thinBorder;
        cell.alignment = { wrapText: true, vertical: 'middle' };
      });
    }

    // ── SHEET 1: BESS Review Matrix ──────────────────────────────────────────
    const ws1 = workbook.addWorksheet('BESS Review Matrix');
    ws1.properties.defaultRowHeight = 15;
    const metaRows = [
      ['Project Name', project.name],
      ['Address', project.address || ''],
      ['Date', project.date ? new Date(project.date).toLocaleDateString('en-AU') : ''],
      ['Overall BESS Score', project.bess_score != null ? `${project.bess_score}%` : ''],
      ['Revision', project.revision || ''],
    ];
    for (const [label, value] of metaRows) {
      const row = ws1.addRow([label, value]);
      row.getCell(1).font = boldFont;
      row.getCell(1).fill = headerFill;
    }
    ws1.addRow([]);

    const cols1 = [
      'Credit Name', 'Credit Requirement', 'Mandatory', 'Responsible Party',
      'Y', 'M (Maybe)', 'N', 'Credit Score %', 'Credit Weight %',
      'Comments GIW', 'Comments Project Team',
    ];
    applyHeaderStyle(ws1.addRow(cols1));
    ws1.columns = [
      { width: 30 }, { width: 40 }, { width: 10 }, { width: 22 },
      { width: 8 }, { width: 10 }, { width: 8 }, { width: 12 }, { width: 13 },
      { width: 35 }, { width: 45 },
    ];

    for (const cat of categoryOrder) {
      const catCredits = categoryMap[cat];
      const avgScore = catCredits.length > 0
        ? Math.round(catCredits.filter(c => c.credit_status === 'Y').reduce((s, c) => s + (c.credit_score || 0), 0) / catCredits.length)
        : 0;
      const catWeight = catCredits[0]?.credit_weight ?? '';
      const catLabel = `${cat}    Score: ${avgScore}%${catWeight ? `  |  Weight: ${catWeight}%` : ''}`;
      const catRow = ws1.addRow([catLabel]);
      ws1.mergeCells(catRow.number, 1, catRow.number, 11);
      catRow.getCell(1).fill = tealFill;
      catRow.getCell(1).font = whiteFont;
      catRow.getCell(1).alignment = { vertical: 'middle' };
      catRow.height = 18;

      for (const credit of catCredits) {
        const teamComments = (commentsByCreditId[credit.id] || [])
          .map(c => `[${c.reviewer_discipline || c.reviewer_email}]: ${c.comment_text}`)
          .join('\n\n');

        const row = ws1.addRow([
          credit.credit_name || '',
          credit.credit_requirement || '',
          credit.mandatory ? 'M' : '',
          credit.responsible_party || '',
          credit.credit_status === 'Y' ? 'Y' : (credit.credit_status === 'ScopedOut' ? 'Scoped Out' : ''),
          credit.is_excellence_target ? 'M' : '',
          credit.credit_status === 'N' ? 'N' : '',
          credit.credit_score != null ? credit.credit_score : '',
          credit.credit_weight != null ? credit.credit_weight : '',
          credit.comments_giw || '',
          teamComments,
        ]);
        row.eachCell((cell) => { cell.border = thinBorder; cell.alignment = { wrapText: true, vertical: 'top' }; });

        const yCell = row.getCell(5);
        if (credit.credit_status === 'Y') {
          yCell.fill = greenFill; yCell.font = { bold: true, color: { argb: 'FF' + GREEN_FONT } };
          yCell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else if (credit.credit_status === 'ScopedOut') {
          yCell.fill = greyFill; yCell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
        }
        const mCell = row.getCell(6);
        if (credit.is_excellence_target) {
          mCell.fill = yellowFill; mCell.font = { bold: true, color: { argb: 'FF7D6608' } };
          mCell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
        const nCell = row.getCell(7);
        if (credit.credit_status === 'N') {
          nCell.fill = orangeFill; nCell.font = { bold: true, color: { argb: 'FF' + ORANGE_FONT } };
          nCell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
        const mandCell = row.getCell(3);
        if (credit.mandatory) {
          mandCell.font = { bold: true, color: { argb: 'FF' + TEAL } };
          mandCell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      }
    }

    // ── ESD Excellence & Innovation at the bottom of Sheet 1 ─────────────────
    const activeOppsSheet1 = (excellenceOpps || [])
      .filter(o => !o.deleted_by_giw)
      .sort((a, b) => {
        const CATS = ['Management','Integrated Water Management','Operational Energy','Indoor Environment Quality','Transport','Waste & Resource Recovery','Urban Ecology','Innovation'];
        const ci = CATS.indexOf(a.category); const di = CATS.indexOf(b.category);
        if (ci !== di) return (ci === -1 ? 99 : ci) - (di === -1 ? 99 : di);
        return (a.current_score ?? 0) - (b.current_score ?? 0);
      });

    const excellenceOppsSheet1 = activeOppsSheet1.filter(o => o.credit_reference !== 'Innovation');
    const innovationOppsSheet1 = activeOppsSheet1.filter(o => o.credit_reference === 'Innovation');

    function addOppRowToSheet1(opp) {
      const teamComments = (commentsByCreditId[opp.credit_id] || [])
        .map(c => `[${c.reviewer_discipline || c.reviewer_email}]: ${c.comment_text}`)
        .join('\n\n');
      const row = ws1.addRow([
        opp.credit_name || '',
        opp.improvement_description || '',
        '',
        '',
        opp.flag === 'Yes' ? 'Y' : '',
        opp.flag === 'Maybe' ? 'M' : '',
        opp.flag === 'No' ? 'N' : '',
        opp.current_score != null ? opp.current_score : '',
        '',
        '',
        teamComments,
      ]);
      row.eachCell((cell) => { cell.border = thinBorder; cell.alignment = { wrapText: true, vertical: 'top' }; });
      const yCell = row.getCell(5);
      if (opp.flag === 'Yes') { yCell.fill = greenFill; yCell.font = { bold: true, color: { argb: 'FF' + GREEN_FONT } }; yCell.alignment = { horizontal: 'center', vertical: 'middle' }; }
      const mCell = row.getCell(6);
      if (opp.flag === 'Maybe') { mCell.fill = yellowFill; mCell.font = { bold: true, color: { argb: 'FF7D6608' } }; mCell.alignment = { horizontal: 'center', vertical: 'middle' }; }
      const nCell = row.getCell(7);
      if (opp.flag === 'No') { nCell.fill = orangeFill; nCell.font = { bold: true, color: { argb: 'FF' + ORANGE_FONT } }; nCell.alignment = { horizontal: 'center', vertical: 'middle' }; }
    }

    if (excellenceOppsSheet1.length > 0) {
      const catRow = ws1.addRow(['ESD Excellence Opportunities']);
      ws1.mergeCells(catRow.number, 1, catRow.number, 11);
      catRow.getCell(1).fill = tealFill; catRow.getCell(1).font = whiteFont;
      catRow.getCell(1).alignment = { vertical: 'middle' }; catRow.height = 18;
      for (const opp of excellenceOppsSheet1) addOppRowToSheet1(opp);
    }

    if (innovationOppsSheet1.length > 0) {
      const catRow = ws1.addRow(['Innovation Credits']);
      ws1.mergeCells(catRow.number, 1, catRow.number, 11);
      catRow.getCell(1).fill = tealFill; catRow.getCell(1).font = whiteFont;
      catRow.getCell(1).alignment = { vertical: 'middle' }; catRow.height = 18;
      for (const opp of innovationOppsSheet1) addOppRowToSheet1(opp);
    }

    // ── SHEET 2: BESS Score Summary ──────────────────────────────────────────
    const ws2 = workbook.addWorksheet('BESS Score Summary');
    ws2.columns = [{ width: 30 }, { width: 12 }, { width: 12 }, { width: 10 }];
    applyHeaderStyle(ws2.addRow(['Category', 'Weight %', 'Score %', 'Pass (Y/N)']));
    let totalWeightedScore = 0, totalWeight = 0;
    for (const cat of categoryOrder) {
      const catCredits = categoryMap[cat];
      const catWeight = catCredits[0]?.credit_weight ?? 0;
      const avgScore = catCredits.length > 0
        ? Math.round(catCredits.filter(c => c.credit_status === 'Y').reduce((s, c) => s + (c.credit_score || 0), 0) / catCredits.length)
        : 0;
      const pass = avgScore >= 50 ? 'Y' : 'N';
      totalWeightedScore += avgScore * catWeight;
      totalWeight += catWeight;
      const row = ws2.addRow([cat, catWeight || '', avgScore || '', pass]);
      row.eachCell((cell) => { cell.border = thinBorder; cell.alignment = { vertical: 'middle' }; });
      row.getCell(4).font = pass === 'Y'
        ? { bold: true, color: { argb: 'FF' + GREEN_FONT } }
        : { bold: true, color: { argb: 'FF' + ORANGE_FONT } };
    }
    const overallScore = totalWeight > 0 ? Math.round(totalWeightedScore / totalWeight) : (project.bess_score || 0);
    const totalRow = ws2.addRow(['TOTAL BESS SCORE', '', `${overallScore}%`, '']);
    totalRow.eachCell((cell) => { cell.fill = tealFill; cell.font = whiteFont; cell.border = thinBorder; });

    // ── SHEET 3: Drawing Requirements ────────────────────────────────────────
    const ws3 = workbook.addWorksheet('Drawing Requirements');
    ws3.columns = [{ width: 20 }, { width: 18 }, { width: 40 }, { width: 22 }, { width: 14 }, { width: 35 }];
    applyHeaderStyle(ws3.addRow(['Credit Reference', 'Type', 'Requirement', 'Discipline', 'Status', 'Notes']));
    for (const dr of (drawingReqs || [])) {
      const row = ws3.addRow([dr.credit_reference || '', dr.drawing_type || '', dr.requirement || '', dr.discipline || '', dr.status || '', dr.notes || '']);
      row.eachCell((cell) => { cell.border = thinBorder; cell.alignment = { wrapText: true, vertical: 'top' }; });
      const statusCell = row.getCell(5);
      if (dr.status === 'Complete') { statusCell.fill = greenFill; statusCell.font = { color: { argb: 'FF' + GREEN_FONT } }; }
      else if (dr.status === 'InProgress') { statusCell.fill = yellowFill; }
    }

    // ── SHEET 4: ESD Excellence Opportunities ────────────────────────────────
    const CATEGORY_ORDER_EXPORT = [
      'Management', 'Integrated Water Management', 'Operational Energy',
      'Indoor Environment Quality', 'Transport', 'Waste & Resource Recovery',
      'Urban Ecology', 'Innovation',
    ];

    // Build credit lookup for BESS status (Y / M)
    const creditById = {};
    for (const c of (credits || [])) { creditById[c.id] = c; }

    const ws4 = workbook.addWorksheet('ESD Excellence Opportunities');
    ws4.columns = [{ width: 16 }, { width: 30 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 50 }, { width: 14 }, { width: 18 }, { width: 45 }];
    applyHeaderStyle(ws4.addRow(['Credit Reference', 'Credit Name', 'BESS Status', 'Current Score %', 'Max Score %', 'Improvement Description', 'Flag', 'Flagged By', 'Comments Project Team']));

    const activeOpps = (excellenceOpps || [])
      .filter(o => !o.deleted_by_giw)
      .sort((a, b) => {
        const ci = CATEGORY_ORDER_EXPORT.indexOf(a.category);
        const di = CATEGORY_ORDER_EXPORT.indexOf(b.category);
        if (ci !== di) return (ci === -1 ? 99 : ci) - (di === -1 ? 99 : di);
        return (a.current_score ?? 0) - (b.current_score ?? 0);
      });

    for (const opp of activeOpps) {
      const linkedCredit = creditById[opp.credit_id];
      const bessStatus = linkedCredit?.credit_status === 'Y' ? 'Y'
        : linkedCredit?.is_excellence_target ? 'M'
        : linkedCredit?.credit_status === 'N' ? 'N'
        : linkedCredit?.credit_status === 'ScopedOut' ? 'Scoped Out'
        : '';

      const teamComments = (commentsByCreditId[opp.credit_id] || [])
        .map(c => `[${c.reviewer_discipline || c.reviewer_email}]: ${c.comment_text}`)
        .join('\n\n');
      const row = ws4.addRow([
        opp.credit_reference || '', opp.credit_name || '', bessStatus,
        opp.current_score != null ? opp.current_score : '',
        opp.max_score != null ? opp.max_score : 100,
        opp.improvement_description || '', opp.flag || 'Unflagged', opp.flagged_by || '', teamComments,
      ]);
      row.eachCell((cell) => { cell.border = thinBorder; cell.alignment = { wrapText: true, vertical: 'top' }; });

      const bessCell = row.getCell(3);
      bessCell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (bessStatus === 'Y') { bessCell.fill = greenFill; bessCell.font = { bold: true, color: { argb: 'FF' + GREEN_FONT } }; }
      else if (bessStatus === 'M') { bessCell.fill = yellowFill; bessCell.font = { bold: true, color: { argb: 'FF7D6608' } }; }
      else if (bessStatus === 'N') { bessCell.fill = orangeFill; bessCell.font = { bold: true, color: { argb: 'FF' + ORANGE_FONT } }; }
      else if (bessStatus === 'Scoped Out') { bessCell.fill = greyFill; bessCell.alignment = { horizontal: 'center', vertical: 'middle' }; }

      const flagCell = row.getCell(7);
      flagCell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (opp.flag === 'Yes') { flagCell.fill = greenFill; flagCell.font = { bold: true, color: { argb: 'FF' + GREEN_FONT } }; }
      else if (opp.flag === 'No') { flagCell.fill = orangeFill; flagCell.font = { bold: true, color: { argb: 'FF' + ORANGE_FONT } }; }
      else if (opp.flag === 'Maybe') { flagCell.fill = yellowFill; flagCell.font = { bold: true, color: { argb: 'FF7D6608' } }; }
    }
    if (activeOpps.length === 0) {
      const row = ws4.addRow(['No improvement opportunities found.']);
      ws4.mergeCells(row.number, 1, row.number, 9);
      row.getCell(1).font = { italic: true, color: { argb: 'FF888888' } };
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const safeName = (project.name || 'export').replace(/[^a-z0-9]/gi, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="BESS_Review_${safeName}.xlsx"`);
    return res.send(Buffer.from(buffer));
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}];
