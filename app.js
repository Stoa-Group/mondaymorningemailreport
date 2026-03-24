/*
 * Monday Morning Report - Stoa Group
 * Data Processing and PDF Export Functionality
 */

// Global variables
let mmrData = [];
let googleReviewsData = [];

// Initialize the report when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    setupPDFExport();
    setupExcelExport();
    loadData();
});

// Setup PDF export functionality
function setupPDFExport() {
    const exportBtn = document.getElementById('export-pdf-btn');
    exportBtn.addEventListener('click', function() {
        exportToPDF();
    });
}

// Setup Excel export functionality
function setupExcelExport() {
    const excelBtn = document.getElementById('export-excel-btn');
    if (!excelBtn) return;
    excelBtn.addEventListener('click', function() {
        exportToExcel();
    });
}

// Build a single-sheet Excel export mirroring the on-screen tables
function exportToExcel() {
    try {
        const wb = XLSX.utils.book_new();

        // Determine sheet name from header week text
        const weekTextEl = document.getElementById('report-date');
        const weekText = weekTextEl ? weekTextEl.textContent.replace('Week Of: ', '').trim() : getWeekEndingDate();
        const sheetName = `MMR - ${weekText}`.substring(0, 31); // Excel sheet name limit

        // Helper to parse cell text into value and number format; also collapses spaces/line breaks
        const parseCell = (text) => {
            const t = (text || '').toString().replace(/\s+/g, ' ').trim();
            if (t === '') return { v: '', z: undefined };
            if (t.startsWith('$')) {
                const num = parseFloat(t.replace(/[$,]/g, ''));
                if (!isNaN(num)) return { v: num, z: '#,##0' };
            }
            if (t.endsWith('%')) {
                const num = parseFloat(t.replace('%', ''));
                if (!isNaN(num)) return { v: num / 100, z: '0.0%' };
            }
            const num = parseFloat(t.replace(/,/g, ''));
            if (!isNaN(num) && /^-?\d+[\d,]*(\.\d+)?$/.test(t)) return { v: num, z: undefined };
            return { v: t, z: undefined };
        };

        // Brand colors from CSS variables
        const COLORS = {
            primaryGreen: 'FF7E8A6B',
            primaryGrey: 'FF757270',
            secondaryGrey: 'FFEFEFF1',
            white: 'FFFFFFFF'
        };

        // Collect rows from each section in order
        const sections = [
            { id: 'occupancy-section', title: 'Occupancy' },
            { id: 'leasing-section', title: 'Leasing' },
            { id: 'renewals-section', title: 'Renewals & Collections' },
            { id: 'rentsf-section', title: 'Rents/SF' },
            { id: 'rents-section', title: 'Rents' },
            { id: 'income-section', title: 'Income' }
        ];

        const aoa = [];
        const merges = [];
        const rowTypes = []; // track styling: 'banner','title','header','body','total','blank'

        // Determine max column span across all tables (to merge banner/title rows cleanly)
        let maxColsAcross = 1;
        sections.forEach((sec) => {
            const sectionEl = document.getElementById(sec.id);
            if (!sectionEl) return;
            const table = sectionEl.querySelector('table');
            if (!table) return;
            const c = table.querySelectorAll('thead th').length || 1;
            if (c > maxColsAcross) maxColsAcross = c;
        });

        // Report banner (brand header like the on-screen header)
        const bannerTitle = 'Monday Morning Report';
        const bannerDate = `Week Of: ${weekText}`;
        const bannerRowIndex1 = aoa.length + 1;
        aoa.push([bannerTitle]);
        rowTypes.push('banner');
        merges.push({ s: { r: bannerRowIndex1 - 1, c: 0 }, e: { r: bannerRowIndex1 - 1, c: maxColsAcross - 1 } });
        const bannerRowIndex2 = aoa.length + 1;
        aoa.push([bannerDate]);
        rowTypes.push('banner');
        merges.push({ s: { r: bannerRowIndex2 - 1, c: 0 }, e: { r: bannerRowIndex2 - 1, c: maxColsAcross - 1 } });
        // Spacer after banner
        aoa.push([]); rowTypes.push('blank');

        const sectionsMeta = [];
        sections.forEach((sec) => {
            const sectionEl = document.getElementById(sec.id);
            if (!sectionEl) return;
            const table = sectionEl.querySelector('table');
            if (!table) return;

            // Section title row
            const titleRow0 = aoa.length;
            const titleRowIndex = aoa.length + 1; // 1-based for merges
            const sectionTitle = sectionEl.querySelector('.section-title')?.textContent?.trim() || sec.title;
            const colCount = table.querySelectorAll('thead th').length || 1;
            aoa.push([sectionTitle]);
            // Merge across all columns for title
            merges.push({ s: { r: titleRowIndex - 1, c: 0 }, e: { r: titleRowIndex - 1, c: Math.max(colCount - 1, 0) } });
            rowTypes.push('title');

            // Header row
            const headerCells = Array.from(table.querySelectorAll('thead th')).map(th => (th.innerText || th.textContent || '').replace(/\s+/g, ' ').trim());
            let headerRow0 = -1;
            if (headerCells.length > 0) { aoa.push(headerCells); rowTypes.push('header'); headerRow0 = aoa.length - 1; }

            // Body rows
            const rows = Array.from(table.querySelectorAll('tbody tr'));
            const firstDataRow0 = aoa.length;
            rows.forEach((tr) => {
                const cells = Array.from(tr.querySelectorAll('td')).map(td => (td.innerText || td.textContent || '').replace(/\s+/g, ' ').trim());
                aoa.push(cells);
                rowTypes.push('body');
            });
            let lastDataRow0 = aoa.length - 1;

            // Total row from tfoot (if present)
            const totalRow = table.querySelector('tfoot tr');
            if (totalRow) {
                const tdsEls = Array.from(totalRow.querySelectorAll('td'));
                const out = [];
                tdsEls.forEach((td) => {
                    const span = Number(td.getAttribute('colspan') || td.colSpan || 1);
                    const txt = (td.innerText || td.textContent || '').replace(/\s+/g, ' ').trim();
                    out.push(txt);
                    if (span > 1) {
                        for (let i = 1; i < span; i++) out.push('');
                    }
                });
                // Force a single blank under Location (second column) so totals begin at column C
                const firstSpan = Number(tdsEls[0] && (tdsEls[0].getAttribute('colspan') || tdsEls[0].colSpan) || 1);
                if (firstSpan < 2) {
                    out.splice(1, 0, '');
                }
                // Normalize row length to header column count
                if (out.length > colCount) {
                    out.length = colCount;
                }
                while (out.length < colCount) out.push('');
                aoa.push(out);
                rowTypes.push('total');
                const totalRow0 = aoa.length - 1;
                lastDataRow0 = totalRow0 - 1;
                sectionsMeta.push({ headerRow0, firstDataRow0, lastDataRow0, totalRow0, colCount });
            }

            // Blank spacer
            aoa.push([]);
            rowTypes.push('blank');
        });

        // Convert AOA to sheet, then apply basic styling and formats
        const ws = XLSX.utils.aoa_to_sheet(aoa);

        // Apply merges (section titles spanning columns)
        if (!ws['!merges']) ws['!merges'] = [];
        ws['!merges'] = ws['!merges'].concat(merges);

        // Insert formulas into total rows for transparency
        sectionsMeta.forEach(meta => {
            if (meta.headerRow0 < 0) return;
            const headers = aoa[meta.headerRow0] || [];
            const unitsCol = headers.findIndex(h => /Total\s*Units/i.test(h));
            if (unitsCol < 0) return;
            const unitsColLetter = XLSX.utils.encode_col(unitsCol);
            const r1 = meta.firstDataRow0 + 1;
            const r2 = meta.lastDataRow0 + 1;
            // Helpful lookups for special ratios in Renewals & Leasing
            const colIndexBy = (regex) => headers.findIndex(h => regex.test(h || ''));
            const inServiceCol = colIndexBy(/In-?Service/i);
            const delinquentCol = colIndexBy(/Delinquent(?!%)/i);
            const expiredCol = colIndexBy(/T-?12\s*Expired/i);
            const renewedCol = colIndexBy(/T-?12\s*Renewed/i);
            const grossCol = colIndexBy(/Gross\s*Leased/i);
            const visitsCol = colIndexBy(/Visits/i);
            const netLeasesCol = colIndexBy(/Net\s*Leases/i);
            for (let c = 0; c < meta.colCount; c++) {
                if (c < 2) continue; // keep label and blank under Location
                const header = headers[c] || '';
                const colLetter = XLSX.utils.encode_col(c);
                const totalAddr = XLSX.utils.encode_cell({ r: meta.totalRow0, c });
                let formula = '';
                if (/Total\s*Units/i.test(header)) {
                    formula = `SUM(${colLetter}${r1}:${colLetter}${r2})`;
                } else if (/(Visits|Gross\s*Leased|Canceled|Denied|Net\s*Leases|Move-Ins|Move-Outs|Delta\s*\(Units\)|Delinquent(?!% )|T-?12\s*Expired|T-?12\s*Renewed|In-?Service)/i.test(header)) {
                    formula = `SUM(${colLetter}${r1}:${colLetter}${r2})`;
                } else if (/Closing\s*Ratio/i.test(header) && grossCol >= 0 && visitsCol >= 0) {
                    const grossL = XLSX.utils.encode_col(grossCol);
                    const visitsL = XLSX.utils.encode_col(visitsCol);
                    formula = `IFERROR(SUM(${grossL}${r1}:${grossL}${r2})/SUM(${visitsL}${r1}:${visitsL}${r2}),0)`;
                } else if (/%\s*Gain/i.test(header) && netLeasesCol >= 0) {
                    const netL = XLSX.utils.encode_col(netLeasesCol);
                    formula = `IFERROR(SUM(${netL}${r1}:${netL}${r2})/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2}),0)`;
                } else if (/%\s*Gain\s*\/\s*Loss/i.test(header)) {
                    // Occupancy section: (Move-Ins - Move-Outs) / Units
                    const miCol = colIndexBy(/Move-?Ins/i);
                    const moCol = colIndexBy(/Move-?Outs/i);
                    if (miCol >= 0 && moCol >= 0) {
                        const miL = XLSX.utils.encode_col(miCol);
                        const moL = XLSX.utils.encode_col(moCol);
                        formula = `IFERROR((SUM(${miL}${r1}:${miL}${r2})-SUM(${moL}${r1}:${moL}${r2}))/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2}),0)`;
                    }
                } else if (/%\s*Delinquent/i.test(header) && delinquentCol >= 0 && inServiceCol >= 0) {
                    const delL = XLSX.utils.encode_col(delinquentCol);
                    const svcL = XLSX.utils.encode_col(inServiceCol);
                    formula = `IFERROR(SUM(${delL}${r1}:${delL}${r2})/SUM(${svcL}${r1}:${svcL}${r2}),0)`;
                } else if (/Renewal\s*Rate/i.test(header) && renewedCol >= 0 && expiredCol >= 0) {
                    const renL = XLSX.utils.encode_col(renewedCol);
                    const expL = XLSX.utils.encode_col(expiredCol);
                    formula = `IFERROR(SUM(${renL}${r1}:${renL}${r2})/SUM(${expL}${r1}:${expL}${r2}),0)`;
                } else if (/%\s*Difference/i.test(header)) {
                    // Build from weighted averages of occupied vs budgeted
                    const occCol = colIndexBy(/Occupied\s*Rent(?!\/SF)|Occupied\s*Rent\/SF/i);
                    const budCol = colIndexBy(/Budgeted\s*Rent(?!\/SF)|Budgeted\s*Rent\/SF/i);
                    if (occCol >= 0 && budCol >= 0) {
                        const oL = XLSX.utils.encode_col(occCol);
                        const bL = XLSX.utils.encode_col(budCol);
                        const avgO = `SUMPRODUCT(${oL}${r1}:${oL}${r2},${unitsColLetter}${r1}:${unitsColLetter}${r2})/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2})`;
                        const avgB = `SUMPRODUCT(${bL}${r1}:${bL}${r2},${unitsColLetter}${r1}:${unitsColLetter}${r2})/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2})`;
                        formula = `IFERROR((${avgO})/(${avgB})-1,0)`;
                    }
                } else if (/%|Occupancy|Leased|Projection|Gain|Ratio|Rate|Difference|vs Budget/i.test(header)) {
                    formula = `SUMPRODUCT(${colLetter}${r1}:${colLetter}${r2},${unitsColLetter}${r1}:${unitsColLetter}${r2})/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2})`;
                } else if (/Rent\/SF/i.test(header)) {
                    formula = `SUMPRODUCT(${colLetter}${r1}:${colLetter}${r2},${unitsColLetter}${r1}:${unitsColLetter}${r2})/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2})`;
                } else if (/Rent|Income/i.test(header)) {
                    formula = `SUMPRODUCT(${colLetter}${r1}:${colLetter}${r2},${unitsColLetter}${r1}:${unitsColLetter}${r2})/SUM(${unitsColLetter}${r1}:${unitsColLetter}${r2})`;
                }
                if (formula) {
                    ws[totalAddr] = ws[totalAddr] || {};
                    ws[totalAddr].f = formula;
                    ws[totalAddr].t = 'n';
                    // Ensure percent-looking totals are formatted as percentages
                    const isPercentLike = /%|Ratio|Rate|Difference|vs Budget|Gain|Projection|Occupancy|Leased/i.test(header);
                    ws[totalAddr].s = ws[totalAddr].s || {};
                    if (isPercentLike) {
                        ws[totalAddr].z = '0.0%';
                        ws[totalAddr].s.numFmt = '0.0%';
                    }
                }
            }
        });

        // Insert per-row formulas for calculated columns (closing ratio, % gain, renewal rate, % delinquent, rent diffs)
        sectionsMeta.forEach(meta => {
            if (meta.headerRow0 < 0) return;
            const headers = aoa[meta.headerRow0] || [];
            const colIndexBy = (regex) => headers.findIndex(h => regex.test(h || ''));
            const visitsCol = colIndexBy(/Visits/i);
            const grossCol = colIndexBy(/Gross\s*Leased/i);
            const netLeasesCol = colIndexBy(/Net\s*Leases/i);
            const unitsCol = colIndexBy(/Total\s*Units/i);
            const closingCol = colIndexBy(/Closing\s*Ratio/i);
            const gainCol = colIndexBy(/%\s*Gain/i);
            const gainLossCol = colIndexBy(/%\s*Gain\s*\/\s*Loss/i);
            const moveInsCol = colIndexBy(/Move-?Ins/i);
            const moveOutsCol = colIndexBy(/Move-?Outs/i);
            const expiredCol = colIndexBy(/T-?12\s*Expired/i);
            const renewedCol = colIndexBy(/T-?12\s*Renewed/i);
            const renewalRateCol = colIndexBy(/Renewal\s*Rate/i);
            const inServiceCol = colIndexBy(/In-?Service/i);
            const delinquentCol = colIndexBy(/Delinquent(?!%)/i);
            const delinquentPctCol = colIndexBy(/%\s*Delinquent/i);
            const occRentCol = colIndexBy(/Occupied\s*Rent(?!\/SF)/i);
            const budgRentCol = colIndexBy(/Budgeted\s*Rent(?!\/SF)/i);
            const diffPctCol = colIndexBy(/%\s*Difference/i);
            const moveInRentCol = colIndexBy(/Move-?In\s*Rent(?!\/SF)/i);
            const moveInVsBudgCol = colIndexBy(/Move-?In\s*vs\s*Budget|Move-?In\s*vs\s*Budget\s*%|Move-?In\s*vs\s*Budget\s*%|Move-?In\s*vs\s*Budget\s*%/i);
            const occRentSFCol = colIndexBy(/Occupied\s*Rent\/SF/i);
            const budgRentSFCol = colIndexBy(/Budgeted\s*Rent\/SF/i);
            const diffPctSFCol = colIndexBy(/%\s*Difference/i); // shared in rentsSF
            const moveInRentSFCol = colIndexBy(/Move-?in\s*Rent\/SF/i);
            const moveInVsBudgSFCol = colIndexBy(/%\s*vs\s*Budget/i);

            for (let r0 = meta.firstDataRow0; r0 <= meta.lastDataRow0; r0++) {
                const r1 = r0 + 1; // 1-based
                const setFormula = (cIdx, f, numFmt) => {
                    if (cIdx < 0) return;
                    const addr = XLSX.utils.encode_cell({ r: r0, c: cIdx });
                    ws[addr] = ws[addr] || {};
                    ws[addr].f = f;
                    ws[addr].t = 'n';
                    ws[addr].s = ws[addr].s || {};
                    if (numFmt) {
                        ws[addr].z = numFmt;
                        ws[addr].s.numFmt = numFmt;
                    }
                };

                if (closingCol >= 0 && grossCol >= 0 && visitsCol >= 0) {
                    const gL = XLSX.utils.encode_col(grossCol);
                    const vL = XLSX.utils.encode_col(visitsCol);
                    setFormula(closingCol, `IFERROR(${gL}${r1}/${vL}${r1},0)`, '0.0%');
                }
                if (gainCol >= 0 && netLeasesCol >= 0 && unitsCol >= 0) {
                    const nL = XLSX.utils.encode_col(netLeasesCol);
                    const uL = XLSX.utils.encode_col(unitsCol);
                    setFormula(gainCol, `IFERROR(${nL}${r1}/${uL}${r1},0)`, '0.0%');
                }
                if (gainLossCol >= 0 && moveInsCol >= 0 && moveOutsCol >= 0 && unitsCol >= 0) {
                    const miL = XLSX.utils.encode_col(moveInsCol);
                    const moL = XLSX.utils.encode_col(moveOutsCol);
                    const uL = XLSX.utils.encode_col(unitsCol);
                    setFormula(gainLossCol, `IFERROR((${miL}${r1}-${moL}${r1})/${uL}${r1},0)`, '0.0%');
                }
                if (renewalRateCol >= 0 && renewedCol >= 0 && expiredCol >= 0) {
                    const renL = XLSX.utils.encode_col(renewedCol);
                    const expL = XLSX.utils.encode_col(expiredCol);
                    setFormula(renewalRateCol, `IFERROR(${renL}${r1}/${expL}${r1},0)`, '0.0%');
                }
                if (delinquentPctCol >= 0 && delinquentCol >= 0 && inServiceCol >= 0) {
                    const delL = XLSX.utils.encode_col(delinquentCol);
                    const svcL = XLSX.utils.encode_col(inServiceCol);
                    setFormula(delinquentPctCol, `IFERROR(${delL}${r1}/${svcL}${r1},0)`, '0.0%');
                }
                if (diffPctCol >= 0 && occRentCol >= 0 && budgRentCol >= 0) {
                    const oL = XLSX.utils.encode_col(occRentCol);
                    const bL = XLSX.utils.encode_col(budgRentCol);
                    setFormula(diffPctCol, `IFERROR((${oL}${r1}-${bL}${r1})/${bL}${r1},0)`, '0.0%');
                }
                if (moveInVsBudgCol >= 0 && moveInRentCol >= 0 && budgRentCol >= 0) {
                    const mL = XLSX.utils.encode_col(moveInRentCol);
                    const bL = XLSX.utils.encode_col(budgRentCol);
                    setFormula(moveInVsBudgCol, `IFERROR((${mL}${r1}-${bL}${r1})/${bL}${r1},0)`, '0.0%');
                }
                if (diffPctSFCol >= 0 && occRentSFCol >= 0 && budgRentSFCol >= 0) {
                    const oL = XLSX.utils.encode_col(occRentSFCol);
                    const bL = XLSX.utils.encode_col(budgRentSFCol);
                    setFormula(diffPctSFCol, `IFERROR((${oL}${r1}-${bL}${r1})/${bL}${r1},0)`, '0.0%');
                }
                if (moveInVsBudgSFCol >= 0 && moveInRentSFCol >= 0 && budgRentSFCol >= 0) {
                    const mL = XLSX.utils.encode_col(moveInRentSFCol);
                    const bL = XLSX.utils.encode_col(budgRentSFCol);
                    setFormula(moveInVsBudgSFCol, `IFERROR((${mL}${r1}-${bL}${r1})/${bL}${r1},0)`, '0.0%');
                }
            }
        });

        // Column widths: derive from content length for efficient fit
        const maxCols = Math.max(...aoa.map(r => r.length));
        const colCharWidths = new Array(maxCols).fill(0);
        aoa.forEach(row => {
            for (let i = 0; i < maxCols; i++) {
                const v = (row[i] === undefined || row[i] === null) ? '' : String(row[i]);
                if (v.length > colCharWidths[i]) colCharWidths[i] = v.length;
            }
        });
        ws['!cols'] = new Array(maxCols).fill(null).map((_, c) => {
            if (c === 0) return { wch: Math.min(Math.max(colCharWidths[c] + 2, 20), 32) };
            if (c === 1) return { wch: Math.min(Math.max(colCharWidths[c] + 2, 16), 26) };
            // Compact numeric columns
            return { wch: Math.min(Math.max(colCharWidths[c] + 2, 10), 16) };
        });

        // Helpers to classify columns per-section based on header text
        const isPercentHeader = (h) => /%|Occupancy|Leased|Projection|Gain|Ratio|Rate|Difference|vs Budget/i.test(h);
        const isIntHeader = (h) => /(Delta \(Units\)|Move-Ins|Move-Outs|Visits|Gross Leased|Canceled|Denied|Net Leases|T-12|Expired|Renewed|In-Service|Delinquent)/i.test(h);
        const isPerSfHeader = (h) => /(Rent\/SF)/i.test(h);
        const isDecimalHeader = (h) => /(Rent|Income)/i.test(h); // numeric decimals, not percent; excludes Rent/SF handled above

        // Cell-by-cell parsing for number formats and detailed styling
        const range = XLSX.utils.decode_range(ws['!ref']);
        let bodyRowCounterSinceHeader = 0; // for zebra striping within each section
        let currentColumnFormats = {}; // per-section map: colIndex -> 'percent'|'int'|'decimal'
        for (let R = range.s.r; R <= range.e.r; R++) {
            const rowType = rowTypes[R] || 'body';
            if (rowType === 'header' || rowType === 'title' || rowType === 'banner') bodyRowCounterSinceHeader = 0; else if (rowType === 'body') bodyRowCounterSinceHeader++;
            if (rowType === 'header') {
                currentColumnFormats = {};
                const headerRow = aoa[R] || [];
                headerRow.forEach((label, cIdx) => {
                    const text = (label || '').toString();
                    if (isIntHeader(text)) currentColumnFormats[cIdx] = 'int';
                    else if (isPercentHeader(text)) currentColumnFormats[cIdx] = 'percent';
                    else if (isPerSfHeader(text)) currentColumnFormats[cIdx] = 'decimal2';
                    else if (isDecimalHeader(text)) currentColumnFormats[cIdx] = 'decimal0';
                });
            }

            for (let C = range.s.c; C <= range.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[addr];
                const text = (aoa[R] && aoa[R][C] !== undefined) ? aoa[R][C] : '';
                if (!cell) continue;

                // Parse numeric/currency/percentage cells
                const parsed = parseCell(text);
                cell.v = parsed.v;
                if (typeof cell.v === 'number') cell.t = 'n';
                if (parsed.z) {
                    cell.z = parsed.z; // fallback
                    cell.s = cell.s || {};
                    if (parsed.z === '#,##0') {
                        cell.s.numFmt = '#,##0';
                    } else if (parsed.z === '0.0%') {
                        cell.s.numFmt = '0.0%';
                    }
                }

                // Apply per-section column formats
                const desired = currentColumnFormats[C];
                if (desired === 'percent' && typeof cell.v === 'number') {
                    // Convert 0-100 values to 0-1 for Excel percent if necessary
                    if (cell.v > 1.1) cell.v = cell.v / 100;
                    cell.z = '0.0%';
                    cell.s = cell.s || {};
                    cell.s.numFmt = '0.0%';
                } else if (desired === 'int' && typeof cell.v === 'number') {
                    cell.v = Math.round(cell.v);
                    cell.z = '0';
                    cell.s = cell.s || {};
                    cell.s.numFmt = '0';
                } else if (desired === 'decimal0' && typeof cell.v === 'number') {
                    // Show whole-dollar values for rent/income metrics (less visual noise)
                    cell.z = '#,##0';
                    cell.s = cell.s || {};
                    cell.s.numFmt = '#,##0';
                } else if (desired === 'decimal2' && typeof cell.v === 'number') {
                    // Per-square-foot metrics should show two decimals
                    cell.z = '#,##0.00';
                    cell.s = cell.s || {};
                    cell.s.numFmt = '#,##0.00';
                }

                // Styling
                cell.s = cell.s || {};
                if (rowType === 'banner') {
                    cell.s = {
                        font: { bold: true, sz: 14, color: { rgb: COLORS.white } },
                        alignment: { horizontal: R % 2 === 0 ? 'left' : 'right', vertical: 'center', wrapText: true },
                        fill: { patternType: 'solid', fgColor: { rgb: COLORS.primaryGreen } }
                    };
                } else if (rowType === 'title') {
                    cell.s = {
                        font: { bold: true, sz: 12, color: { rgb: COLORS.primaryGreen } },
                        alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
                    };
                } else if (rowType === 'header') {
                    cell.s = {
                        font: { bold: true, color: { rgb: COLORS.white } },
                        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                        fill: { patternType: 'solid', fgColor: { rgb: COLORS.primaryGrey } }
                    };
                } else if (rowType === 'total') {
                    cell.s.fill = { patternType: 'solid', fgColor: { rgb: COLORS.secondaryGrey } };
                    cell.s.font = Object.assign({}, cell.s.font || {}, { bold: true });
                    cell.s.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
                } else if (rowType === 'body') {
                    // Zebra striping and centered cells
                    const isEven = bodyRowCounterSinceHeader % 2 === 0;
                    if (isEven) {
                        cell.s.fill = { patternType: 'solid', fgColor: { rgb: 'FFF1F1F1' } };
                    }
                    cell.s.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
                }
            }
        }

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        const fileName = `Monday_Morning_Report_${getCurrentDate()}.xlsx`;
        XLSX.writeFile(wb, fileName);
    } catch (err) {
        console.error('Excel export error:', err);
        showError('Failed to export Excel file.');
    }
}

// (Print button removed; single Export button remains)

// Preflight pass: if a block won't fit, push it to next page
function addSmartPageBreaks(pageHeightIn) {
    // Use the provided PDF page height (inches) to compute pixel page height
    const DPI = 96;
    const pageHeightPx = (pageHeightIn || 11) * DPI; // default to 11in if not provided
    let cursor = 0; // tracks used height within the current page

    // Reset any old markers
    document.querySelectorAll('.page-break-before').forEach(el => el.classList.remove('page-break-before'));

    const blocks = Array.from(document.querySelectorAll('.report-section .keep-together'));
    for (const block of blocks) {
        // Require that the section title and table header both appear at the top of a page
        const titleEl = block.querySelector('.section-title');
        const theadEl = block.querySelector('thead');
        const titleH = titleEl ? Math.ceil(titleEl.getBoundingClientRect().height) : 0;
        const theadH = theadEl ? Math.ceil(theadEl.getBoundingClientRect().height) : 0;
        const guardGap = 8; // small spacing to avoid tight collisions

        const remaining = pageHeightPx - (cursor % pageHeightPx);
        const requiredTopChunk = titleH + theadH + guardGap;

        // If title+header won't fit in the remaining space, start the block on next page
        if (requiredTopChunk > remaining) {
            block.classList.add('page-break-before');
            cursor = 0; // new page
        }

        // Advance cursor by the block's full height (after possibly forcing a page break)
        const blockH = Math.ceil(block.getBoundingClientRect().height);
        cursor += blockH + 12; // maintain inter-section spacing
        if (cursor > pageHeightPx) cursor = cursor % pageHeightPx;
    }
}

// Export report to PDF
function exportToPDF() {
    const element = document.getElementById('report-container');
    const body = document.body;

    // Add class to trigger PDF layout
    body.classList.add('pdf-generating');

    // Wait for browser to render at new size
    setTimeout(() => {
        // Decide PDF page format and pass height to preflight
        const desiredFormat = [11.33, 14.67]; // current custom size; change to 'letter' to use 11in height
        const pageHeightIn = Array.isArray(desiredFormat) ? desiredFormat[1] : 11;

        // Preflight: measure blocks and add page-break-before where needed
        addSmartPageBreaks(pageHeightIn);

        const opt = {
            margin: [0, 0],                   // Match @page margins (no padding)
            filename: `Monday_Morning_Report_${getCurrentDate()}.pdf`,
            image: { type: 'png', quality: 1.0 }, // PNG yields crisper text
            html2canvas: { 
                scale: 3,                      // higher render scale for sharper output
                useCORS: true, 
                allowTaint: true,
                letterRendering: true,
                logging: false,
                dpi: 288                       // higher DPI for sharper PDF
            },
            jsPDF: { 
                unit: 'in', 
                format: desiredFormat,        // Keep in sync with preflight height
                orientation: 'portrait', 
                compress: true 
            },
            pagebreak: {
                mode: ['css', 'legacy'],        // Honor CSS breaks first
                before: ['.page-break-before'], // JS will add as needed
                avoid:  ['.keep-together']      // Try to keep these blocks intact
            },
            enableLinks: false
        };

        // Show loading state
        const btn = document.getElementById('export-pdf-btn');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span>Generating PDF...</span>';
        btn.disabled = true;

        html2pdf().set(opt).from(element).save().then(() => {
            body.classList.remove('pdf-generating');
            btn.innerHTML = originalText;
            btn.disabled = false;
        }).catch(err => {
            console.error('PDF generation error:', err);
            body.classList.remove('pdf-generating');
            btn.innerHTML = originalText;
            btn.disabled = false;
        });
    }, 100);
}

// Generate PDF for a single page
function generateSinglePagePDF(element) {
    return new Promise((resolve, reject) => {
        const opt = {
            margin: [0.2, 0.15],
            filename: 'temp-page.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 2, 
                useCORS: true, 
                allowTaint: true, 
                letterRendering: true, 
                dpi: 192 
            },
            jsPDF: { 
                unit: 'in', 
                format: 'letter', 
                orientation: 'portrait', 
                compress: true 
            }
        };
        
        html2pdf()
            .set(opt)
            .from(element)
            .outputPdf('blob')
            .then(resolve)
            .catch(reject);
    });
}

// Combine multiple PDF blobs into one
async function combinePDFs(pdfBlobs) {
    const { PDFDocument } = PDFLib;
    const mergedPdf = await PDFDocument.create();
    
    for (const blob of pdfBlobs) {
        const arrayBuffer = await blob.arrayBuffer();
        const pdf = await PDFDocument.load(arrayBuffer);
        const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
        pages.forEach((page) => mergedPdf.addPage(page));
    }
    
    const pdfBytes = await mergedPdf.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Monday_Morning_Report_${getCurrentDate()}.pdf`;
    a.click();
    URL.revokeObjectURL(url);
    
    const btn = document.getElementById('export-pdf-btn');
    const originalText = btn.innerHTML;
    btn.innerHTML = originalText;
    btn.disabled = false;
}

// OLD APPROACH - Keeping for reference but replacing with above
function exportToPDF_OLD() {
    const element = document.getElementById('report-container');
    const body = document.body;
    const html = document.documentElement;

    // Add class to trigger scaling
    body.classList.add('pdf-generating');
    
    // Wait for browser to render at new size
    setTimeout(() => {
        // Apply page breaks every 2 sections using h2 elements
        const sections = document.querySelectorAll('.report-section');
        console.log('=== PDF PAGE BREAK LOGGING ===');
        console.log(`Total sections found: ${sections.length}`);
        
        sections.forEach((section, index) => {
            const sectionTitle = section.querySelector('.section-title');
            const sectionText = sectionTitle ? sectionTitle.textContent.trim() : 'Unknown';
            const tableContainer = section.querySelector('.table-container');
            const table = tableContainer ? tableContainer.querySelector('.data-table') : null;
            const thead = table ? table.querySelector('thead') : null;
            const tbody = table ? table.querySelector('tbody') : null;
            
            console.log(`\nSection ${index + 1}: "${sectionText}"`);
            console.log(`  - Index: ${index}, Condition check: (index > 0 && index % 2 === 0) = ${index > 0 && index % 2 === 0}`);
            
            // Log DOM element positions and heights
            if (sectionTitle) {
                const titleRect = sectionTitle.getBoundingClientRect();
                console.log(`  - Section Title Position: top:${titleRect.top}px, height:${titleRect.height}px`);
            }
            
            if (thead) {
                const theadRect = thead.getBoundingClientRect();
                console.log(`  - Table Head Position: top:${theadRect.top}px, height:${theadRect.height}px`);
            }
            
            if (table) {
                const tableRect = table.getBoundingClientRect();
                console.log(`  - Table Position: top:${tableRect.top}px, height:${tableRect.height}px`);
            }
            
            if (index > 0 && index % 2 === 0 && sectionTitle) {
                sectionTitle.classList.add('force-page-break');
                
                // Add explicit page break element before the section
                const pageBreakDiv = document.createElement('div');
                pageBreakDiv.className = 'explicit-page-break';
                pageBreakDiv.style.pageBreakBefore = 'always';
                pageBreakDiv.style.height = '0';
                pageBreakDiv.style.margin = '0';
                pageBreakDiv.style.padding = '0';
                pageBreakDiv.style.display = 'block';
                
                section.parentNode.insertBefore(pageBreakDiv, section);
                
                console.log(`  - ✅ PAGE BREAK APPLIED - This section will start on new page`);
                console.log(`  - Section title element:`, sectionTitle);
                console.log(`  - Class added: force-page-break`);
                console.log(`  - Explicit page break div added before section`);
            } else {
                console.log(`  - No page break (stays with previous section)`);
            }
        });
        
        console.log('\n=== DOM ELEMENT ANALYSIS ===');
        document.querySelectorAll('.force-page-break').forEach((el, index) => {
            const computedStyle = window.getComputedStyle(el);
            console.log(`Force page break element ${index + 1}:`, {
                text: el.textContent.trim(),
                classes: el.className,
                pageBreakBefore: computedStyle.pageBreakBefore,
                breakBefore: computedStyle.breakBefore,
                element: el
            });
            
            // Test if our CSS is working
            if (computedStyle.pageBreakBefore === 'always') {
                console.log(`✅ CSS WORKING: ${el.textContent.trim()} has page-break-before: always`);
            } else {
                console.log(`❌ CSS NOT WORKING: ${el.textContent.trim()} has page-break-before: ${computedStyle.pageBreakBefore}`);
            }
        });
        
        console.log('\n=== TABLE ROW ANALYSIS ===');
        document.querySelectorAll('.data-table tr').forEach((tr, index) => {
            if (index < 5) { // Only log first 5 rows to avoid spam
                const computedStyle = window.getComputedStyle(tr);
                console.log(`Table row ${index + 1}:`, {
                    pageBreakInside: computedStyle.pageBreakInside,
                    breakInside: computedStyle.breakInside,
                    element: tr
                });
                
                // Test if our CSS is working
                if (computedStyle.pageBreakInside === 'avoid') {
                    console.log(`✅ CSS WORKING: Table row ${index + 1} has page-break-inside: avoid`);
                } else {
                    console.log(`❌ CSS NOT WORKING: Table row ${index + 1} has page-break-inside: ${computedStyle.pageBreakInside}`);
                }
            }
        });
        
        console.log('\n=== SECTION ANALYSIS ===');
        document.querySelectorAll('.report-section').forEach((section, index) => {
            const computedStyle = window.getComputedStyle(section);
            console.log(`Section ${index + 1}:`, {
                pageBreakInside: computedStyle.pageBreakInside,
                breakInside: computedStyle.breakInside,
                element: section
            });
            
            // Test if our CSS is working
            if (computedStyle.pageBreakInside === 'avoid') {
                console.log(`✅ CSS WORKING: Section ${index + 1} has page-break-inside: avoid`);
            } else {
                console.log(`❌ CSS NOT WORKING: Section ${index + 1} has page-break-inside: ${computedStyle.pageBreakInside}`);
            }
        });

        const opt = {
            margin: [0.2, 0.15],
            filename: `Monday_Morning_Report_${getCurrentDate()}.pdf`,
            image: { 
                type: 'jpeg', 
                quality: 0.98 
            },
            html2canvas: { 
                scale: 2,
                useCORS: true,
                logging: true,
                allowTaint: true,
                letterRendering: true,
                dpi: 192
            },
            pagebreak: { 
                mode: ['legacy', 'avoid-all'],
                before: ['.force-page-break', '.explicit-page-break'],
                avoid: ['.data-table tr', '.report-section']
            },
            jsPDF: { 
                unit: 'in', 
                format: 'letter', 
                orientation: 'portrait', 
                compress: true 
            },
            enableLinks: false
        };
        
        console.log('\n=== HTML2PDF OPTIONS ===');
        console.log('PDF Format:', opt.jsPDF.format);
        console.log('Page Break Before Selector:', opt.pagebreak.before);
        console.log('Page Break Avoid:', opt.pagebreak.avoid);
        console.log('Page Break Modes:', opt.pagebreak.mode);
        console.log('Margin:', opt.margin);
        console.log('HTML2Canvas Scale:', opt.html2canvas.scale);
        console.log('AutoPaging:', opt.autoPaging || 'undefined');
        console.log('Full options object:', JSON.stringify(opt, null, 2));

        // Apply simple formatting - just ensure tables are visible
        
        // Show loading state
        const btn = document.getElementById('export-pdf-btn');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span>Generating PDF...</span>';
        btn.disabled = true;
        
        console.log('\n=== STARTING PDF GENERATION ===');
        console.log('Element to convert:', element);
        console.log('Element dimensions:', {
            width: element.offsetWidth,
            height: element.offsetHeight
        });
        console.log('Number of sections in element:', element.querySelectorAll('.report-section').length);
        console.log('Number of force-page-break elements:', element.querySelectorAll('.force-page-break').length);
        console.log('Number of table rows:', element.querySelectorAll('.data-table tr').length);
        
        // Create html2pdf instance with logging
        const pdfInstance = html2pdf().set(opt).from(element);
        console.log('HTML2PDF instance created:', pdfInstance);
        
        pdfInstance.save().then(() => {
            console.log('\n=== PDF GENERATION COMPLETE ===');
            console.log('PDF saved successfully');
            body.classList.remove('pdf-generating');
            btn.innerHTML = originalText;
            btn.disabled = false;
        }).catch(err => {
            console.error('\n=== PDF GENERATION ERROR ===');
            console.error('Error details:', err);
            console.error('Error stack:', err.stack);
            body.classList.remove('pdf-generating');
            btn.innerHTML = originalText;
            btn.disabled = false;
        });
    }, 250); // Give more time for scaling to render
}

// No custom formatting needed - html2pdf handles pagination naturally

function inchesToPixels(inches) {
    // Convert inches to pixels using 96 dpi (browser default).
    return inches * 96;
}

// Get current date formatted as YYYY-MM-DD
function getCurrentDate() {
    const now = new Date();
    return now.toISOString().split('T')[0];
}

// Robust numeric parser: handles $, commas, and percent signs
function toNumber(value) {
    if (value === null || value === undefined) return 0;
    if (typeof value === 'number') return value;
    const cleaned = String(value).replace(/[$,%\s,]/g, '');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
}

// Get current week ending date
function getWeekEndingDate() {
    const now = new Date();
    const daysSinceFriday = (now.getDay() + 2) % 7;
    const weekEnding = new Date(now);
    weekEnding.setDate(now.getDate() - daysSinceFriday);
    return weekEnding.toISOString().split('T')[0];
}

// PropertyList API: fetch authoritative Status from the database
var PROPERTY_LIST_API = "https://stoagroupdb-ddre.onrender.com/api/leasing/property-list";

function fetchPropertyListStatus() {
    return fetch(PROPERTY_LIST_API)
        .then(function(res) {
            if (!res.ok) throw new Error('HTTP ' + res.status);
            return res.json();
        })
        .then(function(json) {
            var map = {};
            (json.data || []).forEach(function(p) {
                var name = (p.Property || '').trim().toLowerCase();
                if (name) map[name] = (p.Status || '').trim();
            });
            console.log('[PropertyList] Loaded ' + Object.keys(map).length + ' authoritative statuses from DB');
            return map;
        })
        .catch(function(e) {
            console.warn('[PropertyList] Could not fetch DB statuses, using Domo Status as-is:', e);
            return {};
        });
}

function overlayDbStatus(rows, statusMap) {
    if (!statusMap || Object.keys(statusMap).length === 0) return rows;
    rows.forEach(function(r) {
        var prop = (r.Property || '').toString().trim().toLowerCase();
        if (prop && statusMap[prop]) {
            r.Status = statusMap[prop];
        }
    });
    return rows;
}

// Load data from Domo
function loadData() {
    // Fetch PropertyList (DB authoritative status) in parallel with MMR data
    Promise.all([
        domo.get('/data/v2/MMRDATA?limit=10000'),
        fetchPropertyListStatus()
    ])
        .then(function(results) {
            var data = results[0];
            var dbStatus = results[1];

            console.log("mmrData", data);
            var allData = data;
            
            if (!allData || allData.length === 0) {
                showError('No data received from mmrData dataset.');
                return;
            }

            // Overlay authoritative Status from database before filtering
            overlayDbStatus(allData, dbStatus);
            
            // Filter to only the most recent week
            mmrData = filterToMostRecentWeek(allData);
            
            if (mmrData.length === 0) {
                showError('No data found for the most recent week. Please check your data.');
                return;
            }
            
            // Update report date with the actual ReportDate from data
            var actualReportDate = mmrData[0].ReportDate || getWeekEndingDate();
            document.getElementById('report-date').textContent = 'Week Of: ' + actualReportDate;
            
            // Process and display data
            populateOccupancyTable();
            populateLeasingTable();
            populateRenewalsTable();
            populateRentsTable();
            populateRentsFTable();
            populateIncomeTable();
        })
        .catch(function(error) {
            console.error('Error loading data:', error);
            showError('Failed to load data: ' + (error.message || error.toString()));
        });
    
    // Load Google Reviews data
    // Commented out per request
    /*
    domo.get('/data/v2/googleReviews?limit=5000')
        .then(function(data){
            console.log("googleReviews", data);
            googleReviewsData = data;
            populateReviewsTable();
        })
        .catch(function(error) {
            console.error('Error loading googleReviews:', error);
            googleReviewsData = [];
        });
    */
}

// Filter data to only include the most recent week per property
function filterToMostRecentWeek(data) {
    if (!data || data.length === 0) {
        console.warn('No data provided to filter');
        return [];
    }
    
    const statusFiltered = data.filter(function(row) {
        return row.Status === 'Lease-Up' || row.Status === 'Stabilized';
    });
    
    // Group by property and find most recent ReportDate for each
    const propertyMap = {};
    
    statusFiltered.forEach(function(row) {
        if (!row.Property || !row.ReportDate) {
            return;
        }
        
        const property = row.Property;
        const dateValue = new Date(row.ReportDate);
        
        if (!propertyMap[property] || dateValue > propertyMap[property].maxDate) {
            propertyMap[property] = {
                maxDate: dateValue,
                row: row
            };
        }
    });
    
    // Extract the most recent row for each property
    const filtered = Object.values(propertyMap)
        .map(obj => obj.row)
        .sort(function(a, b) {
            // Sort by BirthOrder (ascending - oldest first)
            const birthOrderA = parseInt(a.BirthOrder || 0);
            const birthOrderB = parseInt(b.BirthOrder || 0);
            return birthOrderA - birthOrderB;
        });
    
    return filtered;
}

// Populate Occupancy Table
function populateOccupancyTable() {
    const tbody = document.getElementById('occupancy-tbody');
    tbody.innerHTML = '';
    
    let totalUnits = 0;
    let totalOccupied = 0;
    let totalMoveIns = 0;
    let totalMoveOuts = 0;
    let budgetedOccupancySum = 0;
    let totalDelta = 0;
    
    mmrData.forEach(function(row) {
        const occupancyPercent = toNumber(row.OccupancyPercent) || 0;
        const budgetedOccupancy = toNumber(row.BudgetedOccupancyPercentCurrentMonth) || 0;
        const moveIns = parseInt(row.MI || 0);
        const moveOuts = parseInt(row.MO || 0);
        const totalUnitsForRow = parseInt(row.TotalUnits || 0);
        const netChange = moveIns - moveOuts;
        const gainLoss = totalUnitsForRow > 0 ? (netChange / totalUnitsForRow) : 0; // decimal; formatter will convert to %
        
        // Calculate delta (units): (Current % * Total) - (Budgeted % * Total)
        const totalUnitsValue = toNumber(row.TotalUnits || 0);
        
        // Handle both decimal (0.939) and percentage (93.9) formats
        let actualOccupancyValue = toNumber(occupancyPercent) || 0;
        let budgetedOccupancyValue = toNumber(budgetedOccupancy) || 0;
        
        // If value is between -1 and 1 (and not 0), it's a decimal; otherwise it's a percentage
        if (actualOccupancyValue > -1 && actualOccupancyValue < 1 && actualOccupancyValue !== 0) {
            // Already a decimal, use as-is
        } else {
            // It's a percentage, convert to decimal
            actualOccupancyValue = actualOccupancyValue / 100;
        }
        
        if (budgetedOccupancyValue > -1 && budgetedOccupancyValue < 1 && budgetedOccupancyValue !== 0) {
            // Already a decimal, use as-is
        } else {
            // It's a percentage, convert to decimal
            budgetedOccupancyValue = budgetedOccupancyValue / 100;
        }
        
        const actualOccupiedUnits = Math.round(totalUnitsValue * actualOccupancyValue);
        const budgetedOccupiedUnits = Math.round(totalUnitsValue * budgetedOccupancyValue);
        let delta = actualOccupiedUnits - budgetedOccupiedUnits;
        // Align with Leasing Analytics Hub: at 0% occupancy, delta vs budget is 0 (not a large negative)
        if (actualOccupiedUnits === 0 || toNumber(occupancyPercent) === 0) delta = 0;
        
        // Use Week4 data for projection
        const week4Proj = toNumber(row.Week4OccPercent) || 0;
        // Use Week8 data (from Week7 as 8th week) for 8 week projection
        const week8Proj = toNumber(row.Week7OccPercent) || 0;
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.Property || ''}</td>
            <td>${row.City || ''}, ${row.State || ''}</td>
            <td class="total-units">${totalUnitsValue}</td>
            <td>${formatPercent(occupancyPercent)}</td>
            <td>${formatPercent(budgetedOccupancy)}</td>
            <td>${delta}</td>
            <td>${formatPercent(week4Proj)}</td>
            <td>${formatPercent(week8Proj)}</td>
            <td>${moveIns}</td>
            <td>${moveOuts}</td>
            <td>${formatPercent(gainLoss)}</td>
        `;
        tbody.appendChild(tr);
        
        totalUnits += totalUnitsValue;
        totalOccupied += actualOccupiedUnits;
        totalMoveIns += toNumber(row.MI || 0);
        totalMoveOuts += toNumber(row.MO || 0);
        budgetedOccupancySum += toNumber(budgetedOccupancy);
        totalDelta += delta;
    });
    
    // Update totals - all averages should be weighted by total units
    // Current Occupancy %: weighted average = (sum of occupied units) / (sum of total units) * 100
    const avgOccupancy = totalUnits > 0 ? (totalOccupied / totalUnits * 100) : 0;
    
    // Budgeted Occupancy %: weighted average = (sum of budgeted occupied units) / (sum of total units) * 100
    const totalBudgetedOccupiedUnits = mmrData.reduce((sum, row) => {
        const totalUnitsValue = toNumber(row.TotalUnits || 0);
        let budgetedPercent = parseFloat(row.BudgetedOccupancyPercentCurrentMonth) || 0;
        // Handle both decimal and percentage formats
        if (budgetedPercent > -1 && budgetedPercent < 1 && budgetedPercent !== 0) {
            // Already a decimal
        } else {
            budgetedPercent = budgetedPercent / 100;
        }
        return sum + Math.round(totalUnitsValue * budgetedPercent);
    }, 0);
    const avgBudgetedOccupancy = totalUnits > 0 ? (totalBudgetedOccupiedUnits / totalUnits * 100) : 0;
    
    // % Gain/Loss: weighted by total units
    const totalGainLoss = totalUnits > 0 ? ((totalMoveIns - totalMoveOuts) / totalUnits) : 0;
    
    // 4 Week Projection: weighted average = (sum of projected occupied units) / (sum of total units) * 100
    const totalWeek4OccupiedUnits = mmrData.reduce((sum, row) => {
        return sum + parseInt(row.Week4OccUnits || 0);
    }, 0);
    const avgWeek4 = totalUnits > 0 ? (totalWeek4OccupiedUnits / totalUnits * 100) : 0;
    
    // 8 Week Projection: weighted average = (sum of projected occupied units) / (sum of total units) * 100
    const totalWeek8OccupiedUnits = mmrData.reduce((sum, row) => {
        return sum + parseInt(row.Week7OccUnits || 0);
    }, 0);
    const avgWeek8 = totalUnits > 0 ? (totalWeek8OccupiedUnits / totalUnits * 100) : 0;
    
    document.getElementById('total-units').textContent = totalUnits;
    document.getElementById('current-occupancy').textContent = formatPercent(avgOccupancy);
    document.getElementById('budgeted-occupancy').textContent = formatPercent(avgBudgetedOccupancy);
    document.getElementById('total-delta').textContent = totalDelta;
    document.getElementById('total-move-ins').textContent = totalMoveIns;
    document.getElementById('total-move-outs').textContent = totalMoveOuts;
    document.getElementById('gain-loss').textContent = formatPercent(totalGainLoss);
    document.getElementById('week-4-proj').textContent = formatPercent(avgWeek4);
    document.getElementById('week-8-proj').textContent = formatPercent(avgWeek8);
}

// Populate Leasing Table
function populateLeasingTable() {
    const tbody = document.getElementById('leasing-tbody');
    tbody.innerHTML = '';
    
    let totals = {
        units: 0,
        visits: 0,
        grossLeased: 0,
        canceled: 0,
        denied: 0,
        netLeases: 0,
        leasedUnits: 0,  // Track total leased units for weighted average
        closingWeightedSum: 0 // sum(closingRatio * totalUnits)
    };
    
    mmrData.forEach(function(row) {
        const totalUnits = toNumber(row.TotalUnits || 0);
        const leasedPercent = toNumber(row.CurrentLeasedPercent) || 0;
        const visits = toNumber(row['1stVisit'] || 0) + toNumber(row.ReturnVisitCount || 0);
        const grossLeased = toNumber(row.Applied || 0);
        const canceled = toNumber(row.Canceled || 0);
        const denied = toNumber(row.Denied || 0);
        const netLeases = toNumber(row.NetLsd || 0);
        const closingRatio = visits > 0 ? (grossLeased / visits) : 0; // decimal
        const gainPercent = totalUnits > 0 ? (netLeases / totalUnits) : 0; // decimal
        
        // Calculate leased units for weighted average
        let leasedPercentValue = leasedPercent;
        // Handle both decimal and percentage formats
        if (leasedPercentValue > -1 && leasedPercentValue < 1 && leasedPercentValue !== 0) {
            // Already a decimal
        } else {
            leasedPercentValue = leasedPercentValue / 100;
        }
        const leasedUnits = Math.round(totalUnits * leasedPercentValue);
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.Property || ''}</td>
            <td>${row.City || ''}, ${row.State || ''}</td>
            <td class="total-units">${totalUnits}</td>
            <td>${formatPercent(leasedPercent)}</td>
            <td>${visits}</td>
            <td>${grossLeased}</td>
            <td>${canceled}</td>
            <td>${denied}</td>
            <td>${netLeases}</td>
            <td>${formatPercent(closingRatio, true)}</td>
            <td>${formatPercent(gainPercent)}</td>
        `;
        tbody.appendChild(tr);
        
        totals.units += totalUnits;
        totals.visits += visits;
        totals.grossLeased += grossLeased;
        totals.canceled += canceled;
        totals.denied += denied;
        totals.netLeases += netLeases;
        totals.leasedUnits += leasedUnits;
        totals.closingWeightedSum += closingRatio * totalUnits;
    });
    
    // Update totals - weighted average for Current Leased %
    const avgLeased = totals.units > 0 ? (totals.leasedUnits / totals.units * 100) : 0;
    const totalClosingRatio = totals.units > 0 ? (totals.closingWeightedSum / totals.units) : 0;
    const totalGain = totals.units > 0 ? (totals.netLeases / totals.units) : 0;
    
    document.getElementById('leasing-total-units').textContent = totals.units;
    document.getElementById('leasing-avg').textContent = formatPercent(avgLeased);
    document.getElementById('leasing-total-visits').textContent = totals.visits;
    document.getElementById('leasing-total-gross').textContent = totals.grossLeased;
    document.getElementById('leasing-total-canceled').textContent = totals.canceled;
    document.getElementById('leasing-total-denied').textContent = totals.denied;
    document.getElementById('leasing-total-net').textContent = totals.netLeases;
    document.getElementById('leasing-total-ratio').textContent = formatPercent(totalClosingRatio, true);
    document.getElementById('leasing-total-gain').textContent = formatPercent(totalGain);
}

// Populate Renewals Table
function populateRenewalsTable() {
    const tbody = document.getElementById('renewals-tbody');
    tbody.innerHTML = '';
    
    let totals = {
        units: 0,
        expired: 0,
        renewed: 0,
        inService: 0,
        delinquent: 0
    };
    
    mmrData.forEach(function(row) {
        const totalUnits = toNumber(row.TotalUnits || 0);
        const inServiceUnits = toNumber(row.InServiceUnits || 0);
        const delinquent = toNumber(row.Delinquent || 0);
        const delinquentPercent = inServiceUnits > 0 ? (delinquent / inServiceUnits) : 0; // decimal
        
        const expired = toNumber(row.T12LeasesExpired || 0);
        const renewed = toNumber(row.T12LeasesRenewed || 0);
        const renewalRate = expired > 0 ? (renewed / expired) : 0; // decimal
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.Property || ''}</td>
            <td>${row.City || ''}, ${row.State || ''}</td>
            <td class="total-units">${totalUnits}</td>
            <td>${expired}</td>
            <td>${renewed}</td>
            <td>${formatPercent(renewalRate)}</td>
            <td>${inServiceUnits}</td>
            <td>${delinquent}</td>
            <td>${formatPercent(delinquentPercent)}</td>
        `;
        tbody.appendChild(tr);
        
        totals.units += totalUnits;
        totals.expired += expired;
        totals.renewed += renewed;
        totals.inService += inServiceUnits;
        totals.delinquent += delinquent;
    });
    
    // Update totals
    const avgRenewalRate = totals.expired > 0 ? (totals.renewed / totals.expired) : 0;
    const avgDelinquent = totals.inService > 0 ? (totals.delinquent / totals.inService) : 0;
    
    document.getElementById('renewals-total-units').textContent = totals.units;
    document.getElementById('renewals-total-expired').textContent = totals.expired;
    document.getElementById('renewals-total-renewed').textContent = totals.renewed;
    document.getElementById('renewals-avg-rate').textContent = formatPercent(avgRenewalRate);
    document.getElementById('renewals-total-service').textContent = totals.inService;
    document.getElementById('renewals-total-delinquent').textContent = totals.delinquent;
    document.getElementById('renewals-avg-delinquent').textContent = formatPercent(avgDelinquent);
}

// Populate Rents Table
function populateRentsTable() {
    const tbody = document.getElementById('rents-tbody');
    tbody.innerHTML = '';
    
    let totals = {
        units: 0,
        sumOccupiedWeighted: 0, // sum(occupiedRent * units)
        sumBudgetedWeighted: 0, // sum(budgetedRent * units)
        sumMoveInWeighted: 0,   // sum(moveInRent * units)
        occupiedRent: 0,        // simple sum for potential reference
        budgetedRent: 0,
        moveInRent: 0
    };
    
    mmrData.forEach(function(row) {
        const occupiedRent = toNumber(row.OccupiedRent || 0);
        const budgetedRent = toNumber(row.BudgetedRent || 0);
        const difference = budgetedRent > 0 ? ((occupiedRent - budgetedRent) / budgetedRent) : 0; // decimal
        const moveInRent = toNumber(row.MoveInRent || 0);
        const moveInDiff = budgetedRent > 0 ? ((moveInRent - budgetedRent) / budgetedRent) : 0; // decimal
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.Property || ''}</td>
            <td>${row.City || ''}, ${row.State || ''}</td>
            <td class="total-units">${row.TotalUnits || 0}</td>
            <td class="currency">${formatCurrency(occupiedRent)}</td>
            <td class="currency">${formatCurrency(budgetedRent)}</td>
            <td class="${difference >= 0 ? 'positive' : 'negative'}">${formatPercent(difference)}</td>
            <td class="currency">${formatCurrency(moveInRent)}</td>
            <td class="${moveInDiff >= 0 ? 'positive' : 'negative'}">${formatPercent(moveInDiff)}</td>
        `;
        tbody.appendChild(tr);
        
        const unitsForRow = parseInt(row.TotalUnits || 0);
        totals.units += unitsForRow;
        totals.occupiedRent += occupiedRent;
        totals.budgetedRent += budgetedRent;
        totals.moveInRent += moveInRent;
        totals.sumOccupiedWeighted += occupiedRent * unitsForRow;
        totals.sumBudgetedWeighted += budgetedRent * unitsForRow;
        totals.sumMoveInWeighted += moveInRent * unitsForRow;
    });
    
    // Update totals
    const avgOccupied = totals.units > 0 ? (totals.sumOccupiedWeighted / totals.units) : 0;
    const avgBudgeted = totals.units > 0 ? (totals.sumBudgetedWeighted / totals.units) : 0;
    const avgMoveIn = totals.units > 0 ? (totals.sumMoveInWeighted / totals.units) : 0;
    const avgDiff = avgBudgeted > 0 ? ((avgOccupied - avgBudgeted) / avgBudgeted) : 0; // decimal
    const avgMoveInDiff = avgBudgeted > 0 ? ((avgMoveIn - avgBudgeted) / avgBudgeted) : 0; // decimal
    
    document.getElementById('rents-total-units').textContent = totals.units;
    document.getElementById('rents-avg-occupied').textContent = formatCurrency(avgOccupied);
    document.getElementById('rents-avg-budgeted').textContent = formatCurrency(avgBudgeted);
    document.getElementById('rents-avg-diff').textContent = formatPercent(avgDiff);
    document.getElementById('rents-avg-movein').textContent = formatCurrency(avgMoveIn);
    document.getElementById('rents-avg-movein-diff').textContent = formatPercent(avgMoveInDiff);
}

// Populate Rents/SF Table
function populateRentsFTable() {
    const tbody = document.getElementById('rentsf-tbody');
    tbody.innerHTML = '';
    
    let totals = {
        units: 0,
        occupiedSum: 0,          // sum of occupiedRentSF (simple)
        budgetedSum: 0,
        moveInSum: 0,
        sumOccupiedWeighted: 0,  // sum(occupiedRentSF * units)
        sumBudgetedWeighted: 0,  // sum(budgetedRentSF * units)
        sumMoveInWeighted: 0     // sum(moveInRentSF * units)
    };
    
    mmrData.forEach(function(row) {
        const occupiedRentSF = toNumber(row.OccupiedRentSF || 0);
        const budgetedRentSF = toNumber(row.BudgetedRentSF || 0);
        const difference = budgetedRentSF > 0 ? ((occupiedRentSF - budgetedRentSF) / budgetedRentSF) : 0; // decimal
        const moveInRentSF = toNumber(row.MoveinRentSF || 0);
        const moveInDiff = budgetedRentSF > 0 ? ((moveInRentSF - budgetedRentSF) / budgetedRentSF) : 0; // decimal
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.Property || ''}</td>
            <td>${row.City || ''}, ${row.State || ''}</td>
            <td class="total-units">${row.TotalUnits || 0}</td>
            <td class="currency">${formatCurrency2(occupiedRentSF)}</td>
            <td class="currency">${formatCurrency2(budgetedRentSF)}</td>
            <td class="${difference >= 0 ? 'positive' : 'negative'}">${formatPercent(difference)}</td>
            <td class="currency">${formatCurrency2(moveInRentSF)}</td>
            <td class="${moveInDiff >= 0 ? 'positive' : 'negative'}">${formatPercent(moveInDiff)}</td>
        `;
        tbody.appendChild(tr);
        
        const unitsForRow = parseInt(row.TotalUnits || 0);
        totals.units += unitsForRow;
        totals.occupiedSum += occupiedRentSF;
        totals.budgetedSum += budgetedRentSF;
        totals.moveInSum += moveInRentSF;
        totals.sumOccupiedWeighted += occupiedRentSF * unitsForRow;
        totals.sumBudgetedWeighted += budgetedRentSF * unitsForRow;
        totals.sumMoveInWeighted += moveInRentSF * unitsForRow;
    });
    
    // Update totals
    const avgOccupiedSF = totals.units > 0 ? (totals.sumOccupiedWeighted / totals.units) : 0;
    const avgBudgetedSF = totals.units > 0 ? (totals.sumBudgetedWeighted / totals.units) : 0;
    const avgMoveInSF = totals.units > 0 ? (totals.sumMoveInWeighted / totals.units) : 0;
    const avgDiffSF = avgBudgetedSF > 0 ? ((avgOccupiedSF - avgBudgetedSF) / avgBudgetedSF) : 0; // decimal
    const avgMoveInDiffSF = avgBudgetedSF > 0 ? ((avgMoveInSF - avgBudgetedSF) / avgBudgetedSF) : 0; // decimal
    
    document.getElementById('rentsf-total-units').textContent = totals.units;
    document.getElementById('rentsf-avg-occupied').textContent = formatCurrency2(avgOccupiedSF);
    document.getElementById('rentsf-avg-budgeted').textContent = formatCurrency2(avgBudgetedSF);
    document.getElementById('rentsf-avg-diff').textContent = formatPercent(avgDiffSF);
    document.getElementById('rentsf-avg-movein').textContent = formatCurrency2(avgMoveInSF);
    document.getElementById('rentsf-avg-compared').textContent = formatPercent(avgMoveInDiffSF);
}

// Populate Income Table
function populateIncomeTable() {
    const tbody = document.getElementById('income-tbody');
    tbody.innerHTML = '';
    
    let totals = {
        units: 0,
        currentIncome: 0,
        budgetedIncome: 0
    };
    
    mmrData.forEach(function(row) {
        const currentIncome = toNumber(row.CurrentMonthIncome || 0);
        const budgetedIncome = toNumber(row.BudgetedIncome || 0);
        const diffPercent = budgetedIncome > 0 ? ((currentIncome - budgetedIncome) / budgetedIncome) : 0; // decimal
        
        // Check if this is missing data/timing variance (currentIncome = 0 and budgetedIncome > 50000)
        const isMissingData = currentIncome === 0 && budgetedIncome > 50000;
        const diffCellClass = isMissingData ? 'missing-data' : (diffPercent >= 0 ? 'positive' : 'negative');
        const tooltipText = isMissingData ? 'Missing data/timing variance' : '';
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.Property || ''}</td>
            <td>${row.City || ''}, ${row.State || ''}</td>
            <td class="total-units">${row.TotalUnits || 0}</td>
            <td class="currency">${formatCurrency(currentIncome)}</td>
            <td class="currency">${formatCurrency(budgetedIncome)}</td>
            <td class="${diffCellClass}" title="${tooltipText}">${formatPercent(diffPercent)}</td>
        `;
        tbody.appendChild(tr);
        
        // Only include in totals if not missing data
        if (!isMissingData) {
            totals.units += parseInt(row.TotalUnits || 0);
            totals.currentIncome += currentIncome;
            totals.budgetedIncome += budgetedIncome;
        } else {
            // Still count units for total units display
            totals.units += parseInt(row.TotalUnits || 0);
        }
    });
    
    // Update totals (excluding missing data properties)
    const totalDiff = totals.budgetedIncome > 0 ? ((totals.currentIncome - totals.budgetedIncome) / totals.budgetedIncome) : 0; // decimal
    
    document.getElementById('income-total-units').textContent = totals.units;
    document.getElementById('income-current-total').textContent = formatCurrency(totals.currentIncome);
    document.getElementById('income-budgeted-total').textContent = formatCurrency(totals.budgetedIncome);
    document.getElementById('income-total-diff').textContent = formatPercent(totalDiff);
}

// Populate Reviews Table
function populateReviewsTable() {
    const tbody = document.getElementById('reviews-tbody');
    tbody.innerHTML = '';
    
    let totalRating = 0;
    let totalUnits = 0;
    
    // Group by property and calculate average rating
    const propertyRatings = {};
    if (googleReviewsData && googleReviewsData.length > 0) {
        googleReviewsData.forEach(function(row) {
            const property = row.Property || '';
            if (!propertyRatings[property]) {
                propertyRatings[property] = { total: 0, count: 0 };
            }
            propertyRatings[property].total += parseFloat(row.rating || 0);
            propertyRatings[property].count += 1;
        });
    }
    
    const sortedProperties = mmrData.map(row => ({
        property: row.Property,
        city: row.City,
        state: row.State,
        totalUnits: parseInt(row.TotalUnits || 0),
        rating: propertyRatings[row.Property] ? 
                (propertyRatings[row.Property].total / propertyRatings[row.Property].count) : 
                (4.5 + Math.random() * 0.5) // Placeholder rating
    })).sort((a, b) => b.rating - a.rating);
    
    sortedProperties.forEach(function(item) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.property || ''}</td>
            <td>${item.city || ''}, ${item.state || ''}</td>
            <td class="total-units">${item.totalUnits}</td>
            <td>${item.rating.toFixed(2)}</td>
        `;
        tbody.appendChild(tr);
        totalUnits += item.totalUnits;
        totalRating += item.rating;
    });
    
    const avgRating = sortedProperties.length > 0 ? totalRating / sortedProperties.length : 0;
    document.getElementById('reviews-total-units').textContent = totalUnits;
    document.getElementById('reviews-avg-rating').textContent = avgRating.toFixed(2);
}

// Format helpers
// asDecimalRatio: when true, value is gross/visits (or similar ratio) - always multiply by 100.
//   Use for Closing Ratio, which can exceed 100% (e.g. 11 leases / 5 visits = 220%).
function formatPercent(value, asDecimalRatio) {
    let percentValue = parseFloat(value) || 0;
    
    if (asDecimalRatio) {
        // Always treat as decimal ratio (e.g. closing ratio = gross/visits)
        percentValue = percentValue * 100;
    } else {
        // Heuristic: values with abs <= 2.0 are decimal ratios; > 2.0 are already percentages
        if (Math.abs(percentValue) <= 2.0 && percentValue !== 0) {
            percentValue = percentValue * 100;
        }
    }
    
    return percentValue.toFixed(1) + '%';
}

function formatCurrency(value) {
    return '$' + (value || 0).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatCurrency2(value) {
    // Format currency with 2 decimal places (for per square foot values)
    return '$' + (value || 0).toFixed(2);
}

function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error';
    errorDiv.textContent = message;
    document.body.insertBefore(errorDiv, document.body.firstChild);
}