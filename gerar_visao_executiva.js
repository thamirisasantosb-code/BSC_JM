const fs = require('fs');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

async function generateExcel() {
    const csvPath = 'Base_Indicadores_BSC.csv';
    const manualPath = 'Indicadores_Manuais.xlsx';
    
    const csvData = fs.readFileSync(csvPath, 'utf8');
    const lines = csvData.trim().split('\n');
    
    const pivot = new Map();

    const parseValue = (val) => {
        if (!val || val === '-') return null;
        let isPercentage = val.includes('%');
        let clean = val.replace(/[><%]/g, '').trim().replace(',', '.');
        let num = parseFloat(clean);
        if (isNaN(num)) return null;
        if (isPercentage) num = num / 100.0;
        return num;
    };

    const periodosSet = new Set();

    // 1. Load CSV Data
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const row = lines[i].split(',').map(c => c.trim());
        
        const periodo = row[0];
        if (periodo && periodo !== 'W21' && (periodo.startsWith('W') || periodo === 'Mai')) periodosSet.add(periodo);

        const milha = row[3] || 'Geral';
        const kpi = row[4];
        
        const targetKpis = [
            'Pickup Sucess (SI) FULL',
            'Pickup Sucess (Ontime) FULL',
            'Aderência ao Perfil FULL',
            '% Utilização Frota Fixa FULL',
            'Aceite de Scheduling',
            'Ad. Config XD',
            'Delivery Success',
            '% Real x D7 FDS (SVC) SPOT',
            'Aceite Scheduling Pré Routing',
            'SDD - % ER',
            '% Utilização Frota Fixa (LM)',
            'Delivery Success XPT',
            '% Real x D7 FDS (XPT) SPOT',
            'Telemetria MM XPT',
            'Ad. Vec Fleet',
            'Telemetria Scorecard',
            'Aderência Treinamentos Safety Driver'
        ];
        if (!targetKpis.includes(kpi)) continue;
        
        const metaRaw = row[5] || '';
        const resultRaw = row[7]; 
        
        const key = kpi;
        if (!pivot.has(key)) {
            pivot.set(key, {
                milha: milha,
                kpi: kpi,
                meta: parseValue(metaRaw),
                isLowerBetter: metaRaw.includes('<'),
                isPercentage: metaRaw.includes('%') || (resultRaw && resultRaw.includes('%')),
                resultados: {}
            });
        }
        
        const dataItem = pivot.get(key);
        dataItem.resultados[periodo] = parseValue(resultRaw);
    }

    // 2. Load Manual Data
    if (fs.existsSync(manualPath)) {
        const workbookManual = XLSX.readFile(manualPath);
        const sheet = workbookManual.Sheets[workbookManual.SheetNames[0]];
        const manualJson = XLSX.utils.sheet_to_json(sheet);

        manualJson.forEach(row => {
            const milha = row['Milha'] || 'Geral';
            const kpi = row['KPI'];
            const targetKpis = [
                'Pickup Sucess (SI) FULL',
                'Pickup Sucess (Ontime) FULL',
                'Aderência ao Perfil FULL',
                '% Utilização Frota Fixa FULL',
                'Aceite de Scheduling',
                'Ad. Config XD',
                'Delivery Success',
                '% Real x D7 FDS (SVC) SPOT',
                'Aceite Scheduling Pré Routing',
                'SDD - % ER',
                '% Utilização Frota Fixa (LM)',
                'Delivery Success XPT',
                '% Real x D7 FDS (XPT) SPOT',
                'Telemetria MM XPT',
                'Ad. Vec Fleet',
                'Telemetria Scorecard',
                'Aderência Treinamentos Safety Driver'
            ];
            if (!kpi || !targetKpis.includes(kpi)) return;
            const metaRaw = row['Meta 4 Pontos'] ? row['Meta 4 Pontos'].toString() : '';
            
            const key = `${milha}|${kpi}`;
            if (!pivot.has(key)) {
                pivot.set(key, {
                    milha: milha,
                    kpi: kpi,
                    meta: parseValue(metaRaw),
                    isLowerBetter: metaRaw.includes('<'),
                    isPercentage: metaRaw.includes('%'),
                    resultados: {}
                });
            }

            const dataItem = pivot.get(key);
            Object.keys(row).forEach(col => {
                if ((col.startsWith('W') || col === 'Mai') && col !== 'W21' && row[col] !== '') {
                    periodosSet.add(col);
                    const resStr = row[col].toString();
                    if (resStr.includes('%')) dataItem.isPercentage = true;
                    dataItem.resultados[col] = parseValue(resStr);
                }
            });
        });
    }

    // 2.5 Load JSON Manual Data (Web UI)
    const manualJsonPath = 'manual_data.json';
    if (fs.existsSync(manualJsonPath)) {
        try {
            const apiDataObj = JSON.parse(fs.readFileSync(manualJsonPath, 'utf8'));
            for (const [milha, indicators] of Object.entries(apiDataObj)) {
                for (const [kpi, weeks] of Object.entries(indicators)) {
                    const key = kpi;
                    if (!pivot.has(key)) {
                        pivot.set(key, {
                            milha: milha,
                            kpi: kpi,
                            meta: null, // Fallback if missing
                            isLowerBetter: false,
                            isPercentage: true,
                            resultados: {}
                        });
                    }
                    const dataItem = pivot.get(key);
                    for (const [week, value] of Object.entries(weeks)) {
                        if (week === 'W21') continue;
                        periodosSet.add(week);
                        const resStr = value.toString();
                        if (resStr.includes('%')) dataItem.isPercentage = true;
                        dataItem.resultados[week] = parseValue(resStr);
                    }
                }
            }
        } catch (e) {
            console.error("Erro ao carregar manual_data.json", e);
        }
    }

    // 3. Determine latest 5 periods
    const sortedPeriodos = Array.from(periodosSet).sort((a, b) => {
        if (a === 'Mai') return -1;
        if (b === 'Mai') return 1;
        let numA = parseInt(a.replace('W', '')) || 0;
        let numB = parseInt(b.replace('W', '')) || 0;
        return numA - numB;
    });

    const last5 = sortedPeriodos.slice(-5);
    const weeksToInclude = last5.filter(w => w !== '');
    
    // Identificar a semana fechada (se a última for W26, a fechada é W25)
    const latestWeek = sortedPeriodos[sortedPeriodos.length - 1];
    let closedWeek = latestWeek;
    if (latestWeek === 'W26') {
        closedWeek = 'W25';
    }

    let statusCurrentWeek = latestWeek;
    let statusPrevWeek = sortedPeriodos[sortedPeriodos.length - 2];
    if (latestWeek === 'W26') {
        statusCurrentWeek = 'W25';
        statusPrevWeek = 'W24';
    }

    const isOutOfTarget = (item, closedWk) => {
        const d = pivot.get(item.kpi);
        if (!d) return true; // se não achar, falhou
        
        const metaVal = d.meta;
        if (metaVal === null || isNaN(metaVal)) return false;
        
        const resVal = d.resultados[closedWk];
        if (resVal === null || resVal === undefined) {
            return true; // sem dados na semana fechada, falhou
        }
        
        const atingiu = d.isLowerBetter ? (resVal <= metaVal) : (resVal >= metaVal);
        return !atingiu;
    };

    // 4. Generate Excel
    const workbook = new ExcelJS.Workbook();
    const sheet1 = workbook.addWorksheet('Atualizacao BSC');

    // 1. Top Header
    sheet1.mergeCells('A1:Q2');
    const titleCell = sheet1.getCell('A1');
    titleCell.value = 'BCS - JM';
    titleCell.font = { name: 'Arial', size: 20, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } }; 
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Common column widths
    sheet1.getColumn(1).width = 30; // FM Indicador
    sheet1.getColumn(2).width = 12; // Objetivo
    for(let i=3; i<=8; i++) sheet1.getColumn(i).width = 10; // Ws + Status
    sheet1.getColumn(9).width = 4;  // Spacer (I)
    sheet1.getColumn(10).width = 30; // LM Indicador
    sheet1.getColumn(11).width = 12; // Objetivo
    for(let i=12; i<=17; i++) sheet1.getColumn(i).width = 10; // Ws + Status

    const drawTable = (startCol, startRow, title, headers, items, milha) => {
        // Title
        const numTableCols = 2 + weeksToInclude.length + 1;
        sheet1.mergeCells(startRow, startCol, startRow, startCol + numTableCols - 1);
        const secTitle = sheet1.getCell(startRow, startCol);
        secTitle.value = title;
        secTitle.font = { bold: true, size: 14 };
        secTitle.alignment = { horizontal: 'center' };
        if(milha === 'First Mile') secTitle.font.color = { argb: 'FF4F81BD' };
        if(milha === 'Last Mile') secTitle.font.color = { argb: 'FF2CA02C' };
        if(milha === 'Safety') secTitle.font.color = { argb: 'FF7030A0' };

        // Headers
        const headerRow = sheet1.getRow(startRow + 1);
        headers.forEach((h, idx) => {
            const cell = sheet1.getCell(startRow + 1, startCol + idx);
            cell.value = h;
            cell.font = { bold: true };
            if(milha === 'First Mile') { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } }; cell.font.color = { argb: 'FF2C5282' }; }
            if(milha === 'Last Mile') { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }; cell.font.color = { argb: 'FF276749' }; }
            if(milha === 'Safety') { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E0EC' } }; cell.font.color = { argb: 'FF553C9A' }; }
            cell.border = { top: {style:'thin'}, bottom: {style:'thin'}, left: {style:'thin'}, right: {style:'thin'} };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        // Rows
        let currRow = startRow + 2;
        items.forEach(item => {
            const row = sheet1.getRow(currRow);
            const kpiKey = item.kpi;
            const d = pivot.get(kpiKey) || { meta: null, isPercentage: true, resultados: {} };
            
            // Meta fallback
            let metaVal = d.meta;
            let metaStr = '-';
            if (metaVal !== null) {
                metaStr = d.isPercentage ? `${(metaVal*100).toFixed(1)}%` : metaVal;
            } else if (item.meta) {
                metaVal = parseFloat(item.meta.replace('%','')) / 100;
                metaStr = item.meta;
            }
            d.meta = metaVal;
            
            row.getCell(startCol).value = item.label;
            row.getCell(startCol+1).value = metaStr;

            const formatVal = (val, isPerc) => val !== null && val !== undefined ? (isPerc ? `${(val*100).toFixed(2)}%` : val) : '-';
            
            weeksToInclude.forEach((wk, idx) => {
                const val = d.resultados[wk];
                row.getCell(startCol + 2 + idx).value = formatVal(val, d.isPercentage);
            });

            // Status Logic
            let statusText = '-';
            let statusColor = 'FF000000';
            
            const vCurrent = d.resultados[statusCurrentWeek];
            const vPrev = d.resultados[statusPrevWeek];
            
            if (vCurrent !== null && vCurrent !== undefined && vPrev !== null && vPrev !== undefined) {
                let melhorou = d.isLowerBetter ? (vCurrent < vPrev) : (vCurrent > vPrev);
                let manteve = vCurrent === vPrev;
                if (melhorou) { statusText = '🟢 Progrediu'; statusColor = 'FF00B050'; }
                else if (manteve) { statusText = '🟡 Manteve'; statusColor = 'FFFFC000'; }
                else { statusText = '🔴 Piorou'; statusColor = 'FFFF0000'; }
            }

            const statusColIdx = startCol + 2 + weeksToInclude.length;
            row.getCell(statusColIdx).value = statusText;
            row.getCell(statusColIdx).font = { color: { argb: statusColor }, bold: true };

            // Cell coloring
            weeksToInclude.forEach((wk, idx) => {
                const cell = row.getCell(startCol + 2 + idx);
                const val = d.resultados[wk];
                if (val !== null && val !== undefined && d.meta !== null && d.meta !== undefined) {
                    let atingiu = d.isLowerBetter ? (val <= d.meta) : (val >= d.meta);
                    cell.font = { color: { argb: atingiu ? 'FF00B050' : 'FFFF0000' }, bold: true };
                }
            });

            for(let i=0; i < numTableCols; i++) {
                row.getCell(startCol+i).border = { top: {style:'thin'}, bottom: {style:'thin'}, left: {style:'thin'}, right: {style:'thin'} };
                row.getCell(startCol+i).alignment = { horizontal: 'center', vertical: 'middle' };
            }
            row.getCell(startCol).alignment = { horizontal: 'left', vertical: 'middle' };

            currRow++;
        });
        return currRow;
    };

    const firstMileStruct = [
        { kpi: 'Pickup Sucess (SI) FULL', label: 'Pickup Sucess (SI) FULL', meta: '97.9%' },
        { kpi: 'Pickup Sucess (Ontime) FULL', label: 'Pickup Sucess (Ontime) FULL', meta: '92.6%' },
        { kpi: 'Aderência ao Perfil FULL', label: 'Aderência ao Perfil FULL', meta: '95%' },
        { kpi: '% Utilização Frota Fixa FULL', label: '% Utilização Frota Fixa FULL', meta: '95%' },
        { kpi: 'Aceite de Scheduling', label: 'Aceite de Scheduling XD', meta: '95%' },
        { kpi: 'Ad. Config XD', label: 'Ad. Config XD', meta: '97%' }
    ];
    
    const lastMileStruct = [
        { kpi: 'Delivery Success', label: 'Delivery Success', meta: '98.2%' },
        { kpi: '% Real x D7 FDS (SVC) SPOT', label: '% Real x D7 FDS (SVC) SPOT', meta: '90%' },
        { kpi: 'Aceite Scheduling Pré Routing', label: 'Aceite Scheduling Pré Routing', meta: '5%' },
        { kpi: 'SDD - % ER', label: 'SDD - % ER', meta: '95%' },
        { kpi: '% Utilização Frota Fixa (LM)', label: '% Utilização Frota Fixa (LM)', meta: '95%' },
        { kpi: 'Delivery Success XPT', label: 'Delivery Success XPT', meta: '98.2%' },
        { kpi: '% Real x D7 FDS (XPT) SPOT', label: '% Real x D7 FDS (XPT) SPOT', meta: '90%' },
        { kpi: 'Telemetria MM XPT', label: 'Telemetria MM XPT', meta: '72%' }
    ];

    const safetyStruct = [
        { kpi: 'Ad. Vec Fleet', label: 'Ad. Vec Fleet', meta: '90%' },
        { kpi: 'Telemetria Scorecard', label: 'Telemetria Scorecard', meta: '90%' },
        { kpi: 'Aderência Treinamentos Safety Driver', label: 'Aderência Treinamentos Safety Driver', meta: '97%' }
    ];

    const filteredFirstMile = firstMileStruct.filter(item => isOutOfTarget(item, closedWeek));
    const filteredLastMile = lastMileStruct.filter(item => isOutOfTarget(item, closedWeek));
    const filteredSafety = safetyStruct.filter(item => isOutOfTarget(item, closedWeek));

    const headersList = ['Indicador', 'Objetivo', ...weeksToInclude.map(w => w === 'W26' ? 'W26 (Prévia)' : w), 'Status'];

    // Draw First Mile (Row 4, Col 1)
    const nextRowAfterFM = drawTable(1, 4, 'First Mile', headersList, filteredFirstMile, 'First Mile');
    
    // Draw Last Mile (Row 4, Col 10)
    drawTable(10, 4, 'Last Mile', headersList, filteredLastMile, 'Last Mile');

    // Draw Safety (Below First Mile)
    const nextRowAfterSafety = drawTable(1, nextRowAfterFM + 2, 'Safety', headersList, filteredSafety, 'Safety');

    // 5. Add Tasks / Actions Table at the bottom
    const actionsPath = 'actions.json';
    if (fs.existsSync(actionsPath)) {
        try {
            const actionsData = JSON.parse(fs.readFileSync(actionsPath, 'utf8'));
            let allActions = [];
            Object.values(actionsData).forEach(arr => {
                if (Array.isArray(arr)) allActions = allActions.concat(arr);
            });

            if (allActions.length > 0) {
                const actionsStartRow = nextRowAfterSafety + 3;
                
                // Title
                sheet1.mergeCells(actionsStartRow, 1, actionsStartRow, 10);
                const titleAction = sheet1.getCell(actionsStartRow, 1);
                titleAction.value = 'Acompanhamento de Ações / Status das Tarefas';
                titleAction.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
                titleAction.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };
                titleAction.alignment = { horizontal: 'center', vertical: 'middle' };

                // Headers
                const actionHeaders = ['Operação', 'Indicador/Geral', 'O que será feito? (Tarefa)', 'Quem? (Responsável)', 'Até Quando? (Prazo)', 'Status'];
                const hRow = sheet1.getRow(actionsStartRow + 1);
                
                sheet1.mergeCells(actionsStartRow + 1, 3, actionsStartRow + 1, 5); // Tarefa spans 3 cols
                sheet1.mergeCells(actionsStartRow + 1, 6, actionsStartRow + 1, 7); // Responsável spans 2 cols
                sheet1.mergeCells(actionsStartRow + 1, 8, actionsStartRow + 1, 9); // Prazo spans 2 cols
                
                const colMaps = [
                    { col: 1, text: 'Operação' },
                    { col: 2, text: 'Indicador/Geral' },
                    { col: 3, text: 'O que será feito? (Tarefa)' },
                    { col: 6, text: 'Quem? (Responsável)' },
                    { col: 8, text: 'Até Quando? (Prazo)' },
                    { col: 10, text: 'Status' }
                ];

                colMaps.forEach(m => {
                    const c = sheet1.getCell(actionsStartRow + 1, m.col);
                    c.value = m.text;
                    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };
                    c.border = { top: {style:'thin'}, bottom: {style:'thin'}, left: {style:'thin'}, right: {style:'thin'} };
                    c.alignment = { horizontal: 'center', vertical: 'middle' };
                });

                let rIdx = actionsStartRow + 2;
                allActions.forEach(a => {
                    const r = sheet1.getRow(rIdx);
                    sheet1.mergeCells(rIdx, 3, rIdx, 5);
                    sheet1.mergeCells(rIdx, 6, rIdx, 7);
                    sheet1.mergeCells(rIdx, 8, rIdx, 9);

                    const opCell = sheet1.getCell(rIdx, 1);
                    const indCell = sheet1.getCell(rIdx, 2);
                    const taskCell = sheet1.getCell(rIdx, 3);
                    const respCell = sheet1.getCell(rIdx, 6);
                    const prazoCell = sheet1.getCell(rIdx, 8);
                    const statusCell = sheet1.getCell(rIdx, 10);

                    opCell.value = a.operation || 'Outros';
                    indCell.value = a.indicator || 'Geral';
                    taskCell.value = a.task;
                    respCell.value = a.owner;
                    prazoCell.value = a.deadline;
                    
                    let st = a.status || 'Não Iniciado';
                    statusCell.value = st;
                    if(st === 'Concluído') statusCell.font = { bold: true, color: { argb: 'FF00B050' } };
                    if(st === 'Atrasado') statusCell.font = { bold: true, color: { argb: 'FFFF0000' } };
                    if(st === 'Em Andamento') statusCell.font = { bold: true, color: { argb: 'FFFFC000' } };

                    [1, 2, 3, 6, 8, 10].forEach(c => {
                        const cell = sheet1.getCell(rIdx, c);
                        cell.border = { top: {style:'thin'}, bottom: {style:'thin'}, left: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    });
                    taskCell.alignment = { horizontal: 'left', vertical: 'middle' };
                    
                    rIdx++;
                });
            }
        } catch (e) {
            console.error("Erro ao carregar actions.json no Excel", e);
        }
    }

    const exportPath = 'Visao_Executiva_BSC_Master_v4.xlsx';
    await workbook.xlsx.writeFile(exportPath);
    return exportPath;
}

module.exports = { generateExcel };
