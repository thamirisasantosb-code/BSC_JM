const fs = require('fs');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const https = require('https');

async function generateExcel() {
    const csvPath = 'C:\\Users\\thamires.santos\\OneDrive - JM DISTRIBUIÇÃO\\Projeto\\BSC\\Bannco de Dados\\Base_Indicadores_BSC.csv';
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

    const formatVal = (val, isPerc) => val !== null && val !== undefined ? (isPerc ? `${(val*100).toFixed(2)}%` : val) : '-';

    const downloadChart = (chartConfig) => {
        return new Promise((resolve, reject) => {
            const encodedConfig = encodeURIComponent(JSON.stringify(chartConfig));
            const url = `https://quickchart.io/chart?w=500&h=300&c=${encodedConfig}`;
            
            https.get(url, (res) => {
                if (res.statusCode !== 200) {
                    reject(new Error(`Failed to download chart: ${res.statusCode}`));
                    return;
                }
                const data = [];
                res.on('data', (chunk) => data.push(chunk));
                res.on('end', () => {
                    resolve(Buffer.concat(data));
                });
            }).on('error', reject);
        });
    };

    const periodosSet = new Set();

    // 1. Load CSV Data
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const row = lines[i].split(',').map(c => c.trim());
        
        const periodo = row[0];
        if (periodo && periodo !== 'W21' && (periodo.startsWith('W') || periodo === 'Jun')) periodosSet.add(periodo);

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
                if ((col.startsWith('W') || col === 'Jun') && col !== 'W21' && row[col] !== '') {
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
        if (a === 'Jun') return -1;
        if (b === 'Jun') return 1;
        let numA = parseInt(a.replace('W', '')) || 0;
        let numB = parseInt(b.replace('W', '')) || 0;
        return numA - numB;
    });

    const last5 = sortedPeriodos.slice(-5);
    const weeksToInclude = last5.filter(w => w !== '');
    
    // Identificar a semana fechada (a última cadastrada WXX é a prévia, a anterior é a fechada)
    const latestWeek = sortedPeriodos[sortedPeriodos.length - 1];
    const closedWeek = sortedPeriodos.length > 1 ? sortedPeriodos[sortedPeriodos.length - 2] : latestWeek;

    let statusCurrentWeek = latestWeek;
    let statusPrevWeek = sortedPeriodos.length > 1 ? sortedPeriodos[sortedPeriodos.length - 2] : latestWeek;

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
    sheet1.mergeCells('A1:S2');
    const titleCell = sheet1.getCell('A1');
    titleCell.value = 'BCS - JM';
    titleCell.font = { name: 'Arial', size: 20, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } }; 
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Common column widths
    sheet1.getColumn(1).width = 30; // FM Indicador
    sheet1.getColumn(2).width = 12; // Objetivo
    for(let i=3; i<=9; i++) sheet1.getColumn(i).width = 10; // Ws + Média + Status
    sheet1.getColumn(10).width = 4;  // Spacer (J)
    sheet1.getColumn(11).width = 30; // LM Indicador
    sheet1.getColumn(12).width = 12; // Objetivo
    for(let i=13; i<=19; i++) sheet1.getColumn(i).width = 10; // Ws + Média + Status

    const drawTable = (startCol, startRow, title, headers, items, milha) => {
        // Title
        const numTableCols = 2 + weeksToInclude.length + 2;
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

            // Média de cada semana, sem colocar a semana de prévia (latestWeek)
            let sumAverage = 0;
            let countAverage = 0;
            weeksToInclude.forEach(wk => {
                if (wk !== latestWeek) {
                    const val = d.resultados[wk];
                    if (val !== null && val !== undefined && !isNaN(val)) {
                        sumAverage += val;
                        countAverage++;
                    }
                }
            });
            let avgVal = countAverage > 0 ? (sumAverage / countAverage) : null;
            let avgStr = avgVal !== null ? formatVal(avgVal, d.isPercentage) : '-';

            // Escrever a média na coluna correspondente
            const avgColIdx = startCol + 2 + weeksToInclude.length;
            row.getCell(avgColIdx).value = avgStr;

            // Status Logic (com base na semana fechada 'closedWeek' contra a meta)
            let statusText = '-';
            let statusColor = 'FF000000';
            
            const closedWeekVal = d.resultados[closedWeek];
            if (closedWeekVal !== null && closedWeekVal !== undefined && !isNaN(closedWeekVal)) {
                let atingiuMeta = false;
                if (d.meta !== null && d.meta !== undefined && !isNaN(d.meta)) {
                    atingiuMeta = d.isLowerBetter ? (closedWeekVal <= d.meta) : (closedWeekVal >= d.meta);
                }

                if (atingiuMeta) {
                    statusText = 'Atingido';
                    statusColor = 'FF00B050'; // Verde
                } else {
                    statusText = 'Abaixo';
                    statusColor = 'FFFF0000'; // Vermelho
                }
            }

            const statusColIdx = startCol + 2 + weeksToInclude.length + 1;
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

    const filteredFirstMile = firstMileStruct.filter(item => isOutOfTarget(item, 'Jun'));
    const filteredLastMile = lastMileStruct.filter(item => isOutOfTarget(item, 'Jun'));
    const filteredSafety = safetyStruct.filter(item => isOutOfTarget(item, 'Jun'));

    const headersList = ['Indicador', 'Objetivo', ...weeksToInclude.map(w => w === latestWeek && w.startsWith('W') ? `${w} (Prévia)` : w), 'Média', 'Status'];

    // Draw First Mile (Row 4, Col 1)
    const nextRowAfterFM = drawTable(1, 4, 'First Mile', headersList, filteredFirstMile, 'First Mile');
    
    // Draw Last Mile (Row 4, Col 11)
    drawTable(11, 4, 'Last Mile', headersList, filteredLastMile, 'Last Mile');

    // Draw Safety (Below First Mile)
    const nextRowAfterSafety = drawTable(1, nextRowAfterFM + 2, 'Safety', headersList, filteredSafety, 'Safety');

    // 5. Add Tasks / Actions Table at the bottom
    const actionsPath = 'actions.json';
    if (fs.existsSync(actionsPath)) {
        try {
            const actionsData = JSON.parse(fs.readFileSync(actionsPath, 'utf8'));
            let allActions = [];
            Object.entries(actionsData).forEach(([op, arr]) => {
                if (Array.isArray(arr)) {
                    arr.forEach(action => {
                        action.operation = op;
                    });
                    allActions = allActions.concat(arr);
                }
            });

            if (allActions.length > 0) {
                const actionsStartRow = nextRowAfterSafety + 3;
                
                // Title
                sheet1.mergeCells(actionsStartRow, 1, actionsStartRow, 11);
                const titleAction = sheet1.getCell(actionsStartRow, 1);
                titleAction.value = 'Acompanhamento de Ações / Status das Tarefas';
                titleAction.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
                titleAction.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };
                titleAction.alignment = { horizontal: 'center', vertical: 'middle' };

                // Headers
                const actionHeaders = ['Operação', 'Indicador/Geral', 'O que será feito? (Tarefa)', 'Quem? (Responsável)', 'Até Quando? (Prazo)', 'Status'];
                const hRow = sheet1.getRow(actionsStartRow + 1);
                
                sheet1.mergeCells(actionsStartRow + 1, 3, actionsStartRow + 1, 5); // Tarefa spans 3 cols (3, 4, 5)
                sheet1.mergeCells(actionsStartRow + 1, 6, actionsStartRow + 1, 7); // Responsável spans 2 cols (6, 7)
                sheet1.mergeCells(actionsStartRow + 1, 8, actionsStartRow + 1, 10); // Prazo spans 3 cols (8, 9, 10)
                
                const colMaps = [
                    { col: 1, text: 'Operação' },
                    { col: 2, text: 'Indicador/Geral' },
                    { col: 3, text: 'O que será feito? (Tarefa)' },
                    { col: 6, text: 'Quem? (Responsável)' },
                    { col: 8, text: 'Até Quando? (Prazo)' },
                    { col: 11, text: 'Status' }
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
                    sheet1.mergeCells(rIdx, 8, rIdx, 10); // Prazo spans 3 cols (8, 9, 10)

                    const opCell = sheet1.getCell(rIdx, 1);
                    const indCell = sheet1.getCell(rIdx, 2);
                    const taskCell = sheet1.getCell(rIdx, 3);
                    const respCell = sheet1.getCell(rIdx, 6);
                    const prazoCell = sheet1.getCell(rIdx, 8);
                    const statusCell = sheet1.getCell(rIdx, 11);

                    opCell.value = a.mile || a.operation || 'Outros';
                    indCell.value = a.indicator || 'Geral';
                    taskCell.value = a.task;
                    respCell.value = a.owner;
                    prazoCell.value = a.deadline;
                    
                    let st = a.status || 'Não Iniciado';
                    statusCell.value = st;
                    if(st === 'Concluído') statusCell.font = { bold: true, color: { argb: 'FF00B050' } };
                    if(st === 'Atrasado') statusCell.font = { bold: true, color: { argb: 'FFFF0000' } };
                    if(st === 'Em Andamento') statusCell.font = { bold: true, color: { argb: 'FFFFC000' } };

                    [1, 2, 3, 6, 8, 11].forEach(c => {
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

    // ==========================================
    // SHEET 2: Fechamento Semanal
    // ==========================================
    const sheet2 = workbook.addWorksheet('Fechamento Semanal');

    // 1. Title Banner
    sheet2.mergeCells('A1:G2');
    const titleCell2 = sheet2.getCell('A1');
    titleCell2.value = `Fechamento Semanal - BSC (${closedWeek})`;
    titleCell2.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };
    titleCell2.alignment = { vertical: 'middle', horizontal: 'center' };

    // 2. Headers
    const headersList2 = ['Milha', 'Indicador', 'Objetivo', 'Média', `${closedWeek} (Fechamento)`, 'Status', `${latestWeek} (Prévia)`];
    const headerRow2 = sheet2.getRow(4);
    headersList2.forEach((h, idx) => {
        const cell = sheet2.getCell(4, idx + 1);
        cell.value = h;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F497D' } };
        cell.border = { top: {style:'thin'}, bottom: {style:'thin'}, left: {style:'thin'}, right: {style:'thin'} };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    headerRow2.height = 25;

    // Set Column Widths
    sheet2.getColumn(1).width = 15; // Milha
    sheet2.getColumn(2).width = 35; // Indicador
    sheet2.getColumn(3).width = 12; // Objetivo
    sheet2.getColumn(4).width = 12; // Média
    sheet2.getColumn(5).width = 18; // Fechamento
    sheet2.getColumn(6).width = 12; // Status
    sheet2.getColumn(7).width = 18; // Prévia

    // Gather all KPIs
    const allKpis = [
        ...firstMileStruct.map(k => ({ ...k, milha: 'First Mile' })),
        ...lastMileStruct.map(k => ({ ...k, milha: 'Last Mile' })),
        ...safetyStruct.map(k => ({ ...k, milha: 'Safety' }))
    ];

    let rowIdx2 = 5;
    allKpis.forEach(item => {
        const row = sheet2.getRow(rowIdx2);
        row.height = 20;

        const kpiKey = item.kpi;
        const d = pivot.get(kpiKey) || { meta: null, isPercentage: true, resultados: {} };

        // Milha cell
        const cellMilha = row.getCell(1);
        cellMilha.value = item.milha;
        cellMilha.alignment = { horizontal: 'center', vertical: 'middle' };
        
        // Add color tags for Milha
        if (item.milha === 'First Mile') {
            cellMilha.font = { bold: true, color: { argb: 'FF2C5282' } };
            cellMilha.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };
        } else if (item.milha === 'Last Mile') {
            cellMilha.font = { bold: true, color: { argb: 'FF276749' } };
            cellMilha.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
        } else if (item.milha === 'Safety') {
            cellMilha.font = { bold: true, color: { argb: 'FF553C9A' } };
            cellMilha.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6E0EC' } };
        }

        // Indicador
        const cellInd = row.getCell(2);
        cellInd.value = item.label;
        cellInd.font = { name: 'Arial', size: 10 };
        cellInd.alignment = { horizontal: 'left', vertical: 'middle' };

        // Meta/Objetivo fallback
        let metaVal = d.meta;
        let metaStr = '-';
        if (metaVal !== null) {
            metaStr = d.isPercentage ? `${(metaVal*100).toFixed(1)}%` : metaVal;
        } else if (item.meta) {
            metaVal = parseFloat(item.meta.replace('%','')) / 100;
            metaStr = item.meta;
        }
        d.meta = metaVal;

        const cellMeta = row.getCell(3);
        cellMeta.value = metaStr;
        cellMeta.font = { bold: true };
        cellMeta.alignment = { horizontal: 'center', vertical: 'middle' };

        // Calculate Average
        let sumAverage = 0;
        let countAverage = 0;
        weeksToInclude.forEach(wk => {
            if (wk !== latestWeek) {
                const val = d.resultados[wk];
                if (val !== null && val !== undefined && !isNaN(val)) {
                    sumAverage += val;
                    countAverage++;
                }
            }
        });
        let avgVal = countAverage > 0 ? (sumAverage / countAverage) : null;
        let avgStr = avgVal !== null ? formatVal(avgVal, d.isPercentage) : '-';
        
        const cellAvg = row.getCell(4);
        cellAvg.value = avgStr;
        cellAvg.font = { bold: true };
        cellAvg.alignment = { horizontal: 'center', vertical: 'middle' };

        // closedWeek (Fechamento)
        const closedVal = d.resultados[closedWeek];
        const cellClosed = row.getCell(5);
        cellClosed.value = closedVal !== null && closedVal !== undefined ? formatVal(closedVal, d.isPercentage) : '-';
        cellClosed.font = { bold: true };
        cellClosed.alignment = { horizontal: 'center', vertical: 'middle' };

        // Color coding for closed week cell against target
        let isLowerBetter = false;
        if (item.meta && item.meta.includes('<')) {
            isLowerBetter = true;
        }
        if (closedVal !== null && closedVal !== undefined && !isNaN(closedVal) && d.meta !== null && d.meta !== undefined && !isNaN(d.meta)) {
            let fail = isLowerBetter ? (closedVal > d.meta) : (closedVal < d.meta);
            if (fail) {
                cellClosed.font = { color: { argb: 'FFFF0000' }, bold: true }; // Vermelho
            } else {
                cellClosed.font = { color: { argb: 'FF00B050' }, bold: true }; // Verde
            }
        }

        // Status Logic
        let statusText = '-';
        let statusColor = 'FF000000';
        if (closedVal !== null && closedVal !== undefined && !isNaN(closedVal)) {
            let atingiuMeta = false;
            if (d.meta !== null && d.meta !== undefined && !isNaN(d.meta)) {
                atingiuMeta = isLowerBetter ? (closedVal <= d.meta) : (closedVal >= d.meta);
            }
            if (atingiuMeta) {
                statusText = 'Atingido';
                statusColor = 'FF00B050';
            } else {
                statusText = 'Abaixo';
                statusColor = 'FFFF0000';
            }
        }
        const cellStatus = row.getCell(6);
        cellStatus.value = statusText;
        cellStatus.font = { bold: true, color: { argb: statusColor } };
        cellStatus.alignment = { horizontal: 'center', vertical: 'middle' };

        // latestWeek (Prévia)
        const latestVal = d.resultados[latestWeek];
        const cellLatest = row.getCell(7);
        cellLatest.value = latestVal !== null && latestVal !== undefined ? formatVal(latestVal, d.isPercentage) : '-';
        cellLatest.alignment = { horizontal: 'center', vertical: 'middle' };

        // Border for all cells in row
        for (let c = 1; c <= 7; c++) {
            row.getCell(c).border = {
                top: {style:'thin'},
                bottom: {style:'thin'},
                left: {style:'thin'},
                right: {style:'thin'}
            };
        }

        rowIdx2++;
    });

    // Calculate chart stats based on allKpis
    let atingidoCount = 0;
    let abaixoCount = 0;

    const milhaStats = {
        'First Mile': { atingido: 0, abaixo: 0 },
        'Last Mile': { atingido: 0, abaixo: 0 },
        'Safety': { atingido: 0, abaixo: 0 }
    };

    allKpis.forEach(item => {
        const kpiKey = item.kpi;
        const d = pivot.get(kpiKey) || { meta: null, isPercentage: true, resultados: {} };
        const closedVal = d.resultados[closedWeek];
        
        let isLowerBetter = false;
        if (item.meta && item.meta.includes('<')) {
            isLowerBetter = true;
        }
        let metaVal = d.meta;
        if (metaVal === null && item.meta) {
            metaVal = parseFloat(item.meta.replace('%','')) / 100;
        }
        
        if (closedVal !== null && closedVal !== undefined && !isNaN(closedVal) && metaVal !== null && metaVal !== undefined && !isNaN(metaVal)) {
            const atingiuMeta = isLowerBetter ? (closedVal <= metaVal) : (closedVal >= metaVal);
            if (atingiuMeta) {
                atingidoCount++;
                if (milhaStats[item.milha]) milhaStats[item.milha].atingido++;
            } else {
                abaixoCount++;
                if (milhaStats[item.milha]) milhaStats[item.milha].abaixo++;
            }
        }
    });

    try {
        const pieConfig = {
            type: 'pie',
            data: {
                labels: ['Atingido', 'Abaixo da Meta'],
                datasets: [{
                    data: [atingidoCount, abaixoCount],
                    backgroundColor: ['#00B050', '#FF0000']
                }]
            },
            options: {
                title: {
                    display: true,
                    text: `Resumo Geral de Indicadores - ${closedWeek}`,
                    fontSize: 16
                }
            }
        };

        const barConfig = {
            type: 'bar',
            data: {
                labels: ['First Mile', 'Last Mile', 'Safety'],
                datasets: [
                    {
                        label: 'Atingido',
                        backgroundColor: '#00B050',
                        data: [
                            milhaStats['First Mile'].atingido,
                            milhaStats['Last Mile'].atingido,
                            milhaStats['Safety'].atingido
                        ]
                    },
                    {
                        label: 'Abaixo',
                        backgroundColor: '#FF0000',
                        data: [
                            milhaStats['First Mile'].abaixo,
                            milhaStats['Last Mile'].abaixo,
                            milhaStats['Safety'].abaixo
                        ]
                    }
                ]
            },
            options: {
                title: {
                    display: true,
                    text: `Status por Milha - ${closedWeek}`,
                    fontSize: 16
                },
                scales: {
                    yAxes: [{
                        ticks: {
                            beginAtZero: true,
                            stepSize: 1
                        }
                    }]
                }
            }
        };

        console.log("Baixando gráficos...");
        const pieChartBuffer = await downloadChart(pieConfig);
        const barChartBuffer = await downloadChart(barConfig);

        const pieImageId = workbook.addImage({
            buffer: pieChartBuffer,
            extension: 'png'
        });
        const barImageId = workbook.addImage({
            buffer: barChartBuffer,
            extension: 'png'
        });

        // Insert side-by-side below the table (table ends around row 20)
        sheet2.addImage(pieImageId, 'B22:D36');
        sheet2.addImage(barImageId, 'E22:G36');
        console.log("Gráficos adicionados ao Excel com sucesso!");
    } catch (err) {
        console.error("Erro ao gerar ou adicionar gráficos no Excel:", err);
    }

    const exportPath = 'Visao_Executiva_BSC_Master_v4.xlsx';
    await workbook.xlsx.writeFile(exportPath);
    return exportPath;
}

module.exports = { generateExcel };
