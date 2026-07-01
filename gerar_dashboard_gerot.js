const fs = require('fs');
const ExcelJS = require('exceljs');

async function main() {
    const csvData = fs.readFileSync('Relatorio_BSC_15062026.csv', 'utf-8');
    const lines = csvData.trim().split('\n');

    // Parse CSV
    const headers = lines[0].split(',').map(h => h.trim());
    const data = [];
    for (let i = 1; i < lines.length; i++) {
        if (!lines[i].trim()) continue;
        const row = lines[i].split(',');
        const obj = {};
        headers.forEach((h, idx) => {
            obj[h] = row[idx] ? row[idx].trim() : '';
        });
        data.push(obj);
    }

    // Filter to W24
    const w24Data = data.filter(d => d['Período'] === 'W24');
    const findKpi = (kpiName) => w24Data.find(d => d['KPI'] === kpiName);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Dashboard', { views: [{ showGridLines: false }] });

    // Grid config
    sheet.columns = [
        { key: 'A', width: 2 },
        { key: 'B', width: 28 },
        { key: 'C', width: 12 },
        { key: 'D', width: 12 },
        { key: 'E', width: 2 },
        { key: 'F', width: 28 },
        { key: 'G', width: 12 },
        { key: 'H', width: 12 },
        { key: 'I', width: 2 },
        { key: 'J', width: 20 },
        { key: 'K', width: 12 },
        { key: 'L', width: 12 },
        { key: 'M', width: 2 },
        { key: 'N', width: 20 },
        { key: 'O', width: 12 },
        { key: 'P', width: 12 },
    ];

    // Background white for the whole visible area
    for (let r = 1; r <= 30; r++) {
        for (let c = 1; c <= 16; c++) {
            sheet.getCell(r, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        }
    }

    // Title
    sheet.mergeCells('B2:P3');
    const title = sheet.getCell('B2');
    title.value = 'GEROT PRÉVIA W24';
    title.font = { name: 'Arial', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
    title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };
    title.alignment = { vertical: 'middle', horizontal: 'center' };

    // Function to draw a block header
    const drawBlockHeader = (range, blockTitle, colorHex, fontColorHex) => {
        sheet.mergeCells(range);
        const cell = sheet.getCell(range.split(':')[0]);
        cell.value = blockTitle;
        cell.font = { bold: true, color: { argb: fontColorHex } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        
        // Add top border
        const cols = range.split(':');
        const startCol = sheet.getCell(cols[0]).col;
        const endCol = sheet.getCell(cols[1]).col;
        const row = sheet.getCell(cols[0]).row;
        for (let c = startCol; c <= endCol; c++) {
            sheet.getCell(row, c).border = { top: { style: 'medium', color: { argb: colorHex } } };
        }
    };

    // Table Headers
    const drawTableHeaders = (row, startColLetter, colorHex, fontColorHex) => {
        const c1 = `${startColLetter}${row}`;
        const c2 = `${String.fromCharCode(startColLetter.charCodeAt(0)+1)}${row}`;
        const c3 = `${String.fromCharCode(startColLetter.charCodeAt(0)+2)}${row}`;
        
        sheet.getCell(c1).value = 'INDICADOR';
        sheet.getCell(c2).value = 'OBJETIVO';
        sheet.getCell(c3).value = 'W24';
        const c4 = `${String.fromCharCode(startColLetter.charCodeAt(0)+3)}${row}`;
        sheet.getCell(c4).value = 'STATUS';

        [c1, c2, c3, c4].forEach(cellRef => {
            const cell = sheet.getCell(cellRef);
            cell.font = { bold: true, color: { argb: fontColorHex }, size: 9 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorHex } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: { style: 'thin', color: { argb: 'FFCCCCCC' } }, bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } }, left: { style: 'thin', color: { argb: 'FFCCCCCC' } }, right: { style: 'thin', color: { argb: 'FFCCCCCC' } } };
        });
    };

    drawBlockHeader('B5:E6', 'FIRST MILE', 'FF4F81BD', 'FF4F81BD');
    drawTableHeaders(7, 'B', 'FFE6F0FA', 'FF4F81BD');

    drawBlockHeader('G5:J6', 'LAST MILE', 'FF2CA02C', 'FF2CA02C');
    drawTableHeaders(7, 'G', 'FFE2EFDA', 'FF2CA02C');

    drawBlockHeader('L5:O6', 'LINE HAUL', 'FFED7D31', 'FFED7D31');
    drawTableHeaders(7, 'L', 'FFFCE4D6', 'FFED7D31');

    drawBlockHeader('Q5:T6', 'SAFETY', 'FF7030A0', 'FF7030A0');
    drawTableHeaders(7, 'Q', 'FFE6E0EC', 'FF7030A0');

    // Mocks / Structures
    const structure = {
        firstMile: [
            { kpi: 'Pickup Sucess (SI) FULL', label: 'Pickup Sucess (SI) Full' },
            { kpi: 'Pickup Sucess (Ontime) FULL', label: 'Pickup Sucess (Ontime) Full' },
            { kpi: '% Utilização Frota Fixa FULL', label: '% Utilização Frota Fixa Full' },
            { kpi: 'Aceite de Scheduling', label: 'Aceite de Scheduling XD' },
            { kpi: 'Ad. Config XD', label: 'Ad. Config XD' }
        ],
        lastMile: [
            { kpi: 'Delivery Success', label: 'Delivery Success' },
            { kpi: '% Real x D7 FDS (SVC) SPOT', label: '% Real x D7 FDS (SVC) SPOT' },
            { kpi: 'Aceite Scheduling Pré Routing', label: 'Aceite Scheduling Pré Routing' },
            { kpi: 'SDD - % ER', label: 'SDD - % ER' },
            { kpi: '% Utilização Frota Fixa (LM)', label: '% Utilização Frota Fixa (LM)' },
            { kpi: 'Delivery Success XPT', label: 'Delivery Success XPT' },
            { kpi: '% Real x D7 FDS (XPT) SPOT', label: '% Real x D7 FDS (XPT) SPOT' },
            { kpi: 'Telemetria MM XPT', label: 'Telemetria MM XPT' },
            { kpi: 'Pet love fluxo de tratativa', label: 'Pet love fluxo de tratativa' }
        ],
        lineHaul: [
            { label: 'Treinamentos', mockMeta: '97%', mockPrevia: '87.50%', success: false }
        ],
        safety: [
            { kpi: 'Ad. Vec Fleet', label: 'Ad. Vec Fleet' },
            { kpi: 'Aderência Treinamentos Safety Driver', label: 'Treinamentos' }
        ]
    };

    const drawDataRows = (startRow, startColLetter, items) => {
        let currRow = startRow;
        items.forEach(item => {
            let metaText = '-';
            let w24Text = '-';
            let isSuccess = false;
            let statusText = '-';
            let badgeColor = 'FF666666';
            let w24Color = 'FF000000';

            if (item.mockMeta) {
                metaText = item.mockMeta;
                w24Text = item.mockW21 || '-';
                statusText = '🔴 Piorou';
                badgeColor = 'FFFF0000';
                w24Color = 'FFFF0000';
            } else {
                const kpiData = findKpi(item.kpi);
                if (kpiData) {
                    metaText = kpiData['Meta 4 Pontos'] ? kpiData['Meta 4 Pontos'].replace('> ', '').replace('< ', '') : '-';
                    w24Text = kpiData['Resultado Final'] || '-';
                    
                    const w24Data = data.find(d => d['KPI'] === item.kpi && d['Período'] === 'W24');
                    const w23Data = data.find(d => d['KPI'] === item.kpi && d['Período'] === 'W23');
                    
                    let metaVal = parseFloat(metaText.replace('%', '').trim());
                    let w24Val = parseFloat(w24Text.replace('%', '').trim());
                    let w23Val = w23Data ? parseFloat((w23Data['Resultado Final'] || '0').replace('%', '').trim()) : NaN;
                    
                    let isLowerBetter = false;
                    if (kpiData['Meta 4 Pontos'] && kpiData['Meta 4 Pontos'].includes('<')) {
                        isLowerBetter = true;
                    }

                    if (!isNaN(metaVal) && !isNaN(w24Val)) {
                        let fail = isLowerBetter ? (w24Val > metaVal) : (w24Val < metaVal);
                        if (fail) w24Color = 'FFFF0000';
                    }
                    
                    if (!isNaN(w24Val) && !isNaN(w23Val)) {
                        let melhorou = isLowerBetter ? (w24Val < w23Val) : (w24Val > w23Val);
                        let manteve = w24Val === w23Val;

                        if (melhorou) {
                            statusText = '🟢 Progrediu';
                            badgeColor = 'FF00B050'; 
                        } else if (manteve) {
                            statusText = '🟡 Manteve';
                            badgeColor = 'FFE67E22';
                        } else {
                            statusText = '🔴 Piorou';
                            badgeColor = 'FFFF0000';
                        }
                    } else if (!isNaN(w24Val)) {
                        statusText = '🟡 S/ Ref';
                        badgeColor = 'FFE67E22';
                    }
                }
            }

            const c1 = `${startColLetter}${currRow}`;
            const c2 = `${String.fromCharCode(startColLetter.charCodeAt(0)+1)}${currRow}`;
            const c3 = `${String.fromCharCode(startColLetter.charCodeAt(0)+2)}${currRow}`;
            const c4 = `${String.fromCharCode(startColLetter.charCodeAt(0)+3)}${currRow}`;

            sheet.getCell(c1).value = item.label;
            sheet.getCell(c1).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
            sheet.getCell(c1).font = { size: 9, bold: true };
            
            sheet.getCell(c2).value = metaText;
            sheet.getCell(c2).alignment = { vertical: 'middle', horizontal: 'center' };
            sheet.getCell(c2).font = { size: 9, bold: true };
            
            sheet.getCell(c3).value = w24Text;
            sheet.getCell(c3).alignment = { vertical: 'middle', horizontal: 'center' };
            sheet.getCell(c3).font = { size: 9, color: { argb: w24Color }, bold: true };

            sheet.getCell(c4).value = statusText;
            sheet.getCell(c4).alignment = { vertical: 'middle', horizontal: 'center' };
            sheet.getCell(c4).font = { size: 8, bold: true, color: { argb: badgeColor } };

            [c1, c2, c3, c4].forEach(cellRef => {
                sheet.getCell(cellRef).border = { top: { style: 'thin', color: { argb: 'FFCCCCCC' } }, bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } }, left: { style: 'thin', color: { argb: 'FFCCCCCC' } }, right: { style: 'thin', color: { argb: 'FFCCCCCC' } } };
            });

            currRow++;
        });
        return currRow;
    };

    drawDataRows(8, 'B', structure.firstMile);
    drawDataRows(8, 'G', structure.lastMile);
    drawDataRows(8, 'L', structure.lineHaul);
    drawDataRows(8, 'Q', structure.safety);

    sheet.mergeCells('B14:P15');
    const dateBox = sheet.getCell('B14');
    const stats = fs.statSync('Relatorio_BSC_15062026.csv');
    const now = stats.mtime;
    const dateStr = now.toLocaleDateString('pt-BR', { day: 'numeric', month: 'short', year: 'numeric' }).replace(' de ', ' de ').replace(' de ', ' de ');
    const timeStr = now.toLocaleTimeString('pt-BR');
    dateBox.value = `Data e Hora da última atualização da base\n${dateStr}, ${timeStr}`;
    dateBox.font = { size: 11, bold: true };
    dateBox.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    dateBox.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    dateBox.border = { top: { style: 'thin', color: { argb: 'FFCCCCCC' } }, bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } }, left: { style: 'thin', color: { argb: 'FFCCCCCC' } }, right: { style: 'thin', color: { argb: 'FFCCCCCC' } } };

    // Task Table Header
    const taskStartRow = 18;
    sheet.mergeCells(`B${taskStartRow}:F${taskStartRow}`);
    sheet.mergeCells(`G${taskStartRow}:O${taskStartRow}`);
    
    sheet.getCell(`B${taskStartRow}`).value = 'OPERAÇÃO';
    sheet.getCell(`G${taskStartRow}`).value = 'TAREFA';
    sheet.getCell(`P${taskStartRow}`).value = 'STATUS';

    [`B${taskStartRow}`, `G${taskStartRow}`, `P${taskStartRow}`].forEach(ref => {
        sheet.getCell(ref).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getCell(ref).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };
        sheet.getCell(ref).alignment = { vertical: 'middle', horizontal: 'center' };
        sheet.getCell(ref).border = { top: { style: 'thin', color: { argb: 'FFCCCCCC' } }, bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } }, left: { style: 'thin', color: { argb: 'FFCCCCCC' } }, right: { style: 'thin', color: { argb: 'FFCCCCCC' } } };
    });

    const tasksMock = [
        { op: 'First Mile', task: 'Realizar follow-up do 1º BIP junto ao cliente.', status: 'ATRASO' },
        { op: 'First Mile', task: 'Acompanhar evolução operacional de PU Maringá.', status: 'EM ANDAMENTO' },
        { op: 'Last Mile', task: 'Treinamentos.', status: 'ATRASO' },
        { op: 'Last Mile', task: 'Viabilidade financeiria (ajuste de tabela e incentivo).', status: 'EM ANDAMENTO' },
        { op: 'Last Mile', task: 'Alinhamento operacional para PDP (escalas).', status: 'EM ANDAMENTO' },
        { op: 'Line Haul', task: 'Treinamentos.', status: 'EM ANDAMENTO' },
        { op: 'Qualidade', task: 'Apresentar Cockpit Qualidade nas próximas agendas.', status: 'EM ANDAMENTO' }
    ];

    let tRow = taskStartRow + 1;
    tasksMock.forEach(t => {
        sheet.mergeCells(`B${tRow}:F${tRow}`);
        sheet.mergeCells(`G${tRow}:O${tRow}`);
        
        sheet.getCell(`B${tRow}`).value = t.op;
        sheet.getCell(`G${tRow}`).value = t.task;
        sheet.getCell(`P${tRow}`).value = t.status;

        sheet.getCell(`B${tRow}`).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        sheet.getCell(`G${tRow}`).alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
        sheet.getCell(`P${tRow}`).alignment = { vertical: 'middle', horizontal: 'center' };

        sheet.getCell(`B${tRow}`).font = { size: 9, bold: true };
        sheet.getCell(`G${tRow}`).font = { size: 9 };
        sheet.getCell(`P${tRow}`).font = { size: 9, bold: true, color: { argb: t.status === 'ATRASO' ? 'FFFF0000' : 'FF00B050' } };
        
        if (t.status === 'ATRASO') {
            sheet.getCell(`P${tRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDECEC' } };
        } else {
            sheet.getCell(`P${tRow}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF1DE' } };
        }

        [`B${tRow}`, `G${tRow}`, `P${tRow}`].forEach(ref => {
            sheet.getCell(ref).border = { top: { style: 'thin', color: { argb: 'FFCCCCCC' } }, bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } }, left: { style: 'thin', color: { argb: 'FFCCCCCC' } }, right: { style: 'thin', color: { argb: 'FFCCCCCC' } } };
        });

        tRow++;
    });

    // Footer
    sheet.mergeCells(`B${tRow}:O${tRow}`);
    const footerLeft = sheet.getCell(`B${tRow}`);
    footerLeft.value = 'RELATÓRIO DE STATUS OPERACIONAL';
    footerLeft.font = { bold: true, size: 10 };
    footerLeft.alignment = { vertical: 'middle', horizontal: 'center' };
    footerLeft.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

    const footerRight = sheet.getCell(`P${tRow}`);
    footerRight.value = `Total de Tarefas: 7\nAtraso: 2`;
    footerRight.font = { bold: true, size: 9, color: { argb: 'FFFF0000' } };
    footerRight.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    footerRight.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

    await workbook.xlsx.writeFile('Dashboard_Gerot_Previa.xlsx');
    console.log('Dashboard_Gerot_Previa.xlsx gerado com sucesso!');
}

main().catch(console.error);
