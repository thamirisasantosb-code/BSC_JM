const fs = require('fs');
const ExcelJS = require('exceljs');

async function createMirror() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Indicadores');

    // Define colunas
    sheet.columns = [
        { header: 'Operação', key: 'op', width: 15 },
        { header: 'KPI', key: 'kpi', width: 40 },
        { header: 'Meta 4 Pontos', key: 'meta', width: 15 },
        { header: 'Jun', key: 'Jun', width: 10 },
        { header: 'W24', key: 'W24', width: 10 },
        { header: 'W25', key: 'W25', width: 10 },
        { header: 'W26', key: 'W26', width: 10 },
        { header: 'W27', key: 'W27', width: 10 },
        { header: 'W28', key: 'W28', width: 10 },
        { header: 'W29', key: 'W29', width: 10 },
        { header: 'W30', key: 'W30', width: 10 }
    ];

    // Estilizar o cabeçalho
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0A246A' } };

    const structure = {
        'First Mile': [
            { kpi: 'Pickup Sucess (SI) FULL', meta: '97.9%' },
            { kpi: 'Pickup Sucess (Ontime) FULL', meta: '92.6%' },
            { kpi: 'Aderência ao Perfil FULL', meta: '95%' },
            { kpi: '% Utilização Frota Fixa FULL', meta: '95%' },
            { kpi: 'Aceite de Scheduling', meta: '95%' },
            { kpi: 'Ad. Config XD', meta: '97%' }
        ],
        'Last Mile': [
            { kpi: 'Delivery Success', meta: '98.2%' },
            { kpi: '% Real x D7 FDS (SVC) SPOT', meta: '90%' },
            { kpi: 'Aceite Scheduling Pré Routing', meta: '5%' },
            { kpi: 'SDD - % ER', meta: '95%' },
            { kpi: 'Delivery Success XPT', meta: '98.2%' },
            { kpi: '% Real x D7 FDS (XPT) SPOT', meta: '90%' },
            { kpi: 'Telemetria MM XPT', meta: '72%' }
        ],
        'Safety': [
            { kpi: 'Ad. Vec Fleet', meta: '90%' },
            { kpi: 'Telemetria Scorecard', meta: '90%' },
            { kpi: 'Aderência Treinamentos Safety Driver', meta: '97%' }
        ]
    };

    // Ler CSV para extrair dados existentes
    const csvPath = 'Base_Indicadores_BSC.csv';
    const csvData = fs.readFileSync(csvPath, 'utf-8');
    const lines = csvData.trim().split('\n');
    
    // Processar cabeçalho
    const header = lines[0].split(',').map(c => c.trim());
    const kpiIdx = header.indexOf('KPI');
    const periodIdx = header.indexOf('Período');
    const resIdx = header.indexOf('Resultado Final');

    // Mapear dados existentes { 'KPI': { 'W21': '...', 'W22': '...' } }
    const existingData = {};

    if (kpiIdx !== -1 && periodIdx !== -1 && resIdx !== -1) {
        for (let i = 1; i < lines.length; i++) {
            if (!lines[i].trim()) continue;
            // Trata o split lidando com aspas se houver
            let row = [];
            let inQuotes = false;
            let current = '';
            for (let char of lines[i]) {
                if (char === '"') inQuotes = !inQuotes;
                else if (char === ',' && !inQuotes) {
                    row.push(current.trim());
                    current = '';
                } else {
                    current += char;
                }
            }
            row.push(current.trim());

            const kpiName = row[kpiIdx];
            const period = row[periodIdx];
            let res = row[resIdx];
            
            if (!kpiName || !period) continue;
            
            if (res && res !== '-') {
                // Tenta limpar e garantir que mantém formato percentual amigável se for o caso
                // No CSV os resultados podem vir já com % ou sem
                if (!existingData[kpiName]) existingData[kpiName] = {};
                existingData[kpiName][period] = res;
            }
        }
    }

    // Preencher as linhas com Operação, KPI, Meta e valores existentes
    for (const [op, kpis] of Object.entries(structure)) {
        for (const item of kpis) {
            const dataForKpi = existingData[item.kpi] || {};
            sheet.addRow({
                op: op,
                kpi: item.kpi,
                meta: item.meta,
                Jun: dataForKpi['Jun'] || '',
                W24: dataForKpi['W24'] || '',
                W25: dataForKpi['W25'] || '',
                W26: dataForKpi['W26'] || '',
                W27: dataForKpi['W27'] || '',
                W28: dataForKpi['W28'] || '',
                W29: dataForKpi['W29'] || '',
                W30: dataForKpi['W30'] || ''
            });
        }
    }

    // Salvar o arquivo
    await workbook.xlsx.writeFile('Espelho_Indicadores_BSC.xlsx');
    console.log('Arquivo Espelho_Indicadores_BSC.xlsx recriado com sucesso com dados do CSV!');
}

createMirror();
