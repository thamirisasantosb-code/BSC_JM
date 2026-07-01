window.userRole = localStorage.getItem('bsc_user_role');
window.userEmail = localStorage.getItem('bsc_user_email');

function logoutRole() {
    window.userRole = null;
    window.userEmail = null;
    localStorage.removeItem('bsc_user_role');
    localStorage.removeItem('bsc_user_email');
    window.location.href = 'login.html';
}

function applyRoleRestrictions() {
    if (!window.userRole) return;
    
    // Update profile button text
    const btnProfile = document.getElementById('btn-profile');
    if (btnProfile) {
        const shortEmail = window.userEmail ? window.userEmail.split('@')[0] : window.userRole;
        btnProfile.innerHTML = `👤 ${shortEmail} ▾`;
    }

    const btnUsers = document.getElementById('btn-users-mgmt');
    const isMaster = window.userRole === 'Master';
    
    if (btnUsers) {
        btnUsers.style.display = isMaster ? 'block' : 'none';
    }

    // Esconder/mostrar painéis baseado na role
    const grids = document.querySelectorAll('.excel-grid');
    grids.forEach(grid => {
        const cols = grid.querySelectorAll('.col-left, .col-right');
        let visibleCols = 0;
        cols.forEach(col => {
            const titleEl = col.querySelector('.section-title');
            if (titleEl) {
                const title = titleEl.textContent.trim();
                if (isMaster || title === window.userRole) {
                    col.style.display = '';
                    visibleCols++;
                } else {
                    col.style.display = 'none';
                }
            }
        });
        // Se a grid ficar vazia, esconde
        grid.style.display = visibleCols === 0 ? 'none' : 'flex';
    });
    
    // Filtrar Tabela Global Tasks (já é chamada no populateGlobalTasks, mas caso precise re-renderizar)
    if (typeof allActions !== 'undefined' && Array.isArray(allActions)) {
        populateGlobalTasks();
    }
}

document.addEventListener('DOMContentLoaded', () => {
    if (!window.userRole) {
        window.location.href = 'login.html';
        return;
    } else {
        applyRoleRestrictions();
    }
    
    // Definimos a estrutura dos cards baseada na imagem
    const structure = {
        firstMile: [
            { kpi: 'Pickup Sucess (SI) FULL', label: 'Pickup Sucess (SI) FULL', meta: '97.9%' },
            { kpi: 'Pickup Sucess (Ontime) FULL', label: 'Pickup Sucess (Ontime) FULL', meta: '92.6%' },
            { kpi: 'Aderência ao Perfil FULL', label: 'Aderência ao Perfil FULL', meta: '95%' },
            { kpi: '% Utilização Frota Fixa FULL', label: '% Utilização Frota Fixa FULL', meta: '95%' },
            { kpi: 'Aceite de Scheduling', label: 'Aceite de Scheduling XD', meta: '95%' },
            { kpi: 'Ad. Config XD', label: 'Ad. Config XD', meta: '97%' }
        ],
        lastMile: [
            { kpi: 'Delivery Success', label: 'Delivery Success', meta: '98.2%' },
            { kpi: '% Real x D7 FDS (SVC) SPOT', label: '% Real x D7 FDS (SVC) SPOT', meta: '90%' },
            { kpi: 'Aceite Scheduling Pré Routing', label: 'Aceite Scheduling Pré Routing', meta: '5%' },
            { kpi: 'SDD - % ER', label: 'SDD - % ER', meta: '95%' },
            { kpi: '% Utilização Frota Fixa (LM)', label: '% Utilização Frota Fixa (LM)', meta: '95%' },
            { kpi: 'Delivery Success XPT', label: 'Delivery Success XPT', meta: '98.2%' },
            { kpi: '% Real x D7 FDS (XPT) SPOT', label: '% Real x D7 FDS (XPT) SPOT', meta: '90%' },
            { kpi: 'Telemetria MM XPT', label: 'Telemetria MM XPT', meta: '72%' }
        ],
        lineHaul: [],
        safety: [
            { kpi: 'Ad. Vec Fleet', label: 'Ad. Vec Fleet', meta: '90%' },
            { kpi: 'Telemetria Scorecard', label: 'Telemetria Scorecard', meta: '90%' },
            { kpi: 'Aderência Treinamentos Safety Driver', label: 'Aderência Treinamentos Safety Driver', meta: '97%' }
        ]
    };

    // Carregar CSV e Excel automaticamente
    Promise.all([
        fetch('Base_Indicadores_BSC.csv?t=' + new Date().getTime()),
        fetch('Indicadores_Manuais.xlsx?t=' + new Date().getTime()),
        fetch('/api/manual-data').catch(() => ({ ok: false }))
    ])
    .then(async ([csvResponse, excelResponse, apiManualResponse]) => {
        if (!csvResponse.ok) throw new Error("Erro ao carregar CSV local");
        
        let lastMod = csvResponse.headers.get('Last-Modified');
        window.csvLastModifiedDate = lastMod ? new Date(lastMod) : new Date();
        
        const csvText = await csvResponse.text();
        
        const csvData = await new Promise(resolve => {
            Papa.parse(csvText, {
                delimiter: ",",
                header: true,
                skipEmptyLines: true,
                complete: function(results) {
                    resolve(results.data);
                }
            });
        });
        
        let manualData = [];
        if (excelResponse.ok) {
            try {
                const arrayBuffer = await excelResponse.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const rawExcelJson = XLSX.utils.sheet_to_json(worksheet);
                
                rawExcelJson.forEach(row => {
                    const kpi = row['KPI'];
                    const meta = row['Meta 4 Pontos'];
                    Object.keys(row).forEach(key => {
                        if ((key.startsWith('W') || key === 'Mai') && key !== 'W21' && row[key] !== '') {
                            manualData.push({
                                'KPI': kpi,
                                'Período': key,
                                'Meta 4 Pontos': meta,
                                'Resultado Final': row[key]
                            });
                        }
                    });
                });
            } catch (e) {
                console.error("Erro ao ler Indicadores_Manuais.xlsx", e);
            }
        }
        
        // 3. Carregar dados manuais da API
        if (apiManualResponse && apiManualResponse.ok) {
            try {
                const apiDataObj = await apiManualResponse.json();
                window.apiManualData = apiDataObj;
                
                for (const [operation, indicators] of Object.entries(apiDataObj)) {
                    for (const [kpi, weeks] of Object.entries(indicators)) {
                        for (const [week, value] of Object.entries(weeks)) {
                            const existingIndex = manualData.findIndex(d => d['KPI'] === kpi && d['Período'] === week);
                            if (existingIndex !== -1) {
                                manualData[existingIndex]['Resultado Final'] = value;
                                manualData[existingIndex]['isManualAPI'] = true;
                            } else {
                                manualData.push({
                                    'KPI': kpi,
                                    'Período': week,
                                    'Meta 4 Pontos': '', 
                                    'Resultado Final': value,
                                    'isManualAPI': true
                                });
                            }
                        }
                    }
                }
            } catch (e) {
                console.error("Erro ao ler dados manuais da API", e);
            }
        }
        
        const combinedData = [...csvData, ...manualData];
        processData(combinedData);
    })
    .catch(err => {
        console.error(err);
        alert("Erro Técnico: " + err.message + "\n\nVerifique se o painel está rodando pelo Iniciar_Painel.bat (http://localhost:3000)");
    });

    let globalLast6Weeks = [];
    let statusCurrentWeek = '';
    let statusPrevWeek = '';

    function isOutOfTarget(item, allData, closedWeek) {
        // Encontrar a linha correspondente ao mês fechado ('Mai' ou qualquer período que não comece com 'W' e não seja 'W21')
        const monthData = [...allData].find(d => d['KPI'] === item.kpi && !d['Período'].startsWith('W') && d['Período'] !== 'W21' && d['Período'].trim() !== '');
        
        if (monthData) {
            const gapPont = monthData['GAP Pontuação'] || '';
            const gapPeso = monthData['GAP Peso Ponderado'] || '';
            
            if (gapPont.includes('🔴') || gapPeso.includes('🔴')) {
                return true;
            }
            if (parseFloat(gapPont.replace(',', '.')) < 0 || parseFloat(gapPeso.replace(',', '.')) < 0) {
                return true;
            }
            return false;
        }

        // Fallback original caso não tenha linha de mês
        const closedData = [...allData].reverse().find(d => d['KPI'] === item.kpi && d['Período'] === closedWeek);
        
        let metaText = '-';
        let isLowerBetter = false;
        if (closedData && closedData['Meta 4 Pontos']) {
            metaText = closedData['Meta 4 Pontos'];
        } else if (item.meta) {
            metaText = item.meta;
        }
        
        if (metaText === '-') return false; 
        
        isLowerBetter = metaText.includes('<');
        const cleanMetaText = metaText.replace(/[><%]/g, '').trim().replace(',', '.');
        const metaVal = parseFloat(cleanMetaText);
        if (isNaN(metaVal)) return false;
        
        const closedResultText = closedData ? (closedData['Resultado Final'] || '') : '';
        if (!closedResultText || closedResultText === '-') {
            return true; // Sem resultado = fora da meta
        }
        
        const closedResultVal = parseFloat(closedResultText.replace(/[><%]/g, '').trim().replace(',', '.'));
        if (isNaN(closedResultVal)) return true;
        
        const atingiu = isLowerBetter ? (closedResultVal <= metaVal) : (closedResultVal >= metaVal);
        return !atingiu;
    }

    function processData(data) {
        window.autoTasks = [];
        
        const periodos = new Set();
        data.forEach(row => {
            if (row['Período'] && row['Período'] !== 'W21' && (row['Período'].startsWith('W') || row['Período'] === 'Mai')) {
                periodos.add(row['Período']);
            }
        });
        
        let sortedPeriodos = Array.from(periodos).sort((a, b) => {
            if (a === 'Mai') return -1;
            if (b === 'Mai') return 1;
            let numA = parseInt(a.replace('W', '')) || 0;
            let numB = parseInt(b.replace('W', '')) || 0;
            return numA - numB;
        });
        
        globalLast6Weeks = sortedPeriodos.slice(-6);
        
        const latestWeek = sortedPeriodos[sortedPeriodos.length - 1];
        let closedWeek = latestWeek;
        if (latestWeek === 'W27') {
            closedWeek = 'W26';
        }
        
        statusCurrentWeek = latestWeek;
        statusPrevWeek = sortedPeriodos[sortedPeriodos.length - 2];
        if (latestWeek === 'W27') {
            statusCurrentWeek = 'W26';
            statusPrevWeek = 'W25';
        }
        
        const w6th = globalLast6Weeks[5];
        if (!w6th) {
            document.querySelectorAll('.dyn-w6').forEach(th => th.style.display = 'none');
        } else {
            document.querySelectorAll('.dyn-w6').forEach(th => th.style.display = '');
        }
        
        const formatHeader = (wk) => wk === 'W27' ? 'W27 (Prévia)' : (wk || '--');
        document.querySelectorAll('.dyn-w1').forEach(th => th.textContent = formatHeader(globalLast6Weeks[0]));
        document.querySelectorAll('.dyn-w2').forEach(th => th.textContent = formatHeader(globalLast6Weeks[1]));
        document.querySelectorAll('.dyn-w3').forEach(th => th.textContent = formatHeader(globalLast6Weeks[2]));
        document.querySelectorAll('.dyn-w4').forEach(th => th.textContent = formatHeader(globalLast6Weeks[3]));
        document.querySelectorAll('.dyn-w5').forEach(th => th.textContent = formatHeader(globalLast6Weeks[4]));
        document.querySelectorAll('.dyn-w6').forEach(th => th.textContent = formatHeader(globalLast6Weeks[5]));
        
        const filteredFirstMile = structure.firstMile.filter(item => isOutOfTarget(item, data, closedWeek));
        const filteredLastMile = structure.lastMile.filter(item => isOutOfTarget(item, data, closedWeek));
        const filteredSafety = structure.safety.filter(item => isOutOfTarget(item, data, closedWeek));
        
        populateTable('table-first-mile', filteredFirstMile, data);
        populateTable('table-last-mile', filteredLastMile, data);
        populateTable('table-safety', filteredSafety, data);
        populateGlobalTasks();
        updateCentralPanel();
    }

    function populateTable(tableId, structureList, allData) {
        const tbody = document.querySelector(`#${tableId} tbody`);
        tbody.innerHTML = '';

        const w1 = globalLast6Weeks[0];
        const w2 = globalLast6Weeks[1];
        const w3 = globalLast6Weeks[2];
        const w4 = globalLast6Weeks[3];
        const w5 = globalLast6Weeks[4];
        const w6 = globalLast6Weeks[5];

        structureList.forEach(item => {
            let metaText = '-';
            let statusText = '-';
            let statusClass = '';
            let metaVal = NaN;
            let isLowerBetter = false;

            const getKpiData = (week) => [...allData].reverse().find(d => d['KPI'] === item.kpi && d['Período'] === week);
            
            const w6Data = getKpiData(w6);
            const w5Data = getKpiData(w5);
            const w4Data = getKpiData(w4);
            const w3Data = getKpiData(w3);
            const w2Data = getKpiData(w2);
            const w1Data = getKpiData(w1);
            
            let w1Text = w1Data ? (w1Data['Resultado Final'] || '-') : '-';
            let w2Text = w2Data ? (w2Data['Resultado Final'] || '-') : '-';
            let w3Text = w3Data ? (w3Data['Resultado Final'] || '-') : '-';
            let w4Text = w4Data ? (w4Data['Resultado Final'] || '-') : '-';
            let w5Text = w5Data ? (w5Data['Resultado Final'] || '-') : '-';
            let w6Text = w6Data ? (w6Data['Resultado Final'] || '-') : '-';

            const anyData = w6Data || w5Data || w4Data || w3Data || w2Data || w1Data || [...allData].reverse().find(d => d['KPI'] === item.kpi);

            if (anyData && anyData['Meta 4 Pontos']) {
                metaText = anyData['Meta 4 Pontos'].replace('> ', '').replace('< ', '');
                if (anyData['Meta 4 Pontos'].includes('<')) {
                    isLowerBetter = true;
                }
            } else if (item.meta) {
                metaText = item.meta.replace('> ', '').replace('< ', '');
                if (item.meta.includes('<')) {
                    isLowerBetter = true;
                }
            }
            metaVal = parseFloat(metaText.replace('%', '').trim());

            const currentWeekData = getKpiData(statusCurrentWeek);
            const prevWeekData = getKpiData(statusPrevWeek);

            let currentWeekText = currentWeekData ? (currentWeekData['Resultado Final'] || '-') : '-';
            let prevWeekText = prevWeekData ? (prevWeekData['Resultado Final'] || '-') : '-';

            let currentVal = parseFloat(currentWeekText.replace('%', '').trim());
            let prevVal = parseFloat(prevWeekText.replace('%', '').trim());

            if (!isNaN(currentVal) && !isNaN(prevVal)) {
                let melhorou = isLowerBetter ? (currentVal < prevVal) : (currentVal > prevVal);
                let manteve = currentVal === prevVal;

                if (melhorou) {
                    statusText = '🟢 Progrediu';
                    statusClass = 'text-green';
                } else if (manteve) {
                    statusText = '🟡 Manteve';
                    statusClass = 'text-yellow';
                } else {
                    statusText = '🔴 Piorou';
                    statusClass = 'text-red-val';
                }
            } else if (!isNaN(currentVal)) {
                statusText = '🟡 S/ Ref';
                statusClass = 'text-yellow';
            }

            const tr = document.createElement('tr');
            
            const tdKpi = document.createElement('td');
            tdKpi.textContent = item.label;
            
            const tdMeta = document.createElement('td');
            tdMeta.textContent = metaText;
            
            const applyColor = (td, valText) => {
                if (valText === '-' || isNaN(metaVal)) return;
                let val = parseFloat(valText.replace('%','').trim());
                if (!isNaN(val)) {
                    let fail = isLowerBetter ? (val > metaVal) : (val < metaVal);
                    if (fail) {
                        td.classList.add('text-red-val');
                    } else {
                        td.classList.add('text-green-val');
                    }
                }
            };

            const makeEditable = (td, weekVal) => {
                if (window.userRole !== 'Master') return;
                if (!weekVal) return;
                td.title = "Clique para editar/inserir valor manualmente";
                td.style.cursor = "pointer";
                td.addEventListener('click', (e) => {
                    if (td.querySelector('input')) return; 
                    const currentVal = td.textContent === '-' ? '' : td.textContent;
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.value = currentVal;
                    input.style.width = '100%';
                    input.style.boxSizing = 'border-box';
                    input.style.textAlign = 'center';
                    input.style.border = '2px solid #0078D4'; 
                    input.style.padding = '4px';
                    input.style.outline = 'none';
                    td.innerHTML = '';
                    td.appendChild(input);
                    
                    setTimeout(() => {
                        input.focus();
                        input.setSelectionRange(input.value.length, input.value.length);
                    }, 10);

                    const saveVal = () => {
                        let newVal = input.value.trim();
                        if (newVal !== '' && newVal !== '-' && !newVal.includes('%')) {
                            const numCheck = parseFloat(newVal.replace(',', '.'));
                            if (!isNaN(numCheck)) {
                                newVal = newVal + ' %';
                            }
                        }

                        const opNameMap = {
                            'table-first-mile': 'First Mile',
                            'table-last-mile': 'Last Mile',
                            'table-safety': 'Safety'
                        };
                        const opName = opNameMap[tableId] || 'Outros';

                        fetch('/api/manual-data', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ operation: opName, indicator: item.kpi, week: weekVal, value: newVal || '' })
                        }).then(res => {
                            if (!res.ok) throw new Error('Servidor não encontrou a rota. Você reiniciou o Iniciar_Painel.bat?');
                            location.reload();
                        }).catch(err => {
                            alert('Erro ao salvar valor: ' + err.message);
                            td.textContent = currentVal || '-';
                        });
                    };

                    input.addEventListener('blur', saveVal);
                    input.addEventListener('keydown', (e) => {
                        if (e.key === 'Enter') {
                            input.removeEventListener('blur', saveVal);
                            saveVal();
                        }
                    });
                });
            };

            const tdW1 = document.createElement('td');
            tdW1.textContent = w1Text;
            applyColor(tdW1, w1Text);
            makeEditable(tdW1, w1);
            if (w1Data && w1Data.isManualAPI) tdW1.setAttribute('data-manual', 'true');

            const tdW2 = document.createElement('td');
            tdW2.textContent = w2Text;
            applyColor(tdW2, w2Text);
            makeEditable(tdW2, w2);
            if (w2Data && w2Data.isManualAPI) tdW2.setAttribute('data-manual', 'true');

            const tdW3 = document.createElement('td');
            tdW3.textContent = w3Text;
            applyColor(tdW3, w3Text);
            makeEditable(tdW3, w3);
            if (w3Data && w3Data.isManualAPI) tdW3.setAttribute('data-manual', 'true');

            const tdW4 = document.createElement('td');
            tdW4.textContent = w4Text;
            applyColor(tdW4, w4Text);
            makeEditable(tdW4, w4);
            if (w4Data && w4Data.isManualAPI) tdW4.setAttribute('data-manual', 'true');

            const tdW5 = document.createElement('td');
            tdW5.textContent = w5Text;
            applyColor(tdW5, w5Text);
            makeEditable(tdW5, w5);
            if (w5Data && w5Data.isManualAPI) tdW5.setAttribute('data-manual', 'true');
            if (!w5) {
                tdW5.style.display = 'none';
            }

            const tdW6 = document.createElement('td');
            tdW6.textContent = w6Text;
            applyColor(tdW6, w6Text);
            makeEditable(tdW6, w6);
            if (w6Data && w6Data.isManualAPI) tdW6.setAttribute('data-manual', 'true');
            if (!w6) {
                tdW6.style.display = 'none';
            }

            const tdStatus = document.createElement('td');
            tdStatus.textContent = statusText;
            tdStatus.className = statusClass;
            tdStatus.style.fontWeight = '700';

            tdKpi.style.cursor = 'pointer';
            tdKpi.title = "Clique para abrir uma Ação/Tarefa";
            tdKpi.style.color = '#0078D4';
            tdKpi.style.textDecoration = 'underline';
            
            tdKpi.addEventListener('click', () => {
                const opNameMap = {
                    'table-first-mile': 'First Mile',
                    'table-last-mile': 'Last Mile',
                    'table-safety': 'Safety'
                };
                const opName = opNameMap[tableId] || 'Outros';
                
                let existingAction = null;
                if (allActions && allActions[opName]) {
                    existingAction = allActions[opName].find(m => m.indicator === item.label || m.indicator === `⚠️ ${item.label}`);
                }
                
                if (existingAction) {
                    openActionModalForEdit({
                        opName: opName,
                        indicator: item.label,
                        isAuto: false,
                        rawAction: existingAction
                    });
                } else {
                    if (window.userRole !== 'Master') {
                        alert('Apenas o usuário Master pode criar novas ações.');
                        return;
                    }
                    openActionModalForEdit({
                        opName: opName,
                        indicator: item.label,
                        isAuto: true
                    });
                }
            });

            tr.appendChild(tdKpi);
            tr.appendChild(tdMeta);
            tr.appendChild(tdW1);
            tr.appendChild(tdW2);
            tr.appendChild(tdW3);
            tr.appendChild(tdW4);
            tr.appendChild(tdW5);
            tr.appendChild(tdW6);
            tr.appendChild(tdStatus);
            
            tbody.appendChild(tr);
        });
    }

    function updateCentralPanel() {
        document.getElementById('last-update').textContent = `1 de jul. de 2026, 02:06:34`;
    }
});

let currentOperation = '';
let allActions = {};

// Modal
function openActionModal(opName) {
    currentOperation = opName;
    document.getElementById('modal-title').textContent = `Ações: ${opName}`;
    
    const selectInd = document.getElementById('action-indicator');
    selectInd.innerHTML = '<option value="Geral">Geral (Sem indicador específico)</option>';
    
    const tableIdMap = {
        'First Mile': 'table-first-mile',
        'Last Mile': 'table-last-mile',
        'Safety': 'table-safety'
    };
    
    const tableId = tableIdMap[opName];
    if (tableId) {
        if (window.autoTasks) {
            window.autoTasks.filter(t => t.operation === opName).forEach(t => {
                const opt = document.createElement('option');
                opt.value = t.indicator;
                opt.textContent = t.indicator;
                selectInd.appendChild(opt);
            });
        }
    }
    
    document.getElementById('action-modal').style.display = 'block';
    loadActions();
}

function closeActionModal() {
    document.getElementById('action-modal').style.display = 'none';
    currentOperation = '';
}

window.onclick = function(event) {
    const modal = document.getElementById('action-modal');
    if (event.target === modal) {
        closeActionModal();
    }
}

// Fetch Actions
async function loadActions() {
    try {
        const response = await fetch('/api/actions');
        if (response.ok) {
            allActions = await response.json();
            populateGlobalTasks();
        }
    } catch (err) {
        console.error('Erro ao carregar ações. O servidor está rodando?', err);
    }
}

// Abrir modal para edição de uma tarefa (manual ou auto)
function openActionModalForEdit(actionData) {
    currentOperation = actionData.opName;
    document.getElementById('modal-title').textContent = `Editar Ação: ${actionData.opName}`;
    
    document.getElementById('action-id').value = '';
    document.getElementById('action-is-auto').value = 'false';
    document.getElementById('action-indicator').innerHTML = '<option value="Geral">Geral</option>';
    document.getElementById('action-task').value = '';
    document.getElementById('action-mile').value = actionData.opName || 'First Mile';
    document.getElementById('action-owner').value = '';
    document.getElementById('action-start-date').value = '';
    document.getElementById('action-deadline').value = '';
    document.getElementById('action-status').value = 'Em andamento';
    
    if (actionData.isAuto) {
        document.getElementById('action-indicator').innerHTML = `<option value="${actionData.indicator}">${actionData.indicator}</option>`;
        document.getElementById('action-deadline').value = actionData.deadline || '';
        document.getElementById('action-status').value = 'Em andamento';
        document.getElementById('action-is-auto').value = 'true';
        if (document.getElementById('btn-delete-action')) document.getElementById('btn-delete-action').style.display = 'none';
    } else {
        const raw = actionData.rawAction;
        document.getElementById('action-id').value = raw.id || '';
        
        let indName = raw.indicator ? raw.indicator.replace('⚠️ ', '') : 'Geral';
        document.getElementById('action-indicator').innerHTML = `<option value="${indName}">${indName}</option>`;
        
        document.getElementById('action-task').value = raw.task || '';
        document.getElementById('action-mile').value = raw.mile || actionData.opName || 'First Mile';
        document.getElementById('action-owner').value = raw.owner || '';
        
        if (raw.startDate) document.getElementById('action-start-date').value = raw.startDate.split('T')[0];
        if (raw.deadline) document.getElementById('action-deadline').value = raw.deadline.split('T')[0];
        
        document.getElementById('action-status').value = raw.status || 'Em andamento';
        if (document.getElementById('btn-delete-action')) {
            document.getElementById('btn-delete-action').style.display = (window.userRole === 'Master') ? 'block' : 'none';
        }
        
        const historyList = document.getElementById('action-history-list');
        if (historyList) {
            historyList.innerHTML = '';
            if (raw.history && raw.history.length > 0) {
                raw.history.forEach(h => {
                    const d = new Date(h.date).toLocaleString('pt-BR');
                    const div = document.createElement('div');
                    div.style.marginBottom = '5px';
                    div.innerHTML = `<strong>${d}</strong> - <em>${h.user}</em>: ${h.change}`;
                    historyList.appendChild(div);
                });
            } else {
                historyList.innerHTML = '<em>Nenhum histórico registrado.</em>';
            }
        }
    }
    
    // Disable indicator name edit for non-master
    const isMaster = window.userRole === 'Master';
    document.getElementById('action-indicator').disabled = !isMaster;
    
    // The rest of the fields remain enabled so the user can modify the task.
    document.getElementById('action-task').disabled = false;
    document.getElementById('action-mile').disabled = false;
    document.getElementById('action-owner').disabled = false;
    document.getElementById('action-start-date').disabled = false;
    document.getElementById('action-deadline').disabled = false;
    document.getElementById('action-status').disabled = false;
    
    document.getElementById('action-modal').style.display = 'block';
}

// Save Action (POST ou PUT)
async function saveAction() {
    const id = document.getElementById('action-id').value;
    const isAuto = document.getElementById('action-is-auto').value === 'true';
    
    const actionData = {
        indicator: document.getElementById('action-indicator').value,
        task: document.getElementById('action-task').value,
        mile: document.getElementById('action-mile').value,
        owner: document.getElementById('action-owner').value,
        startDate: document.getElementById('action-start-date').value,
        deadline: document.getElementById('action-deadline').value,
        status: document.getElementById('action-status').value
    };
    
    if (!actionData.task) {
        alert("Descreva a tarefa!");
        return;
    }
    
    const method = id ? 'PUT' : 'POST';
    const url = id 
        ? `/api/actions/${encodeURIComponent(currentOperation)}/${id}`
        : `/api/actions`;

    // Append audit info
    if (id) {
        actionData.userEmail = window.userEmail || window.userRole;
        actionData.changeDescription = `Atualizou a Ação/Tarefa`;
    } else {
        actionData.userEmail = window.userEmail || window.userRole;
    }

    const bodyPayload = id ? actionData : { operation: currentOperation, action: actionData, userEmail: window.userEmail || window.userRole };

    try {
        const response = await fetch(url, {
            method: method,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(bodyPayload)
        });
        
        if (response.ok) {
            closeActionModal();
            loadActions(); 
        } else {
            alert('Erro ao salvar ação.');
        }
    } catch (err) {
        console.error(err);
        alert('Erro de conexão ao salvar.');
    }
}

// Complete Action explicitly
function completeAction() {
    document.getElementById('action-status').value = 'Concluído';
    saveAction();
}

async function deleteCurrentAction() {
    const id = document.getElementById('action-id').value;
    if (!id) {
        alert("Esta ação ainda não foi salva, portanto não pode ser excluída.");
        return;
    }
    
    if (!confirm('Deseja realmente excluir esta ação?')) return;
    
    try {
        const response = await fetch(`/api/actions/${encodeURIComponent(currentOperation)}/${id}`, {
            method: 'DELETE'
        });
        if (response.ok) {
            closeActionModal();
            loadActions();
        } else {
            alert('Erro ao excluir ação.');
        }
    } catch (err) {
        console.error(err);
        alert('Erro de conexão ao excluir.');
    }
}

// Render Global Tasks Table
function populateGlobalTasks() {
    const tbody = document.getElementById('table-tasks');
    if (!tbody) return;
    tbody.innerHTML = '';
    
    let delayed = 0;
    let total = 0;
    
    const combinedActions = [];
    
    if (window.autoTasks) {
        window.autoTasks.forEach(a => {
            let alreadyHasAction = false;
            if (allActions && allActions[a.operation]) {
                alreadyHasAction = allActions[a.operation].some(m => 
                    m.indicator === a.indicator || m.indicator === `⚠️ ${a.indicator}`
                );
            }
            if (!alreadyHasAction) {
                combinedActions.push({
                    opName: a.operation,
                    indicator: a.indicator,
                    task: a.task,
                    status: a.status,
                    deadline: a.deadline,
                    isAuto: true,
                    rawAction: null 
                });
            }
        });
    }

    for (const [opName, actions] of Object.entries(allActions)) {
        actions.forEach(a => {
            combinedActions.push({
                opName: opName,
                indicator: a.indicator,
                task: a.task,
                status: a.status,
                deadline: a.deadline,
                owner: a.owner,
                isAuto: false,
                rawAction: a 
            });
        });
    }
    
    const isMaster = window.userRole === 'Master';
    let filteredActions = combinedActions;
    if (!isMaster && window.userRole) {
        filteredActions = combinedActions.filter(a => a.opName === window.userRole);
    }
    
    filteredActions.forEach(a => {
        total++;
        const tr = document.createElement('tr');
        
        const tdOp = document.createElement('td');
        tdOp.innerHTML = `<span style="font-weight: bold; color: #0A246A;">${a.opName}</span>`;
        
        const tdInd = document.createElement('td');
        tdInd.textContent = a.indicator ? a.indicator.replace('⚠️ ', '') : '-';

        const tdTask = document.createElement('td');
        tdTask.textContent = a.task || '-';
        
        const tdStatus = document.createElement('td');
        
        let finalStatus = a.status || 'Em andamento';
        if (finalStatus !== 'Concluído') {
            let sDateStr = a.startDate || a.createdAt;
            if (sDateStr) {
                let sDate = new Date(sDateStr);
                let today = new Date();
                sDate.setHours(0,0,0,0);
                today.setHours(0,0,0,0);
                
                let timeDiff = today.getTime() - sDate.getTime();
                let ageDays = Math.round(timeDiff / (1000 * 3600 * 24));
                
                if (ageDays >= 8) {
                    finalStatus = 'Atrasado';
                } else if (ageDays === 7) {
                    finalStatus = 'Atenção';
                } else {
                    finalStatus = 'Em andamento';
                }
            }
        }
        
        if (finalStatus.toUpperCase() === 'ATRASADO' || finalStatus.toUpperCase() === 'ATRASO') {
            tdStatus.innerHTML = `<span style="color: red; font-weight: bold;">Atrasado</span>`;
            delayed++;
        } else if (finalStatus.toUpperCase() === 'CRÍTICO' || finalStatus.toUpperCase() === 'CRITICO') {
            tdStatus.innerHTML = `<span style="color: #FF4500; font-weight: bold;">Crítico</span>`;
        } else if (finalStatus.toUpperCase() === 'ATENÇÃO' || finalStatus.toUpperCase() === 'ATENCAO') {
            tdStatus.innerHTML = `<span style="color: #E6A800; font-weight: bold;">Atenção</span>`;
        } else {
            const formattedStatus = finalStatus.charAt(0).toUpperCase() + finalStatus.slice(1).toLowerCase();
            tdStatus.innerHTML = `<span style="color: #00B050; font-weight: bold;">${formattedStatus}</span>`;
        }
        
        const tdOwner = document.createElement('td');
        tdOwner.textContent = a.owner || '-';

        const tdDeadline = document.createElement('td');
        if (a.deadline) {
            const dlDate = new Date(a.deadline);
            dlDate.setMinutes(dlDate.getMinutes() + dlDate.getTimezoneOffset());
            tdDeadline.textContent = dlDate.toLocaleDateString('pt-BR');
        } else {
            tdDeadline.textContent = '-';
        }

        tr.appendChild(tdOp);
        tr.appendChild(tdInd);
        tr.appendChild(tdTask);
        tr.appendChild(tdOwner);
        tr.appendChild(tdDeadline);
        tr.appendChild(tdStatus);
        
        tr.style.cursor = 'pointer';
        tr.title = "Dê dois cliques para editar esta tarefa";
        tr.addEventListener('dblclick', () => {
            openActionModalForEdit(a);
        });
        
        tbody.appendChild(tr);
    });
    
    window.totalTasks = total;
    window.delayedTasks = delayed;
    
    if (total === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;">Nenhuma tarefa encontrada.</td></tr>';
    }

    const totalEl = document.getElementById('total-tasks');
    const delayedEl = document.getElementById('delayed-tasks');
    if (totalEl) totalEl.textContent = total;
    if (delayedEl) delayedEl.textContent = delayed;
}

window.toggleTasksTable = function() {
    const tbody = document.getElementById('table-tasks');
    const btn = document.getElementById('btn-toggle-tasks');
    if (tbody.style.display === 'none') {
        tbody.style.display = '';
        btn.textContent = 'Minimizar';
    } else {
        tbody.style.display = 'none';
        btn.textContent = 'Maximizar';
    }
};

async function deleteAction(id) {
    if (!confirm('Deseja realmente excluir esta ação?')) return;
    
    try {
        const response = await fetch(`/api/actions/${encodeURIComponent(currentOperation)}/${id}`, {
            method: 'DELETE'
        });
        if (response.ok) {
            loadActions();
        }
    } catch (err) {
        console.error(err);
    }
}

function renderActions() {
    const tbody = document.getElementById('modal-actions-tbody');
    tbody.innerHTML = '';
    
    const actions = allActions[currentOperation] || [];
    
    if (actions.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;">Nenhuma ação registrada.</td></tr>';
        return;
    }

    actions.forEach(a => {
        const tr = document.createElement('tr');
        
        const tdInd = document.createElement('td'); 
        tdInd.textContent = a.indicator ? a.indicator.replace('⚠️ ', '') : 'Geral';
        const tdTask = document.createElement('td'); tdTask.textContent = a.task;
        const tdOwner = document.createElement('td'); tdOwner.textContent = a.owner;
        const tdDead = document.createElement('td'); 
        tdDead.textContent = a.deadline ? new Date(a.deadline).toLocaleDateString('pt-BR') : '-';
        const tdStat = document.createElement('td'); tdStat.textContent = a.status;
        
        const tdBtn = document.createElement('td');
        tdBtn.innerHTML = `<button class="btn-delete" onclick="deleteAction('${a.id}')">Excluir</button>`;
        
        tr.appendChild(tdInd);
        tr.appendChild(tdTask);
        tr.appendChild(tdOwner);
        tr.appendChild(tdDead);
        tr.appendChild(tdStat);
        tr.appendChild(tdBtn);
        
        tbody.appendChild(tr);
    });
}

// Load actions and set current date on page load
document.addEventListener("DOMContentLoaded", () => {
    const accessDateEl = document.getElementById('current-access-date');
    if (accessDateEl) {
        const now = new Date();
        accessDateEl.textContent = now.toLocaleDateString('pt-BR') + ' às ' + now.toLocaleTimeString('pt-BR');
    }
    loadActions();
});

// Take Screenshot and Download
async function downloadScreenshot() {
    const btn = document.querySelector('.btn-print');
    const oldText = btn.innerHTML;
    btn.innerHTML = '⏳ Gerando...';
    btn.disabled = true;
    
    try {
        const container = document.querySelector('.dashboard-container');
        
        // Esconder temporariamente as linhas de ação durante o print
        const tbody = document.getElementById('table-tasks');
        const oldDisplay = tbody.style.display;
        tbody.style.display = 'none';
        
        const canvas = await html2canvas(container, {
            scale: 2,
            useCORS: true,
            logging: false
        });
        
        // Restaurar exibição
        tbody.style.display = oldDisplay;
        
        canvas.toBlob(async function(blob) {
            try {
                // Copy to clipboard
                const item = new ClipboardItem({ "image/png": blob });
                await navigator.clipboard.write([item]);
                
                // Download file
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `Print_Gerot_${new Date().getTime()}.png`;
                a.click();
                URL.revokeObjectURL(url);
                
                alert("O Print do dashboard foi baixado no seu computador e também copiado para sua área de transferência (Ctrl+V)!");
            } catch (err) {
                console.error(err);
                
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `Print_Gerot_${new Date().getTime()}.png`;
                a.click();
                URL.revokeObjectURL(url);
                
                alert("O Print do dashboard foi baixado no seu computador!");
            } finally {
                btn.innerHTML = oldText;
                btn.disabled = false;
            }
        });
    } catch (err) {
        console.error(err);
        alert('Erro ao gerar print da tela.');
        btn.innerHTML = oldText;
        btn.disabled = false;
    }
}

// Export to Excel
function exportToExcel() {
    fetch('/api/export-excel')
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                const a = document.createElement('a');
                a.href = data.downloadUrl;
                a.download = data.downloadUrl.replace('/', '');
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            } else {
                alert("Erro ao gerar Excel: " + (data.error || "Desconhecido"));
            }
        })
        .catch(err => {
            console.error(err);
            alert("Erro de conexão ao tentar gerar Excel. Verifique se o Iniciar_Painel.bat está rodando.");
        });
}

window.editAction = function(actionData) {
    const actionStr = encodeURIComponent(JSON.stringify(actionData));
    window.open(`plano_acao.html?action=${actionStr}`, '_blank');
};

window.toggleManualHighlight = () => { document.body.classList.toggle('highlight-manual'); };

// ================= USER MANAGEMENT ================= //
async function openUsersModal() {
    document.getElementById('users-modal').style.display = 'block';
    await loadUsers();
}

function closeUsersModal() {
    document.getElementById('users-modal').style.display = 'none';
}

async function loadUsers() {
    const tbody = document.getElementById('users-tbody');
    tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">Carregando...</td></tr>';
    
    try {
        const res = await fetch('/api/users');
        const users = await res.json();
        
        tbody.innerHTML = '';
        users.forEach(u => {
            const tr = document.createElement('tr');
            
            const tdEmail = document.createElement('td');
            tdEmail.textContent = u.email;
            
            const tdDate = document.createElement('td');
            tdDate.textContent = new Date(u.createdAt).toLocaleDateString('pt-BR');
            
            const tdRole = document.createElement('td');
            const selRole = document.createElement('select');
            selRole.innerHTML = `
                <option value="Master" ${u.role === 'Master' ? 'selected' : ''}>Master</option>
                <option value="First Mile" ${u.role === 'First Mile' ? 'selected' : ''}>First Mile</option>
                <option value="Last Mile" ${u.role === 'Last Mile' ? 'selected' : ''}>Last Mile</option>
                <option value="Safety" ${u.role === 'Safety' ? 'selected' : ''}>Safety</option>
                <option value="SRM" ${u.role === 'SRM' ? 'selected' : ''}>SRM</option>
                <option value="XPT" ${u.role === 'XPT' ? 'selected' : ''}>XPT</option>
                <option value="Nenhuma" ${u.role === 'Nenhuma' ? 'selected' : ''}>Nenhuma</option>
            `;
            selRole.onchange = (e) => updateUser(u.id, { role: e.target.value });
            tdRole.appendChild(selRole);
            
            const tdStatus = document.createElement('td');
            const selStatus = document.createElement('select');
            selStatus.innerHTML = `
                <option value="Aprovado" ${u.status === 'Aprovado' ? 'selected' : ''}>Aprovado</option>
                <option value="Pendente" ${u.status === 'Pendente' ? 'selected' : ''}>Pendente</option>
                <option value="Bloqueado" ${u.status === 'Bloqueado' ? 'selected' : ''}>Bloqueado</option>
            `;
            selStatus.onchange = (e) => updateUser(u.id, { status: e.target.value });
            tdStatus.appendChild(selStatus);
            
            const tdAction = document.createElement('td');
            tdAction.textContent = "Salva auto.";
            tdAction.style.color = "#888";
            tdAction.style.fontSize = "12px";
            
            tr.appendChild(tdEmail);
            tr.appendChild(tdDate);
            tr.appendChild(tdRole);
            tr.appendChild(tdStatus);
            tr.appendChild(tdAction);
            
            tbody.appendChild(tr);
        });
    } catch (err) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:red;">Erro ao carregar usuários.</td></tr>';
    }
}

async function updateUser(id, payload) {
    try {
        const res = await fetch(`/api/users/${id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });
        if (!res.ok) {
            alert('Erro ao atualizar usuário.');
        } else {
            // Optional: Show a tiny success feedback
        }
    } catch (err) {
        console.error(err);
        alert('Erro ao conectar ao servidor.');
    }
}
// ================================================== //