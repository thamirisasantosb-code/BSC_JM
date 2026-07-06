var allActions = {};
window.userRole = localStorage.getItem('bsc_user_role');
window.userEmail = localStorage.getItem('bsc_user_email');

// Interceptar fetch global para injetar token JWT
const originalFetch = window.fetch;
window.fetch = function(url, options = {}) {
    const token = localStorage.getItem('bsc_token');
    const isApi = typeof url === 'string' && url.includes('/api/');
    
    if (isApi) {
        if (!options.headers) options.headers = {};
        if (token && !options.headers['Authorization']) {
            options.headers['Authorization'] = 'Bearer ' + token;
        }
        if (!options.headers['Content-Type'] && !(options.body instanceof FormData) && typeof options.body === 'string') {
            options.headers['Content-Type'] = 'application/json';
        }
    }
    
    return originalFetch(url, options).then(response => {
        if (isApi && (response.status === 401 || response.status === 403)) {
            logoutRole();
            throw new Error('Sessão expirada. Por favor, faça login novamente.');
        }
        return response;
    });
};

// ==================== TOAST NOTIFICATIONS ====================
window.showToast = function(message, type = 'success', duration = 4000) {
    const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
    const container = document.getElementById('toast-container');
    if (!container) { console.warn(message); return; }
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <span class="toast-icon">${icons[type] || 'ℹ️'}</span>
        <span class="toast-msg">${message}</span>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>`;
    container.appendChild(toast);
    setTimeout(() => {
        toast.classList.add('toast-out');
        setTimeout(() => toast.remove(), 280);
    }, duration);
};

// ==================== CONFIRM MODAL ====================
window.showConfirm = function({ title = 'Confirmar ação', message = 'Tem certeza?', icon = '⚠️', okLabel = 'Confirmar', cancelLabel = 'Cancelar' }) {
    return new Promise((resolve) => {
        const modal = document.getElementById('confirm-modal');
        if (!modal) { resolve(window.confirm(message)); return; }
        document.getElementById('confirm-icon').textContent = icon;
        document.getElementById('confirm-title').textContent = title;
        document.getElementById('confirm-msg').textContent = message;
        document.getElementById('confirm-ok').textContent = okLabel;
        document.getElementById('confirm-cancel').textContent = cancelLabel;
        modal.classList.add('open');
        const close = (result) => {
            modal.classList.remove('open');
            okBtn.removeEventListener('click', handleOk);
            cancelBtn.removeEventListener('click', handleCancel);
            resolve(result);
        };
        const okBtn = document.getElementById('confirm-ok');
        const cancelBtn = document.getElementById('confirm-cancel');
        const handleOk = () => close(true);
        const handleCancel = () => close(false);
        okBtn.addEventListener('click', handleOk);
        cancelBtn.addEventListener('click', handleCancel);
    });
};

// Fechar modais com tecla Escape
document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    const actionModal = document.getElementById('action-modal');
    if (actionModal && actionModal.style.display !== 'none') { closeActionModal(); return; }
    const confirmModal = document.getElementById('confirm-modal');
    if (confirmModal && confirmModal.classList.contains('open')) {
        confirmModal.classList.remove('open'); return;
    }
    const kpiModal = document.getElementById('kpi-detail-modal');
    if (kpiModal && kpiModal.classList.contains('open')) { kpiModal.classList.remove('open'); }
});

// ==================== FILTER TASKS TABLE ====================
window.filterTasksTable = function() {
    const search = (document.getElementById('tasks-search')?.value || '').toLowerCase();
    const opFilter = (document.getElementById('filter-op')?.value || '').toLowerCase();
    const statusFilter = (document.getElementById('filter-status')?.value || '').toLowerCase();
    const rows = document.querySelectorAll('#table-tasks tr');
    let visible = 0;
    rows.forEach(tr => {
        const text = tr.textContent.toLowerCase();
        const op = (tr.dataset.op || '').toLowerCase();
        const status = (tr.dataset.status || '').toLowerCase();
        const matchSearch = !search || text.includes(search);
        const matchOp = !opFilter || op.includes(opFilter);
        const matchStatus = !statusFilter || status.includes(statusFilter);
        const show = matchSearch && matchOp && matchStatus;
        tr.style.display = show ? '' : 'none';
        if (show) visible++;
    });
    const countEl = document.getElementById('tasks-count');
    if (countEl) countEl.textContent = `${visible} de ${rows.length} ações`;
};

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
    window.dashboardStructure = structure;

    window.getKpiNameByLabel = function(label) {
        if (!window.dashboardStructure) return label;
        const all = [
            ...window.dashboardStructure.firstMile,
            ...window.dashboardStructure.lastMile,
            ...window.dashboardStructure.safety
        ];
        const match = all.find(item => item.label === label || item.kpi === label);
        return match ? match.kpi : label;
    };

    window.getKpiHistoricalData = function(kpiName, week) {
        if (!window.dashboardIndexedData) return null;
        const dbKpiName = window.getKpiNameByLabel(kpiName);
        return window.dashboardIndexedData[dbKpiName]?.[week] || null;
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
        window.dashboardCombinedData = combinedData;
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
        // Encontrar a linha correspondente ao mês fechado usando o índice O(1)
        const monthData = window.dashboardIndexedMonthData?.[item.kpi];
        
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

        // Fallback original usando o índice O(1)
        const closedData = window.dashboardIndexedData?.[item.kpi]?.[closedWeek];
        
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
        
        // Criar índices O(1) para otimizar busca de dados
        window.dashboardIndexedData = {};
        window.dashboardIndexedMonthData = {};
        data.forEach(d => {
            if (!d || !d['KPI']) return;
            const kpi = d['KPI'];
            const p = d['Período'] || '';
            
            if (!window.dashboardIndexedData[kpi]) {
                window.dashboardIndexedData[kpi] = {};
            }
            window.dashboardIndexedData[kpi][p] = d;

            if (!p.startsWith('W') && p !== 'W21' && p.trim() !== '') {
                window.dashboardIndexedMonthData[kpi] = d;
            }
        });
        
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
        window.globalLast6Weeks = globalLast6Weeks;
        
        window.latestWeek = sortedPeriodos[sortedPeriodos.length - 1];
        window.closedWeek = sortedPeriodos.length > 1 ? sortedPeriodos[sortedPeriodos.length - 2] : window.latestWeek;
        
        const latestWeek = window.latestWeek;
        const closedWeek = window.closedWeek;
        
        statusCurrentWeek = latestWeek;
        statusPrevWeek = sortedPeriodos[sortedPeriodos.length - 2];
        
        const w6th = globalLast6Weeks[5];
        if (!w6th) {
            document.querySelectorAll('.dyn-w6').forEach(th => th.style.display = 'none');
        } else {
            document.querySelectorAll('.dyn-w6').forEach(th => th.style.display = '');
        }
        
        const formatHeader = (wk) => wk === latestWeek ? `${wk} (Prévia)` : (wk || '--');
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

        const latestWeek = window.latestWeek;
        const closedWeek = window.closedWeek;

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

            const getKpiData = (week) => window.dashboardIndexedData?.[item.kpi]?.[week];
            
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

            const anyData = w6Data || w5Data || w4Data || w3Data || w2Data || w1Data || (window.dashboardIndexedData?.[item.kpi] ? Object.values(window.dashboardIndexedData[item.kpi]).pop() : null);

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

            // 1. Calcular a média das semanas fechadas (desconsiderando a prévia/latestWeek)
            let sumAverage = 0;
            let countAverage = 0;
            globalLast6Weeks.forEach(wk => {
                if (wk !== latestWeek) {
                    const pData = getKpiData(wk);
                    const pText = pData ? (pData['Resultado Final'] || '-') : '-';
                    const pVal = parseFloat(pText.replace(/[><%]/g, '').trim().replace(',', '.'));
                    if (!isNaN(pVal)) {
                        sumAverage += pVal;
                        countAverage++;
                    }
                }
            });
            let avgVal = countAverage > 0 ? (sumAverage / countAverage) : null;
            let avgStr = '-';
            if (avgVal !== null) {
                if (metaText.includes('%')) {
                    avgStr = `${avgVal.toFixed(2).replace('.', ',')}%`;
                } else {
                    avgStr = avgVal.toFixed(2).replace('.', ',');
                }
            }

            // 2. Status com base na semana fechada (closedWeek, ex: W26) contra o objetivo
            const closedData = getKpiData(closedWeek);
            const closedText = closedData ? (closedData['Resultado Final'] || '-') : '-';
            const closedVal = parseFloat(closedText.replace(/[><%]/g, '').trim().replace(',', '.'));

            if (!isNaN(closedVal)) {
                let atingiuMeta = false;
                if (!isNaN(metaVal)) {
                    atingiuMeta = isLowerBetter ? (closedVal <= metaVal) : (closedVal >= metaVal);
                }

                if (atingiuMeta) {
                    statusText = 'Atingido';
                    statusClass = 'text-green-val'; // Verde
                } else {
                    statusText = 'Abaixo';
                    statusClass = 'text-red-val'; // Vermelho
                }
            } else {
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
                            if (!res.ok) throw new Error('Servidor não respondeu corretamente.');
                            
                            // Atualizar dados na memória local
                            let existing = window.dashboardCombinedData.find(d => d['KPI'] === item.kpi && d['Período'] === weekVal);
                            if (existing) {
                                existing['Resultado Final'] = newVal;
                                existing['isManualAPI'] = true;
                            } else {
                                window.dashboardCombinedData.push({
                                    'KPI': item.kpi,
                                    'Período': weekVal,
                                    'Meta 4 Pontos': item.meta || '',
                                    'Resultado Final': newVal,
                                    'isManualAPI': true
                                });
                            }
                            
                            // Re-processar e re-renderizar
                            processData(window.dashboardCombinedData);
                            showToast('Valor salvo com sucesso!', 'success');
                        }).catch(err => {
                            showToast('Erro ao salvar valor: ' + err.message, 'error');
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
            tdKpi.title = "Clique para ver gráfico e ações deste indicador";
            tdKpi.style.color = '#0078D4';
            tdKpi.style.textDecoration = 'underline';
            
            tdKpi.addEventListener('click', () => {
                if (typeof window.openActionForKpi === 'function') {
                    window.openActionForKpi(item.label, item.kpi);
                }
            });


            const tdMedia = document.createElement('td');
            tdMedia.textContent = avgStr;
            tdMedia.style.fontWeight = '700';
            tdMedia.classList.add('col-media');

            tr.appendChild(tdKpi);
            tr.appendChild(tdMeta);
            tr.appendChild(tdW1);
            tr.appendChild(tdW2);
            tr.appendChild(tdW3);
            tr.appendChild(tdW4);
            tr.appendChild(tdW5);
            tr.appendChild(tdW6);
            tr.appendChild(tdMedia);
            tr.appendChild(tdStatus);
            
            tbody.appendChild(tr);
        });
    }

    function updateCentralPanel() {
        if (window.csvLastModifiedDate) {
            const dateStr = window.csvLastModifiedDate.toLocaleDateString('pt-BR', { day: 'numeric', month: 'short', year: 'numeric' });
            const timeStr = window.csvLastModifiedDate.toLocaleTimeString('pt-BR');
            document.getElementById('last-update').textContent = `${dateStr}, ${timeStr}`;
        } else {
            document.getElementById('last-update').textContent = `1 de jul. de 2026, 02:06:34`;
        }
    }
});

let currentOperation = '';

// Retorna o nome da operação (First Mile, Last Mile, Safety) para um KPI
window.getOpNameForKpi = function(kpiLabel) {
    if (!window.dashboardStructure) return 'First Mile';
    const { firstMile = [], lastMile = [], safety = [] } = window.dashboardStructure;
    if (firstMile.some(i => i.label === kpiLabel || i.kpi === kpiLabel)) return 'First Mile';
    if (lastMile.some(i => i.label === kpiLabel || i.kpi === kpiLabel))  return 'Last Mile';
    if (safety.some(i => i.label === kpiLabel || i.kpi === kpiLabel))    return 'Safety';
    return 'First Mile'; // fallback
};

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
    if (typeof window.updateActionModalIndicatorDetail === 'function') {
        window.updateActionModalIndicatorDetail();
    }
    loadActions();
}

function closeActionModal() {
    document.getElementById('action-modal').style.display = 'none';
    currentOperation = '';
    if (window._actionDetailChart) {
        window._actionDetailChart.destroy();
        window._actionDetailChart = null;
    }
    if (window._actionDetailMonthChart) {
        window._actionDetailMonthChart.destroy();
        window._actionDetailMonthChart = null;
    }
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
    document.getElementById('action-owner').value = '';
    document.getElementById('action-start-date').value = '';
    document.getElementById('action-deadline').value = '';
    document.getElementById('action-status').value = 'Em andamento';
    
    // Default to the operation name for new action
    const targetMile = actionData.opName || 'First Mile';
    const initialCheckboxes = document.querySelectorAll('#mile-checkboxes input[type="checkbox"]');
    initialCheckboxes.forEach(cb => {
        cb.checked = (cb.value === targetMile);
    });
    updateSelectedMilesText();
    
    if (actionData.isAuto) {
        const selectInd = document.getElementById('action-indicator');
        selectInd.innerHTML = `<option value="${actionData.indicator}">${actionData.indicator}</option>`;
        selectInd.value = actionData.indicator;
        document.getElementById('action-deadline').value = actionData.deadline || '';
        document.getElementById('action-status').value = 'Em andamento';
        document.getElementById('action-is-auto').value = 'true';
        if (document.getElementById('btn-delete-action')) document.getElementById('btn-delete-action').style.display = 'none';
    } else {
        const raw = actionData.rawAction;
        document.getElementById('action-id').value = raw.id || '';
        
        let indName = raw.indicator ? raw.indicator.replace('⚠️ ', '') : 'Geral';
        const selectInd = document.getElementById('action-indicator');
        selectInd.innerHTML = `<option value="${indName}">${indName}</option>`;
        selectInd.value = indName;
        
        document.getElementById('action-task').value = raw.task || '';
        document.getElementById('action-owner').value = raw.owner || '';
        
        const rawMile = raw.mile || actionData.opName || 'First Mile';
        const selectedMiles = typeof rawMile === 'string' ? rawMile.split(',').map(s => s.trim()) : (Array.isArray(rawMile) ? rawMile : []);
        const editCheckboxes = document.querySelectorAll('#mile-checkboxes input[type="checkbox"]');
        editCheckboxes.forEach(cb => {
            cb.checked = selectedMiles.includes(cb.value);
        });
        updateSelectedMilesText();
        
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
    setMileMultiselectDisabled(false);
    document.getElementById('action-owner').disabled = false;
    document.getElementById('action-start-date').disabled = false;
    document.getElementById('action-deadline').disabled = false;
    document.getElementById('action-status').disabled = false;
    
    document.getElementById('action-modal').style.display = 'block';
    if (typeof window.updateActionModalIndicatorDetail === 'function') {
        window.updateActionModalIndicatorDetail();
    }
}

window.openActionForKpi = function(kpiLabel, kpiKey) {
    const opName = (typeof window.getOpNameForKpi === 'function')
        ? window.getOpNameForKpi(kpiLabel)
        : 'First Mile';

    // Verificar se já existe ação para este KPI
    let existingAction = null;
    if (typeof allActions !== 'undefined') {
        for (const op in allActions) {
            const found = (allActions[op] || []).find(m => {
                const ind = m.indicator ? m.indicator.replace('⚠️ ', '').trim().toLowerCase() : '';
                const l1 = kpiLabel.replace('⚠️ ', '').trim().toLowerCase();
                const l2 = kpiKey.replace('⚠️ ', '').trim().toLowerCase();
                return ind === l1 || ind === l2;
            });
            if (found) { existingAction = found; break; }
        }
    }

    if (existingAction) {
        openActionModalForEdit({ opName, indicator: kpiLabel, isAuto: false, rawAction: existingAction });
    } else {
        if (window.userRole !== 'Master') {
            alert('Apenas o usuário Master pode criar novas ações.');
            return;
        }
        openActionModalForEdit({ opName, indicator: kpiLabel, isAuto: true });
    }
};

window.openGlobalActionModal = function() {
    if (window.userRole !== 'Master' && window.userRole !== 'First Mile' && window.userRole !== 'Last Mile' && window.userRole !== 'Safety') {
        alert('Você não tem permissão para cadastrar ações.');
        return;
    }
    
    const opName = window.userRole === 'Master' ? 'First Mile' : window.userRole;
    
    // Reset form fields
    document.getElementById('action-id').value = '';
    document.getElementById('action-is-auto').value = 'false';
    
    // Popular indicators dropdown with ALL indicators for that operation or Geral
    const selectInd = document.getElementById('action-indicator');
    selectInd.innerHTML = '<option value="Geral">Geral (Sem indicador específico)</option>';
    
    // Add all indicators of the current operation
    if (window.dashboardStructure) {
        const structureMap = {
            'First Mile': window.dashboardStructure.firstMile,
            'Last Mile': window.dashboardStructure.lastMile,
            'Safety': window.dashboardStructure.safety
        };
        const list = structureMap[opName] || [];
        list.forEach(item => {
            const opt = document.createElement('option');
            opt.value = item.kpi;
            opt.textContent = item.label;
            selectInd.appendChild(opt);
        });
    }
    
    document.getElementById('action-task').value = '';
    document.getElementById('action-owner').value = '';
    document.getElementById('action-start-date').value = new Date().toISOString().split('T')[0];
    document.getElementById('action-deadline').value = '';
    document.getElementById('action-status').value = 'Em andamento';
    
    // Set checked checkbox for current operation
    const checkboxes = document.querySelectorAll('#mile-checkboxes input[type="checkbox"]');
    checkboxes.forEach(cb => {
        cb.checked = (cb.value === opName);
    });
    if (typeof updateSelectedMilesText === 'function') {
        updateSelectedMilesText();
    }
    
    if (document.getElementById('btn-delete-action')) {
        document.getElementById('btn-delete-action').style.display = 'none';
    }
    
    // Enable fields
    document.getElementById('action-indicator').disabled = false;
    document.getElementById('action-task').disabled = false;
    if (typeof setMileMultiselectDisabled === 'function') {
        setMileMultiselectDisabled(false);
    }
    document.getElementById('action-owner').disabled = false;
    document.getElementById('action-start-date').disabled = false;
    document.getElementById('action-deadline').disabled = false;
    document.getElementById('action-status').disabled = false;
    
    currentOperation = opName;
    document.getElementById('modal-title').textContent = `Nova Ação: ${opName}`;
    
    document.getElementById('action-modal').style.display = 'block';
    if (typeof window.updateActionModalIndicatorDetail === 'function') {
        window.updateActionModalIndicatorDetail();
    }
};

// Save Action (POST ou PUT)
async function saveAction() {
    const id = document.getElementById('action-id').value;
    const isAuto = document.getElementById('action-is-auto').value === 'true';
    
    const selectedMiles = Array.from(document.querySelectorAll('#mile-checkboxes input[type="checkbox"]:checked'))
                               .map(cb => cb.value)
                               .join(', ');
    
    const actionData = {
        indicator: document.getElementById('action-indicator').value,
        task: document.getElementById('action-task').value,
        mile: selectedMiles || currentOperation || 'First Mile',
        owner: document.getElementById('action-owner').value,
        startDate: document.getElementById('action-start-date').value,
        deadline: document.getElementById('action-deadline').value,
        status: document.getElementById('action-status').value
    };
    
    if (!actionData.task) {
        showToast('Descreva a tarefa antes de salvar!', 'warning');
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
            showToast('Ação salva com sucesso! ✅', 'success');
        } else {
            showToast('Erro ao salvar ação. Tente novamente.', 'error');
        }
    } catch (err) {
        console.error(err);
        showToast('Erro de conexão ao salvar.', 'error');
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
        showToast('Esta ação ainda não foi salva e não pode ser excluída.', 'warning');
        return;
    }
    
    const confirmed = await showConfirm({
        title: 'Excluir Ação',
        message: 'Esta ação será excluída permanentemente. Não é possível desfazer!',
        icon: '🗑️',
        okLabel: 'Sim, excluir',
        cancelLabel: 'Cancelar'
    });
    if (!confirmed) return;
    
    try {
        const response = await fetch(`/api/actions/${encodeURIComponent(currentOperation)}/${id}`, {
            method: 'DELETE'
        });
        if (response.ok) {
            closeActionModal();
            loadActions();
            showToast('Ação excluída com sucesso.', 'success');
        } else {
            showToast('Erro ao excluir ação. Tente novamente.', 'error');
        }
    } catch (err) {
        console.error(err);
        showToast('Erro de conexão ao excluir.', 'error');
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
                    mile: a.operation,
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
                mile: a.mile,
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
        tr.dataset.op = (a.mile || a.opName || '').toLowerCase();
        tr.dataset.status = (a.status || '').toLowerCase();
        
        const tdOp = document.createElement('td');
        tdOp.innerHTML = `<span style="font-weight: bold; color: #0A246A;">${a.mile || a.opName}</span>`;
        
        const tdInd = document.createElement('td');
        const kpiClean = a.indicator ? a.indicator.replace('⚠️ ', '') : '-';
        tdInd.textContent = kpiClean;
        if (a.indicator && kpiClean !== '-' && kpiClean !== 'Geral') {
            tdInd.style.cursor = 'pointer';
            tdInd.style.color = '#0078D4';
            tdInd.style.textDecoration = 'underline';
            tdInd.title = "Clique para abrir a tela de ação deste indicador";
            tdInd.addEventListener('click', (e) => {
                e.stopPropagation(); // Evita abrir o modal de edição da ação
                if (typeof window.openActionForKpi === 'function') {
                    const dbKpiName = window.getKpiNameByLabel(kpiClean);
                    window.openActionForKpi(kpiClean, dbKpiName);
                }
            });
        }

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
        tr.title = "Clique para editar esta tarefa";
        tr.addEventListener('click', () => {
            openActionModalForEdit(a);
        });
        
        tbody.appendChild(tr);
    });
    
    // Atualizar contagem de ações após renderizar
    const countEl = document.getElementById('tasks-count');
    if (countEl) countEl.textContent = `${filteredActions.length} ações`;
    // Re-aplicar filtros se algum estiver ativo
    const searchVal = document.getElementById('tasks-search')?.value;
    const opVal = document.getElementById('filter-op')?.value;
    const statusVal = document.getElementById('filter-status')?.value;
    if (searchVal || opVal || statusVal) filterTasksTable();
    
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

// ================= MULTISELECT MILHAS ================= //
window.toggleMileDropdown = function(event) {
    event.stopPropagation();
    const dropdown = document.getElementById('mile-checkboxes');
    if (dropdown.style.display === 'none' || !dropdown.style.display) {
        dropdown.style.display = 'flex';
    } else {
        dropdown.style.display = 'none';
    }
};

window.updateSelectedMilesText = function() {
    const checked = Array.from(document.querySelectorAll('#mile-checkboxes input[type="checkbox"]:checked'))
                         .map(cb => cb.value);
    const textSpan = document.getElementById('selected-miles-text');
    if (checked.length === 0) {
        textSpan.textContent = 'Selecione as milhas';
    } else {
        textSpan.textContent = checked.join(', ');
    }
};

window.setMileMultiselectDisabled = function(disabled) {
    const selectBox = document.querySelector('.custom-multiselect .select-box');
    const checkboxes = document.querySelectorAll('#mile-checkboxes input[type="checkbox"]');
    if (selectBox) {
        if (disabled) {
            selectBox.style.pointerEvents = 'none';
            selectBox.style.backgroundColor = '#f5f5f5';
            selectBox.style.color = '#888';
        } else {
            selectBox.style.pointerEvents = 'auto';
            selectBox.style.backgroundColor = '#fff';
            selectBox.style.color = '#333';
        }
    }
    checkboxes.forEach(cb => {
        cb.disabled = disabled;
    });
};

// Fechar o dropdown de milhas ao clicar fora dele
document.addEventListener('click', (event) => {
    const dropdown = document.getElementById('mile-checkboxes');
    if (dropdown && dropdown.style.display === 'flex') {
        const multiselect = document.querySelector('.custom-multiselect');
        if (multiselect && !multiselect.contains(event.target)) {
            dropdown.style.display = 'none';
        }
    }
});

window.updateActionModalIndicatorDetail = function() {
    const select = document.getElementById('action-indicator');
    const detailPane = document.getElementById('action-modal-indicator-detail');
    
    if (!select || !detailPane) return;
    
    const selectedKpi = select.value;
    if (!selectedKpi || selectedKpi === 'Geral') {
        detailPane.style.display = 'none';
        return;
    }
    
    detailPane.style.display = 'flex';
    document.getElementById('action-detail-kpi-name').textContent = selectedKpi.replace('⚠️ ', '');

    // Fetch Monthly details
    const monthTableWrap = document.getElementById('action-detail-month-table-wrap');
    monthTableWrap.innerHTML = '<div class="kpi-month-loading">⏳ Carregando dados mensais...</div>';

    // Clear previous chart
    if (window._actionDetailChart) {
        window._actionDetailChart.destroy();
        window._actionDetailChart = null;
    }

    fetch('/api/fechamento-mensal').then(r => r.json()).then(json => {
        const allMonthData = json.data || {};
        const MONTHS = json.months || [];

        const cleanName = (name) => name ? name.replace('⚠️', '').replace(/\s+/g, ' ').trim().toLowerCase() : '';
        const sk = cleanName(selectedKpi);

        const monthRows = [];
        MONTHS.forEach(mes => {
            const rows = allMonthData[mes] || [];
            const match = rows.find(r => {
                const rk = cleanName(r.kpi);
                return rk === sk || rk.includes(sk) || sk.includes(rk);
            });
            if (match) monthRows.push({ mes, ...match });
        });

        if (monthRows.length === 0) {
            monthTableWrap.innerHTML = '<div style="text-align:center; padding:12px; color:#94A3B8; font-size:12px;">📭 Sem histórico mensal para este indicador.</div>';
            document.getElementById('action-detail-kpi-meta').textContent = 'Meta: -';
            return;
        }

        // Get Meta text
        let metaText = '-';
        const firstWithMeta = monthRows.find(r => r.meta && r.meta !== '-');
        if (firstWithMeta) {
            metaText = String(firstWithMeta.meta).replace(/\s*[🟢🔴]/g,'').trim();
        }
        document.getElementById('action-detail-kpi-meta').textContent = metaText !== '-' ? `Meta: ${metaText}` : 'Meta: -';

        // Render Table
        monthTableWrap.innerHTML = `
            <table class="kpi-month-table" style="width:100%; border-collapse:collapse; font-size:11.5px;">
                <thead>
                    <tr style="background:#F8FAFC; border-bottom:1.5px solid #E2E8F0; text-align:left;">
                        <th style="padding:4px 12px; font-size:9.5px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:0.5px;">Mês</th>
                        <th style="padding:4px 12px; font-size:9.5px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:0.5px;">Objetivo</th>
                        <th style="padding:4px 12px; font-size:9.5px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:0.5px;">Resultado</th>
                        <th style="padding:4px 12px; font-size:9.5px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:0.5px; text-align:center;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${monthRows.map(r => {
                        const mStr = String(r.meta || '-').replace(/\s*[🟢🔴]/g,'').trim();
                        const rStr = String(r.resultado || '-').replace(/\s*[🟢🔴]/g,'').trim();
                        const mV = parseFloat(mStr.replace(/[%><]/g,'').replace(',','.'));
                        const rV = parseFloat(rStr.replace(/[%><]/g,'').replace(',','.'));
                        const isLower = mStr.includes('<');
                        
                        let statusHtml = '<span style="color:#94A3B8; font-weight:600;">–</span>';
                        let resColor = '#64748B';
                        if (!isNaN(mV) && !isNaN(rV)) {
                            const ok = isLower ? rV <= mV : rV >= mV;
                            statusHtml = ok
                                ? '<span style="background:#DEF7EC; color:#03543F; font-size:9.5px; font-weight:700; padding:2px 8px; border-radius:12px; display:inline-block; min-width:60px; text-align:center;">Atingido</span>'
                                : '<span style="background:#FDE8E8; color:#9B1C1C; font-size:9.5px; font-weight:700; padding:2px 8px; border-radius:12px; display:inline-block; min-width:60px; text-align:center;">Abaixo</span>';
                            resColor = ok ? '#16A34A' : '#DC2626';
                        }
                        
                        return `<tr style="border-bottom: 1px solid #F1F5F9; transition: background 0.2s;">
                            <td style="padding:4px 12px; font-weight:600; color:#1E293B;">${r.mes}</td>
                            <td style="padding:4px 12px; color:#64748B;">${mStr}</td>
                            <td style="padding:4px 12px; font-weight:700; color:${resColor};">${rStr}</td>
                            <td style="padding:4px 12px; text-align:center;">${statusHtml}</td>
                        </tr>`;
                    }).join('')}
                </tbody>
            </table>`;

        // Render Monthly Line Chart
        const monthlyLabels = monthRows.map(r => r.mes);
        const monthlyValues = monthRows.map(r => {
            const v = parseFloat(String(r.resultado || '').replace(/[%><\s🟢🔴]/g,'').replace(',','.'));
            return isNaN(v) ? null : v;
        });
        const monthlyMetas = monthRows.map(r => {
            const v = parseFloat(String(r.meta || '').replace(/[%><\s🟢🔴]/g,'').replace(',','.'));
            return isNaN(v) ? null : v;
        });

        try {
            if (typeof Chart !== 'undefined') {
                const ctx = document.getElementById('action-detail-trend-chart').getContext('2d');
                window._actionDetailChart = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: monthlyLabels,
                        datasets: [
                            {
                                label: 'Resultado',
                                data: monthlyValues,
                                borderColor: '#2563EB',
                                backgroundColor: 'rgba(37,99,235,.07)',
                                pointBackgroundColor: monthlyValues.map((v, i) => {
                                    if (v === null) return '#9CA3AF';
                                    const m = monthlyMetas[i];
                                    if (m === null || isNaN(m)) return '#2563EB';
                                    return v >= m ? '#16A34A' : '#DC2626';
                                }),
                                pointRadius: 5,
                                tension: 0.25,
                                fill: true,
                                spanGaps: true
                            },
                            {
                                label: 'Meta',
                                data: monthlyMetas,
                                borderColor: '#F59E0B',
                                borderDash: [5,5],
                                borderWidth: 2,
                                pointRadius: 0,
                                fill: false
                            }
                        ]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { display: false },
                            tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${c.parsed.y !== null ? c.parsed.y.toFixed(2).replace('.', ',') + '%' : '-'}` } }
                        },
                        scales: {
                            x: { grid: { color: '#F8FAFC' }, ticks: { font: { size: 10 } } },
                            y: { grid: { color: '#F8FAFC' }, ticks: { font: { size: 9 }, callback: v => v + '%' } }
                        }
                    }
                });
            }
        } catch (chartErr) {
            console.error("Erro ao renderizar gráfico mensal de linha no modal de ação:", chartErr);
        }
    }).catch((err) => {
        console.error("Erro ao buscar histórico mensal no modal de ação:", err);
        monthTableWrap.innerHTML = '<div style="text-align:center; padding:12px; color:#DC2626; font-size:11px;">⚠️ Erro ao carregar histórico mensal.</div>';
    });
};

window.toggleMediaColumn = function() {
    document.body.classList.toggle('show-media');
    const btn = document.getElementById('btn-toggle-media');
    if (document.body.classList.contains('show-media')) {
        btn.textContent = '[-] Média';
        btn.style.backgroundColor = '#d32f2f'; // Vermelho para recolher
    } else {
        btn.textContent = '[+] Média';
        btn.style.backgroundColor = '#107c41'; // Verde para expandir
    }
};
// ====================================================== //