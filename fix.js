const fs = require('fs');
const txt = fs.readFileSync('script.js', 'utf8');
const lines = txt.split('\n');
const newLines = lines.slice(0, 192); // Keep until line 192
newLines.push(`// Modal
function openActionModal(opName) {
    currentOperation = opName;
    document.getElementById('modal-title').textContent = \`Ações: \${opName}\`;
    
    // Popular dropdown de indicadores negativos
    const selectInd = document.getElementById('action-indicator');
    selectInd.innerHTML = '<option value="Geral">Geral (Sem indicador específico)</option>';
    
    const tableIdMap = {
        'First Mile': 'table-first-mile',
        'Last Mile': 'table-last-mile',
        'Line Haul': 'table-line-haul',
        'Safety': 'table-safety'
    };
    
    const tableId = tableIdMap[opName];
    if (tableId) {
        const tbody = document.querySelector(\`#\${tableId} tbody\`);
        if (tbody) {
            const rows = tbody.querySelectorAll('tr');
            rows.forEach(tr => {
                const cells = tr.querySelectorAll('td');
                if (cells.length > 0) {
                    const indicatorName = cells[0].textContent.trim();
                    const statusHtml = cells[cells.length - 1].innerHTML.toLowerCase();
                    if (statusHtml.includes('piorou')) {
                        const opt = document.createElement('option');
                        opt.value = indicatorName;
                        opt.textContent = \`⚠️ \${indicatorName}\`;
                        selectInd.appendChild(opt);
                    }
                }
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

// Fechar modal ao clicar fora
window.onclick = function(event) {
    const modal = document.getElementById('action-modal');
    if (event.target === modal) {
        closeActionModal();
    }
}

// Fetch Actions
async function loadActions() {
    try {
        const response = await fetch('http://localhost:3000/api/actions');
        if (response.ok) {
            allActions = await response.json();
            renderActions();
            populateGlobalTasks();
        }
    } catch (err) {
        console.error('Erro ao carregar ações. O servidor está rodando?', err);
    }
}

// Render Global Tasks Table
function populateGlobalTasks() {
    const tbody = document.getElementById('table-tasks');
    if (!tbody) return;
    tbody.innerHTML = '';
    
    let delayed = 0;
    let total = 0;

    const iconMap = {
        'First Mile': '709/709790.png',
        'Last Mile': '679/679720.png',
        'Line Haul': '3063/3063822.png',
        'Safety': '1161/1161388.png'
    };
    
    for (const [opName, actions] of Object.entries(allActions)) {
        actions.forEach(a => {
            total++;
            const tr = document.createElement('tr');
            
            const tdOp = document.createElement('td');
            const iconStr = iconMap[opName] || '1161/1161388.png';
            tdOp.innerHTML = \`<img src="https://cdn-icons-png.flaticon.com/512/\${iconStr}" alt="" style="width:16px;height:16px; border-radius:50%; background:var(--primary-blue); padding:2px;"> \${opName}\`;
            
            const tdTask = document.createElement('td');
            tdTask.textContent = a.indicator && a.indicator !== 'Geral' 
                ? \`[\${a.indicator.replace('⚠️ ', '')}] \${a.task}\` 
                : a.task;
            
            const tdStatus = document.createElement('td');
            if (a.status.toUpperCase() === 'ATRASADO') {
                tdStatus.innerHTML = \`<span class="status-badge status-atraso"><span class="dot red"></span> ATRASO</span>\`;
                delayed++;
            } else {
                tdStatus.innerHTML = \`<span class="status-badge status-andamento"><span class="dot green"></span> \${a.status.toUpperCase()}</span>\`;
            }
            
            tr.appendChild(tdOp);
            tr.appendChild(tdTask);
            tr.appendChild(tdStatus);
            tbody.appendChild(tr);
        });
    }
    
    if (total === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;">Nenhuma ação registrada nas operações.</td></tr>';
    }

    const totalEl = document.getElementById('total-tasks');
    const delayedEl = document.getElementById('delayed-tasks');
    if (totalEl) totalEl.textContent = total;
    if (delayedEl) delayedEl.textContent = delayed;
}

// Save Action
async function saveAction() {
    const indicator = document.getElementById('action-indicator').value;
    const task = document.getElementById('action-task').value;
    const owner = document.getElementById('action-owner').value;
    const deadline = document.getElementById('action-deadline').value;
    const status = document.getElementById('action-status').value;

    if (!task || !owner) {
        alert('Por favor, preencha Tarefa e Responsável!');
        return;
    }

    const newAction = { indicator, task, owner, deadline, status };

    try {
        const response = await fetch('http://localhost:3000/api/actions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ operation: currentOperation, action: newAction })
        });
        
        if (response.ok) {
            // Limpa form
            document.getElementById('action-indicator').value = 'Geral';
            document.getElementById('action-task').value = '';
            document.getElementById('action-owner').value = '';
            document.getElementById('action-deadline').value = '';
            loadActions();
        }
    } catch (err) {
        alert('Erro ao salvar. Verifique se o servidor Iniciar_Painel.bat está rodando.');
    }
}

// Delete Action
async function deleteAction(id) {
    if (!confirm('Deseja realmente excluir esta ação?')) return;
    
    try {
        const response = await fetch(\`http://localhost:3000/api/actions/\${currentOperation}/\${id}\`, {
            method: 'DELETE'
        });
        if (response.ok) {
            loadActions();
        }
    } catch (err) {
        console.error(err);
    }
}

// Render Actions
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
        tdBtn.innerHTML = \`<button class="btn-delete" onclick="deleteAction('\${a.id}')">Excluir</button>\`;
        
        tr.appendChild(tdInd);
        tr.appendChild(tdTask);
        tr.appendChild(tdOwner);
        tr.appendChild(tdDead);
        tr.appendChild(tdStat);
        tr.appendChild(tdBtn);
        
        tbody.appendChild(tr);
    });
}

// Generate PPT
async function generatePPT() {
    const btn = document.getElementById('btn-generate-ppt');
    const oldText = btn.innerHTML;
    btn.innerHTML = '⏳ Gerando (Aguarde)...';
    btn.disabled = true;
    
    try {
        // Tira print do dashboard principal apenas
        const container = document.querySelector('.dashboard-container');
        const canvas = await html2canvas(container, {
            scale: 2,
            useCORS: true,
            logging: false
        });
        
        const base64Image = canvas.toDataURL('image/png');
        
        const response = await fetch('http://localhost:3000/api/generate-ppt', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ screenshotBase64: base64Image })
        });
        
        if (response.ok) {
            const data = await response.json();
            window.open(\`http://localhost:3000\${data.downloadUrl}\`, '_blank');
        } else {
            alert('Erro ao gerar PPT no servidor.');
        }
    } catch (err) {
        console.error(err);
        alert('Erro. Certifique-se de acessar pelo servidor Iniciar_Painel.bat.');
    } finally {
        btn.innerHTML = oldText;
        btn.disabled = false;
    }
}

// Load actions on page load
document.addEventListener("DOMContentLoaded", () => {
    loadActions();
});
`);
fs.writeFileSync('script.js', newLines.join('\n'));
