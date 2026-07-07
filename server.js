const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const { generateExcel } = require('./gerar_visao_executiva');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const nodemailer = require('nodemailer');

const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.gmail.com',
    port: parseInt(process.env.SMTP_PORT) || 587,
    secure: process.env.SMTP_SECURE === 'true',
    auth: {
        user: process.env.SMTP_USER || '',
        pass: process.env.SMTP_PASS || ''
    }
});

async function enviarEmailRecuperacao(email, novaSenha) {
    const assunto = 'Recuperação de Senha - BSC JM';
    const corpo = `
Olá,

Você solicitou a recuperação de senha no painel BSC JM Distribuição.
Sua nova senha temporária é: ${novaSenha}

Atenção: Para acessar o painel com esta nova senha, o usuário Master precisará aprovar seu acesso novamente na tela de Gestão de Usuários.

Atenciosamente,
Equipe JM Distribuição
    `;
    
    if (process.env.SMTP_USER && process.env.SMTP_PASS) {
        try {
            await transporter.sendMail({
                from: `"BSC JM Distribuição" <${process.env.SMTP_USER}>`,
                to: email,
                subject: assunto,
                text: corpo
            });
            console.log(`E-mail de recuperação enviado para: ${email}`);
            return true;
        } catch (err) {
            console.error('Erro ao enviar e-mail via SMTP:', err);
        }
    }
    
    const logMsg = `[E-MAIL SIMULADO] Enviado para: ${email}\nAssunto: ${assunto}\nConteúdo:\n${corpo}\n========================================\n`;
    fs.appendFileSync('recuperacao_senha_log.txt', logMsg, 'utf8');
    console.log(logMsg);
    return false;
}

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || __dirname;

const ACTIONS_FILE = path.join(DATA_DIR, 'actions.json');
const JWT_SECRET = process.env.JWT_SECRET || 'BSC_SECRET_JM_2026_PORTAL_KEY';

// Middleware de Autenticação JWT
const authMiddleware = (req, res, next) => {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];
    if (!token) {
        return res.status(401).json({ error: 'Token de autenticação não fornecido.' });
    }
    try {
        const decoded = jwt.verify(token, JWT_SECRET);
        req.user = decoded;
        next();
    } catch (err) {
        return res.status(403).json({ error: 'Sua sessão expirou ou o token é inválido.' });
    }
};

// Middleware de Autorização de Função (Roles)
const requireRole = (allowedRoles) => {
    return (req, res, next) => {
        if (!req.user || !allowedRoles.includes(req.user.role)) {
            return res.status(403).json({ error: 'Permissão negada. Apenas funções autorizadas.' });
        }
        next();
    };
};

// Middleware
app.use(cors({ exposedHeaders: ['Last-Modified'] }));
app.use(bodyParser.json({ limit: '50mb' }));

// Custom route to serve CSV from Bannco de Dados folder
app.get('/Base_Indicadores_BSC.csv', (req, res) => {
    const csvFilePath = path.join(DATA_DIR, 'Bannco de Dados', 'Base_Indicadores_BSC.csv');
    if (!fs.existsSync(csvFilePath)) {
        return res.status(404).send('Arquivo CSV não encontrado.');
    }
    res.sendFile(csvFilePath);
});

// API: Fechamento Mensal - lê todos os CSVs da pasta Bannco de Dados
// Cache em memória para evitar leitura de disco a cada requisição
let _monthlyCache = null;
let _monthlyCacheTime = 0;
const MONTHLY_CACHE_TTL = 2 * 60 * 1000; // 2 minutos

app.get('/api/fechamento-mensal', authMiddleware, (req, res) => {
    // Servir do cache se ainda válido
    if (_monthlyCache && Date.now() - _monthlyCacheTime < MONTHLY_CACHE_TTL) {
        return res.json(_monthlyCache);
    }

    const dbDir = path.join(DATA_DIR, 'Bannco de Dados');
    const MONTHLY_LABELS = ['Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Jan'];

    try {
        const files = fs.readdirSync(dbDir).filter(f => f.endsWith('.csv'));
        const allRows = [];

        files.forEach(file => {
            const content = fs.readFileSync(path.join(dbDir, file), 'utf8');
            const lines = content.split('\n').slice(1); // skip header
            lines.forEach(line => {
                if (!line.trim()) return;
                const cols = line.split(',').map(c => c.trim());
                const periodo = cols[0];
                if (!periodo || !MONTHLY_LABELS.includes(periodo)) return;

                const milha = cols[3] || '';
                const kpi = cols[4] || '';
                const meta4 = cols[5] || '';
                const resultado = cols[7] || '';

                if (!kpi) return;

                allRows.push({ periodo, milha, kpi, meta: meta4, resultado });
            });
        });

        // Deduplicate: same periodo+kpi - keep first
        const seen = new Set();
        const deduped = allRows.filter(r => {
            const key = `${r.periodo}|${r.kpi}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });

        // Group by month
        const byMonth = {};
        deduped.forEach(r => {
            if (!byMonth[r.periodo]) byMonth[r.periodo] = [];
            byMonth[r.periodo].push(r);
        });

        // Order months correctly
        const orderedMonths = MONTHLY_LABELS.filter(m => byMonth[m]);

        const result = { months: orderedMonths, data: byMonth };

        // Armazenar no cache
        _monthlyCache = result;
        _monthlyCacheTime = Date.now();

        res.json(result);
    } catch (err) {
        console.error('Erro ao ler dados de fechamento:', err);
        res.status(500).json({ error: err.message });
    }
});

// Desabilitar cache para arquivos HTML para garantir que correções de layout e JS sejam recarregadas
app.use((req, res, next) => {
    const isHtml = req.path.endsWith('.html') || req.path === '/' || req.path === '/index.html';
    if (isHtml) {
        res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, private');
        res.setHeader('Pragma', 'no-cache');
        res.setHeader('Expires', '0');
    }
    next();
});

app.use(express.static(__dirname));

// Ensure DATA_DIR exists if custom
if (DATA_DIR !== __dirname && !fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

// Ensure actions.json exists
if (!fs.existsSync(ACTIONS_FILE)) {
    fs.writeFileSync(ACTIONS_FILE, JSON.stringify({}, null, 2), 'utf8');
}

const USERS_FILE = path.join(DATA_DIR, 'users.json');

// Ensure users.json exists
if (!fs.existsSync(USERS_FILE)) {
    const defaultUsers = [
        {
            "id": "1",
            "email": "thamiris.santos@jmdistribuicao.com.br",
            "password": bcrypt.hashSync("JM@2026", 10),
            "role": "Master",
            "status": "Aprovado",
            "createdAt": new Date().toISOString()
        }
    ];
    fs.writeFileSync(USERS_FILE, JSON.stringify(defaultUsers, null, 2), 'utf8');
} else {
    // Migração de senhas em texto puro para hash bcrypt
    try {
        let users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        let modified = false;
        users = users.map(user => {
            if (user.password && !user.password.startsWith('$2a$') && !user.password.startsWith('$2b$')) {
                user.password = bcrypt.hashSync(user.password, 10);
                modified = true;
            }
            return user;
        });
        if (modified) {
            fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
            console.log('Migração: Senhas em texto puro atualizadas para hash bcrypt.');
        }
    } catch (e) {
        console.error('Erro na migração de senhas:', e);
    }
}

const MANUAL_DATA_FILE = path.join(DATA_DIR, 'manual_data.json');

// Ensure manual_data.json exists
if (!fs.existsSync(MANUAL_DATA_FILE)) {
    fs.writeFileSync(MANUAL_DATA_FILE, JSON.stringify({}, null, 2), 'utf8');
}

// GET all actions
app.get('/api/actions', authMiddleware, (req, res) => {
    try {
        const data = fs.readFileSync(ACTIONS_FILE, 'utf8');
        res.json(JSON.parse(data));
    } catch (error) {
        res.status(500).json({ error: 'Erro ao ler ações' });
    }
});

// GET manual data
app.get('/api/manual-data', authMiddleware, (req, res) => {
    try {
        const data = fs.readFileSync(MANUAL_DATA_FILE, 'utf8');
        res.json(JSON.parse(data));
    } catch (error) {
        res.status(500).json({ error: 'Erro ao ler manual data' });
    }
});

// POST manual data (Apenas Master)
app.post('/api/manual-data', authMiddleware, requireRole(['Master']), (req, res) => {
    try {
        const { operation, indicator, week, value } = req.body;
        const data = JSON.parse(fs.readFileSync(MANUAL_DATA_FILE, 'utf8'));
        
        if (!data[operation]) data[operation] = {};
        if (!data[operation][indicator]) data[operation][indicator] = {};
        
        data[operation][indicator][week] = value;
        
        fs.writeFileSync(MANUAL_DATA_FILE, JSON.stringify(data, null, 2), 'utf8');
        res.status(200).json({ success: true });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao salvar manual data' });
    }
});

// POST new action (Apenas Master)
app.post('/api/actions', authMiddleware, requireRole(['Master']), (req, res) => {
    try {
        const { operation, action, userEmail } = req.body;
        const data = JSON.parse(fs.readFileSync(ACTIONS_FILE, 'utf8'));
        
        if (!data[operation]) {
            data[operation] = [];
        }
        
        action.id = Date.now().toString();
        action.createdAt = new Date().toISOString();
        action.history = [{
            date: action.createdAt,
            user: userEmail || 'Sistema',
            change: 'Ação Criada'
        }];
        
        data[operation].push(action);
        
        fs.writeFileSync(ACTIONS_FILE, JSON.stringify(data, null, 2), 'utf8');
        res.status(201).json(action);
    } catch (error) {
        res.status(500).json({ error: 'Erro ao salvar ação' });
    }
});

// DELETE an action (Apenas Master)
app.delete('/api/actions/:operation/:id', authMiddleware, requireRole(['Master']), (req, res) => {
    try {
        const { operation, id } = req.params;
        const data = JSON.parse(fs.readFileSync(ACTIONS_FILE, 'utf8'));
        
        if (data[operation]) {
            data[operation] = data[operation].filter(a => a.id !== id);
            fs.writeFileSync(ACTIONS_FILE, JSON.stringify(data, null, 2), 'utf8');
        }
        
        res.status(200).json({ success: true });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao deletar ação' });
    }
});

// PUT (update) an action
app.put('/api/actions/:operation/:id', authMiddleware, (req, res) => {
    try {
        const { operation, id } = req.params;
        const { userEmail, changeDescription, ...updatedAction } = req.body;
        const data = JSON.parse(fs.readFileSync(ACTIONS_FILE, 'utf8'));
        
        if (data[operation]) {
            const index = data[operation].findIndex(a => a.id === id);
            if (index !== -1) {
                const oldAction = data[operation][index];
                if (!oldAction.history) oldAction.history = [];
                
                oldAction.history.push({
                    date: new Date().toISOString(),
                    user: userEmail || 'Sistema',
                    change: changeDescription || 'Ação Atualizada'
                });
                
                // Merge old with new
                data[operation][index] = { ...oldAction, ...updatedAction, history: oldAction.history };
                fs.writeFileSync(ACTIONS_FILE, JSON.stringify(data, null, 2), 'utf8');
                return res.status(200).json(data[operation][index]);
            }
        }
        
        res.status(404).json({ error: 'Ação não encontrada' });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao atualizar ação' });
    }
});

// ======================= AUTH & USERS ======================= //

// Login
app.post('/api/login', (req, res) => {
    try {
        const email = (req.body.email || '').trim().toLowerCase();
        const password = (req.body.password || '').trim();
        const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        
        const user = users.find(u => u.email.trim().toLowerCase() === email);
        
        if (!user || !bcrypt.compareSync(password, user.password)) {
            return res.status(401).json({ error: 'Email ou senha inválidos' });
        }
        if (user.status !== 'Aprovado') {
            return res.status(403).json({ error: 'Seu acesso ainda está pendente de aprovação.' });
        }
        
        // Gerar token JWT válido por 24 horas
        const token = jwt.sign(
            { id: user.id, email: user.email, role: user.role, visibility: user.visibility || [] },
            JWT_SECRET,
            { expiresIn: '24h' }
        );
        
        res.status(200).json({
            token,
            email: user.email,
            role: user.role,
            status: user.status,
            visibility: user.visibility || []
        });
    } catch (error) {
        res.status(500).json({ error: 'Erro no servidor' });
    }
});

// Register
app.post('/api/register', (req, res) => {
    try {
        const email = (req.body.email || '').trim().toLowerCase();
        const password = (req.body.password || '').trim();
        const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        
        if (users.find(u => u.email.trim().toLowerCase() === email)) {
            return res.status(400).json({ error: 'Email já cadastrado' });
        }
        
        const newUser = {
            id: Date.now().toString(),
            email,
            password: bcrypt.hashSync(password, 10),
            role: 'Nenhuma',
            status: 'Pendente',
            createdAt: new Date().toISOString()
        };
        
        users.push(newUser);
        fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
        
        res.status(201).json({ success: true, message: 'Cadastro realizado. Aguarde aprovação.' });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao registrar' });
    }
});

// Get Users (Apenas Master)
app.get('/api/users', authMiddleware, requireRole(['Master']), (req, res) => {
    try {
        const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        // Return without passwords
        const safeUsers = users.map(u => ({ id: u.id, email: u.email, role: u.role, status: u.status, visibility: u.visibility || [], createdAt: u.createdAt }));
        res.status(200).json(safeUsers);
    } catch (error) {
        res.status(500).json({ error: 'Erro ao buscar usuários' });
    }
});

// Update User (Role/Status) (Apenas Master)
app.put('/api/users/:id', authMiddleware, requireRole(['Master']), (req, res) => {
    try {
        const { id } = req.params;
        const { role, status, visibility } = req.body;
        const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        
        const index = users.findIndex(u => u.id === id);
        if (index === -1) {
            return res.status(404).json({ error: 'Usuário não encontrado' });
        }
        
        if (role) users[index].role = role;
        if (status) users[index].status = status;
        if (visibility !== undefined) users[index].visibility = visibility;
        
        fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
        res.status(200).json({ success: true });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao atualizar usuário' });
    }
});

// POST Forgot Password
app.post('/api/forgot-password', (req, res) => {
    try {
        const email = (req.body.email || '').trim().toLowerCase();
        const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        
        const index = users.findIndex(u => u.email.trim().toLowerCase() === email);
        if (index === -1) {
            return res.status(404).json({ error: 'E-mail não cadastrado.' });
        }
        
        const novaSenha = 'JM-' + Math.random().toString(36).substring(2, 8).toUpperCase();
        
        users[index].password = bcrypt.hashSync(novaSenha, 10);
        users[index].status = 'Pendente'; // Força aprovação pelo Master
        
        fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
        
        enviarEmailRecuperacao(email, novaSenha);
        
        res.status(200).json({ success: true, message: 'Nova senha enviada por e-mail. Seu acesso foi redefinido para Pendente e precisará ser aprovado pelo usuário Master.' });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao processar recuperação de senha.' });
    }
});

// POST Create User (Apenas Master)
app.post('/api/users', authMiddleware, requireRole(['Master']), (req, res) => {
    try {
        const email = (req.body.email || '').trim().toLowerCase();
        const password = (req.body.password || '').trim();
        const role = req.body.role || 'Nenhuma';
        const visibility = req.body.visibility || [];
        const status = req.body.status || 'Aprovado';
        
        const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
        
        if (users.find(u => u.email.trim().toLowerCase() === email)) {
            return res.status(400).json({ error: 'E-mail já cadastrado.' });
        }
        
        const newUser = {
            id: Date.now().toString(),
            email,
            password: bcrypt.hashSync(password, 10),
            role,
            status,
            visibility,
            createdAt: new Date().toISOString()
        };
        
        users.push(newUser);
        fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
        res.status(201).json({ success: true });
    } catch (error) {
        res.status(500).json({ error: 'Erro ao criar usuário.' });
    }
});

// ========================================================== //

// POST Generate PPT
app.post('/api/generate-ppt', authMiddleware, async (req, res) => {
    try {
        const { screenshotBase64, meetingData } = req.body;
        const pres = new pptxgen();

        // Slide 1: Screenshot
        if (screenshotBase64) {
            const slide1 = pres.addSlide();
            slide1.addText('Resultado do Dashboard', { x: 0.5, y: 0.5, fontSize: 24, bold: true, color: '003366' });
            
            // Clean base64 string
            const base64Data = screenshotBase64.replace(/^data:image\/\w+;base64,/, "");
            slide1.addImage({ data: 'image/png;base64,' + base64Data, x: 0.5, y: 1.2, w: 9, h: 4.5, sizing: { type: 'contain', w: 9, h: 4.5 } });
        }

        // Action Slides
        const actionsData = JSON.parse(fs.readFileSync(ACTIONS_FILE, 'utf8'));
        for (const [operation, actions] of Object.entries(actionsData)) {
            if (actions.length > 0) {
                const slide = pres.addSlide();
                slide.addText(`Ações: ${operation}`, { x: 0.5, y: 0.5, fontSize: 24, bold: true, color: '003366' });
                
                const tableRows = [
                    [{ text: 'Indicador', options: { bold: true, fill: '003366', color: 'FFFFFF' } },
                     { text: 'Tarefa', options: { bold: true, fill: '003366', color: 'FFFFFF' } },
                     { text: 'Responsável', options: { bold: true, fill: '003366', color: 'FFFFFF' } },
                     { text: 'Prazo', options: { bold: true, fill: '003366', color: 'FFFFFF' } },
                     { text: 'Status', options: { bold: true, fill: '003366', color: 'FFFFFF' } }]
                ];

                actions.forEach(a => {
                    tableRows.push([
                        { text: a.indicator || 'Geral' },
                        { text: a.task },
                        { text: a.owner },
                        { text: a.deadline },
                        { text: a.status }
                    ]);
                });

                slide.addTable(tableRows, { x: 0.5, y: 1.2, w: 9 });
            }
        }

        const fileName = `Reuniao_Gerot_${Date.now()}.pptx`;
        const filePath = path.join(__dirname, fileName);
        
        await pres.writeFile({ fileName: filePath });
        
        res.status(200).json({ success: true, downloadUrl: `/${fileName}` });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Erro ao gerar PPT' });
    }
});

// GET Generate Excel Master
app.get('/api/export-excel', authMiddleware, async (req, res) => {
    try {
        const filePath = await generateExcel();
        // Return download URL pointing to static file
        res.status(200).json({ success: true, downloadUrl: `/${filePath}` });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Erro ao gerar Excel' });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
});
