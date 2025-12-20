import { makeWASocket, useMultiFileAuthState, DisconnectReason } from '@whiskeysockets/baileys';
import express from 'express';
import pino from 'pino';
import fs from 'fs';
import QRCode from 'qrcode';
import path from 'path';

const app = express();
app.use(express.json()); 

// --- CORREÇÃO PARA O EXECUTÁVEL (.EXE) ---
// Quando roda em .exe, precisamos pegar a pasta onde o arquivo está
const pastaAtual = process.cwd(); 

// Variáveis Globais
let sock;
let qrCodeImage = null; 
let statusConexao = "DESCONECTADO"; 

async function connectToWhatsApp() {
    const caminhoAuth = path.join(pastaAtual, 'auth_info_baileys');

    // Garante que a pasta de autenticação existe
    if (!fs.existsSync(caminhoAuth)) {
        fs.mkdirSync(caminhoAuth);
    }

    const { state, saveCreds } = await useMultiFileAuthState(caminhoAuth);

    sock = makeWASocket({
        logger: pino({ level: 'silent' }),
        printQRInTerminal: false,
        auth: state,
    });

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;
        
        if (qr) {
            statusConexao = "AGUARDANDO_QR";
            try {
                qrCodeImage = await QRCode.toDataURL(qr);
                console.log('⚡ QR Code novo gerado! Acesse http://localhost:3000');
            } catch (err) {
                console.error('Erro ao gerar imagem do QR:', err);
            }
        }

        if (connection === 'close') {
            statusConexao = "DESCONECTADO";
            const shouldReconnect = lastDisconnect.error?.output?.statusCode !== DisconnectReason.loggedOut;
            console.log('🔴 Conexão fechada. Reconectando...', shouldReconnect);
            
            if (shouldReconnect) {
                connectToWhatsApp();
            } else {
                console.log('❌ Desconectado (Logout). Apague a pasta "auth_info_baileys".');
            }
        } else if (connection === 'open') {
            statusConexao = "CONECTADO";
            qrCodeImage = null;
            console.log('✅ Bot conectado e pronto!');
        }
    });
    
    // Log de grupos (Opcional)
    sock.ev.on('messages.upsert', async (m) => {
        if (m.type === 'notify') {
            const msg = m.messages[0];
            if (!msg.key.fromMe && msg.key.remoteJid.includes('@g.us')) {
                console.log(`📢 GRUPO ID: ${msg.key.remoteJid}`);
            }
        }
    });
}

// ROTAS
app.get('/status', (req, res) => {
    res.json({ status: statusConexao, temQrCode: qrCodeImage !== null });
});

app.get('/', (req, res) => {
    res.setHeader('Content-Type', 'text/html');
    if (statusConexao === "CONECTADO") {
        res.send('<h1>✅ Conectado! Pode fechar.</h1>');
    } else if (qrCodeImage) {
        res.send(`<h1>Escaneie:</h1><br><img src="${qrCodeImage}" />`);
    } else {
        res.send('<h2>🔄 Iniciando... atualize em instantes.</h2>');
    }
});

app.post('/enviar-escala', async (req, res) => {
    const { caminhoImagem, grupoId, legenda } = req.body;

    if (!sock || statusConexao !== "CONECTADO") {
        return res.status(500).json({ erro: 'Bot desconectado.' });
    }

    try {
        if (!fs.existsSync(caminhoImagem)) {
            return res.status(400).json({ erro: 'Imagem não encontrada no disco.' });
        }

        await sock.sendMessage(grupoId, { 
            image: { url: caminhoImagem }, 
            caption: legenda || "Escala atualizada"
        });

        res.json({ sucesso: true });
    } catch (err) {
        console.error(err);
        res.status(500).json({ erro: err.message });
    }
});

// Start
connectToWhatsApp().then(() => {
    app.listen(3000, () => console.log('🚀 Servidor rodando na porta 3000'));
});