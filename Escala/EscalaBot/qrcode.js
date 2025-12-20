import express from 'express';
import QRCode from 'qrcode';

const app = express();
const port = process.env.PORT || 3000;

let qrCodeImage = null;

// Função chamada pelo index.js quando o WhatsApp manda um QR novo
export const mostrarQRCode = async (text) => {
    try {
        // Converte o texto do QR em uma imagem Base64 para exibir no navegador
        qrCodeImage = await QRCode.toDataURL(text);
        console.log('✅ QR Code atualizado! Acesse a URL do Render para escanear.');
    } catch (err) {
        console.error('Erro ao gerar imagem do QR Code:', err);
    }
};

// Rota principal (ao acessar https://seu-app.onrender.com)
app.get('/', (req, res) => {
    res.setHeader('Content-Type', 'text/html');

    if (qrCodeImage) {
        res.send(`
            <html>
                <head>
                    <title>WhatsApp Bot QR</title>
                    <meta http-equiv="refresh" content="10"> <style>
                        body { display: flex; justify-content: center; align-items: center; height: 100vh; background: #f0f2f5; font-family: sans-serif; flex-direction: column; }
                        h1 { color: #333; }
                        img { border: 10px solid white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
                    </style>
                </head>
                <body>
                    <h1>Escaneie para conectar</h1>
                    <img src="${qrCodeImage}" alt="QR Code" />
                    <p>Atualiza automaticamente a cada 10 segundos.</p>
                </body>
            </html>
        `);
    } else {
        res.send(`
            <html>
                <head>
                    <meta http-equiv="refresh" content="3">
                    <style>body { display: flex; justify-content: center; align-items: center; height: 100vh; font-family: sans-serif; }</style>
                </head>
                <body>
                    <h2>🔄 Aguardando geração do QR Code...</h2>
                    <p>Se o bot já estiver conectado, esta tela não mudará.</p>
                </body>
            </html>
        `);
    }
});

// Inicia o servidor Web (O Render exige isso para manter o app vivo)
app.listen(port, () => {
    console.log(`🌐 Servidor Web rodando na porta ${port}`);
});
