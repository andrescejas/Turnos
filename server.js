const http = require('http');
const fs = require('fs');
const path = require('path');
const { runAuditEngine } = require('./auditEngine');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = path.join(__dirname, 'public');
const TURNOS_FILE = path.join(__dirname, 'Turnos.xlsx');

let cachedData = null;
let lastModifiedTime = 0;
let isCalculating = false;

function refreshData() {
    try {
        isCalculating = true;
        console.log(`[${new Date().toLocaleTimeString()}] Procesando Turnos.xlsx...`);
        delete require.cache[require.resolve('./auditEngine')];
        const { runAuditEngine } = require('./auditEngine');
        cachedData = runAuditEngine();
        lastModifiedTime = fs.existsSync(TURNOS_FILE) ? fs.statSync(TURNOS_FILE).mtimeMs : Date.now();
        console.log(`[${new Date().toLocaleTimeString()}] Datos actualizados con éxito. Total turnos: ${cachedData.totalTurnos}`);
    } catch (err) {
        console.error("Error al procesar auditoría:", err.message);
    } finally {
        isCalculating = false;
    }
}

// Initial load
refreshData();

// Watch Turnos.xlsx for changes
if (fs.existsSync(TURNOS_FILE)) {
    fs.watchFile(TURNOS_FILE, { interval: 2000 }, (curr, prev) => {
        if (curr.mtimeMs !== prev.mtimeMs) {
            console.log(`[${new Date().toLocaleTimeString()}] Detectado cambio en Turnos.xlsx. Recalculando...`);
            refreshData();
        }
    });
}

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.svg': 'image/svg+xml'
};

const server = http.createServer((req, res) => {
    // CORS headers for local dev
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }

    const url = new URL(req.url, `http://${req.headers.host}`);
    const pathname = url.pathname;

    // API endpoints
    if (pathname === '/api/data') {
        if (!cachedData || req.url.includes('force=true')) {
            refreshData();
        }
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ ok: true, data: cachedData, lastModifiedTime }));
        return;
    }

    if (pathname === '/api/refresh') {
        refreshData();
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ ok: true, message: 'Datos recalculados', timestamp: new Date().toISOString() }));
        return;
    }

    // Static files
    let filePath = path.join(PUBLIC_DIR, pathname === '/' ? 'index.html' : pathname);

    // Prevent directory traversal
    if (!filePath.startsWith(PUBLIC_DIR)) {
        res.writeHead(403, { 'Content-Type': 'text/plain' });
        res.end('Acceso denegado');
        return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    fs.readFile(filePath, (err, content) => {
        if (err) {
            if (err.code === 'ENOENT') {
                res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
                res.end(`<h1>404 No Encontrado</h1><p>El archivo ${pathname} no existe.</p>`);
            } else {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
                res.end(`Error de servidor: ${err.code}`);
            }
        } else {
            res.writeHead(200, { 'Content-Type': contentType });
            res.end(content);
        }
    });
});

server.listen(PORT, () => {
    console.log(`\n======================================================`);
    console.log(`🚀 Dashboard de Turnos & Capacidad Activo`);
    console.log(`📡 URL Local: http://localhost:${PORT}`);
    console.log(`👀 Monitoreando: Turnos.xlsx en tiempo real`);
    console.log(`======================================================\n`);
});
