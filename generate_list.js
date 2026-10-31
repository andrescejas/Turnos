const xlsx = require('xlsx');

const TURNOS_FILE = 'Turnos.xlsx';
const wbTurnos = xlsx.readFile(TURNOS_FILE);
const turnosData = xlsx.utils.sheet_to_json(wbTurnos.Sheets[wbTurnos.SheetNames[0]], {defval: ""});

const START_MONTH = 8; // August

const results = [];

function pad(n) { return n < 10 ? '0'+n : n; }
function formatTime(excelTime) {
    let totalSeconds = Math.round(excelTime * 86400);
    let hours = Math.floor(totalSeconds / 3600);
    let minutes = Math.floor((totalSeconds % 3600) / 60);
    return `${pad(hours)}:${pad(minutes)}`;
}

turnosData.forEach(row => {
    let rawDate = row['Fecha'];
    let d;
    if (typeof rawDate === 'number') {
        const parsed = xlsx.SSF.parse_date_code(rawDate);
        d = new Date(parsed.y, parsed.m - 1, parsed.d);
    } else {
        d = new Date(rawDate);
    }
    
    if (isNaN(d.getTime())) return;

    const month = d.getMonth() + 1; // 1-12
    const year = d.getFullYear();
    const pacienteRaw = row['Paciente'] ? row['Paciente'].toString() : '';
    const pacienteLower = pacienteRaw.toLowerCase();
    const asistio = (row['Asistio'] || '').toString().toLowerCase();

    // Check if cancelled
    const isCancelled = pacienteLower.includes('cancela') || asistio.includes('cancela');

    if (!isCancelled && year === 2026 && month >= START_MONTH && month <= 12) {
        // Infer modality
        let modalidad = 'Presencial';
        if (pacienteLower.match(/\b(v|virtual|video)\b/)) {
            modalidad = 'Virtual';
        } else if (pacienteLower.match(/\b(p|presencial)\b/)) {
            modalidad = 'Presencial';
        }

        // Clean name
        const cleanName = pacienteLower.replace(/\b(ctrol|ctrl|ctl|control|p|v|1 vez|vez|presencial|virtual|video)\b/g, ' ').replace(/-/g, ' ').replace(/\s+/g, ' ').trim();
        // Capitalize names
        const capitalizedName = cleanName.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');

        let timeStr = "";
        if (typeof row['Hora'] === 'number') {
            timeStr = formatTime(row['Hora']);
        } else {
            timeStr = row['Hora'].toString();
        }

        results.push({
            year, month, day: d.getDate(),
            dateObj: d,
            timeStr: timeStr,
            timeVal: typeof row['Hora'] === 'number' ? row['Hora'] : 0,
            name: capitalizedName,
            modalidad
        });
    }
});

// Sort by month, day, time
results.sort((a, b) => {
    if (a.month !== b.month) return a.month - b.month;
    if (a.day !== b.day) return a.day - b.day;
    return a.timeVal - b.timeVal;
});

// Group by month
const grouped = {};
results.forEach(r => {
    if (!grouped[r.month]) grouped[r.month] = [];
    grouped[r.month].push(r);
});

const monthNames = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

let markdown = `# Listado de Turnos (Agosto - Diciembre 2026)\n\n`;

Object.keys(grouped).sort((a, b) => a - b).forEach(m => {
    markdown += `## ${monthNames[m]} 2026\n`;
    markdown += `| Fecha | Hora | Paciente | Modalidad |\n`;
    markdown += `|---|---|---|---|\n`;
    grouped[m].forEach(t => {
        const dateStr = `${pad(t.day)}/${pad(t.month)}/${t.year}`;
        markdown += `| ${dateStr} | ${t.timeStr} | ${t.name} | ${t.modalidad} |\n`;
    });
    markdown += `\n`;
});

const fs = require('fs');
const artifactPath = 'C:\\Users\\Ceci\\.gemini\\antigravity-ide\\brain\\672d4eb4-f747-451f-99f6-f9dcfc808fb3\\listado_turnos.md';
fs.writeFileSync(artifactPath, markdown);
console.log('Listado escrito en ' + artifactPath);
