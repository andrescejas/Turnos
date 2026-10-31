const xlsx = require('xlsx');

const PACIENTES_FILE = 'Pacientes.xlsx';
const TURNOS_FILE = 'Turnos.xlsx';

function normalize(str) {
    if (!str) return "";
    return str.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

const wbPacientes = xlsx.readFile(PACIENTES_FILE);
const pacData = xlsx.utils.sheet_to_json(wbPacientes.Sheets[wbPacientes.SheetNames[0]], {defval: ""});

const wbTurnos = xlsx.readFile(TURNOS_FILE);
const turnosData = xlsx.utils.sheet_to_json(wbTurnos.Sheets[wbTurnos.SheetNames[0]], {defval: ""});

// Parse dates properly
turnosData.forEach(row => {
    let rawDate = row['Fecha'];
    let d;
    if (typeof rawDate === 'number') {
        const parsed = xlsx.SSF.parse_date_code(rawDate);
        d = new Date(parsed.y, parsed.m - 1, parsed.d);
    } else {
        d = new Date(rawDate);
    }
    row._dateObj = d;
});

const now = new Date();
const currentMonth = 8; // August 2026
const currentYear = 2026;

const report = {};

// 1. Tasa de Cancelación (por mes)
const monthlyStatus = {};
turnosData.forEach(row => {
    if (isNaN(row._dateObj.getTime())) return;
    const m = row._dateObj.getMonth() + 1;
    const y = row._dateObj.getFullYear();
    const key = `${y}-${m.toString().padStart(2, '0')}`;
    
    if (!monthlyStatus[key]) monthlyStatus[key] = { total: 0, cancelados: 0 };
    
    const pacienteRaw = normalize(row['Paciente']);
    const asistio = normalize(row['Asistio']);
    const isCancelled = pacienteRaw.includes('cancela') || asistio.includes('cancela');
    
    monthlyStatus[key].total++;
    if (isCancelled) monthlyStatus[key].cancelados++;
});
report.cancelacion = monthlyStatus;

// 2. Ingresos del mes actual
const ingresos = {};
turnosData.forEach(row => {
    if (isNaN(row._dateObj.getTime())) return;
    const m = row._dateObj.getMonth() + 1;
    const y = row._dateObj.getFullYear();
    
    const pacienteRaw = normalize(row['Paciente']);
    const asistio = normalize(row['Asistio']);
    const isCancelled = pacienteRaw.includes('cancela') || asistio.includes('cancela');
    
    if (m === currentMonth && y === currentYear && !isCancelled) {
        const medio = (row['Medio de Pago'] || 'No especificado').toString().trim().toUpperCase();
        const importe = parseFloat(row['Importe']) || 0;
        if (!ingresos[medio]) ingresos[medio] = 0;
        ingresos[medio] += importe;
    }
});
report.ingresos = ingresos;

// 3. Ratio de Modalidad (mes actual)
let presencial = 0;
let virtual = 0;
turnosData.forEach(row => {
    if (isNaN(row._dateObj.getTime())) return;
    const m = row._dateObj.getMonth() + 1;
    const y = row._dateObj.getFullYear();
    
    const pacienteRaw = normalize(row['Paciente']);
    const asistio = normalize(row['Asistio']);
    const isCancelled = pacienteRaw.includes('cancela') || asistio.includes('cancela');
    
    if (m === currentMonth && y === currentYear && !isCancelled) {
        if (pacienteRaw.match(/\b(v|virtual|video)\b/)) {
            virtual++;
        } else {
            presencial++;
        }
    }
});
report.modalidad = { presencial, virtual };

// 4. Retención ("1 vez" a "Control")
// Definiremos retención: buscar pacientes cuya frecuencia es "1 vez", y ver si tienen > 1 turno agendado (pasado o futuro).
// Mejor aún, buscamos todos los pacientes, calculamos cuántos turnos tuvieron. Si la frecuencia actual es "control", ya se retuvieron. Si es "1 vez", tal vez no.
// Como Pacientes.xlsx tiene "Frecuencia", veamos cuántos "1 vez" hay en total.
let total1VezHistorico = 0;
let retenidos = 0;

// First, get all canonical patient names from pacData
const canonicalPatients = pacData.map(row => {
    return normalize(row['Nombre'] + " " + row['Apellido']);
});
const stringSimilarity = require('string-similarity');

// Analizamos todos los turnos para agrupar por paciente CANÓNICO
const turnosPorPaciente = {};
turnosData.forEach(row => {
    if (isNaN(row._dateObj.getTime())) return;
    const pacienteRaw = normalize(row['Paciente']);
    const asistio = normalize(row['Asistio']);
    const isCancelled = pacienteRaw.includes('cancela') || asistio.includes('cancela');
    
    if (!isCancelled) {
        const cleanName = pacienteRaw.replace(/\b(ctrol|ctrl|ctl|control|p|v|1 vez|vez|presencial|virtual|video)\b/g, ' ').replace(/-/g, ' ').replace(/\s+/g, ' ').trim();
        
        let bestKey = cleanName;
        if (canonicalPatients.length > 0) {
            const match = stringSimilarity.findBestMatch(cleanName, canonicalPatients);
            if (match.bestMatch.rating > 0.55) {
                bestKey = match.bestMatch.target;
            } else {
                // fallback first/last name
                const fallbackMatch = canonicalPatients.find(p => {
                    const parts = p.split(' ');
                    return parts.length >= 2 && cleanName.includes(parts[0]) && cleanName.includes(parts[parts.length-1]);
                });
                if (fallbackMatch) {
                    bestKey = fallbackMatch;
                }
            }
        }
        
        if (!turnosPorPaciente[bestKey]) turnosPorPaciente[bestKey] = [];
        turnosPorPaciente[bestKey].push(row._dateObj);
    }
});

// Chequear contra Pacientes.xlsx. Aquellos que dicen "1 vez" y tienen 1 solo turno = no volvieron.
// Aquellos que antes decían "1 vez" y ahora son "control", es más difícil saber quiénes eran. Asumamos que todos los activos que tienen >= 2 turnos son retenidos, y los que tienen 1 turno son no retenidos? No, solo importa medir la conversión reciente.
// Aproximación: contar cuántos pacientes en Pacientes.xlsx dicen "1 vez". Esos son los pendientes de retener.
// Para medir "retención", veamos pacientes que tuvieron su primer turno este año. ¿Cuántos tuvieron un segundo turno?
let nuevos2026 = 0;
let convertidos2026 = 0;
Object.keys(turnosPorPaciente).forEach(p => {
    const dates = turnosPorPaciente[p].sort((a,b) => a-b);
    if (dates[0].getFullYear() === 2026) {
        nuevos2026++;
        if (dates.length > 1) {
            convertidos2026++;
        }
    }
});
report.retencion = { nuevos2026, convertidos2026 };

// 5. Mapa de Calor Horario (Top 3 y Bottom 3)
const horasCount = {};
turnosData.forEach(row => {
    if (isNaN(row._dateObj.getTime())) return;
    const pacienteRaw = normalize(row['Paciente']);
    const asistio = normalize(row['Asistio']);
    const isCancelled = pacienteRaw.includes('cancela') || asistio.includes('cancela');
    
    if (!isCancelled && row._dateObj.getFullYear() === 2026) { // Analyze current year
        let timeStr = "";
        if (typeof row['Hora'] === 'number') {
            let totalSeconds = Math.round(row['Hora'] * 86400);
            let hours = Math.floor(totalSeconds / 3600);
            timeStr = `${hours < 10 ? '0'+hours : hours}:00`; // round to hour for heatmap
        } else {
            timeStr = row['Hora'].toString().substring(0, 2) + ":00";
        }
        horasCount[timeStr] = (horasCount[timeStr] || 0) + 1;
    }
});
const sortedHoras = Object.entries(horasCount).sort((a,b) => b[1] - a[1]);
report.calor = {
    top3: sortedHoras.slice(0, 3),
    bottom3: sortedHoras.slice(-3) // least requested among those that have at least 1 turn
};

// 6. Frecuencia y Riesgo (>40 días sin asistir y sin turno futuro)
const pacientesEnRiesgo = [];
pacData.forEach(row => {
    const activo = normalize(row['Activo']);
    const frecuencia = normalize(row['Frecuencia']);
    if (activo === 'si' && frecuencia.includes('control') && !frecuencia.includes('demanda') && !frecuencia.includes('dos meses')) {
        const fullName = normalize(row['Nombre'] + " " + row['Apellido']);
        
        // Find directly in turnosPorPaciente using the canonical name
        if (!turnosPorPaciente[fullName]) return;
        
        const dates = turnosPorPaciente[fullName].sort((a,b) => b-a); // Descending
        const lastDate = dates[0];
        const futureDates = dates.filter(d => d >= now);
        
        if (futureDates.length === 0 && lastDate < now) {
            const diffTime = Math.abs(now - lastDate);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            if (diffDays > 40) {
                pacientesEnRiesgo.push({ name: row['Nombre'] + " " + row['Apellido'], days: diffDays });
            }
        }
    }
});
report.riesgo = pacientesEnRiesgo.sort((a,b) => b.days - a.days);

console.log(JSON.stringify(report));
