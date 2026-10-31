const xlsx = require('xlsx');
const stringSimilarity = require('string-similarity');

const TURNOS_FILE = 'Turnos.xlsx';
const VALIDACION_FILE = 'Validacion_Controles_Actualizado.xlsx';

function normalize(str) {
    if (!str) return "";
    return str.toString().toLowerCase().trim()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// 1. Read Pacientes
const wb = xlsx.readFile(TURNOS_FILE);
const pacData = xlsx.utils.sheet_to_json(wb.Sheets['Pacientes'], {defval: ""});

const patientsToControl = [];
const allPatients = [];

pacData.forEach(row => {
    const activo = normalize(row['Activo']);
    const frecuencia = normalize(row['Frecuencia']);
    
    let fullName = normalize(row['Nombre'] + " " + row['Apellido']);
    const patientObj = {
        name: fullName,
        firstName: normalize(row['Nombre']),
        lastName: normalize(row['Apellido']),
        originalName: (row['Nombre'] + " " + row['Apellido']).trim(),
        is1vez: frecuencia.includes('1 vez'),
        hasFutureTurnos: false
    };
    
    allPatients.push(patientObj);
    
    if (activo === 'si') {
        if (frecuencia.includes('control') && !frecuencia.includes('demanda') && !frecuencia.includes('dos meses') && !frecuencia.includes('finalizado')) {
            patientsToControl.push(patientObj);
        }
    }
});

// Read inactive for alert
const inactivePatients = pacData.filter(row => {
    const activo = normalize(row['Activo']);
    return activo !== 'si';
}).map(row => normalize(row['Nombre'] + " " + row['Apellido']));


// 2. Read Old Validacion
let oldComments = {};
try {
    const wbVal = xlsx.readFile(VALIDACION_FILE);
    const valData = xlsx.utils.sheet_to_json(wbVal.Sheets[wbVal.SheetNames[0]], {defval: ""});
    valData.forEach(row => {
        const name = normalize(row['Paciente']);
        oldComments[name] = {
            comentario: row['Comentario'] || "",
            seguimiento: row['Seguimiento'] || ""
        };
    });
} catch(e) {
    console.log("No previous validacion found or error reading it.");
}

// 3. Read Turnos
const turnosData = xlsx.utils.sheet_to_json(wb.Sheets['Turnos'], {defval: ""});

const monthlyCounts = {};
const dailyCounts = {};
const dailyCancelledCounts = {};
const patientAssistance = {};
const unmatchedAlerts = new Set();

const START_MONTH = 8; // August

turnosData.forEach(row => {
    let rawDate = row['Fecha'];
    let d;
    if (typeof rawDate === 'number') {
        const parsed = xlsx.SSF.parse_date_code(rawDate);
        d = new Date(parsed.y, parsed.m - 1, parsed.d);
    } else {
        d = new Date(rawDate);
    }
    
    if (isNaN(d.getTime())) return; // skip invalid

    const month = d.getMonth() + 1; // 1-12
    const year = d.getFullYear();
    const pacienteRaw = normalize(row['Paciente']);
    const asistio = normalize(row['Asistio']);

    // Check if cancelled
    const isCancelled = pacienteRaw.includes('cancela') || asistio.includes('cancela');

    // Count capacity (all non-cancelled)
    const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
    const dayKey = `${year}-${month.toString().padStart(2, '0')}-${d.getDate().toString().padStart(2, '0')}`;

    if (!isCancelled) {
        monthlyCounts[monthKey] = (monthlyCounts[monthKey] || 0) + 1;
        
        if (year === 2026 && month >= 8 && month <= 12) {
            dailyCounts[dayKey] = (dailyCounts[dayKey] || 0) + 1;
        }
    } else {
        if (year === 2026 && month >= 8 && month <= 12) {
            dailyCancelledCounts[dayKey] = (dailyCancelledCounts[dayKey] || 0) + 1;
        }
    }

    // Check inactive alert
    if (!isCancelled && d >= new Date()) { // future turno
        const match = stringSimilarity.findBestMatch(pacienteRaw, inactivePatients);
        if (match.bestMatch.rating > 0.8) {
            // we will handle this below
        }
    }

    // Record assistance for controls
    if (!isCancelled && month >= START_MONTH && year === 2026) { // evaluate from August 2026
        const monthStr = month.toString();
        
        // Find matching patient in patientsToControl
        if (patientsToControl.length > 0) {
            const cleanPacienteRaw = pacienteRaw.replace(/\b(ctrol|ctrl|ctr|ctl|control|p|v|1 vez|vez|presencial|virtual)\b/g, ' ').replace(/-/g, ' ').replace(/\s+/g, ' ').trim();
            if (!cleanPacienteRaw) return;
            
            let candidates = patientsToControl.filter(p => {
                const mainLastName = p.lastName.split(' ')[0];
                const rawWords = cleanPacienteRaw.split(' ');
                // Check if any word in the agenda text is highly similar to the main last name
                return rawWords.some(w => stringSimilarity.compareTwoStrings(w, mainLastName) > 0.75);
            });
            let bestName = null;
            let matched = false;

            if (candidates.length > 0) {
                const candidateNames = candidates.map(c => c.name);
                const match = stringSimilarity.findBestMatch(cleanPacienteRaw, candidateNames);
                if (match.bestMatch.rating > 0.45) { // Lower threshold since last name already matched
                    bestName = match.bestMatch.target;
                    matched = true;
                } else {
                    const fallbackMatch = candidates.find(c => cleanPacienteRaw.includes(c.firstName));
                    if (fallbackMatch) {
                        bestName = fallbackMatch.name;
                        matched = true;
                    }
                }
            }

            if (matched) {
                if (!patientAssistance[bestName]) {
                    patientAssistance[bestName] = new Set();
                }
                patientAssistance[bestName].add(month);
            } else {
                // Check if it matches any non-control patient in the DB before throwing an alert
                let candidatesAll = allPatients.filter(p => {
                    const mainLastName = p.lastName.split(' ')[0];
                    if (!mainLastName) return false;
                    const rawWords = cleanPacienteRaw.split(' ');
                    return rawWords.some(w => stringSimilarity.compareTwoStrings(w, mainLastName) > 0.75);
                });
                
                let matchedAny = false;
                if (candidatesAll.length > 0) {
                    const candidateNamesAll = candidatesAll.map(c => c.name);
                    const matchAll = stringSimilarity.findBestMatch(cleanPacienteRaw, candidateNamesAll);
                    if (matchAll.bestMatch.rating > 0.45) {
                        matchedAny = true;
                    }
                }
                
                if (!matchedAny) {
                    unmatchedAlerts.add(cleanPacienteRaw);
                }
            }
        }
    }
});

// Capacity Output
let capacityMsg = "Análisis de Capacidad:\n";
Object.keys(monthlyCounts).sort().forEach(m => {
    const used = monthlyCounts[m];
    const margin = 64 - used;
    capacityMsg += `- Mes ${m}: ${used} turnos ocupados. Margen: ${margin} lugares libres.\n`;
});

capacityMsg += "\nAnálisis Diario (Días más cargados - Ordenados por fecha):\n";
const weekdays = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
Object.keys(dailyCounts).sort().forEach(day => {
    const parts = day.split('-');
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const d = parseInt(parts[2], 10);
    const dateObj = new Date(y, m, d);
    const weekdayStr = weekdays[dateObj.getDay()];
    const formattedDate = `${parts[2]}.${parts[1]}.${parts[0]}`;
    capacityMsg += `- ${weekdayStr} ${formattedDate} = ${dailyCounts[day]}\n`;
});

capacityMsg += "\nTurnos Cancelados por Día (Agosto - Diciembre 2026):\n";
Object.keys(dailyCancelledCounts).sort().forEach(day => {
    const parts = day.split('-');
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const d = parseInt(parts[2], 10);
    const dateObj = new Date(y, m, d);
    const weekdayStr = weekdays[dateObj.getDay()];
    const formattedDate = `${parts[2]}.${parts[1]}.${parts[0]}`;
    capacityMsg += `- ${weekdayStr} ${formattedDate} = ${dailyCancelledCounts[day]} cancelados\n`;
});
if (Object.keys(dailyCancelledCounts).length === 0) {
    capacityMsg += "- Ninguno\n";
}

// Controls Audit
const newValidacion = [];
const missingAlerts = [];

patientsToControl.forEach(p => {
    const comm = oldComments[p.name] || {comentario: "", seguimiento: ""};
    
    // Ignore desestimar
    if (normalize(comm.comentario).includes('desestimar') || normalize(comm.seguimiento).includes('desestimar')) {
        return; // perfect agenda assumed
    }

    const attendedMonths = patientAssistance[p.name] || new Set();
    
    let expectedMonths = [8, 9, 10, 11, 12];
    if (p.is1vez) {
        // If "1 vez" (just started recently), only demand from the month they first attended, or something similar
        // Without more context, let's say from the earliest month they attended or just mark it as missing only after they started.
        if (attendedMonths.size > 0) {
            const first = Math.min(...Array.from(attendedMonths));
            expectedMonths = expectedMonths.filter(m => m >= first);
        } else {
            // no turnos yet
            expectedMonths = [8, 9, 10, 11, 12];
        }
    }

    const missing = expectedMonths.filter(m => !attendedMonths.has(m));
    const monthNames = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

    if (missing.length > 0 && missing.length < expectedMonths.length) {
        // If they attended at least once but missing some
        const missingStr = missing.map(m => monthNames[m]).join(', ');
        newValidacion.push({
            'Paciente': p.originalName,
            'Estado': 'Incompleto (Faltan Meses)',
            'Meses Faltantes': missingStr,
            'Arranque (Si es nuevo)': p.is1vez ? "1 vez" : "",
            'Comentario': comm.comentario,
            'Seguimiento': comm.seguimiento
        });
        missingAlerts.push(`${p.originalName} faltan meses: ${missingStr}`);
    } else if (missing.length === expectedMonths.length) {
        newValidacion.push({
            'Paciente': p.originalName,
            'Estado': 'Sin Asistencia',
            'Meses Faltantes': 'Todos desde Agosto',
            'Arranque (Si es nuevo)': p.is1vez ? "1 vez" : "",
            'Comentario': comm.comentario,
            'Seguimiento': comm.seguimiento
        });
    }
});

// Inactive alert check
const inactivesWithFuture = new Set();
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
    const pacienteRaw = normalize(row['Paciente']);
    const isCancelled = pacienteRaw.includes('cancela') || normalize(row['Asistio']).includes('cancela');
    
    if (!isCancelled && d >= new Date()) { // future
        const match = stringSimilarity.findBestMatch(pacienteRaw, inactivePatients);
        if (match.bestMatch.rating > 0.8) {
            inactivesWithFuture.add(match.bestMatch.target);
        }
    }
});

// Write to new Validacion
const newWb = xlsx.utils.book_new();
const newWs = xlsx.utils.json_to_sheet(newValidacion);
xlsx.utils.book_append_sheet(newWb, newWs, "Validacion");
console.log(capacityMsg);
console.log("----");
if (inactivesWithFuture.size > 0) {
    console.log("Alerta Inactivos (con turnos futuros):", Array.from(inactivesWithFuture).join(', '));
} else {
    console.log("Alerta Inactivos: Ninguno.");
}

console.log("----");
if (unmatchedAlerts.size > 0) {
    console.log("ALERTA APELLIDOS NO COINCIDENTES:");
    Array.from(unmatchedAlerts).forEach(name => console.log(`- ${name}`));
}
try {
    xlsx.writeFile(newWb, VALIDACION_FILE);
    console.log("Validación generada.");
} catch (e) {
    if (e.code === 'EBUSY') {
        console.log("ERROR: El archivo Validacion_Controles_Actualizado.xlsx está abierto. Por favor ciérralo e intenta de nuevo.");
    } else {
        console.log("Error al guardar archivo:", e.message);
    }
}
