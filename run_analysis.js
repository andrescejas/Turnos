const xlsx = require('xlsx');
const stringSimilarity = require('string-similarity');
const fs = require('fs');

// Constants
const TARGET_MONTHS = [8, 9, 10, 11, 12]; // August to December
const TARGET_YEAR = 2026;
const MAX_TURNS_PER_MONTH = 64;

const NOISE_WORDS = ["ctrol", "ctrl", "ctr", "ctl", "control", "p", "v", "1 vez", "vez", "presencial", "virtual"];

function cleanText(text) {
    if (!text) return "";
    let cleaned = text.toString().toLowerCase();
    
    // Remove punctuation
    cleaned = cleaned.replace(/[.,-]/g, ' ');

    NOISE_WORDS.forEach(word => {
        const regex = new RegExp(`\\b${word}\\b`, 'gi');
        cleaned = cleaned.replace(regex, ' ');
    });
    
    // Removing extra spaces
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    return cleaned;
}

function normalize(text) {
    if (!text) return "";
    return text.toString().toLowerCase().trim();
}

function isDateInTarget(dateObj) {
    if (!dateObj) return false;
    return dateObj.y === TARGET_YEAR && TARGET_MONTHS.includes(dateObj.m); // xlsx.SSF.parse_date_code returns 1-12 for m
}

function parseExcelDate(excelDate) {
    if (!excelDate) return null;
    if (typeof excelDate === 'number') {
        return xlsx.SSF.parse_date_code(excelDate);
    }
    // Handle string dates if any, assume DD/MM/YYYY
    const parts = excelDate.toString().split('/');
    if (parts.length === 3) {
        return {
            y: parseInt(parts[2]),
            m: parseInt(parts[1]),
            d: parseInt(parts[0])
        };
    }
    return null;
}

function matchPatient(turnoText, patient, isCompoundSurname) {
    const turnoClean = cleanText(turnoText);
    const turnoTokens = turnoClean.split(' ');
    
    const pNombre = normalize(patient.Nombre);
    const pApellido = normalize(patient.Apellido);
    
    if (!pNombre && !pApellido) return false;

    const apellidos = isCompoundSurname ? pApellido.split(' ') : [pApellido];
    const nombres = pNombre.split(' ');

    // Must match at least one part of surname and one part of name with similarity > 0.75
    let matchedApellido = false;
    for (const ap of apellidos) {
        if (!ap) continue;
        const match = stringSimilarity.findBestMatch(ap, turnoTokens);
        if (match.bestMatch && match.bestMatch.rating > 0.75) {
            matchedApellido = true;
            break;
        }
    }

    let matchedNombre = false;
    for (const nom of nombres) {
        if (!nom) continue;
        const match = stringSimilarity.findBestMatch(nom, turnoTokens);
        if (match.bestMatch && match.bestMatch.rating > 0.75) {
            matchedNombre = true;
            break;
        }
    }

    return matchedApellido && matchedNombre;
}

function runAnalysis() {
    console.log("Loading files...");
    const wbTurnos = xlsx.readFile('Turnos.xlsx');
    const sheetTurnos = xlsx.utils.sheet_to_json(wbTurnos.Sheets['Turnos']);
    const sheetPacientes = xlsx.utils.sheet_to_json(wbTurnos.Sheets['Pacientes']);

    let oldValidacion = [];
    try {
        const wbOld = xlsx.readFile('Validacion_Controles_Actualizado.xlsx');
        if (wbOld.Sheets['Validacion']) {
            oldValidacion = xlsx.utils.sheet_to_json(wbOld.Sheets['Validacion']);
        }
    } catch (e) {
        console.log("No old validation file found or readable, starting fresh.");
    }

    const oldDataMap = {};
    oldValidacion.forEach(row => {
        if (row.Paciente) {
            oldDataMap[row.Paciente.toLowerCase()] = {
                Comentario: row.Comentario || '',
                Seguimiento: row.Seguimiento || ''
            };
        }
    });

    console.log("Processing patients...");
    const pacientes = sheetPacientes.map((p, idx) => {
        return {
            ...p,
            _id: idx,
            Activo: normalize(p.Activo),
            Frecuencia: normalize(p.Frecuencia),
            Nombre: p.Nombre,
            Apellido: p.Apellido,
            _nombreCompleto: `${p.Apellido} ${p.Nombre}`.trim(),
            turnos: []
        };
    });

    const activeControls = pacientes.filter(p => p.Activo === 'si' && p.Frecuencia === 'control');
    const allOtherPacientes = pacientes.filter(p => !(p.Activo === 'si' && p.Frecuencia === 'control'));

    // Turns processing
    console.log("Processing turns...");
    const validTurnos = [];
    let unmatchedTurnos = [];

    sheetTurnos.forEach(row => {
        if (!row.Paciente) return;
        const tText = row.Paciente.toString().toLowerCase();
        
        if (tText.includes("cancelado") || tText.includes("cancela")) return;

        const date = parseExcelDate(row.Fecha);
        if (!date) return;

        const turnoObj = {
            rawText: row.Paciente,
            date: date,
            dateString: `${date.d.toString().padStart(2, '0')}/${date.m.toString().padStart(2, '0')}/${date.y}`,
            isTarget: isDateInTarget(date)
        };

        validTurnos.push(turnoObj);

        // Find matching patient
        let matched = false;
        
        // Check active controls first
        for (const p of activeControls) {
            const isComp = (p.Apellido || '').split(' ').length > 1;
            if (matchPatient(row.Paciente, p, isComp)) {
                p.turnos.push(turnoObj);
                matched = true;
                break; // stop searching
            }
        }

        if (!matched) {
            // Check other patients
            for (const p of allOtherPacientes) {
                const isComp = (p.Apellido || '').split(' ').length > 1;
                if (matchPatient(row.Paciente, p, isComp)) {
                    p.turnos.push(turnoObj);
                    matched = true;
                    break;
                }
            }
        }

        if (!matched && turnoObj.isTarget) {
            unmatchedTurnos.push(turnoObj);
        }
    });

    // Auditoria Inconsistencias (Pacientes de Control)
    console.log("Auditing inconsistences...");
    const inconsistencias = [];
    
    activeControls.forEach(p => {
        const oldData = oldDataMap[p._nombreCompleto.toLowerCase()] || {};
        const comentario = oldData.Comentario || '';
        const seguimiento = oldData.Seguimiento || '';
        
        if (comentario.toLowerCase().includes('desestimar') || seguimiento.toLowerCase().includes('desestimar')) {
            // Perfect attendance assumed
            return; 
        }

        const is1Vez = p._nombreCompleto.toLowerCase().includes('1 vez'); // Or check in turns? Let's check turns for '1 vez' or start of year
        let firstTurnMonth = 8; // Default start check from Aug
        
        // find first turn this year
        let firstTurnThisYear = p.turnos.filter(t => t.date.y === TARGET_YEAR).sort((a,b) => (a.date.m - b.date.m) || (a.date.d - b.date.d))[0];
        if (firstTurnThisYear && firstTurnThisYear.date.m > 8) {
            firstTurnMonth = firstTurnThisYear.date.m;
        }

        const monthsPresent = new Set(p.turnos.filter(t => isDateInTarget(t.date)).map(t => t.date.m));
        
        const requiredMonths = TARGET_MONTHS.filter(m => m >= firstTurnMonth);
        const missingMonths = requiredMonths.filter(m => !monthsPresent.has(m));

        if (missingMonths.length > 0) {
            inconsistencias.push({
                Paciente: p._nombreCompleto,
                Estado: 'Activo - Control',
                'Meses Faltantes': missingMonths.join(', '),
                'Arranque (Si es nuevo)': firstTurnThisYear ? `${firstTurnThisYear.date.m}/${firstTurnThisYear.date.y}` : '',
                Comentario: comentario,
                Seguimiento: seguimiento
            });
        }
    });

    // Alerta de Inactivos
    const inactivosConFuturos = [];
    allOtherPacientes.filter(p => p.Activo === 'no' || !p.Activo).forEach(p => {
        const futureTurns = p.turnos.filter(t => t.date.y > TARGET_YEAR || (t.date.y === TARGET_YEAR && t.date.m >= 8));
        if (futureTurns.length > 0) {
            inactivosConFuturos.push({
                Paciente: p._nombreCompleto,
                'Turnos Futuros': futureTurns.map(t => t.dateString).join(', ')
            });
        }
    });

    // Unmatched Turnos Alert
    const alertaApellidos = unmatchedTurnos.map(t => ({
        Turno: t.rawText,
        Fecha: t.dateString
    }));

    // Output validacion
    const newWb = xlsx.utils.book_new();
    const wsValidacion = xlsx.utils.json_to_sheet(inconsistencias.length ? inconsistencias : [{Paciente: 'No hay pacientes con faltas'}]);
    const wsInactivos = xlsx.utils.json_to_sheet(inactivosConFuturos.length ? inactivosConFuturos : [{Paciente: 'No hay inactivos con turnos'}]);
    const wsNoMatch = xlsx.utils.json_to_sheet(alertaApellidos.length ? alertaApellidos : [{Turno: 'No hay turnos huerfanos'}]);

    xlsx.utils.book_append_sheet(newWb, wsValidacion, 'Validacion');
    xlsx.utils.book_append_sheet(newWb, wsNoMatch, 'Alerta_Apellidos');
    xlsx.utils.book_append_sheet(newWb, wsInactivos, 'Alerta_Inactivos');

    xlsx.writeFile(newWb, 'Validacion_Controles_Actualizado.xlsx');

    // Capacity analysis
    console.log("Capacity Analysis:");
    const turnsByMonth = {8:0, 9:0, 10:0, 11:0, 12:0};
    const turnsByDayAndMonth = {};

    validTurnos.forEach(t => {
        if (isDateInTarget(t.date)) {
            turnsByMonth[t.date.m]++;
            
            // For day of week, construct Date
            // Date takes month 0-indexed
            const jsDate = new Date(t.date.y, t.date.m - 1, t.date.d);
            const days = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
            const dayName = days[jsDate.getDay()];
            
            const dayKey = `${dayName} ${t.dateString}`;
            if (!turnsByDayAndMonth[t.date.m]) turnsByDayAndMonth[t.date.m] = {};
            if (!turnsByDayAndMonth[t.date.m][dayKey]) turnsByDayAndMonth[t.date.m][dayKey] = 0;
            turnsByDayAndMonth[t.date.m][dayKey]++;
        }
    });

    const report = {
        capacity: turnsByMonth,
        margins: Object.keys(turnsByMonth).reduce((acc, m) => {
            acc[m] = MAX_TURNS_PER_MONTH - turnsByMonth[m];
            return acc;
        }, {}),
        turnsByDay: turnsByDayAndMonth,
        unmatched: alertaApellidos
    };

    fs.writeFileSync('report.json', JSON.stringify(report, null, 2));
    console.log("Done.");
}

runAnalysis();
