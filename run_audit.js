const xlsx = require('xlsx');
const stringSimilarity = require('string-similarity');

function normalizeStr(str) {
    if (!str) return '';
    return str.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function cleanApptStr(str) {
    let s = normalizeStr(str);
    const noise = ["ctrol", "ctrl", "ctr", "ctl", "control", "presencial", "virtual", "1 vez", "vez", " p ", " v ", " p", " v"];
    // replace exact words
    noise.forEach(n => {
        // use regex to replace whole words
        const regex = new RegExp('\\b' + n + '\\b', 'g');
        s = s.replace(regex, ' ');
    });
    // handle 'p' and 'v' at the end or start
    s = s.replace(/\bp\b/g, ' ').replace(/\bv\b/g, ' ');
    return s.replace(/\s+/g, ' ').trim();
}

function checkCancel(str) {
    if (!str) return false;
    const s = normalizeStr(str);
    return s.includes("cancela");
}

function wordMatch(targetWords, textWords) {
    for (const tw of targetWords) {
        if (tw.length < 2) continue;
        for (const word of textWords) {
            if (stringSimilarity.compareTwoStrings(tw, word) > 0.75) {
                return true;
            }
        }
    }
    return false;
}

function isPatientMatch(apptStr, paciente) {
    const cleaned = cleanApptStr(apptStr);
    const textWords = cleaned.split(/\s+/).filter(w => w.length > 1);
    
    const nombreWords = normalizeStr(paciente.Nombre).split(/\s+/).filter(w => w.length > 1);
    const apellidoWords = normalizeStr(paciente.Apellido).split(/\s+/).filter(w => w.length > 1);

    if (apellidoWords.length === 0 || nombreWords.length === 0) return false;

    const matchApellido = wordMatch(apellidoWords, textWords);
    const matchNombre = wordMatch(nombreWords, textWords);

    return matchApellido && matchNombre;
}

function runAudit() {
    const wb = xlsx.readFile('d:/MisProyectos/Turnos/Turnos.xlsx');
    const pacientesData = xlsx.utils.sheet_to_json(wb.Sheets['Pacientes']);
    const turnosData = xlsx.utils.sheet_to_json(wb.Sheets['Turnos']);

    const wbOld = xlsx.readFile('d:/MisProyectos/Turnos/Validacion_Controles_Actualizado.xlsx');
    const oldValidacion = xlsx.utils.sheet_to_json(wbOld.Sheets['Validacion']);

    // Build old validation map
    const oldValMap = {};
    oldValidacion.forEach(row => {
        if (row.Paciente) {
            oldValMap[normalizeStr(row.Paciente)] = {
                comentario: row.Comentario || '',
                seguimiento: row.Seguimiento || '',
                arranque: row['Arranque (Si es nuevo)'] || ''
            };
        }
    });

    // Parse Turnos
    const parsedTurnos = turnosData.map(t => {
        let dateObj = null;
        if (typeof t.Fecha === 'number') {
            const d = xlsx.SSF.parse_date_code(t.Fecha);
            dateObj = new Date(d.y, d.m - 1, d.d);
        }
        return {
            ...t,
            dateObj,
            isCancelled: checkCancel(t.Paciente) || checkCancel(t.Validacion) || checkCancel(t['Event ID']),
            is1vez: normalizeStr(t.Paciente).includes("1 vez") || normalizeStr(t.Validacion).includes("1 vez")
        };
    }).filter(t => t.dateObj && !t.isCancelled);

    // Identify active control patients
    const allPacientes = pacientesData.map(p => {
        const frecStr = normalizeStr(p.Frecuencia);
        const isA_Demanda = frecStr.includes('demanda');
        const isDosMeses = frecStr.includes('dos meses');
        return {
            ...p,
            isActivo: normalizeStr(p.Activo) === 'si',
            isControl: frecStr.includes('control') && !isA_Demanda && !isDosMeses,
            isFinalizado: frecStr.includes('finalizado'),
            is1vez: frecStr.includes('1 vez') || normalizeStr(p.Comentario).includes('1 vez')
        };
    });

    const activeControls = allPacientes.filter(p => p.isActivo && p.isControl && !p.isFinalizado);

    // Assign appointments to patients
    const apptsByPatient = new Map();
    const unmatchedAppts = [];

    parsedTurnos.forEach(t => {
        const apptStr = (t.Paciente || '') + ' ' + (t.Validacion || '');
        let matchedPatient = null;
        
        // Match against active controls first
        for (const p of activeControls) {
            if (isPatientMatch(apptStr, p)) {
                matchedPatient = p;
                break;
            }
        }

        // If not, match against all others
        if (!matchedPatient) {
            for (const p of allPacientes) {
                if (!p.isActivo || !p.isControl || p.isFinalizado) {
                    if (isPatientMatch(apptStr, p)) {
                        matchedPatient = p;
                        break;
                    }
                }
            }
        }

        if (matchedPatient) {
            if (!apptsByPatient.has(matchedPatient)) {
                apptsByPatient.set(matchedPatient, []);
            }
            apptsByPatient.get(matchedPatient).push(t);
            // propagate 1 vez flag
            if (t.is1vez) matchedPatient.is1vez = true;
        } else {
            // Only care about future unmatched for alert? Actually the prompt says "ponlo en Alerta"
            if (t.dateObj.getFullYear() === 2026 && t.dateObj.getMonth() >= 7 && t.dateObj.getMonth() <= 11) {
                unmatchedAppts.push(t);
            }
        }
    });

    const monthsToCheck = [
        {m: 7, name: 'Agosto'},
        {m: 8, name: 'Septiembre'},
        {m: 9, name: 'Octubre'},
        {m: 10, name: 'Noviembre'},
        {m: 11, name: 'Diciembre'}
    ];

    const results = [];
    const inactivosConTurnos = [];

    // Check inactive patients with future appointments (Aug-Dec 2026)
    allPacientes.forEach(p => {
        if (!p.isActivo) {
            const appts = apptsByPatient.get(p) || [];
            const futureAppts = appts.filter(t => t.dateObj.getFullYear() === 2026 && t.dateObj.getMonth() >= 7);
            if (futureAppts.length > 0) {
                inactivosConTurnos.push(`${p.Nombre} ${p.Apellido}`);
            }
        }
    });

    // Check active controls
    activeControls.forEach(p => {
        const patientName = `${p.Nombre} ${p.Apellido}`;
        const pNorm = normalizeStr(patientName);
        const oldData = oldValMap[pNorm] || {comentario: '', seguimiento: ''};

        // Perfect agenda check
        if (normalizeStr(oldData.comentario).includes('desestimar') || normalizeStr(oldData.seguimiento).includes('desestimar')) {
            return; // Perfect agenda
        }

        const appts = apptsByPatient.get(p) || [];
        
        // Find first appointment month in 2026
        let firstMonth2026 = -1;
        appts.forEach(t => {
            if (t.dateObj.getFullYear() === 2026) {
                if (firstMonth2026 === -1 || t.dateObj.getMonth() < firstMonth2026) {
                    firstMonth2026 = t.dateObj.getMonth();
                }
            }
        });

        // check if '1 vez'
        const is1Vez = p.is1vez || normalizeStr(oldData.comentario).includes('1 vez') || normalizeStr(oldData.arranque).includes('1 vez');

        const expectedMonths = [];
        monthsToCheck.forEach(m => {
            // If 1 vez, require only from their first appointment
            if (is1Vez && firstMonth2026 !== -1 && m.m < firstMonth2026) {
                return; // skip months before their first appointment
            }
            expectedMonths.push(m);
        });

        const missing = [];
        expectedMonths.forEach(m => {
            const hasAppt = appts.some(t => t.dateObj.getFullYear() === 2026 && t.dateObj.getMonth() === m.m);
            if (!hasAppt) {
                missing.push(m.name);
            }
        });

        if (missing.length > 0) {
            results.push({
                Paciente: patientName,
                Estado: 'Incompleto (Faltan Meses)',
                'Meses Faltantes': missing.join(', '),
                'Arranque (Si es nuevo)': is1Vez ? '1 vez' : '',
                Comentario: oldData.comentario,
                Seguimiento: oldData.seguimiento
            });
        }
    });

    // Save results to excel
    const newWb = xlsx.utils.book_new();
    const newWs = xlsx.utils.json_to_sheet(results.length > 0 ? results : [{Paciente: 'Ninguno con faltas'}]);
    xlsx.utils.book_append_sheet(newWb, newWs, 'Validacion');

    // Add Alertas sheets
    const uniqueUnmatched = [...new Set(unmatchedAppts.map(t => (t.Paciente || '') + ' | ' + (t.Validacion || '') + ' (' + t.dateObj.toLocaleDateString('es-AR') + ')'))];
    const wsAlertasApe = xlsx.utils.json_to_sheet(uniqueUnmatched.map(u => ({'Alerta: Apellidos No Coincidentes': u})));
    xlsx.utils.book_append_sheet(newWb, wsAlertasApe, 'Alerta_Apellidos');

    const uniqueInactivos = [...new Set(inactivosConTurnos)];
    const wsAlertasInac = xlsx.utils.json_to_sheet(uniqueInactivos.map(i => ({'Alerta: Inactivos con Turno': i})));
    xlsx.utils.book_append_sheet(newWb, wsAlertasInac, 'Alerta_Inactivos');

    xlsx.writeFile(newWb, 'd:/MisProyectos/Turnos/Validacion_Controles_Actualizado.xlsx');

    // Analysis of capacity (Aug to Dec 2026)
    const capByMonth = {7:0, 8:0, 9:0, 10:0, 11:0};
    const dayStats = {}; // month -> dayOfWeek -> count
    const cancelStats = {}; // month -> dayOfWeek -> count

    turnosData.forEach(t => {
        let dateObj = null;
        if (typeof t.Fecha === 'number') {
            const d = xlsx.SSF.parse_date_code(t.Fecha);
            dateObj = new Date(d.y, d.m - 1, d.d);
        }
        if (!dateObj || dateObj.getFullYear() !== 2026) return;
        const m = dateObj.getMonth();
        if (m >= 7 && m <= 11) {
            const dateStr = `${String(dateObj.getDate()).padStart(2,'0')}.${String(m+1).padStart(2,'0')}.${dateObj.getFullYear()}`;
            const days = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
            const dayName = days[dateObj.getDay()];
            const key = `${dayName} ${dateStr}`;

            const isCanc = checkCancel(t.Paciente) || checkCancel(t.Validacion) || checkCancel(t['Event ID']);
            if (!isCanc) {
                capByMonth[m]++;
                if (!dayStats[m]) dayStats[m] = {};
                dayStats[m][key] = (dayStats[m][key] || 0) + 1;
            } else {
                if (!cancelStats[m]) cancelStats[m] = {};
                cancelStats[m][key] = (cancelStats[m][key] || 0) + 1;
            }
        }
    });

    console.log("=== CAPACIDAD Y MARGEN (Tope 64) ===");
    monthsToCheck.forEach(m => {
        const occ = capByMonth[m.m];
        console.log(`${m.name} 2026: ${occ} ocupados, Margen para '1 vez': ${64 - occ}`);
    });

    console.log("\n=== CANTIDAD DADOS POR DIA ===");
    monthsToCheck.forEach(m => {
        console.log(`\n-- ${m.name} --`);
        if (dayStats[m.m]) {
            const keys = Object.keys(dayStats[m.m]).sort((a,b) => {
                const da = a.split(' ')[1].split('.');
                const db = b.split(' ')[1].split('.');
                return new Date(da[2], da[1]-1, da[0]) - new Date(db[2], db[1]-1, db[0]);
            });
            keys.forEach(k => {
                console.log(`- ${k} = ${dayStats[m.m][k]}`);
            });
        }
    });

    console.log("\n=== CANTIDAD CANCELADOS POR DIA ===");
    monthsToCheck.forEach(m => {
        console.log(`\n-- ${m.name} --`);
        if (cancelStats[m.m]) {
             const keys = Object.keys(cancelStats[m.m]).sort((a,b) => {
                const da = a.split(' ')[1].split('.');
                const db = b.split(' ')[1].split('.');
                return new Date(da[2], da[1]-1, da[0]) - new Date(db[2], db[1]-1, db[0]);
            });
            keys.forEach(k => {
                console.log(`- ${k} = ${cancelStats[m.m][k]}`);
            });
        }
    });

    console.log("\n=== ALERTA: APELLIDOS NO COINCIDENTES ===");
    uniqueUnmatched.forEach(u => console.log(u));

    console.log("\n=== ALERTA: PACIENTES INACTIVOS CON TURNOS ===");
    uniqueInactivos.forEach(i => console.log(i));
}

runAudit();
