const xlsx = require('xlsx');
const stringSimilarity = require('string-similarity');
const fs = require('fs');

const TURNOS_PATH = 'd:/MisProyectos/Turnos/Turnos.xlsx';
const OUTPUT_PATH = 'd:/MisProyectos/Turnos/Validacion_Controles_Actualizado.xlsx';
const CECILIA_VAL_PATH = 'd:/MisProyectos/Turnos/Validacion_Controles_ComentarioCecilia.xlsx';
const CAPACIDAD_MAXIMA = 64;

function normalizeStr(str) {
    if (!str) return '';
    return str.toString()
        .toLowerCase()
        .trim()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
}

function cleanApptText(str) {
    let s = normalizeStr(str);
    s = s.replace(/[^a-z0-9 ]/g, ' ');
    const noise = ['ctrol', 'ctrl', 'ctr', 'ctl', 'control', '1 vez', 'vez', 'presencial', 'virtual', 'p', 'v'];
    noise.forEach(n => {
        const regex = new RegExp('\\b' + n + '\\b', 'gi');
        s = s.replace(regex, ' ');
    });
    return s.replace(/\s+/g, ' ').trim();
}

function isCancelled(t) {
    const p = normalizeStr(t.Paciente);
    const v = normalizeStr(t.Validacion);
    const e = normalizeStr(t['Event ID']);
    return p.includes('cancela') || v.includes('cancela') || e.includes('cancela');
}

function detectModality(t) {
    const combined = normalizeStr(`${t.Paciente} ${t.Validacion}`);
    if (/\b(presencial|p)\b/.test(combined) || combined.includes('- presencial') || combined.endsWith(' p')) {
        return 'presencial';
    }
    if (/\b(virtual|v)\b/.test(combined) || combined.includes('- virtual') || combined.endsWith(' v')) {
        return 'virtual';
    }
    return 'sin_definir';
}

function parseDate(rawDate) {
    if (typeof rawDate === 'number') {
        const d = xlsx.SSF.parse_date_code(rawDate);
        return { y: d.y, m: d.m, d: d.d };
    }
    if (!rawDate) return null;
    const str = rawDate.toString().trim();
    const parts = str.split(/[\/\-\.]/);
    if (parts.length === 3) {
        if (parts[2].length === 4) {
            return { y: parseInt(parts[2], 10), m: parseInt(parts[1], 10), d: parseInt(parts[0], 10) };
        } else if (parts[0].length === 4) {
            return { y: parseInt(parts[0], 10), m: parseInt(parts[1], 10), d: parseInt(parts[2], 10) };
        }
    }
    return null;
}

function runAudit() {
    console.log("Cargando archivos...");
    const wb = xlsx.readFile(TURNOS_PATH);
    const sheetPacientes = xlsx.utils.sheet_to_json(wb.Sheets['Pacientes'], { defval: '' });
    const sheetTurnos = xlsx.utils.sheet_to_json(wb.Sheets['Turnos'], { defval: '' });

    // 1. Rescatar comentarios existentes
    const oldCommentsMap = {};

    function loadOldComments(filePath, sheetName, pacCol, comCol, segCol) {
        try {
            if (!fs.existsSync(filePath)) return;
            const oldWb = xlsx.readFile(filePath);
            const targetSheet = sheetName || oldWb.SheetNames[0];
            if (!oldWb.Sheets[targetSheet]) return;
            const rows = xlsx.utils.sheet_to_json(oldWb.Sheets[targetSheet], { defval: '' });
            rows.forEach(r => {
                const pac = normalizeStr(r[pacCol]);
                if (pac) {
                    if (!oldCommentsMap[pac]) {
                        oldCommentsMap[pac] = {
                            comentario: r[comCol] || '',
                            seguimiento: r[segCol] || ''
                        };
                    } else {
                        if (r[comCol] && !oldCommentsMap[pac].comentario) oldCommentsMap[pac].comentario = r[comCol];
                        if (r[segCol] && !oldCommentsMap[pac].seguimiento) oldCommentsMap[pac].seguimiento = r[segCol];
                    }
                }
            });
        } catch (e) {
            console.warn(`Aviso comentarios: ${e.message}`);
        }
    }

    loadOldComments(OUTPUT_PATH, 'Validacion', 'Paciente', 'Comentario', 'Seguimiento');
    loadOldComments(CECILIA_VAL_PATH, 'Validacion Controles', 'Paciente', 'Comentario', '__EMPTY');

    // 2. Base de pacientes
    const allPatients = sheetPacientes.map((p, idx) => {
        const nomWords = normalizeStr(p.Nombre).split(/\s+/).filter(w => w.length >= 3);
        const apeWords = normalizeStr(p.Apellido).split(/\s+/).filter(w => w.length >= 3);
        const activoNorm = normalizeStr(p.Activo);
        const frecNorm = normalizeStr(p.Frecuencia);

        const isActivo = activoNorm === 'si';
        const isDemanda = frecNorm.includes('demanda');
        const isDosMeses = frecNorm.includes('dos meses');
        const isFinalizado = frecNorm.includes('finalizado');
        const isControl = frecNorm.includes('control') && !isDemanda && !isDosMeses && !isFinalizado;

        return {
            id: idx,
            originalNombre: String(p.Nombre).trim(),
            originalApellido: String(p.Apellido).trim(),
            fullName: `${String(p.Apellido).trim()} ${String(p.Nombre).trim()}`.trim(),
            nombreNorm: normalizeStr(p.Nombre),
            apellidoNorm: normalizeStr(p.Apellido),
            nomWords,
            apeWords,
            isActivo,
            isControl,
            isFinalizado,
            turnos2026: []
        };
    });

    const controlPatients = allPatients.filter(p => p.isActivo && p.isControl);
    const otherPatients = allPatients.filter(p => !(p.isActivo && p.isControl));

    function matchPatientInList(tokens, patientList) {
        for (const p of patientList) {
            if (p.apeWords.length === 0 || p.nomWords.length === 0) continue;

            let matchApe = false;
            for (const ap of p.apeWords) {
                for (const t of tokens) {
                    if (stringSimilarity.compareTwoStrings(ap, t) > 0.75) {
                        matchApe = true;
                        break;
                    }
                }
                if (matchApe) break;
            }
            if (!matchApe) continue;

            let matchNom = false;
            for (const nom of p.nomWords) {
                for (const t of tokens) {
                    if (stringSimilarity.compareTwoStrings(nom, t) > 0.75) {
                        matchNom = true;
                        break;
                    }
                }
                if (matchNom) break;
            }

            if (matchNom) return p;
        }
        return null;
    }

    const targetMonths = [8, 9, 10, 11, 12];
    const targetYear = 2026;

    const monthlyCounts = { 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };
    const monthlyModality = {
        8: { presencial: 0, virtual: 0, sin_definir: 0 },
        9: { presencial: 0, virtual: 0, sin_definir: 0 },
        10: { presencial: 0, virtual: 0, sin_definir: 0 },
        11: { presencial: 0, virtual: 0, sin_definir: 0 },
        12: { presencial: 0, virtual: 0, sin_definir: 0 }
    };
    const dailyCounts = {}; // month -> { dayKey: count }
    const unmatchedAlerts = [];

    sheetTurnos.forEach(row => {
        if (!row.Paciente) return;
        if (isCancelled(row)) return;

        const dateObj = parseDate(row.Fecha);
        if (!dateObj) return;

        const rawText = String(row.Paciente).trim();
        const cleaned = cleanApptText(rawText);
        const tokens = cleaned.split(/\s+/).filter(w => w.length >= 2);

        let matchedPatient = matchPatientInList(tokens, controlPatients);

        if (matchedPatient) {
            if (dateObj.y === targetYear) {
                matchedPatient.turnos2026.push(dateObj);
            }
        } else {
            const matchedOther = matchPatientInList(tokens, otherPatients);
            if (matchedOther) {
                if (dateObj.y === targetYear) {
                    matchedOther.turnos2026.push(dateObj);
                }
            } else {
                if (dateObj.y === targetYear && targetMonths.includes(dateObj.m)) {
                    const dateStr = `${String(dateObj.d).padStart(2, '0')}/${String(dateObj.m).padStart(2, '0')}/${dateObj.y}`;
                    unmatchedAlerts.push({
                        Turno: rawText,
                        Fecha: dateStr,
                        rawDate: dateObj
                    });
                }
            }
        }

        if (dateObj.y === targetYear && targetMonths.includes(dateObj.m)) {
            monthlyCounts[dateObj.m]++;
            const mod = detectModality(row);
            monthlyModality[dateObj.m][mod]++;

            const jsDate = new Date(Date.UTC(dateObj.y, dateObj.m - 1, dateObj.d));
            const weekdays = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
            const dayName = weekdays[jsDate.getUTCDay()];
            const formattedDate = `${String(dateObj.d).padStart(2, '0')}.${String(dateObj.m).padStart(2, '0')}.${dateObj.y}`;
            const dayKey = `- ${dayName} ${formattedDate}`;

            if (!dailyCounts[dateObj.m]) dailyCounts[dateObj.m] = {};
            dailyCounts[dateObj.m][dayKey] = (dailyCounts[dateObj.m][dayKey] || 0) + 1;
        }
    });

    function getComment(p) {
        const variants = [
            normalizeStr(`${p.originalApellido} ${p.originalNombre}`),
            normalizeStr(`${p.originalNombre} ${p.originalApellido}`)
        ];
        for (const v of variants) {
            if (oldCommentsMap[v]) return oldCommentsMap[v];
        }
        const keys = Object.keys(oldCommentsMap);
        for (const k of keys) {
            const kWords = k.split(' ');
            const matchA = p.apeWords.some(a => kWords.some(kw => stringSimilarity.compareTwoStrings(a, kw) > 0.75));
            const matchN = p.nomWords.some(n => kWords.some(kw => stringSimilarity.compareTwoStrings(n, kw) > 0.75));
            if (matchA && matchN) return oldCommentsMap[k];
        }
        return { comentario: '', seguimiento: '' };
    }

    const inconsistencias = [];

    controlPatients.forEach(p => {
        const comm = getComment(p);
        const commNorm = normalizeStr(comm.comentario);
        const segNorm = normalizeStr(comm.seguimiento);

        if (commNorm.includes('desestimar') || segNorm.includes('desestimar')) {
            return;
        }

        p.turnos2026.sort((a, b) => (a.m - b.m) || (a.d - b.d));
        const firstTurn = p.turnos2026[0];

        let firstCheckMonth = 8;
        if (firstTurn && firstTurn.m > 8) {
            firstCheckMonth = firstTurn.m;
        }

        const attendedMonths = new Set(p.turnos2026.filter(t => targetMonths.includes(t.m)).map(t => t.m));
        const requiredMonths = targetMonths.filter(m => m >= firstCheckMonth);
        const missingMonths = requiredMonths.filter(m => !attendedMonths.has(m));

        if (missingMonths.length > 0) {
            inconsistencias.push({
                Paciente: p.fullName,
                Estado: 'Activo - Control',
                'Meses Faltantes': missingMonths.join(', '),
                'Arranque (Si es nuevo)': firstTurn ? `${firstTurn.m}/${firstTurn.y}` : '',
                Comentario: comm.comentario,
                Seguimiento: comm.seguimiento
            });
        }
    });

    const inactivosConTurnos = [];
    allPatients.filter(p => !p.isActivo).forEach(p => {
        const futureTurns = p.turnos2026.filter(t => t.m >= 8 && t.m <= 12);
        if (futureTurns.length > 0) {
            inactivosConTurnos.push({
                Paciente: p.fullName,
                'Turnos Futuros': futureTurns.map(t => `${String(t.d).padStart(2, '0')}/${String(t.m).padStart(2, '0')}/${t.y}`).join(', ')
            });
        }
    });

    unmatchedAlerts.sort((a, b) => (a.rawDate.m - b.rawDate.m) || (a.rawDate.d - b.rawDate.d));
    const alertaApellidosRows = unmatchedAlerts.map(u => ({
        Turno: u.Turno,
        Fecha: u.Fecha
    }));

    // Escribir archivo Excel
    const newWb = xlsx.utils.book_new();
    const wsValidacion = xlsx.utils.json_to_sheet(inconsistencias.length ? inconsistencias : [{ Paciente: 'Sin inconsistencias' }]);
    const wsAlertaApellidos = xlsx.utils.json_to_sheet(alertaApellidosRows.length ? alertaApellidosRows : [{ Turno: 'Sin alertas' }]);
    const wsAlertaInactivos = xlsx.utils.json_to_sheet(inactivosConTurnos.length ? inactivosConTurnos : [{ Paciente: 'Sin alertas' }]);

    xlsx.utils.book_append_sheet(newWb, wsValidacion, 'Validacion');
    xlsx.utils.book_append_sheet(newWb, wsAlertaApellidos, 'Alerta_Apellidos');
    xlsx.utils.book_append_sheet(newWb, wsAlertaInactivos, 'Alerta_Inactivos');

    xlsx.writeFile(newWb, OUTPUT_PATH);

    const nombresMeses = {
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    };

    console.log("\n=======================================================");
    console.log("       REPORTE DE AUDITORÍA Y CAPACIDAD DE AGENDA      ");
    console.log("=======================================================\n");

    console.log("1. ANÁLISIS DE CAPACIDAD MENSUAL (Tope: 64 por mes):");
    console.log("-------------------------------------------------------");
    targetMonths.forEach(m => {
        const count = monthlyCounts[m];
        const margen = CAPACIDAD_MAXIMA - count;
        const pres = monthlyModality[m].presencial;
        const virt = monthlyModality[m].virtual;
        const modInfo = `(${pres} presenciales, ${virt} virtuales)`;

        let estado = "";
        if (margen < 0) {
            estado = `🔴 Sobrecupo: ${margen} (Límite excedido)`;
        } else if (margen === 0) {
            estado = `🟡 Capacidad Completa`;
        } else {
            estado = `🟢 Margen: +${margen} lugares libres`;
        }

        console.log(`* ${nombresMeses[m]} 2026: ${count} ${modInfo} -> ${estado}`);
    });

    console.log("\n2. CANTIDAD POR DÍA Y POR MES (Sin cancelados):");
    console.log("-------------------------------------------------------");
    targetMonths.forEach(m => {
        console.log(`\n--- ${nombresMeses[m]} 2026 ---`);
        const days = dailyCounts[m] || {};
        const sortedDayKeys = Object.keys(days).sort((a, b) => {
            const dateA = a.match(/\d{2}\.\d{2}\.\d{4}/);
            const dateB = b.match(/\d{2}\.\d{2}\.\d{4}/);
            if (!dateA || !dateB) return 0;
            const [dA, mA, yA] = dateA[0].split('.').map(Number);
            const [dB, mB, yB] = dateB[0].split('.').map(Number);
            return (mA - mB) || (dA - dB);
        });

        sortedDayKeys.forEach(k => {
            console.log(`${k} = ${days[k]}`);
        });
    });

    console.log("\n3. ALERTA DE APELLIDOS NO COINCIDENTES (No hallados en la base):");
    console.log("-------------------------------------------------------");
    if (alertaApellidosRows.length === 0) {
        console.log("No hay turnos con apellidos no coincidentes.");
    } else {
        alertaApellidosRows.forEach(a => {
            console.log(`- ${a.Fecha}: "${a.Turno}"`);
        });
    }

    console.log("\n4. ALERTA DE PACIENTES INACTIVOS CON TURNOS A FUTURO:");
    console.log("-------------------------------------------------------");
    if (inactivosConTurnos.length === 0) {
        console.log("No hay pacientes inactivos con turnos a futuro.");
    } else {
        inactivosConTurnos.forEach(i => {
            console.log(`- ${i.Paciente}: ${i['Turnos Futuros']}`);
        });
    }

    console.log("\n5. PACIENTES CON INCONSISTENCIAS DE CONTROL (Meses Faltantes):");
    console.log("-------------------------------------------------------");
    console.log(`Total de pacientes con huecos: ${inconsistencias.length}`);
    inconsistencias.slice(0, 10).forEach(inc => {
        console.log(`* ${inc.Paciente} -> Faltan: [${inc['Meses Faltantes']}] ${inc.Comentario ? `(Nota: ${inc.Comentario})` : ''}`);
    });
    if (inconsistencias.length > 10) {
        console.log(`... y ${inconsistencias.length - 10} pacientes más detallados en el Excel.`);
    }

    console.log(`\n✓ Archivo Excel generado: ${OUTPUT_PATH}`);
}

runAudit();
