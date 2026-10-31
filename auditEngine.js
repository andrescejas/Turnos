const xlsx = require('xlsx');
const stringSimilarity = require('string-similarity');
const fs = require('fs');
const path = require('path');

const TURNOS_PATH = path.join(__dirname, 'Turnos.xlsx');
const OUTPUT_PATH = path.join(__dirname, 'Validacion_Controles_Actualizado.xlsx');
const CECILIA_VAL_PATH = path.join(__dirname, 'Validacion_Controles_ComentarioCecilia.xlsx');
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
    const st = normalizeStr(t['Estado Turno']);
    return p.includes('cancela') || v.includes('cancela') || e.includes('cancela') || st.includes('cancela');
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

function formatHora(rawHora) {
    if (typeof rawHora === 'number') {
        const totalMinutes = Math.round(rawHora * 24 * 60);
        const hours = Math.floor(totalMinutes / 60) % 24;
        const minutes = totalMinutes % 60;
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')} hs`;
    }
    if (!rawHora) return '--:--';
    return String(rawHora).trim();
}

function getCanonicalEstadoTurno(row, dateObj) {
    if (isCancelled(row)) return 'CANCELADO';
    const rawSt = row['Estado Turno'] ? String(row['Estado Turno']).trim().toUpperCase() : '';
    if (rawSt === 'ATENDIDO') return 'ATENDIDO';
    if (rawSt === 'AUSENTE') return 'AUSENTE';
    if (rawSt === 'CANCELADO') return 'CANCELADO';
    
    const now = new Date();
    const refYear = now.getFullYear();
    const refMonth = now.getMonth() + 1;
    const refDay = now.getDate();
    
    const isPast = (dateObj.y < refYear) ||
                   (dateObj.y === refYear && dateObj.m < refMonth) ||
                   (dateObj.y === refYear && dateObj.m === refMonth && dateObj.d <= refDay);
                   
    return isPast ? 'ATENDIDO' : 'AGENDADO';
}

function isTurnoAtendido(t, rawEstadoTurno) {
    if (rawEstadoTurno) {
        const st = String(rawEstadoTurno).trim().toUpperCase();
        if (st === 'ATENDIDO') return true;
        if (st === 'AUSENTE' || st === 'CANCELADO') return false;
    }
    const now = new Date();
    const refYear = now.getFullYear();
    const refMonth = now.getMonth() + 1;
    const refDay = now.getDate();
    
    if (t.y < refYear) return true;
    if (t.y > refYear) return false;
    if (t.m < refMonth) return true;
    if (t.m > refMonth) return false;
    return t.d <= refDay;
}

function runAuditEngine() {
    let filePath = TURNOS_PATH;
    if (!fs.existsSync(filePath)) {
        const candidates = ['Turnos.xlsx', 'Turnos (1).xlsx', 'Turnos-ant.xlsx'];
        for (const c of candidates) {
            const p = path.join(__dirname, c);
            if (fs.existsSync(p)) {
                filePath = p;
                break;
            }
        }
    }

    if (!fs.existsSync(filePath)) {
        throw new Error("No se encontró el archivo Turnos.xlsx en " + TURNOS_PATH);
    }

    const wb = xlsx.readFile(filePath);
    const sheetPacientes = xlsx.utils.sheet_to_json(wb.Sheets['Pacientes'], { defval: '' });
    const sheetTurnos = xlsx.utils.sheet_to_json(wb.Sheets['Turnos'], { defval: '' });

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

    const allPatients = sheetPacientes.map((p, idx) => {
        const nomWords = normalizeStr(p.Nombre).split(/\s+/).filter(w => w.length >= 3);
        const apeWords = normalizeStr(p.Apellido).split(/\s+/).filter(w => w.length >= 3);
        const activoNorm = normalizeStr(p.Activo);
        const frecNorm = normalizeStr(p.Frecuencia);
        const frecDias = parseInt(p['Frecuencia (dias)'], 10) || 0;

        const isActivo = activoNorm === 'si';
        const isDemanda = frecNorm.includes('demanda') || frecDias === 0;
        const isDosMeses = frecNorm.includes('dos meses') || frecDias === 60;
        const isQuinceDias = frecDias === 15;
        const isFinalizado = frecNorm.includes('finalizado');
        const isControl = frecDias === 30 || frecDias === 15 || (frecNorm.includes('control') && !isDemanda && !isDosMeses && !isFinalizado);

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
            isDosMeses,
            isQuinceDias,
            isFinalizado,
            frecuencia: String(p.Frecuencia || '').trim(),
            frecuenciaDias: frecDias,
            frecuenciaTexto: frecDias > 0 ? `Cada ${frecDias} días` : (p.Frecuencia || 'A demanda'),
            obraSocial: String(p['Obra Social '] || p['Obra Social'] || '').trim(),
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

    const allMesesNombres = {
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    };

    function getWeekOfMonth(y, m, d) {
        let week = 1;
        for (let day = 1; day < d; day++) {
            const dt = new Date(Date.UTC(y, m - 1, day));
            if (dt.getUTCDay() === 0) week++;
        }
        return week;
    }

    function getMonthWeeksMeta(y, m) {
        const daysInMonth = new Date(Date.UTC(y, m, 0)).getUTCDate();
        const weeks = {};
        for (let d = 1; d <= daysInMonth; d++) {
            const w = getWeekOfMonth(y, m, d);
            if (!weeks[w]) weeks[w] = { weekNum: w, startDay: d, endDay: d };
            else weeks[w].endDay = d;
        }
        return Object.values(weeks).map(w => ({
            weekNum: w.weekNum,
            label: `Semana ${w.weekNum}`,
            dateRange: `${String(w.startDay).padStart(2, '0')}/${String(m).padStart(2, '0')} - ${String(w.endDay).padStart(2, '0')}/${String(m).padStart(2, '0')}`,
            startDay: w.startDay,
            endDay: w.endDay
        }));
    }

    const weeklyData = {};
    [7, 8, 9, 10, 11, 12].forEach(m => {
        const weeksMeta = getMonthWeeksMeta(targetYear, m);
        const weeksObj = {};
        weeksMeta.forEach(w => {
            weeksObj[w.weekNum] = {
                ...w,
                count: 0,
                presencial: 0,
                virtual: 0,
                atendidos: 0,
                ausentes: 0,
                cancelados: 0,
                agendados: 0,
                pagados: 0,
                deben: 0
            };
        });
        weeklyData[m] = {
            mes: m,
            mesNombre: allMesesNombres[m],
            totalTurnos: 0,
            totalAtendidos: 0,
            totalAusentes: 0,
            totalCancelados: 0,
            totalAgendados: 0,
            weeks: weeksObj
        };
    });

    const monthlyCounts = { 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };
    const monthlyAtendidos = { 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };
    const monthlyAusentes = { 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };
    const monthlyCancelados = { 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };
    const monthlyAgendados = { 8: 0, 9: 0, 10: 0, 11: 0, 12: 0 };

    let totalCanceladosGlobal = 0;
    let totalAtendidosGlobal = 0;
    let totalAusentesGlobal = 0;
    let totalAgendadosGlobal = 0;

    const monthlyModality = {
        8: { presencial: 0, virtual: 0, sin_definir: 0 },
        9: { presencial: 0, virtual: 0, sin_definir: 0 },
        10: { presencial: 0, virtual: 0, sin_definir: 0 },
        11: { presencial: 0, virtual: 0, sin_definir: 0 },
        12: { presencial: 0, virtual: 0, sin_definir: 0 }
    };
    const dailyCounts = {}; // month -> array of { dayKey, date, dayName, count, presencial, virtual }
    const dailyModalityMap = {}; // month -> { dayKey: { presencial, virtual } }
    const unmatchedAlerts = [];

    sheetTurnos.forEach(row => {
        if (!row.Paciente) return;
        const dateObj = parseDate(row.Fecha);
        if (!dateObj) return;

        const isCanc = isCancelled(row);
        const canonicalEstado = getCanonicalEstadoTurno(row, dateObj);

        // Si es cancelado: se contabiliza para auditoría pero se descarta de la capacidad efectiva
        if (isCanc || canonicalEstado === 'CANCELADO') {
            if (dateObj.y === targetYear) {
                if (targetMonths.includes(dateObj.m)) {
                    monthlyCancelados[dateObj.m] = (monthlyCancelados[dateObj.m] || 0) + 1;
                    totalCanceladosGlobal++;
                }
                if (weeklyData[dateObj.m]) {
                    const w = getWeekOfMonth(dateObj.y, dateObj.m, dateObj.d);
                    if (weeklyData[dateObj.m].weeks[w]) {
                        weeklyData[dateObj.m].weeks[w].cancelados++;
                        weeklyData[dateObj.m].totalCancelados++;
                    }
                }
            }
            return;
        }

        const rawText = String(row.Paciente).trim();
        const cleaned = cleanApptText(rawText);
        const tokens = cleaned.split(/\s+/).filter(w => w.length >= 2);
        const mod = detectModality(row);

        const hora = formatHora(row.Hora);
        let rawImporte = row.Importe;
        let importe = null;
        if (rawImporte !== undefined && rawImporte !== null && rawImporte !== '') {
            const num = Number(rawImporte);
            if (!isNaN(num)) importe = num;
        }
        const medioPago = row['Medio de Pago'] ? String(row['Medio de Pago']).trim().toUpperCase() : '';
        const rawEstadoPago = row['Estado Pago'] !== undefined && row['Estado Pago'] !== null ? String(row['Estado Pago']).trim() : '';
        const rawEstadoTurno = row['Estado Turno'] !== undefined && row['Estado Turno'] !== null ? String(row['Estado Turno']).trim() : '';
        const comentario = row['Comentario'] !== undefined && row['Comentario'] !== null ? String(row['Comentario']).trim() : '';
        const atendido = canonicalEstado === 'ATENDIDO';
        const fechaStr = `${String(dateObj.d).padStart(2, '0')}/${String(dateObj.m).padStart(2, '0')}/${dateObj.y}`;

        const turnoObj = {
            ...dateObj,
            hora,
            rawHora: row.Hora,
            mod,
            rawText,
            importe,
            medioPago,
            estadoPagoExcel: rawEstadoPago,
            estadoTurno: rawEstadoTurno,
            canonicalEstado,
            comentario,
            atendido,
            fechaStr
        };

        let matchedPatient = matchPatientInList(tokens, controlPatients);

        if (matchedPatient) {
            if (dateObj.y === targetYear) {
                matchedPatient.turnos2026.push(turnoObj);
            }
        } else {
            const matchedOther = matchPatientInList(tokens, otherPatients);
            if (matchedOther) {
                if (dateObj.y === targetYear) {
                    matchedOther.turnos2026.push(turnoObj);
                }
            } else {
                if (dateObj.y === targetYear && targetMonths.includes(dateObj.m)) {
                    const dateStr = `${String(dateObj.d).padStart(2, '0')}/${String(dateObj.m).padStart(2, '0')}/${dateObj.y}`;
                    unmatchedAlerts.push({
                        Turno: rawText,
                        Fecha: dateStr,
                        rawDate: dateObj,
                        Modalidad: mod
                    });
                }
            }
        }

        if (dateObj.y === targetYear && targetMonths.includes(dateObj.m)) {
            monthlyCounts[dateObj.m]++;
            monthlyModality[dateObj.m][mod]++;
            if (canonicalEstado === 'ATENDIDO') {
                monthlyAtendidos[dateObj.m]++;
                totalAtendidosGlobal++;
            } else if (canonicalEstado === 'AUSENTE') {
                monthlyAusentes[dateObj.m]++;
                totalAusentesGlobal++;
            } else if (canonicalEstado === 'AGENDADO') {
                monthlyAgendados[dateObj.m]++;
                totalAgendadosGlobal++;
            }

            const jsDate = new Date(Date.UTC(dateObj.y, dateObj.m - 1, dateObj.d));
            const weekdays = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
            const dayName = weekdays[jsDate.getUTCDay()];
            const formattedDate = `${String(dateObj.d).padStart(2, '0')}.${String(dateObj.m).padStart(2, '0')}.${dateObj.y}`;
            const dayKey = `${dayName} ${formattedDate}`;

            if (!dailyModalityMap[dateObj.m]) dailyModalityMap[dateObj.m] = {};
            if (!dailyModalityMap[dateObj.m][dayKey]) {
                dailyModalityMap[dateObj.m][dayKey] = {
                    dayKey,
                    formattedDate,
                    dayName,
                    d: dateObj.d,
                    m: dateObj.m,
                    y: dateObj.y,
                    count: 0,
                    presencial: 0,
                    virtual: 0
                };
            }
            dailyModalityMap[dateObj.m][dayKey].count++;
            if (mod === 'presencial') dailyModalityMap[dateObj.m][dayKey].presencial++;
            if (mod === 'virtual') dailyModalityMap[dateObj.m][dayKey].virtual++;
        }

        // Registro semanal (incluyendo julio para la comparación del mes anterior)
        if (dateObj.y === targetYear && weeklyData[dateObj.m]) {
            const w = getWeekOfMonth(dateObj.y, dateObj.m, dateObj.d);
            if (weeklyData[dateObj.m].weeks[w]) {
                weeklyData[dateObj.m].weeks[w].count++;
                weeklyData[dateObj.m].totalTurnos++;
                if (mod === 'presencial') weeklyData[dateObj.m].weeks[w].presencial++;
                else if (mod === 'virtual') weeklyData[dateObj.m].weeks[w].virtual++;

                if (canonicalEstado === 'ATENDIDO') {
                    weeklyData[dateObj.m].weeks[w].atendidos++;
                    weeklyData[dateObj.m].totalAtendidos++;
                } else if (canonicalEstado === 'AUSENTE') {
                    weeklyData[dateObj.m].weeks[w].ausentes++;
                    weeklyData[dateObj.m].totalAusentes++;
                } else if (canonicalEstado === 'AGENDADO') {
                    weeklyData[dateObj.m].weeks[w].agendados++;
                    weeklyData[dateObj.m].totalAgendados++;
                }

                const normEst = (rawEstadoPago || '').toUpperCase();
                if (normEst === 'PAGADO') weeklyData[dateObj.m].weeks[w].pagados++;
                else if (normEst.startsWith('DEBE')) weeklyData[dateObj.m].weeks[w].deben++;
            }
        }
    });

    targetMonths.forEach(m => {
        const daysObj = dailyModalityMap[m] || {};
        dailyCounts[m] = Object.values(daysObj).sort((a, b) => a.d - b.d);
    });

    const weeklyAnalysis = targetMonths.map(m => {
        const cur = weeklyData[m];
        const prev = weeklyData[m - 1];
        const weeksList = Object.values(cur.weeks);
        const maxCount = Math.max(...weeksList.map(w => w.count));
        const semanaPico = weeksList.find(w => w.count === maxCount && maxCount > 0);
        const promSem = weeksList.length > 0 ? Math.round((cur.totalTurnos / weeksList.length) * 10) / 10 : 0;
        const prevProm = prev && Object.keys(prev.weeks).length > 0 ? Math.round((prev.totalTurnos / Object.keys(prev.weeks).length) * 10) / 10 : 0;
        const delta = prev ? cur.totalTurnos - prev.totalTurnos : 0;
        const deltaPct = prev && prev.totalTurnos > 0 ? Math.round((delta / prev.totalTurnos) * 1000) / 10 : 0;

        return {
            mes: m,
            mesNombre: cur.mesNombre,
            totalTurnos: cur.totalTurnos,
            totalAtendidos: cur.totalAtendidos,
            totalAusentes: cur.totalAusentes,
            totalCancelados: cur.totalCancelados,
            totalAgendados: cur.totalAgendados,
            promedioSemanal: promSem,
            semanas: weeksList,
            semanaPico: semanaPico ? { semana: semanaPico.label, turnos: semanaPico.count, rango: semanaPico.dateRange } : null,
            comparacionMesAnterior: {
                mesAnteriorNombre: prev ? prev.mesNombre : '',
                mesAnteriorTotal: prev ? prev.totalTurnos : 0,
                mesAnteriorPromedio: prevProm,
                deltaTurnos: delta,
                deltaPorcentual: deltaPct,
                tendencia: delta > 0 ? 'sube' : (delta < 0 ? 'baja' : 'igual')
            }
        };
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

        const attendedMonths = new Set(p.turnos2026.filter(t => targetMonths.includes(t.m) && t.canonicalEstado === 'ATENDIDO').map(t => t.m));
        const requiredMonths = targetMonths.filter(m => m >= firstCheckMonth);
        const missingMonths = requiredMonths.filter(m => !attendedMonths.has(m));

        if (missingMonths.length > 0) {
            inconsistencias.push({
                Paciente: p.fullName,
                Estado: 'Activo - Control',
                'Meses Faltantes': missingMonths.join(', '),
                mesesFaltantesArr: missingMonths,
                'Arranque (Si es nuevo)': firstTurn ? `${firstTurn.m}/${firstTurn.y}` : '',
                Comentario: comm.comentario,
                Seguimiento: comm.seguimiento,
                Frecuencia: p.frecuencia,
                ObraSocial: p.obraSocial
            });
        }
    });

    const inactivosConTurnos = [];
    allPatients.filter(p => !p.isActivo).forEach(p => {
        const futureTurns = p.turnos2026.filter(t => t.m >= 8 && t.m <= 12);
        if (futureTurns.length > 0) {
            inactivosConTurnos.push({
                Paciente: p.fullName,
                'Turnos Futuros': futureTurns.map(t => `${String(t.d).padStart(2, '0')}/${String(t.m).padStart(2, '0')}/${t.y}`).join(', '),
                turnosArr: futureTurns
            });
        }
    });

    const sinControlConTurnos = [];
    allPatients.forEach(p => {
        const turnosDesdeAgosto = (p.turnos2026 || []).filter(t => t.m >= 8);
        if (turnosDesdeAgosto.length > 0) {
            const noTieneControl = (p.frecuenciaDias === 0) || (!p.isControl);
            if (noTieneControl) {
                sinControlConTurnos.push({
                    Paciente: p.fullName,
                    Activo: p.isActivo ? 'Sí' : 'No',
                    Frecuencia: p.frecuencia || 'Sin Frecuencia',
                    FrecuenciaDias: p.frecuenciaDias,
                    TotalTurnos: turnosDesdeAgosto.length,
                    TurnosDetalle: turnosDesdeAgosto.map(t => `${t.fechaStr} (${t.hora})`).join(', '),
                    turnosArr: turnosDesdeAgosto,
                    Motivo: p.frecuenciaDias === 0 
                        ? 'Frecuencia (días) = 0' 
                        : `Frecuencia: "${p.frecuencia || 'Sin control'}"`
                });
            }
        }
    });

    unmatchedAlerts.sort((a, b) => (a.rawDate.m - b.rawDate.m) || (a.rawDate.d - b.rawDate.d));

    const nombresMeses = {
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    };

    const capacitySummary = targetMonths.map(m => {
        const count = monthlyCounts[m];
        const margen = CAPACIDAD_MAXIMA - count;
        const pres = monthlyModality[m].presencial;
        const virt = monthlyModality[m].virtual;
        const sinDef = monthlyModality[m].sin_definir;

        let status = 'ok';
        if (margen < 0) status = 'overflow';
        else if (margen === 0) status = 'full';

        return {
            month: m,
            monthName: nombresMeses[m],
            year: targetYear,
            count,
            limit: CAPACIDAD_MAXIMA,
            margin: margen,
            status,
            presencial: pres,
            virtual: virt,
            sin_definir: sinDef,
            atendidos: monthlyAtendidos[m],
            ausentes: monthlyAusentes[m],
            cancelados: monthlyCancelados[m] || 0,
            agendados: monthlyAgendados[m] || 0,
            pctOcupacion: Math.round((count / CAPACIDAD_MAXIMA) * 100)
        };
    });

    let totalTurnos = 0;
    let totalPresencial = 0;
    let totalVirtual = 0;
    capacitySummary.forEach(c => {
        totalTurnos += c.count;
        totalPresencial += c.presencial;
        totalVirtual += c.virtual;
    });

    const targetMonthsEvaluated = [10, 11, 12];
    const mesesNombresMap = { 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre" };
    const weekdaysNames = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];

    let totalRecaudadoGlobal = 0;
    let totalAtencionesGlobal = 0;
    let totalPendientesCobroGlobal = 0;

    const pacientesHistorial = allPatients.map(p => {
        const comm = getComment(p);
        const turnos = p.turnos2026 || [];
        turnos.sort((a, b) => (a.m - b.m) || (a.d - b.d));

        // Turnos desde agosto (m >= 8) categorizados formalmente por estado canónico
        const turnosDesdeAgosto = turnos.filter(t => t.m >= 8);
        const atendidos = turnosDesdeAgosto.filter(t => t.canonicalEstado === 'ATENDIDO');
        const ausentes = turnosDesdeAgosto.filter(t => t.canonicalEstado === 'AUSENTE');
        const futuros = turnosDesdeAgosto.filter(t => t.canonicalEstado === 'AGENDADO');

        let totalAbonado = 0;
        let cantPagas = 0;
        let cantPendientes = 0;

        const controlesAtendidos = atendidos.map(t => {
            const jsDate = new Date(Date.UTC(t.y, t.m - 1, t.d));
            const dayName = weekdaysNames[jsDate.getUTCDay()];

            let estadoPago = 'debe';
            let pagoLabel = 'Debe (Sin pago registrado)';
            let badgeClass = 'red';
            const normEst = (t.estadoPagoExcel || '').toUpperCase();

            if (normEst === 'PAGADO') {
                estadoPago = 'pagado';
                badgeClass = 'green';
                cantPagas++;
                if (t.importe !== null && t.importe > 0) {
                    totalAbonado += t.importe;
                    totalRecaudadoGlobal += t.importe;
                    pagoLabel = `Pagó: $${t.importe.toLocaleString('es-AR')}${t.medioPago ? ` (${t.medioPago})` : ''}`;
                } else if (t.importe === 0) {
                    pagoLabel = `Pagado: $0${t.medioPago ? ` (${t.medioPago})` : ''}`;
                } else {
                    pagoLabel = `Pagado${t.medioPago ? ` (${t.medioPago})` : ''}`;
                }
            } else if (normEst.startsWith('DEBE')) {
                estadoPago = 'debe';
                badgeClass = 'red';
                cantPendientes++;
                totalPendientesCobroGlobal++;
                if (t.importe !== null && t.importe > 0) {
                    pagoLabel = `Debe: $${t.importe.toLocaleString('es-AR')}`;
                } else {
                    pagoLabel = 'Debe (Cobro pendiente)';
                }
            } else if (normEst === 'BONIFICADO') {
                estadoPago = 'bonificado';
                pagoLabel = 'Bonificado';
                badgeClass = 'gray';
            } else {
                // Compatibilidad hacia atrás (si Estado Pago está vacío en meses anteriores)
                if (t.importe !== null && t.importe > 0) {
                    estadoPago = 'pagado';
                    totalAbonado += t.importe;
                    cantPagas++;
                    totalRecaudadoGlobal += t.importe;
                    pagoLabel = `Pagó: $${t.importe.toLocaleString('es-AR')}${t.medioPago ? ` (${t.medioPago})` : ''}`;
                    badgeClass = 'green';
                } else if (t.importe === 0) {
                    estadoPago = 'bonificado';
                    pagoLabel = 'Bonificado ($0)';
                    badgeClass = 'gray';
                } else {
                    cantPendientes++;
                    totalPendientesCobroGlobal++;
                }
            }

            totalAtencionesGlobal++;

            return {
                ...t,
                dayName,
                estadoPago,
                pagoLabel,
                badgeClass,
                comentario: t.comentario || '',
                importeFormatted: t.importe !== null && t.importe > 0 ? `$${t.importe.toLocaleString('es-AR')}` : '-'
            };
        });

        const turnosAusentes = ausentes.map(t => {
            const jsDate = new Date(Date.UTC(t.y, t.m - 1, t.d));
            const dayName = weekdaysNames[jsDate.getUTCDay()];
            return {
                ...t,
                dayName,
                estadoPago: 'ausente',
                pagoLabel: 'Ausente (Inasistencia)',
                badgeClass: 'rose',
                comentario: t.comentario || '',
                importeFormatted: t.importe !== null && t.importe > 0 ? `$${t.importe.toLocaleString('es-AR')}` : '-'
            };
        });

        const turnosProximos = futuros.map(t => {
            const jsDate = new Date(Date.UTC(t.y, t.m - 1, t.d));
            const dayName = weekdaysNames[jsDate.getUTCDay()];
            
            let estadoPago = 'futuro';
            let pagoLabel = 'Agendado a futuro';
            let badgeClass = 'blue';
            const normEst = (t.estadoPagoExcel || '').toUpperCase();

            if (normEst === 'PAGADO') {
                estadoPago = 'pagado';
                badgeClass = 'green';
                pagoLabel = `Pagó adelantado${t.importe ? `: $${t.importe.toLocaleString('es-AR')}` : ''}${t.medioPago ? ` (${t.medioPago})` : ''}`;
                if (t.importe !== null && t.importe > 0) {
                    totalAbonado += t.importe;
                    totalRecaudadoGlobal += t.importe;
                    cantPagas++;
                }
            } else if (normEst.startsWith('DEBE')) {
                estadoPago = 'debe';
                badgeClass = 'red';
                pagoLabel = `Debe${t.importe ? `: $${t.importe.toLocaleString('es-AR')}` : ''}`;
                cantPendientes++;
                totalPendientesCobroGlobal++;
            }

            return {
                ...t,
                dayName,
                estadoPago,
                pagoLabel,
                badgeClass,
                comentario: t.comentario || '',
                importeFormatted: t.importe !== null && t.importe > 0 ? `$${t.importe.toLocaleString('es-AR')}` : '-'
            };
        });

        // Última atención médica efectiva y próxima fecha sugerida
        let ultimaAtencion = null;
        let proximaFechaSugeridaStr = '';
        if (controlesAtendidos.length > 0) {
            ultimaAtencion = controlesAtendidos[controlesAtendidos.length - 1];
            if (p.frecuenciaDias > 0) {
                const lastDate = new Date(ultimaAtencion.y, ultimaAtencion.m - 1, ultimaAtencion.d);
                const nextDate = new Date(lastDate.getTime() + p.frecuenciaDias * 24 * 60 * 60 * 1000);
                proximaFechaSugeridaStr = `${String(nextDate.getDate()).padStart(2, '0')}/${String(nextDate.getMonth() + 1).padStart(2, '0')}/${nextDate.getFullYear()}`;
            }
        }

        // Proyección de meses venideros (Octubre, Noviembre, Diciembre):
        const proyeccionMeses = targetMonthsEvaluated.map(mesNum => {
            const mesNombre = mesesNombresMap[mesNum];

            // 1. ¿Ya se atendió en este mes?
            const at = controlesAtendidos.find(t => t.m === mesNum);
            if (at) {
                const comTxt = at.comentario ? ` • 💬 ${at.comentario}` : '';
                return {
                    mes: mesNum,
                    mesNombre,
                    estado: 'atendido',
                    titulo: `Atendido`,
                    detalle: `${at.fechaStr} (${at.hora}) • ${at.pagoLabel}${comTxt}`,
                    badge: at.badgeClass || 'green',
                    badgeLabel: 'Atendido',
                    turno: at
                };
            }

            // 2. ¿Tuvo inasistencia registrada (Ausente) en este mes?
            const aus = turnosAusentes.find(t => t.m === mesNum);
            if (aus) {
                const comTxt = aus.comentario ? ` • 💬 ${aus.comentario}` : '';
                return {
                    mes: mesNum,
                    mesNombre,
                    estado: 'ausente',
                    titulo: `Inasistencia (Ausente)`,
                    detalle: `${aus.fechaStr} (${aus.hora}) • Registrado como Ausente${comTxt}`,
                    badge: 'rose',
                    badgeLabel: 'Ausente',
                    turno: aus
                };
            }

            // 3. ¿Tiene turno futuro agendado en este mes?
            const fut = turnosProximos.find(t => t.m === mesNum);
            if (fut) {
                const comTxt = fut.comentario ? ` • 💬 ${fut.comentario}` : '';
                return {
                    mes: mesNum,
                    mesNombre,
                    estado: 'agendado',
                    titulo: `Turno Confirmado`,
                    detalle: `${fut.fechaStr} a las ${fut.hora} (${fut.mod})${comTxt}`,
                    badge: fut.badgeClass || 'blue',
                    badgeLabel: 'Agendado',
                    turno: fut
                };
            }

            // 4. Si no tiene turno ni atención en este mes:
            if (!p.isActivo) {
                return {
                    mes: mesNum,
                    mesNombre,
                    estado: 'inactivo',
                    titulo: `Inactivo`,
                    detalle: `Paciente dado de baja`,
                    badge: 'gray',
                    badgeLabel: 'Inactivo'
                };
            }

            if (p.frecuenciaDias === 30 || p.frecuenciaDias === 15) {
                return {
                    mes: mesNum,
                    mesNombre,
                    estado: 'debe_venir',
                    titulo: `Debe Venir`,
                    detalle: `Control cada ${p.frecuenciaDias} días (Sin turno agendado)`,
                    badge: 'red',
                    badgeLabel: 'Pendiente'
                };
            }

            if (p.frecuenciaDias === 60) {
                const lastM = ultimaAtencion ? ultimaAtencion.m : 7;
                const diffMonths = mesNum - lastM;
                if (diffMonths >= 2 || (diffMonths % 2 === 0)) {
                    return {
                        mes: mesNum,
                        mesNombre,
                        estado: 'debe_venir',
                        titulo: `Debe Venir`,
                        detalle: `Control cada 60 días (Corresponde este mes)`,
                        badge: 'red',
                        badgeLabel: 'Pendiente'
                    };
                } else {
                    return {
                        mes: mesNum,
                        mesNombre,
                        estado: 'no_corresponde',
                        titulo: `No corresponde`,
                        detalle: `Control cada 60 días (Intervalo en curso)`,
                        badge: 'gray',
                        badgeLabel: 'No Corresponde'
                    };
                }
            }

            if (p.frecuenciaDias === 0) {
                return {
                    mes: mesNum,
                    mesNombre,
                    estado: 'demanda',
                    titulo: `A Demanda`,
                    detalle: `No tiene frecuencia fija`,
                    badge: 'gray',
                    badgeLabel: 'A demanda'
                };
            }

            return {
                mes: mesNum,
                mesNombre,
                estado: 'debe_venir',
                titulo: `Debe Venir`,
                detalle: `Frecuencia: ${p.frecuenciaTexto}`,
                badge: 'yellow',
                badgeLabel: 'Revisar'
            };
        });

        return {
            id: p.id,
            fullName: p.fullName,
            nombre: p.originalNombre,
            apellido: p.originalApellido,
            isActivo: p.isActivo,
            isControl: p.isControl,
            frecuencia: p.frecuencia,
            frecuenciaDias: p.frecuenciaDias,
            frecuenciaTexto: p.frecuenciaTexto,
            obraSocial: p.obraSocial,
            comentario: comm.comentario,
            seguimiento: comm.seguimiento,
            controlesAtendidos,
            turnosAusentes,
            turnosProximos,
            proyeccionMeses,
            totalAbonado,
            totalAbonadoFormatted: `$${totalAbonado.toLocaleString('es-AR')}`,
            cantAtendidos: controlesAtendidos.length,
            cantAusencias: turnosAusentes.length,
            cantPagas,
            cantPendientesPago: cantPendientes,
            proximoTurno: turnosProximos.length > 0 ? turnosProximos[0] : null,
            proximaFechaSugerida: proximaFechaSugeridaStr,
            ultimaAtencion
        };
    });

    return {
        timestamp: new Date().toISOString(),
        capacidadMaxima: CAPACIDAD_MAXIMA,
        totalTurnos,
        totalPresencial,
        totalVirtual,
        totalAtendidosGlobal,
        totalAusentesGlobal,
        totalCanceladosGlobal,
        totalAgendadosGlobal,
        totalPacientes: allPatients.length,
        totalControlActivos: controlPatients.length,
        totalRecaudado: totalRecaudadoGlobal,
        totalAtenciones: totalAtencionesGlobal,
        totalPendientesCobro: totalPendientesCobroGlobal,
        capacitySummary,
        dailyCounts,
        weeklyAnalysis,
        inconsistencias,
        inactivosConTurnos,
        sinControlConTurnos,
        unmatchedAlerts,
        pacientesHistorial
    };
}

module.exports = { runAuditEngine, CAPACIDAD_MAXIMA };
