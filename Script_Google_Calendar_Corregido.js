const CALENDAR_NAME = "Pacientes";
const SHEET_NAME = "Turnos";
const TZ = "America/Argentina/Buenos_Aires";

// --- Menú Personalizado para ejecutar desde el Excel ---
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Sincronizar Agenda')
      .addItem('Descargar Turnos desde Google Calendar', 'exportarTurnos2026')
      .addToUi();
}

// --- 1. Lógica Matemática de Similitud (Fuzzy Matching) nativa ---
function compareTwoStrings(first, second) {
  first = String(first || '').replace(/\s+/g, '');
  second = String(second || '').replace(/\s+/g, '');

  if (first === second) return 1;
  if (first.length < 2 || second.length < 2) return 0;

  let firstBigrams = new Map();
  for (let i = 0; i < first.length - 1; i++) {
    const bigram = first.slice(i, i + 2);
    firstBigrams.set(bigram, (firstBigrams.get(bigram) || 0) + 1);
  }

  let intersectionSize = 0;
  for (let i = 0; i < second.length - 1; i++) {
    const bigram = second.slice(i, i + 2);
    const count = firstBigrams.get(bigram) || 0;
    if (count > 0) {
      firstBigrams.set(bigram, count - 1);
      intersectionSize++;
    }
  }

  return (2.0 * intersectionSize) / (first.length - 1 + second.length - 1);
}

// Normalización sin dependencias externas (quita tildes y convierte a minúsculas)
function normalize(str) {
  if (!str) return "";
  const map = {
    'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ü': 'u',
    'à': 'a', 'è': 'e', 'ì': 'i', 'ò': 'o', 'ù': 'u',
    'Á': 'a', 'É': 'e', 'Í': 'i', 'Ó': 'o', 'Ú': 'u', 'Ü': 'u'
  };
  return str.toString().toLowerCase().trim().replace(/[áéíóúüàèìòùÁÉÍÓÚÜ]/g, match => map[match]);
}

function cleanAgendaName(rawName) {
  const norm = normalize(rawName);
  return norm.replace(/\b(ctrol|ctrl|ctr|ctl|control|p|v|1 vez|vez|presencial|virtual)\b/g, ' ')
    .replace(/-/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function exportarTurnos2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 2. Cargar Pacientes Maestros ---
  const shPacientes = ss.getSheetByName("Pacientes");
  if (!shPacientes) throw new Error('No existe la hoja "Pacientes".');

  const pacData = shPacientes.getDataRange().getValues();
  const allPacientes = [];

  // Asumimos fila 1 como encabezados. Col C (index 2)=Nombre, Col D (index 3)=Apellido
  for (let i = 1; i < pacData.length; i++) {
    const nombre = String(pacData[i][2] || "").trim();
    const apellido = String(pacData[i][3] || "").trim();
    if (nombre || apellido) {
      allPacientes.push({
        nombre: nombre,
        apellido: apellido,
        fullName: (nombre + " " + apellido).trim()
      });
    }
  }

  // --- 3. Preparar Hoja Destino (Turnos) ---
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`No existe la hoja "${SHEET_NAME}".`);

  const calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (!calendars.length) throw new Error(`No se encontró el calendario "${CALENDAR_NAME}".`);
  const cal = calendars[0];

  const headers = [
    "Fecha",               // Col A (1)
    "Hora",                // Col B (2)
    "Paciente",            // Col C (3)
    "Validacion",          // Col D (4) - OK o Revisar / Paciente Nuevo
    "Importe",             // Col E (5)
    "Medio de Pago",       // Col F (6)
    "Estado Pago",         // Col G (7) - PAGADO / DEBE
    "Estado Turno",        // Col H (8) - NUEVA (Atendido, Cancelado, etc.)
    "Comentario",          // Col I (9) - Desplazada
    "Event ID",            // Col J (10) - Desplazada
    "Ultima Modificación"  // Col K (11) - Desplazada
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  } else if (!sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].includes("Ultima Modificación")) {
    sh.getRange(1, 11).setValue("Ultima Modificación");
  }

  const eventRows = new Map();
  const unverifiedRows = new Map();
  const lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    // Leemos las 11 columnas
    const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();

    data.forEach((row, index) => {
      const cellDate = row[0]; // Fecha (Col A)
      const pacienteName = row[2]; // Paciente (Col C)
      const currentValidacion = row[3]; // Validacion (Col D)
      const eventId = row[9];  // ID del evento (Col J - index 9)
      const rowNum = index + 2;

      // Intentar extraer el año para no alterar turnos viejos (pre-2026)
      if (pacienteName) {
        let y = 0;
        if (cellDate) {
          try {
            const d = new Date(cellDate);
            if (!isNaN(d.getFullYear())) {
              y = d.getFullYear();
            }
          } catch (e) { }
          if (y === 0 && typeof cellDate === 'string') {
            const parts = cellDate.split('.');
            if (parts.length === 3) y = parseInt(parts[2], 10);
          }
        }
        unverifiedRows.set(rowNum, { validacion: currentValidacion, year: y });
      }

      if (eventId && cellDate) {
        let dateStr = "";
        try {
          const d = new Date(cellDate);
          dateStr = Utilities.formatDate(d, TZ, "yyyy-MM-dd");
        } catch (e) {
          dateStr = String(cellDate);
        }

        const compositeKey = String(eventId) + "_" + dateStr;
        eventRows.set(compositeKey, {
          rowNumber: index + 2,
          lastUpdated: row[10] ? String(row[10]) : "", // Ultima Modificación (Col K - index 10)
          paciente: String(row[2] || ""),
          validacion: String(row[3] || "")
        });
      }
    });
  }

  const rowsToAdd = [];

  // --- 4. Alcance Multianual (Desde 2026 hacia el futuro) ---
  const currentYear = new Date().getFullYear();
  // El límite superior será al menos el año actual + 2 (o 2027 como mínimo)
  const endYear = Math.max(currentYear + 2, 2027);

  // Comienza siempre en 2026 sin importar en qué año estemos
  for (let year = 2026; year <= endYear; year++) {
    for (let month = 0; month < 12; month++) {
      const start = new Date(year, month, 1, 0, 0, 0);
      const end = new Date(year, month + 1, 1, 0, 0, 0);

      const events = cal.getEvents(start, end);

      for (const ev of events) {
        const pacienteOriginal = ev.getTitle() || "";

        if (pacienteOriginal === "Virtual" || pacienteOriginal === "Presencial - Consultorio") {
          continue;
        }

        const id = String(ev.getId());
        const evStart = ev.getStartTime();

        const dateStr = Utilities.formatDate(evStart, TZ, "yyyy-MM-dd");
        const hora = Utilities.formatDate(evStart, TZ, "HH:mm");
        const fecha = new Date(evStart.getFullYear(), evStart.getMonth(), evStart.getDate());
        const lastUpdated = Utilities.formatDate(ev.getLastUpdated(), TZ, "yyyy-MM-dd HH:mm:ss");

        // --- 5. Lógica de Limpieza y Matching Nativo ---
        const cleanName = cleanAgendaName(pacienteOriginal);
        const rawWords = cleanName.split(' ');

        let matchedPatient = null;
        let bestRating = 0;

        for (const p of allPacientes) {
          const normNombre = normalize(p.nombre);
          const normApellido = normalize(p.apellido);

          const mainLastName = normApellido.split(' ')[0];
          if (!mainLastName) continue;

          // Exigir similitud > 0.75 en al menos una parte del apellido de la agenda
          const surnameMatches = rawWords.some(w => compareTwoStrings(w, mainLastName) > 0.75);
          if (surnameMatches) {
            const fullTarget = normNombre + " " + normApellido;
            const rating = compareTwoStrings(cleanName, fullTarget);
            if (rating > bestRating) {
              bestRating = rating;
              matchedPatient = p;
            }
          }
        }

        let finalName = pacienteOriginal;
        let estadoVal = "Revisar / Paciente Nuevo";

        if (matchedPatient && bestRating > 0.45) {
          finalName = matchedPatient.fullName;
          
          // Rescatar dinámicamente cualquier palabra extra del turno original
          const rawOriginalWords = pacienteOriginal.split(/\s+/);
          const extraWords = [];
          
          const fullNameNormWords = normalize(matchedPatient.fullName).split(/\s+/);
          
          for (const rawW of rawOriginalWords) {
             const normW = normalize(rawW);
             // Ignorar espacios vacíos o símbolos sueltos (ej: guiones)
             if (!normW || !normW.match(/[a-z0-9]/i)) continue;
             
             // Si es la palabra cancela, la ignoramos acá porque se procesa de forma fija abajo
             if (normW.includes("cancela")) continue;
             
             let isNamePart = false;
             
             // Tolerancia matemática: buscamos si esta palabra pertenece al nombre real
             if (normW.length <= 2) {
                 isNamePart = fullNameNormWords.includes(normW);
             } else {
                 for (const nameW of fullNameNormWords) {
                     if (nameW.length > 2 && compareTwoStrings(normW, nameW) > 0.75) {
                         isNamePart = true;
                         break;
                     } else if (normW === nameW) {
                         isNamePart = true;
                         break;
                     }
                 }
             }
             
             if (!isNamePart) {
                 // Embellecer algunas abreviaturas conocidas si aparecen
                 if (normW === 'p') extraWords.push("Presencial");
                 else if (normW === 'v') extraWords.push("Virtual");
                 else if (normW === 'ctrl' || normW === 'ctrol' || normW === 'ctr' || normW === 'ctl') extraWords.push("Control");
                 else extraWords.push(rawW);
             }
          }
          
          if (extraWords.length > 0) {
              finalName += " - " + extraWords.join(" ");
          }
          
          // Preservar la etiqueta de cancelado si estaba en el calendario
          if (normalize(pacienteOriginal).includes('cancela')) {
              finalName += " - Cancelado";
          }
          
          estadoVal = "Nombre Validado";
        }

        const compositeKey = id + "_" + dateStr;

        if (eventRows.has(compositeKey)) {
          const existing = eventRows.get(compositeKey);
          unverifiedRows.delete(existing.rowNumber); // Lo quitamos de los "huérfanos"

          // Actualizar la fila SOLO si el nombre corregido, la validación o el lastUpdated cambiaron
          // Se preservan intactas las columnas manuales: Importe (5), Medio de Pago (6), Estado Pago (7), Estado Turno (8) y Comentario (9)
          if (existing.lastUpdated !== lastUpdated || existing.paciente !== finalName || existing.validacion !== estadoVal) {
            sh.getRange(existing.rowNumber, 1).setValue(fecha);
            sh.getRange(existing.rowNumber, 2).setValue(hora);
            sh.getRange(existing.rowNumber, 3).setValue(finalName);
            sh.getRange(existing.rowNumber, 4).setValue(estadoVal);
            sh.getRange(existing.rowNumber, 11).setValue(lastUpdated); // Col K (11)
          }

        } else {
          rowsToAdd.push([
            fecha,        // Col A (1)
            hora,         // Col B (2)
            finalName,    // Col C (3)
            estadoVal,    // Col D (4)
            "",           // Col E (5) - Importe
            "",           // Col F (6) - Medio de Pago
            "",           // Col G (7) - Estado Pago
            "",           // Col H (8) - Estado Turno
            "",           // Col I (9) - Comentario
            id,           // Col J (10) - Event ID
            lastUpdated   // Col K (11) - Ultima Modificación
          ]);
        }
      }
    }
  }

  // --- 6. Detectar Turnos Huérfanos/Borrados (Solo de 2026 en adelante) ---
  for (const [rowNum, info] of unverifiedRows.entries()) {
    if (info.year >= 2026 && info.validacion !== "No está en calendario") {
      sh.getRange(rowNum, 4).setValue("No está en calendario");
    }
  }

  if (rowsToAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rowsToAdd.length, 11).setValues(rowsToAdd);
  }

  const totalRows = sh.getLastRow();
  if (totalRows > 1) {
    sh.getRange(2, 1, totalRows - 1, 11)
      .sort([
        { column: 1, ascending: true },
        { column: 2, ascending: true }
      ]);

    sh.getRange(2, 1, totalRows - 1, 1).setNumberFormat("dddd dd.MM.yyyy");
    sh.getRange(2, 2, totalRows - 1, 1).setNumberFormat("HH:mm");
  }
}
