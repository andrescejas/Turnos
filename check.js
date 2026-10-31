const xlsx = require('xlsx');
const wbTurnos = xlsx.readFile('Turnos.xlsx');
const turnosData = xlsx.utils.sheet_to_json(wbTurnos.Sheets[wbTurnos.SheetNames[0]], {defval: ''});

function normalize(str) {
    if (!str) return '';
    return str.toString().toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

console.log('Turnos para Maria Emilia Cenci:');
let found = 0;
turnosData.forEach(row => {
    const pacienteRaw = normalize(row['Paciente']);
    if (pacienteRaw.includes('cenci') || pacienteRaw.includes('emilia')) {
        let rawDate = row['Fecha'];
        let d;
        if (typeof rawDate === 'number') {
            const parsed = xlsx.SSF.parse_date_code(rawDate);
            d = new Date(parsed.y, parsed.m - 1, parsed.d);
        } else {
            d = new Date(rawDate);
        }
        
        let dateStr = isNaN(d.getTime()) ? rawDate : d.toISOString().split('T')[0];
        console.log('- Fecha: ' + dateStr + ', Paciente exacto en excel: "' + row['Paciente'] + '", Asistio: "' + row['Asistio'] + '"');
        found++;
    }
});
if (found === 0) console.log('No se encontraron turnos con ese nombre.');
