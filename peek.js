const xlsx = require('xlsx');

function peek(file) {
    try {
        const wb = xlsx.readFile(file);
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, {header: 1, raw: false});
        console.log(`--- ${file} ---`);
        console.log("Columns:", data[0]);
        console.log("Row 1:", data[1]);
        console.log("Row 2:", data[2]);
    } catch(e) {
        console.log(`Error reading ${file}: ${e.message}`);
    }
}

peek('Pacientes.xlsx');
peek('Turnos.xlsx');
peek('Validacion_Controles_Actualizado.xlsx');
