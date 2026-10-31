const xlsx = require('xlsx');
const wbPacientes = xlsx.readFile('Pacientes.xlsx');
const pacData = xlsx.utils.sheet_to_json(wbPacientes.Sheets[wbPacientes.SheetNames[0]], {defval: ''});

const namesToCheck = 'luciano rivata, constanza obregon, lis gaia, carolina rubiolo, macarena diaz, malena loza, gonzalo lopez nole, jose vijande, david rubin, lara giuliana nievas, laura paulino, natalia marquinez, santino bongiovani, lautaro almada, jaled alcaraz, elian perez, aldana correa, alejo mendieta, melina salome caminos, constanza martinez, emma diaz'.split(', ').map(n => n.trim().toLowerCase());

console.log('Buscando coincidencias parciales en Pacientes.xlsx...');
namesToCheck.forEach(name => {
    const parts = name.split(' ');
    const possibleMatches = pacData.filter(row => {
        const full = (row['Nombre'] + ' ' + row['Apellido']).toLowerCase();
        return parts.some(p => p.length > 3 && full.includes(p));
    }).map(row => row['Nombre'] + ' ' + row['Apellido']);
    
    if (possibleMatches.length > 0) {
        console.log("- '" + name + "' podría ser: " + possibleMatches.join(' O '));
    } else {
        console.log("- '" + name + "': Sin rastro en la base de datos.");
    }
});
