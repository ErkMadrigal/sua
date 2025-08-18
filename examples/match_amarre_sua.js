const xlsx = require("xlsx");

// Función para leer columna F de un archivo
function readColumnF(filename) {
  const workbook = xlsx.readFile(filename);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  
  // Extrae solo la columna F
  return data
    .map(row => row[5]) // Columna F → índice 5
    .filter(value => value); // Quita vacíos
}

// Leer SUA y DATA
const suaNames = readColumnF("sua.xls").map(n => n.toString().trim().toUpperCase());
const dataNames = readColumnF("data.xlsx");

// Revisa coincidencias
dataNames.forEach(name => {
  if (name) {
    const formatted = name.toString().trim().toUpperCase();
    if (!suaNames.includes(formatted)) {
      console.log("No encontrado:", name);
    }
  }
});
