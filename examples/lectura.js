const xlsx = require("xlsx");

// Carga el archivo
const workbook = xlsx.readFile("sua.xls");

// Selecciona la primera hoja
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convierte la hoja a formato JSON (cada fila es un objeto)
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Recorre todas las filas
data.forEach((row, index) => {
  const value = row[5]; // Columna F (los índices empiezan en 0)
  if (value) {
    console.log(`Fila ${index + 1}: ${value}`);
  }
});
