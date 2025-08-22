const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const session = require("express-session");

const app = express();
const upload = multer({ dest: "uploads/" });

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));

app.use(session({
  secret: 'supersecretkey',
  resave: false,
  saveUninitialized: true,
  cookie: { maxAge: 60 * 60 * 1000 } // 1 hora
}));

// Asegúrate de tener una carpeta 'downloads' en 'public'
const downloadsDir = path.join(__dirname, "public/downloads");
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir, { recursive: true });
}

// Normalizador según opciones
function normalizeText(text, removeAccents, trimSpaces) {
  if (!text) return "";
  let t = text.toString();

  if (removeAccents) {
    t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // quita acentos
  }
  if (trimSpaces) {
    t = t.replace(/\s+/g, " ").trim(); // quita espacios dobles y extremos
  }

  return t.toUpperCase(); // comparar sin importar mayúsculas/minúsculas
}

function readFullSheet(filePath, sheetName) {
  const wb = xlsx.readFile(filePath);
  const sheet = sheetName ? wb.Sheets[sheetName] : wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return [];
  return xlsx.utils.sheet_to_json(sheet, { header: 1 }); // Array de arrays (filas completas)
}

function getColIndex(columnLetter) {
  return columnLetter.toUpperCase().charCodeAt(0) - 65;
}

app.get("/", (req, res) => {
  res.render("index", { missing: null, matches: null, showMissing: false, showMatches: false, exportFilename: null });
});

app.post("/upload", upload.fields([{ name: "sua" }, { name: "data" }]), (req, res) => {
  try {
    const { suaSheet, suaColumn, dataSheet, dataColumn, removeAccents, trimSpaces, showMissing, showMatches, exportResults } = req.body;

    const removeAcc = !!removeAccents;
    const trimSp = !!trimSpaces;
    const showMiss = !!showMissing;
    const showMatch = !!showMatches;
    const expRes = !!exportResults;

    const hasFiles = req.files && req.files["sua"] && req.files["data"];

    if (hasFiles) {
      req.session.suaFull = readFullSheet(req.files["sua"][0].path, suaSheet);
      req.session.dataFull = readFullSheet(req.files["data"][0].path, dataSheet);

      // Borrar archivos subidos
      fs.unlinkSync(req.files["sua"][0].path);
      fs.unlinkSync(req.files["data"][0].path);
    }

    if (!req.session.suaFull || !req.session.dataFull) {
      return res.render("index", { missing: null, matches: null, showMissing: showMiss, showMatches: showMatch, exportFilename: null });
      // O res.send("Por favor sube los archivos primero.");
    }

    const suaFull = req.session.suaFull;
    const dataFull = req.session.dataFull;

    const suaColIndex = getColIndex(suaColumn || "F");
    const dataColIndex = getColIndex(dataColumn || "F");

    // Extraer nombres raw de SUA (asumiendo header en fila 0)
    const rawSuaNames = suaFull.slice(1).map(row => row[suaColIndex]).filter(val => val != null && val.toString().trim() !== "");

    const suaSet = new Set(rawSuaNames.map(text => normalizeText(text, removeAcc, trimSp)));

    // Procesar DATA
    let missing = [];
    let matches = [];
    let matchingRows = dataFull.length > 0 ? [dataFull[0]] : []; // Incluir headers
    let missingRows = dataFull.length > 0 ? [dataFull[0]] : [];

    for (let i = 1; i < dataFull.length; i++) {
      const row = dataFull[i];
      const originalName = row[dataColIndex] ? row[dataColIndex].toString() : "";
      const name = normalizeText(originalName, removeAcc, trimSp);
      if (name === "") continue;

      if (suaSet.has(name)) {
        matches.push(originalName);
        matchingRows.push(row);
      } else {
        missing.push(originalName);
        missingRows.push(row);
      }
    }

    let exportFilename = null;
    if (expRes && (matchingRows.length > 1 || missingRows.length > 1)) {
      const newWb = xlsx.utils.book_new();

      if (matchingRows.length > 1) {
        const matchWs = xlsx.utils.aoa_to_sheet(matchingRows);
        xlsx.utils.book_append_sheet(newWb, matchWs, "Coincidentes");
      }

      if (missingRows.length > 1) {
        const missWs = xlsx.utils.aoa_to_sheet(missingRows);
        xlsx.utils.book_append_sheet(newWb, missWs, "No Coincidentes");
      }

      exportFilename = `results_${Date.now()}.xlsx`;
      const exportPath = path.join(downloadsDir, exportFilename);
      xlsx.writeFile(newWb, exportPath);
    }

    res.render("index", { missing, matches, showMissing: showMiss, showMatches: showMatch, exportFilename });
  } catch (err) {
    console.error(err);
    // Borrar archivos en caso de error
    if (req.files && req.files["sua"]) fs.unlinkSync(req.files["sua"][0].path);
    if (req.files && req.files["data"]) fs.unlinkSync(req.files["data"][0].path);
    res.send("❌ Error procesando los archivos");
  }
});

// Ruta para descargar
app.get("/download/:file", (req, res) => {
  const filePath = path.join(downloadsDir, req.params.file);
  if (fs.existsSync(filePath)) {
    res.download(filePath, (err) => {
      if (err) console.error(err);
      // Borrar después de descargar
      fs.unlink(filePath, (err) => { if (err) console.error(err); });
    });
  } else {
    res.send("Archivo no encontrado");
  }
});

// ===== PUERTO =====
const PORT = process.env.PORT || 5000;
app.listen(PORT, () =>
  console.log(`🚀 Servidor corriendo en http://localhost:${PORT}`)
);