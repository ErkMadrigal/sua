const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));

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

function readColumn(filePath, sheetName, columnLetter = "F", removeAccents, trimSpaces) {
  const wb = xlsx.readFile(filePath);
  const sheet = sheetName ? wb.Sheets[sheetName] : wb.Sheets[wb.SheetNames[0]];
  if (!sheet) return [];

  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  const colIndex = columnLetter.toUpperCase().charCodeAt(0) - 65;

  return data
    .map(row => normalizeText(row[colIndex], removeAccents, trimSpaces))
    .filter(val => val !== "");
}

app.get("/", (req, res) => {
  res.render("index", { results: null });
});

app.post("/upload", upload.fields([{ name: "sua" }, { name: "data" }]), (req, res) => {
  try {
    const { suaSheet, suaColumn, dataSheet, dataColumn, removeAccents, trimSpaces } = req.body;

    const suaNames = readColumn(
      req.files.sua[0].path,
      suaSheet,
      suaColumn || "F",
      !!removeAccents,
      !!trimSpaces
    );

    const dataNames = readColumn(
      req.files.data[0].path,
      dataSheet,
      dataColumn || "F",
      !!removeAccents,
      !!trimSpaces
    );

    const missing = dataNames.filter(name => !suaNames.includes(name));

    res.render("index", { results: missing });
  } catch (err) {
    console.error(err);
    res.send("❌ Error procesando los archivos");
  }
});


// ===== PUERTO =====
const PORT = process.env.PORT || 5000;
app.listen(PORT, () =>
  console.log(`🚀 Servidor corriendo en http://localhost:${PORT}`)
);
