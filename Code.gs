// ============================================================
// Code.gs — Backend Google Apps Script
// Prompt Manager SPA — ABM (Altas, Bajas, Modificaciones)
// ============================================================

// ── Configuración ──────────────────────────────────────────
const SHEET_NAME = "Prompts";
const HEADERS = ["Categoria", "Nombre prompt", "Prompt", "Ejemplos"];

// ── Servir el frontend ─────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Prompt Manager")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── Configuración inicial de la hoja ───────────────────────
/**
 * Ejecutar UNA VEZ para crear la hoja con los encabezados.
 * Menú: Extensiones > Apps Script > ejecutar setupSheet
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Si la hoja no existe, crearla
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Escribir encabezados en la fila 1
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);

  // Formato de los encabezados
  headerRange
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  // Ajustar ancho de columnas
  sheet.setColumnWidth(1, 180);  // Categoria
  sheet.setColumnWidth(2, 220);  // Nombre prompt
  sheet.setColumnWidth(3, 450);  // Prompt
  sheet.setColumnWidth(4, 350);  // Ejemplos

  // Congelar fila de encabezados
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    "✅ Hoja \"" + SHEET_NAME + "\" configurada correctamente con las columnas:\n" +
    HEADERS.join(", ")
  );
}

// ── CRUD — Leer todos los registros ────────────────────────
function getPrompts() {
  const sheet = _getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return []; // Solo encabezados o vacía

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  return data.map(function (row, index) {
    return {
      id: index + 2, // Fila real en la hoja (fila 2 en adelante)
      categoria: row[0],
      nombre: row[1],
      prompt: row[2],
      ejemplos: row[3]
    };
  });
}

// ── CRUD — Crear un registro (Alta) ────────────────────────
function addPrompt(data) {
  const sheet = _getSheet();
  sheet.appendRow([
    data.categoria,
    data.nombre,
    data.prompt,
    data.ejemplos
  ]);
  return { success: true, message: "Prompt creado correctamente." };
}

// ── CRUD — Actualizar un registro (Modificación) ───────────
function updatePrompt(data) {
  const sheet = _getSheet();
  const rowIndex = data.id; // Fila real en la hoja

  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
    return { success: false, message: "Fila no encontrada." };
  }

  sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([
    [data.categoria, data.nombre, data.prompt, data.ejemplos]
  ]);

  return { success: true, message: "Prompt actualizado correctamente." };
}

// ── CRUD — Eliminar un registro (Baja) ─────────────────────
function deletePrompt(rowId) {
  const sheet = _getSheet();

  if (rowId < 2 || rowId > sheet.getLastRow()) {
    return { success: false, message: "Fila no encontrada." };
  }

  sheet.deleteRow(rowId);
  return { success: true, message: "Prompt eliminado correctamente." };
}

// ── Obtener categorías únicas ──────────────────────────────
function getCategories() {
  const sheet = _getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const unique = [...new Set(data.map(function (r) { return r[0]; }).filter(Boolean))];
  return unique.sort();
}

// ── Utilidad privada ───────────────────────────────────────
function _getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(
      'La hoja "' + SHEET_NAME + '" no existe. Ejecuta setupSheet() primero.'
    );
  }
  return sheet;
}

// ── Rellenar con Datos de Ejemplo ──────────────────────────
/**
 * Ejecuta esta función para rellenar la tabla con datos de prueba.
 * Menú: Extensiones > Apps Script > seleccionar e iniciar "insertDummyData"
 */
function insertDummyData() {
  const sheet = _getSheet();
  
  // Limpiar cualquier dato previo (excepto encabezados)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  const dummyData = [
    [
      "Marketing",
      "Generador de copies para Instagram",
      "Actúa como un redactor creativo de marketing. Escribe 3 opciones de copias llamativas para Instagram sobre el producto: {producto}. Incluye emojis, llamadas a la acción (CTA) y hashtags relevantes.",
      "Ejemplo de producto: Auriculares inalámbricos con cancelación de ruido."
    ],
    [
      "Desarrollo",
      "Refactorizador y optimizador de código JS",
      "Analiza el siguiente fragmento de código JavaScript y refactorízalo para que sea más legible, eficiente y siga las mejores prácticas modernas (ES6+). Explica brevemente los cambios realizados.\n\nCódigo:\n{codigo}",
      "Ejemplo de código: function sumar(a,b){ var res = 0; res = a + b; return res; }"
    ],
    [
      "Redacción",
      "Corrector de estilo y tono",
      "Reescribe el siguiente texto para que tenga un tono profesional, claro y persuasivo. Corrige cualquier error gramatical u ortográfico sin cambiar la idea principal del mensaje.\n\nTexto:\n{texto}",
      "Ejemplo de texto: hola keria saber si me podes mandar el presupuesto de la pag web q te pedi la semana pasada gracias"
    ],
    [
      "Productividad",
      "Resumen ejecutivo de notas",
      "Lee las siguientes notas desorganizadas de una reunión y genera un resumen ejecutivo estructurado en: 1) Puntos clave discutidos, 2) Decisiones tomadas, y 3) Tareas asignadas con sus respectivos responsables.\n\nNotas:\n{notas}",
      "Ejemplo de notas: Carlos dice que el diseño estará listo el martes. Ana necesita revisar la base de datos."
    ]
  ];

  // Insertar registros
  sheet.getRange(2, 1, dummyData.length, dummyData[0].length).setValues(dummyData);
  
  SpreadsheetApp.getUi().alert("✅ Se han insertado " + dummyData.length + " prompts de ejemplo correctamente.");
}

