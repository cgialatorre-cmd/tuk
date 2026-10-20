// Guarda cada respuesta de la encuesta de TUK IA en esta hoja de Google.
// Cómo instalarlo (una sola vez):
//  1. sheets.new → nombra la hoja «TUK IA · encuesta».
//  2. Extensiones → Apps Script → borra lo que haya y pega este archivo → Guardar.
//  3. Implementar → Nueva implementación → tipo «App web» →
//     Ejecutar como: «Yo» · Quién tiene acceso: «Cualquier usuario» → Implementar → Autorizar.
//  4. Copia la URL que termina en /exec y pásasela a Claude (va en config.js).

const COLUMNAS = ["fecha", "rol", "frecuencia", "usos", "problemas", "privacidad", "usaria", "camino",
                  "especial", "celular", "pago", "sugerencia", "origen", "version"];

function doPost(e) {
  const datos = JSON.parse(e.postData.contents || "{}");
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = libro.getSheetByName("Respuestas") || libro.insertSheet("Respuestas");
  if (hoja.getLastRow() === 0) hoja.appendRow(COLUMNAS);
  const limpio = (v) => {
    let s = Array.isArray(v) ? v.join(", ") : String(v === undefined || v === null ? "" : v);
    s = s.slice(0, 500);
    return /^[=+\-@]/.test(s) ? "'" + s : s;   // evita que un texto se ejecute como fórmula
  };
  hoja.appendRow(COLUMNAS.map((c) => (c === "fecha" ? new Date() : limpio(datos[c]))));
  return ContentService.createTextOutput("ok");
}

function doGet() {
  return ContentService.createTextOutput("La encuesta de TUK IA está en https://cgialatorre-cmd.github.io/tuk/");
}
