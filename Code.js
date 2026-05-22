const APP_META = {
  name: "CRM VENTA-BOOKING",
  spreadsheetId: "REPLACE_WITH_SHEET_ID",
  version: "1.0.0"
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui
    .createMenu("CRM Venta Booking")
    .addItem("Auditar contactos (IA)", "codexAbrirAuditoriaContactosDialog")
    .addItem("Autocompletado de celdas (IA)", "codexAbrirAutocompletadoDialog")
    .addToUi();

  // Solo cambios visuales al abrir. La auditoria se lanza unicamente con boton/menu.
  try {
    codexAplicarDisenoVisualAlAbrir();
  } catch (_err) {
    // Evitar bloquear apertura por problemas auxiliares.
  }
}

function onEdit(e) {
  // Si cambia un email, dejamos estado en MAL hasta la siguiente auditoria manual.
  try {
    codexMarcarPendienteSiCambiaEmail_(e);
    if (e && e.range && e.range.getRow() >= 2) {
      const sheet = e.range.getSheet();
      const cfg = getSheetEmailConfig_(sheet);
      const editedCol = e.range.getColumn();
      // Solo forzamos refresco visual cuando se toca EMAIL o CORREO REVISADO.
      if (editedCol === cfg.emailCol || editedCol === cfg.iaCol) {
        codexAplicarColorFilaPorEstadoEnFila_(sheet, e.range.getRow());
      }
    }
  } catch (_err) {
    // No interrumpir la edicion del usuario por un fallo auxiliar.
  }
}

function codexPing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    ok: true,
    app: APP_META,
    activeSpreadsheetId: ss ? ss.getId() : null,
    activeSpreadsheetName: ss ? ss.getName() : null,
    generatedAt: new Date().toISOString()
  };
}
