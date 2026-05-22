const EMAIL_IA_CFG = {
  menuName: "CRM Venta Booking",
  panelSheetName: "PANEL CRM",
  header: "CORREO REVISADO",
  legacyHeader: "CORREO REVISADO IA",
  mergeHeader: "Merge status",
  statusGood: "BIEN",
  statusBad: "MAL",
  statusFixed: "CORREGIDO",
  validationValues: ["BIEN", "MAL", "CORREGIDO"],
  batchSize: 260,
  uiStepSize: 180,
  autocompleteStepSize: 14,
  maxAutocompleteWebLookupsPerSheet: 6,
  maxSecondsPerRun: 260,
  maxWebLookupsPerSheet: 3,
  fastOpenBudgetMs: 5000,
  autocompleteNotFoundText: "IA NO ENCUENTRA",
  autocompleteIgnoreHeaders: [
    "merge status",
    "correo revisado",
    "correo revisado ia",
    "correo enviado",
    "llamada"
  ],
  colors: {
    good: "#D9EAD3",
    bad: "#F4CCCC",
    fixed: "#D9E2F3"
  }
};

function codexAplicarDisenoVisualAlAbrir() {
  const started = Date.now();
  const triggerInfo = codexDesactivarTriggersAuditoriaAutomatica_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const active = ss ? ss.getActiveSheet() : null;
  let activeSheetConfigured = "";
  let timedOut = false;

  if (active && (Date.now() - started) < EMAIL_IA_CFG.fastOpenBudgetMs) {
    const cfg = getSheetEmailConfig_(active);
    if (cfg.emailCol && cfg.iaCol) {
      if (cfg.iaCol && !hasCodexColorRules_(active, cfg.iaCol)) {
        applyColorRules_(active, cfg.iaCol);
      }
      activeSheetConfigured = active.getName();
    }
  } else {
    timedOut = true;
  }

  const elapsed = Date.now() - started;
  return {
    ok: true,
    manualAuditOnly: true,
    removedAutoTriggers: triggerInfo.removed,
    activeSheetConfigured: activeSheetConfigured,
    elapsedMs: elapsed,
    targetMs: EMAIL_IA_CFG.fastOpenBudgetMs,
    withinTarget: elapsed <= EMAIL_IA_CFG.fastOpenBudgetMs,
    timedOut: timedOut
  };
}

function codexDesactivarTriggersAuditoriaAutomatica_() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "codexRevisarEmailsBatch") {
      ScriptApp.deleteTrigger(t);
      removed += 1;
    }
  });
  return { removed: removed };
}

function codexConfigurarRevisionEmailsIA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = [];

  ss.getSheets().forEach((sheet) => {
    const started = Date.now();
    let cfg = getSheetEmailConfig_(sheet);
    if (!cfg.emailCol) {
      report.push({ sheet: sheet.getName(), status: "SKIP", reason: "Sin columna EMAIL" });
      return;
    }

    cfg = ensureReviewAndMergeColumns_(sheet, cfg, { fullSetup: true });
    applyColorRules_(sheet, cfg.iaCol);

    report.push({
      sheet: sheet.getName(),
      status: "OK",
      emailCol: cfg.emailCol,
      callCol: cfg.callCol,
      reviewCol: cfg.iaCol,
      mergeCol: cfg.mergeCol,
      elapsedMs: Date.now() - started
    });
  });

  codexActualizarPanelEstadoContactos();
  return { ok: true, configuredSheets: report.length, details: report };
}

function codexAsegurarMergeStatusEnTodasHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];

  ss.getSheets().forEach((sheet) => {
    let cfg = getSheetEmailConfig_(sheet);
    if (!cfg.emailCol) return;

    const beforeMerge = cfg.mergeCol;
    cfg = ensureReviewAndMergeColumns_(sheet, cfg, { fullSetup: false });
    out.push({
      sheet: sheet.getName(),
      mergeCol: cfg.mergeCol,
      movedToLastColumn: !!beforeMerge && beforeMerge !== cfg.mergeCol
    });
  });

  return { ok: true, details: out };
}

function ensureReviewAndMergeColumns_(sheet, cfg, options) {
  const opts = options || {};
  let out = cfg || getSheetEmailConfig_(sheet);

  if (!out.iaCol) {
    const insertAfter = out.callCol || out.emailCol;
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue(EMAIL_IA_CFG.header);
    out = getSheetEmailConfig_(sheet);
  }
  sheet.getRange(1, out.iaCol).setValue(EMAIL_IA_CFG.header);

  if (!out.mergeCol) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    const mergeCol = sheet.getLastColumn();
    sheet.getRange(1, mergeCol).setValue(EMAIL_IA_CFG.mergeHeader);
    out = getSheetEmailConfig_(sheet);
  } else {
    sheet.getRange(1, out.mergeCol).setValue(EMAIL_IA_CFG.mergeHeader);
    if (out.mergeCol !== sheet.getLastColumn()) {
      moveColumnToEnd_(sheet, out.mergeCol);
      out = getSheetEmailConfig_(sheet);
    }
  }

  if (opts.fullSetup) {
    applyStatusValidation_(sheet, out.iaCol);
    normalizeStatusColumn_(sheet, out.iaCol);
  }

  return out;
}

function applyStatusValidation_(sheet, iaCol) {
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const validationRows = Math.max(maxRows - 1, 1);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(EMAIL_IA_CFG.validationValues, true)
    .setAllowInvalid(false)
    .setHelpText("Estado de revision del email")
    .build();
  sheet.getRange(2, iaCol, validationRows, 1).setDataValidation(validation);
}

function normalizeStatusColumn_(sheet, iaCol) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  if (lastRow < 2) return;
  const current = sheet.getRange(2, iaCol, lastRow - 1, 1).getDisplayValues();
  const normalized = current.map((r) => {
    const raw = String(r[0] || "").trim();
    if (!raw) return [""];
    return [normalizeLegacyStatus_(raw)];
  });
  if (normalized.length) {
    sheet.getRange(2, iaCol, normalized.length, 1).setValues(normalized);
  }
}

function moveColumnToEnd_(sheet, colIndex) {
  if (!colIndex || colIndex === sheet.getLastColumn()) return;
  const totalRows = sheet.getMaxRows();
  const source = sheet.getRange(1, colIndex, totalRows, 1);
  sheet.insertColumnAfter(sheet.getLastColumn());
  source.copyTo(sheet.getRange(1, sheet.getLastColumn(), totalRows, 1));
  sheet.deleteColumn(colIndex);
}

function codexRevisarEmailsBatch() {
  const removed = codexDesactivarTriggersAuditoriaAutomatica_();
  return {
    ok: true,
    manualAuditOnly: true,
    message: "Modo manual activo: usa el boton 'Auditar contactos (IA)'.",
    removedAutoTriggers: removed.removed,
    processedRows: 0,
    changedCells: 0,
    correctedEmails: 0,
    webLookups: 0,
    details: []
  };
}

function codexProgramarRevisionCada15Minutos() {
  const removed = codexDesactivarTriggersAuditoriaAutomatica_();
  return {
    ok: true,
    manualAuditOnly: true,
    message: "La auditoria automatica queda desactivada. Usa el boton de auditoria manual.",
    removedAutoTriggers: removed.removed
  };
}

function codexAbrirAuditoriaContactosDialog() {
  const html = HtmlService.createHtmlOutputFromFile("UI_AUDIT_PANEL")
    .setWidth(860)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, "Auditar contactos - CRM Venta Booking");
}

function codexAbrirAutocompletadoDialog() {
  const html = HtmlService.createHtmlOutputFromFile("UI_AUTOCOMPLETE_PANEL")
    .setWidth(860)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, "Autocompletado IA - CRM Venta Booking");
}

function buildWebBudgetBySheet_(sheetNames, perSheetBudget) {
  const out = {};
  const budget = Math.max(0, Number(perSheetBudget || 0));
  (sheetNames || []).forEach((name) => {
    out[name] = budget;
  });
  return out;
}

function getRemainingWebBudget_(state, sheetName, fallbackBudget) {
  if (!state.webBudgetBySheet) {
    state.webBudgetBySheet = buildWebBudgetBySheet_(state.sheetNames || [], fallbackBudget);
  }
  if (typeof state.webBudgetBySheet[sheetName] === "undefined") {
    state.webBudgetBySheet[sheetName] = Math.max(0, Number(fallbackBudget || 0));
  }
  return Math.max(0, Number(state.webBudgetBySheet[sheetName] || 0));
}

function codexUiIniciarAuditoriaContactos() {
  codexConfigurarRevisionEmailsIA();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets()
    .map((s) => s.getName())
    .filter((name) => {
      const sh = ss.getSheetByName(name);
      const cfg = getSheetEmailConfig_(sh);
      return !!cfg.emailCol;
    });

  const totalRows = sheets.reduce((acc, name) => {
    const sh = ss.getSheetByName(name);
    return acc + Math.max(sh.getLastRow() - 1, 0);
  }, 0);

  const state = {
    startedAt: new Date().toISOString(),
    sheetNames: sheets,
    sheetIndex: 0,
    rowCursor: 2,
    totalRows: totalRows,
    doneRows: 0,
    done: totalRows === 0,
    webBudgetBySheet: buildWebBudgetBySheet_(sheets, EMAIL_IA_CFG.maxWebLookupsPerSheet),
    stats: {
      changedCells: 0,
      correctedEmails: 0,
      webLookups: 0,
      good: 0,
      bad: 0,
      fixed: 0
    },
    recentEvents: []
  };

  PropertiesService.getScriptProperties().setProperty("EMAIL_AUDIT_UI_STATE", JSON.stringify(state));

  return {
    ok: true,
    totalRows: state.totalRows,
    doneRows: state.doneRows,
    percent: state.totalRows ? 0 : 100,
    done: state.done,
    currentSheet: state.sheetNames[0] || "",
    stats: state.stats,
    recentEvents: state.recentEvents
  };
}

function codexUiAuditarContactosPaso() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("EMAIL_AUDIT_UI_STATE");
  if (!raw) {
    return { ok: false, error: "No hay auditoria iniciada" };
  }

  const state = JSON.parse(raw);
  if (state.done) {
    return buildUiAuditResponse_(state, "Completado");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = CacheService.getScriptCache();
  const started = Date.now();
  let message = "";

  while (!state.done && (Date.now() - started) < 20000) {
    if (state.sheetIndex >= state.sheetNames.length) {
      state.done = true;
      break;
    }

    const sheetName = state.sheetNames[state.sheetIndex];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    const cfg = getSheetEmailConfig_(sheet);
    if (!cfg.emailCol || !cfg.iaCol) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    if (state.rowCursor > lastRow) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    const chunk = Math.min(EMAIL_IA_CFG.uiStepSize, lastRow - state.rowCursor + 1);
    if (chunk <= 0) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    const chunkStart = state.rowCursor;
    const chunkEnd = state.rowCursor + chunk - 1;
    const chunkBudget = getRemainingWebBudget_(state, sheetName, EMAIL_IA_CFG.maxWebLookupsPerSheet);
    const chunkResult = processAuditChunk_(sheet, cfg, state.rowCursor, chunk, cache, chunkBudget);
    state.webBudgetBySheet[sheetName] = Math.max(0, chunkBudget - Number(chunkResult.webLookups || 0));
    SpreadsheetApp.flush();

    state.doneRows += chunk;
    state.stats.changedCells += chunkResult.changedCells;
    state.stats.correctedEmails += chunkResult.correctedEmails;
    state.stats.webLookups += chunkResult.webLookups;
    state.stats.good += chunkResult.good;
    state.stats.bad += chunkResult.bad;
    state.stats.fixed += chunkResult.fixed;

    const events = (chunkResult.events || []).map((e) => {
      const emailInfo = e.correctedEmail ? (" -> " + e.correctedEmail) : "";
      return sheetName + " | Fila " + e.row + " | " + e.status + emailInfo;
    });
    state.recentEvents = events.concat(state.recentEvents || []).slice(0, 12);

    message =
      sheetName +
      " | Filas " +
      chunkStart +
      "-" +
      chunkEnd +
      " | BIEN " +
      chunkResult.good +
      " | MAL " +
      chunkResult.bad +
      " | CORREGIDO " +
      chunkResult.fixed;

    state.rowCursor += chunk;
    if (state.rowCursor > lastRow) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
    }

    if (state.sheetIndex >= state.sheetNames.length) {
      state.done = true;
      break;
    }

    // Un paso por llamada para que el UI pinte el porcentaje de forma fluida.
    break;
  }

  if (state.done) {
    codexActualizarPanelEstadoContactos();
  }

  props.setProperty("EMAIL_AUDIT_UI_STATE", JSON.stringify(state));
  return buildUiAuditResponse_(state, message || "Procesando...");
}

function codexUiObtenerEstadoAuditoria() {
  const raw = PropertiesService.getScriptProperties().getProperty("EMAIL_AUDIT_UI_STATE");
  if (!raw) return { ok: false, error: "Sin estado" };
  const state = JSON.parse(raw);
  return buildUiAuditResponse_(state, state.done ? "Completado" : "En progreso");
}

function buildUiAuditResponse_(state, message) {
  const total = Number(state.totalRows || 0);
  const doneRows = Math.min(Number(state.doneRows || 0), total || Number(state.doneRows || 0));
  const percent = total > 0 ? Math.min(100, Math.floor((doneRows / total) * 100)) : 100;
  const currentSheet = state.sheetNames && state.sheetNames[state.sheetIndex] ? state.sheetNames[state.sheetIndex] : "";

  return {
    ok: true,
    done: !!state.done,
    percent: percent,
    doneRows: doneRows,
    totalRows: total,
    currentSheet: currentSheet,
    message: message || "",
    stats: state.stats || {},
    recentEvents: state.recentEvents || []
  };
}

function codexUiIniciarAutocompletadoCeldas() {
  codexConfigurarRevisionEmailsIA();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets()
    .map((s) => s.getName())
    .filter((name) => {
      const sh = ss.getSheetByName(name);
      const cfg = getSheetEmailConfig_(sh);
      return !!cfg.emailCol;
    });

  const totalRows = sheets.reduce((acc, name) => {
    const sh = ss.getSheetByName(name);
    return acc + Math.max(sh.getLastRow() - 1, 0);
  }, 0);

  const state = {
    startedAt: new Date().toISOString(),
    sheetNames: sheets,
    sheetIndex: 0,
    rowCursor: 2,
    totalRows: totalRows,
    doneRows: 0,
    done: totalRows === 0,
    webBudgetBySheet: buildWebBudgetBySheet_(sheets, EMAIL_IA_CFG.maxAutocompleteWebLookupsPerSheet),
    stats: {
      cellsFilled: 0,
      iaNoEncuentra: 0,
      webLookups: 0
    },
    recentEvents: []
  };

  PropertiesService.getScriptProperties().setProperty("EMAIL_AUTOCOMPLETE_UI_STATE", JSON.stringify(state));

  return {
    ok: true,
    totalRows: state.totalRows,
    doneRows: state.doneRows,
    percent: state.totalRows ? 0 : 100,
    done: state.done,
    currentSheet: state.sheetNames[0] || "",
    stats: state.stats,
    recentEvents: state.recentEvents
  };
}

function codexUiAutocompletarPaso() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("EMAIL_AUTOCOMPLETE_UI_STATE");
  if (!raw) return { ok: false, error: "No hay autocompletado iniciado" };

  const state = JSON.parse(raw);
  if (state.done) return buildUiAutocompleteResponse_(state, "Completado");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = CacheService.getScriptCache();
  const started = Date.now();
  let message = "";

  while (!state.done && (Date.now() - started) < 20000) {
    if (state.sheetIndex >= state.sheetNames.length) {
      state.done = true;
      break;
    }

    const sheetName = state.sheetNames[state.sheetIndex];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    let cfg = getSheetEmailConfig_(sheet);
    if (!cfg.emailCol) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }
    cfg = ensureReviewAndMergeColumns_(sheet, cfg, { fullSetup: false });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    if (state.rowCursor > lastRow) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    const chunk = Math.min(EMAIL_IA_CFG.autocompleteStepSize, lastRow - state.rowCursor + 1);
    if (chunk <= 0) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
      continue;
    }

    const chunkStart = state.rowCursor;
    const chunkEnd = state.rowCursor + chunk - 1;
    const chunkBudget = getRemainingWebBudget_(state, sheetName, EMAIL_IA_CFG.maxAutocompleteWebLookupsPerSheet);
    const chunkResult = processAutocompleteChunk_(
      sheet,
      cfg,
      state.rowCursor,
      chunk,
      cache,
      chunkBudget
    );
    state.webBudgetBySheet[sheetName] = Math.max(0, chunkBudget - Number(chunkResult.webLookups || 0));
    SpreadsheetApp.flush();

    state.doneRows += chunk;
    state.stats.cellsFilled += chunkResult.cellsFilled;
    state.stats.iaNoEncuentra += chunkResult.notFound;
    state.stats.webLookups += chunkResult.webLookups;

    const events = (chunkResult.events || []).map((e) => {
      return sheetName + " | Fila " + e.row + " | " + e.header + " = " + e.value;
    });
    state.recentEvents = events.concat(state.recentEvents || []).slice(0, 12);

    message =
      sheetName +
      " | Filas " +
      chunkStart +
      "-" +
      chunkEnd +
      " | Celdas " +
      chunkResult.cellsFilled +
      " | IA NO ENCUENTRA " +
      chunkResult.notFound;

    state.rowCursor += chunk;
    if (state.rowCursor > lastRow) {
      state.sheetIndex += 1;
      state.rowCursor = 2;
    }

    if (state.sheetIndex >= state.sheetNames.length) {
      state.done = true;
      break;
    }
    break;
  }

  props.setProperty("EMAIL_AUTOCOMPLETE_UI_STATE", JSON.stringify(state));
  return buildUiAutocompleteResponse_(state, message || "Procesando...");
}

function buildUiAutocompleteResponse_(state, message) {
  const total = Number(state.totalRows || 0);
  const doneRows = Math.min(Number(state.doneRows || 0), total || Number(state.doneRows || 0));
  const percent = total > 0 ? Math.min(100, Math.floor((doneRows / total) * 100)) : 100;
  const currentSheet = state.sheetNames && state.sheetNames[state.sheetIndex] ? state.sheetNames[state.sheetIndex] : "";

  return {
    ok: true,
    done: !!state.done,
    percent: percent,
    doneRows: doneRows,
    totalRows: total,
    currentSheet: currentSheet,
    message: message || "",
    stats: state.stats || {},
    recentEvents: state.recentEvents || []
  };
}

function codexUiObtenerEstadoAutocompletado() {
  const raw = PropertiesService.getScriptProperties().getProperty("EMAIL_AUTOCOMPLETE_UI_STATE");
  if (!raw) return { ok: false, error: "Sin estado" };
  const state = JSON.parse(raw);
  return buildUiAutocompleteResponse_(state, state.done ? "Completado" : "En progreso");
}

function processAutocompleteChunk_(sheet, cfg, startRow, chunkSize, cache, maxWebLookups) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1 || chunkSize <= 0) {
    return { cellsFilled: 0, notFound: 0, webLookups: 0, events: [] };
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map((x) => String(x || "").trim());
  const values = sheet.getRange(startRow, 1, chunkSize, lastCol).getDisplayValues();
  const formulas = sheet.getRange(startRow, 1, chunkSize, lastCol).getFormulas();

  const excludeCols = {};
  if (cfg.mergeCol) excludeCols[cfg.mergeCol] = true;
  if (cfg.iaCol) excludeCols[cfg.iaCol] = true;

  let webBudget = Math.max(0, Number(maxWebLookups || 0));
  let webLookups = 0;
  let cellsFilled = 0;
  let notFound = 0;
  const updates = [];
  const events = [];

  for (let r = 0; r < chunkSize; r++) {
    const rowNumber = startRow + r;
    const name = cfg.nameCol ? String(values[r][cfg.nameCol - 1] || "") : "";
    const municipio = cfg.municipioCol ? String(values[r][cfg.municipioCol - 1] || "") : "";
    const provincia = cfg.provinciaCol ? String(values[r][cfg.provinciaCol - 1] || "") : "";
    const query = buildAutocompleteQuery_(name, municipio, provincia);

    let webText = "";
    const rowHasBlank = values[r].some((cell, idx) => {
      const col = idx + 1;
      if (excludeCols[col]) return false;
      if (formulas[r][idx]) return false;
      const header = String(headers[idx] || "").trim().toLowerCase();
      if (!header || EMAIL_IA_CFG.autocompleteIgnoreHeaders.indexOf(header) !== -1) return false;
      return String(cell || "").trim() === "";
    });

    if (rowHasBlank && query && webBudget > 0) {
      webText = fetchWebTextForQuery_(query, cache);
      webBudget -= 1;
      webLookups += 1;
    }

    for (let c = 0; c < lastCol; c++) {
      const col = c + 1;
      if (excludeCols[col]) continue;
      if (formulas[r][c]) continue;

      const rawHeader = String(headers[c] || "").trim();
      const header = rawHeader.toLowerCase();
      if (!header || EMAIL_IA_CFG.autocompleteIgnoreHeaders.indexOf(header) !== -1) continue;

      const cur = String(values[r][c] || "").trim();
      if (cur) continue;

      let next = "";
      if (webText) {
        next = guessAutocompleteValue_(rawHeader, webText, values[r], cfg);
      }
      if (!next) {
        next = EMAIL_IA_CFG.autocompleteNotFoundText;
        notFound += 1;
      }

      updates.push({ row: rowNumber, col: col, value: next });
      cellsFilled += 1;
      if (events.length < 20) {
        events.push({ row: rowNumber, header: rawHeader || ("COL " + col), value: next });
      }
    }

  }

  applyCellUpdatesBatched_(sheet, updates);

  return {
    cellsFilled: cellsFilled,
    notFound: notFound,
    webLookups: webLookups,
    events: events
  };
}

function applyCellUpdatesBatched_(sheet, updates) {
  if (!updates || !updates.length) return;

  const byRow = {};
  updates.forEach((u) => {
    if (!u || !u.row || !u.col) return;
    const key = String(u.row);
    if (!byRow[key]) byRow[key] = [];
    byRow[key].push(u);
  });

  Object.keys(byRow).forEach((rowKey) => {
    const row = Number(rowKey);
    const cells = byRow[rowKey].sort((a, b) => Number(a.col || 0) - Number(b.col || 0));
    let idx = 0;

    while (idx < cells.length) {
      let end = idx;
      while (end + 1 < cells.length && Number(cells[end + 1].col) === Number(cells[end].col) + 1) {
        end += 1;
      }

      const startCol = Number(cells[idx].col);
      const block = cells.slice(idx, end + 1).map((c) => c.value);
      sheet.getRange(row, startCol, 1, block.length).setValues([block]);
      idx = end + 1;
    }
  });
}

function buildAutocompleteQuery_(name, municipio, provincia) {
  const parts = [name, municipio, provincia, "contacto oficial"]
    .map((x) => String(x || "").trim())
    .filter((x) => !!x);
  if (!parts.length) return "";
  return parts.join(" ");
}

function fetchWebTextForQuery_(query, cache) {
  const key = "acweb:" + query.toLowerCase();
  const cached = cache.get(key);
  if (cached) return cached === "-" ? "" : cached;

  try {
    const url = "https://duckduckgo.com/html/?q=" + encodeURIComponent(query);
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      method: "get",
      headers: { "User-Agent": "Mozilla/5.0" }
    });
    if (resp.getResponseCode() !== 200) {
      cache.put(key, "-", 21600);
      return "";
    }
    const html = resp.getContentText() || "";
    cache.put(key, html || "-", 21600);
    return html;
  } catch (_err) {
    return "";
  }
}

function guessAutocompleteValue_(header, webText, rowValues, cfg) {
  const h = String(header || "").trim().toLowerCase();
  if (!webText) return "";

  if (h.indexOf("email") !== -1) {
    const email = extractEmailCandidate_(webText);
    if (email && isEmailValidWithDns_(email, CacheService.getScriptCache())) return email;
  }

  if (h.indexOf("telefono") !== -1 || h.indexOf("tel") !== -1 || h.indexOf("movil") !== -1) {
    const phoneMatch = webText.match(/(?:\+34[\s-]?)?(?:\d[\s-]?){9,12}\d/);
    if (phoneMatch) return String(phoneMatch[0] || "").replace(/\s+/g, " ").trim();
  }

  if (h.indexOf("web") !== -1 || h.indexOf("sitio") !== -1 || h.indexOf("url") !== -1) {
    const urlMatch = webText.match(/https?:\/\/[^\s"'<>()]+/i);
    if (urlMatch) return String(urlMatch[0] || "").trim();
  }

  if (h.indexOf("nombre contacto") !== -1 || h.indexOf("contacto") !== -1) {
    const emailCol = cfg.emailCol ? Number(cfg.emailCol) : 0;
    if (emailCol > 0 && rowValues[emailCol - 1]) {
      const email = extractEmailCandidate_(rowValues[emailCol - 1]);
      if (email && email.indexOf("@") !== -1) {
        const local = email.split("@")[0].replace(/[._-]+/g, " ").trim();
        if (local) return local.toUpperCase();
      }
    }
  }

  return "";
}

function codexMostrarPanelEstadoContactos() {
  codexActualizarPanelEstadoContactos();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName(EMAIL_IA_CFG.panelSheetName);
  if (panel) ss.setActiveSheet(panel);
  return { ok: true };
}

function codexActualizarPanelEstadoContactos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = collectStatusSummary_();

  let panel = ss.getSheetByName(EMAIL_IA_CFG.panelSheetName);
  if (!panel) panel = ss.insertSheet(EMAIL_IA_CFG.panelSheetName);

  panel.clear({ contentsOnly: true });
  panel.clearConditionalFormatRules();

  panel.getRange("A1:F1").merge();
  panel.getRange("A1").setValue("CRM Venta Booking - Estado de contactos");

  panel.getRange("A3").setValue("TOTAL CONTACTOS");
  panel.getRange("A4").setValue("BIEN");
  panel.getRange("A5").setValue("MAL");
  panel.getRange("A6").setValue("CORREGIDO");
  panel.getRange("A7").setValue("ULTIMA ACTUALIZACION");

  panel.getRange("B3").setValue(summary.total);
  panel.getRange("B4").setValue(summary.good);
  panel.getRange("B5").setValue(summary.bad);
  panel.getRange("B6").setValue(summary.fixed);
  panel.getRange("B7").setValue(new Date());

  panel.getRange("D3").setValue("PORCENTAJE BIEN");
  panel.getRange("D4").setValue(summary.total ? Math.round((summary.good / summary.total) * 10000) / 100 : 0);
  panel.getRange("D5").setValue("PORCENTAJE MAL");
  panel.getRange("D6").setValue(summary.total ? Math.round((summary.bad / summary.total) * 10000) / 100 : 0);
  panel.getRange("D7").setValue("PORCENTAJE CORREGIDO");
  panel.getRange("E7").setValue(summary.total ? Math.round((summary.fixed / summary.total) * 10000) / 100 : 0);

  panel.getRange("A9:F9").setValues([["Pestana", "Contactos", "BIEN", "MAL", "CORREGIDO", "% BIEN"]]);

  if (summary.rows.length) {
    const table = summary.rows.map((r) => [
      r.sheet,
      r.total,
      r.good,
      r.bad,
      r.fixed,
      r.total ? Math.round((r.good / r.total) * 10000) / 100 : 0
    ]);
    panel.getRange(10, 1, table.length, 6).setValues(table);
  }

  panel.setFrozenRows(9);
  panel.setColumnWidths(1, 1, 230);
  panel.setColumnWidths(2, 4, 120);
  panel.setColumnWidths(6, 1, 100);

  panel.getRange("A1").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#FFF2CC");
  panel.getRange("A3:A7").setFontWeight("bold");
  panel.getRange("B4").setBackground(EMAIL_IA_CFG.colors.good);
  panel.getRange("B5").setBackground(EMAIL_IA_CFG.colors.bad);
  panel.getRange("B6").setBackground(EMAIL_IA_CFG.colors.fixed);
  panel.getRange("A9:F9").setFontWeight("bold").setBackground("#D9EAD3");
  panel.getRange("D4:E7").setNumberFormat("0.00");
  panel.getRange("F10:F").setNumberFormat("0.00");

  return {
    ok: true,
    totals: {
      total: summary.total,
      good: summary.good,
      bad: summary.bad,
      fixed: summary.fixed
    },
    rows: summary.rows.length
  };
}

function collectStatusSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = [];
  let total = 0;
  let good = 0;
  let bad = 0;
  let fixed = 0;

  ss.getSheets().forEach((sheet) => {
    const cfg = getSheetEmailConfig_(sheet);
    if (!cfg.emailCol || !cfg.iaCol) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      rows.push({ sheet: sheet.getName(), total: 0, good: 0, bad: 0, fixed: 0 });
      return;
    }

    const statuses = sheet.getRange(2, cfg.iaCol, lastRow - 1, 1).getDisplayValues();
    let sGood = 0;
    let sBad = 0;
    let sFixed = 0;

    statuses.forEach((r) => {
      const v = normalizeLegacyStatus_(r[0]);
      if (v === EMAIL_IA_CFG.statusGood) sGood++;
      else if (v === EMAIL_IA_CFG.statusFixed) sFixed++;
      else sBad++;
    });

    const sTotal = lastRow - 1;
    total += sTotal;
    good += sGood;
    bad += sBad;
    fixed += sFixed;

    rows.push({
      sheet: sheet.getName(),
      total: sTotal,
      good: sGood,
      bad: sBad,
      fixed: sFixed
    });
  });

  rows.sort((a, b) => b.total - a.total);
  return { total: total, good: good, bad: bad, fixed: fixed, rows: rows };
}

function codexMarcarPendienteSiCambiaEmail_(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const sheet = e.range.getSheet();
  const cfg = getSheetEmailConfig_(sheet);
  if (!cfg.emailCol || !cfg.iaCol) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  if (e.range.getColumn() !== cfg.emailCol) return;

  sheet.getRange(e.range.getRow(), cfg.iaCol).setValue(EMAIL_IA_CFG.statusBad);
}

function codexAplicarColorFilaPorEstadoEnFila_(_sheet, _row) {
  // El formato condicional ya actualiza el color automaticamente.
}

function processAuditChunk_(sheet, cfg, startRow, chunkSize, cache, maxWebLookups) {
  const emails = sheet.getRange(startRow, cfg.emailCol, chunkSize, 1).getDisplayValues();
  const current = sheet.getRange(startRow, cfg.iaCol, chunkSize, 1).getDisplayValues();

  const names = cfg.nameCol ? sheet.getRange(startRow, cfg.nameCol, chunkSize, 1).getDisplayValues() : [];
  const municipios = cfg.municipioCol ? sheet.getRange(startRow, cfg.municipioCol, chunkSize, 1).getDisplayValues() : [];
  const provincias = cfg.provinciaCol ? sheet.getRange(startRow, cfg.provinciaCol, chunkSize, 1).getDisplayValues() : [];

  const out = [];
  const nextEmails = emails.map((r) => [String(r[0] || "").trim()]);
  let changed = 0;
  let corrected = 0;
  let webLookups = 0;
  let good = 0;
  let bad = 0;
  let fixed = 0;
  let budget = Math.max(0, Number(maxWebLookups || 0));
  const events = [];

  for (let i = 0; i < chunkSize; i++) {
    const cur = normalizeLegacyStatus_(current[i][0]);
    const name = names[i] ? String(names[i][0] || "") : "";
    const municipio = municipios[i] ? String(municipios[i][0] || "") : "";
    const provincia = provincias[i] ? String(provincias[i][0] || "") : "";

    const result = evaluateEmailRow_(emails[i][0], name, municipio, provincia, cache, budget);

    let next = result.status;
    if (EMAIL_IA_CFG.validationValues.indexOf(next) === -1) {
      next = EMAIL_IA_CFG.statusBad;
    }

    if (result.usedWebLookup) {
      budget -= 1;
      webLookups += 1;
      if (budget < 0) budget = 0;
    }

    if (result.correctedEmail) {
      nextEmails[i] = [result.correctedEmail];
      corrected++;
    }

    if (next !== cur) changed++;
    out.push([next]);

    if (next === EMAIL_IA_CFG.statusGood) good++;
    else if (next === EMAIL_IA_CFG.statusFixed) fixed++;
    else bad++;

    if (events.length < 25) {
      events.push({
        row: startRow + i,
        status: next,
        correctedEmail: result.correctedEmail || ""
      });
    }
  }

  if (out.length) {
    sheet.getRange(startRow, cfg.iaCol, chunkSize, 1).setValues(out);
  }

  if (corrected > 0) {
    sheet.getRange(startRow, cfg.emailCol, chunkSize, 1).setValues(nextEmails);
  }

  return {
    processedRows: chunkSize,
    changedCells: changed,
    correctedEmails: corrected,
    webLookups: webLookups,
    good: good,
    bad: bad,
    fixed: fixed,
    events: events
  };
}

function evaluateEmailRow_(rawEmail, name, municipio, provincia, cache, webBudget) {
  const email = extractEmailCandidate_(rawEmail);
  const query = buildContactQuery_(name, municipio, provincia);

  if (email && isEmailValidWithDns_(email, cache)) {
    if (webBudget > 0 && query) {
      const appears = emailAppearsOnWeb_(email, query, cache);
      if (appears) {
        return { status: EMAIL_IA_CFG.statusGood, correctedEmail: "", usedWebLookup: true };
      }

      const foundFromWeb = findEmailOnWeb_(query, cache);
      if (foundFromWeb && isEmailValidWithDns_(foundFromWeb, cache)) {
        if (foundFromWeb.toLowerCase() === email.toLowerCase()) {
          return { status: EMAIL_IA_CFG.statusGood, correctedEmail: "", usedWebLookup: true };
        }
        return { status: EMAIL_IA_CFG.statusFixed, correctedEmail: foundFromWeb, usedWebLookup: true };
      }
      return { status: EMAIL_IA_CFG.statusBad, correctedEmail: "", usedWebLookup: true };
    }

    return { status: EMAIL_IA_CFG.statusGood, correctedEmail: "", usedWebLookup: false };
  }

  if (webBudget > 0) {
    if (query) {
      const found = findEmailOnWeb_(query, cache);
      if (found && isEmailValidWithDns_(found, cache)) {
        if (email && found.toLowerCase() === email.toLowerCase()) {
          return { status: EMAIL_IA_CFG.statusGood, correctedEmail: "", usedWebLookup: true };
        }
        return { status: EMAIL_IA_CFG.statusFixed, correctedEmail: found, usedWebLookup: true };
      }
      return { status: EMAIL_IA_CFG.statusBad, correctedEmail: "", usedWebLookup: true };
    }
  }

  return { status: EMAIL_IA_CFG.statusBad, correctedEmail: "", usedWebLookup: false };
}

function buildContactQuery_(name, municipio, provincia) {
  const parts = [name, municipio, provincia, "email contacto"]
    .map((x) => String(x || "").trim())
    .filter((x) => !!x);

  if (!parts.length) return "";
  return parts.join(" ");
}

function emailAppearsOnWeb_(email, query, cache) {
  const key = "webok:" + email.toLowerCase() + "|" + query.toLowerCase();
  const cached = cache.get(key);
  if (cached === "1") return true;
  if (cached === "0") return false;

  try {
    const q = query + " " + email;
    const url = "https://duckduckgo.com/html/?q=" + encodeURIComponent(q);
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      method: "get",
      headers: { "User-Agent": "Mozilla/5.0" }
    });

    if (resp.getResponseCode() !== 200) {
      cache.put(key, "0", 21600);
      return false;
    }

    const html = String(resp.getContentText() || "").toLowerCase();
    const found = html.indexOf(email.toLowerCase()) !== -1;
    cache.put(key, found ? "1" : "0", 21600);
    return found;
  } catch (_err) {
    return false;
  }
}

function findEmailOnWeb_(query, cache) {
  const key = "webq:" + query.toLowerCase();
  const cached = cache.get(key);
  if (cached) return cached === "-" ? "" : cached;

  try {
    const url = "https://duckduckgo.com/html/?q=" + encodeURIComponent(query);
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      method: "get",
      headers: { "User-Agent": "Mozilla/5.0" }
    });

    if (resp.getResponseCode() !== 200) {
      cache.put(key, "-", 21600);
      return "";
    }

    const html = resp.getContentText() || "";
    const match = html.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/i);
    const found = match ? String(match[0] || "").trim().toLowerCase() : "";

    if (found) {
      cache.put(key, found, 21600);
      return found;
    }

    cache.put(key, "-", 21600);
    return "";
  } catch (_err) {
    return "";
  }
}

function getSheetEmailConfig_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map((h) => String(h || "").trim());

  const emailCol = findHeaderColumn_(headers, (h) => {
    const s = h.toLowerCase();
    return s.indexOf("email") !== -1 || s.indexOf("e-mail") !== -1;
  });

  const callCol = findHeaderColumn_(headers, (h) => h.toLowerCase().indexOf("llamada") !== -1);
  const iaCol = findHeaderColumn_(headers, (h) => {
    const s = h.toLowerCase().trim();
    return s === EMAIL_IA_CFG.header.toLowerCase() || s === EMAIL_IA_CFG.legacyHeader.toLowerCase();
  });

  const mergeCol = findHeaderColumn_(headers, (h) => h.toLowerCase().trim() === EMAIL_IA_CFG.mergeHeader.toLowerCase());
  const nameCol = findHeaderColumn_(headers, (h) => {
    const s = h.toLowerCase();
    return s.indexOf("nombre") !== -1 && s.indexOf("contacto") === -1;
  });
  const municipioCol = findHeaderColumn_(headers, (h) => {
    const s = h.toLowerCase();
    return s.indexOf("municipio") !== -1 || s.indexOf("poblacion") !== -1;
  });
  const provinciaCol = findHeaderColumn_(headers, (h) => h.toLowerCase().indexOf("provincia") !== -1);

  return {
    emailCol: emailCol,
    callCol: callCol,
    iaCol: iaCol,
    mergeCol: mergeCol,
    nameCol: nameCol,
    municipioCol: municipioCol,
    provinciaCol: provinciaCol
  };
}

function findHeaderColumn_(headers, predicateFn) {
  for (let i = 0; i < headers.length; i++) {
    if (predicateFn(headers[i])) return i + 1;
  }
  return 0;
}

function normalizeLegacyStatus_(value) {
  const v = String(value || "").trim().toUpperCase();
  if (v === EMAIL_IA_CFG.statusGood) return EMAIL_IA_CFG.statusGood;
  if (v === EMAIL_IA_CFG.statusBad) return EMAIL_IA_CFG.statusBad;
  if (v === EMAIL_IA_CFG.statusFixed) return EMAIL_IA_CFG.statusFixed;

  if (v === "REVISADO") return EMAIL_IA_CFG.statusGood;
  if (v === "NO") return EMAIL_IA_CFG.statusBad;
  if (v === "PTE") return EMAIL_IA_CFG.statusBad;
  if (v === "ERROR") return EMAIL_IA_CFG.statusBad;

  return EMAIL_IA_CFG.statusBad;
}

function extractEmailCandidate_(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw) return "";

  const low = raw.toLowerCase();
  const placeholders = ["no disponible", "sin email", "n/a", "na", "pendiente", "-", "none", "no email"];
  if (placeholders.indexOf(low) !== -1) return "";

  const parts = raw.split(/[\s,;|/]+/g);
  for (let i = 0; i < parts.length; i++) {
    const token = parts[i].trim().replace(/\.+$/g, "");
    if (token.indexOf("@") !== -1) return token.toLowerCase();
  }
  return "";
}

function checkEmailSyntax_(email) {
  if (!email || email.indexOf("@") === -1) return { ok: false, reason: "missing_at" };
  if ((email.match(/@/g) || []).length !== 1) return { ok: false, reason: "multi_at" };
  if (email.length > 254) return { ok: false, reason: "length" };

  const pattern = /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/;
  if (!pattern.test(email)) return { ok: false, reason: "regex" };

  const domain = email.split("@")[1].toLowerCase();
  const typoDomains = ["gamil.com", "gmai.com", "hotnail.com", "hotmal.com", "outlok.com", "yaho.com"];
  if (typoDomains.indexOf(domain) !== -1) return { ok: false, reason: "domain_typo" };

  return { ok: true, reason: "ok" };
}

function isEmailValidWithDns_(email, cache) {
  const syntax = checkEmailSyntax_(email);
  if (!syntax.ok) return false;

  const domain = email.split("@")[1].toLowerCase();
  const dns = checkDomainDns_(domain, cache);

  // Si DNS no responde, no damos por malo el contacto para evitar falsos negativos.
  return dns !== false;
}

function checkDomainDns_(domain, cache) {
  const key = "dns:" + domain;
  const fromCache = cache.get(key);
  if (fromCache === "1") return true;
  if (fromCache === "0") return false;

  try {
    const mxUrl = "https://dns.google/resolve?name=" + encodeURIComponent(domain) + "&type=MX";
    const mxResp = UrlFetchApp.fetch(mxUrl, { muteHttpExceptions: true });
    if (mxResp.getResponseCode() === 200) {
      const mxData = JSON.parse(mxResp.getContentText() || "{}");
      if (mxData.Status === 0 && mxData.Answer && mxData.Answer.length) {
        cache.put(key, "1", 21600);
        return true;
      }
    }

    const aUrl = "https://dns.google/resolve?name=" + encodeURIComponent(domain) + "&type=A";
    const aResp = UrlFetchApp.fetch(aUrl, { muteHttpExceptions: true });
    if (aResp.getResponseCode() === 200) {
      const aData = JSON.parse(aResp.getContentText() || "{}");
      if (aData.Status === 0 && aData.Answer && aData.Answer.length) {
        cache.put(key, "1", 21600);
        return true;
      }
      cache.put(key, "0", 21600);
      return false;
    }

    return null;
  } catch (_err) {
    return null;
  }
}

function applyColorRules_(sheet, iaCol) {
  const allRules = sheet.getConditionalFormatRules();
  const colLetter = columnToLetter_(iaCol);

  const formulas = {
    good: `=$${colLetter}2="${EMAIL_IA_CFG.statusGood}"`,
    bad: `=$${colLetter}2="${EMAIL_IA_CFG.statusBad}"`,
    fixed: `=$${colLetter}2="${EMAIL_IA_CFG.statusFixed}"`
  };

  const codexFormulas = [
    formulas.good,
    formulas.bad,
    formulas.fixed,
    `=$${colLetter}2="REVISADO"`,
    `=$${colLetter}2="NO"`,
    `=$${colLetter}2="PTE"`,
    `=$${colLetter}2="ERROR"`
  ].reduce((acc, f) => {
    const key = String(f || "").toLowerCase().replace(/\s+/g, "");
    if (key) acc[key] = true;
    return acc;
  }, {});

  const kept = allRules.filter((rule) => {
    const cond = rule.getBooleanCondition();
    if (!cond) return true;
    if (String(cond.getCriteriaType()) !== "CUSTOM_FORMULA") return true;
    const values = cond.getCriteriaValues() || [];
    if (!values.length) return true;
    const formula = String(values[0] || "").toLowerCase().replace(/\s+/g, "");
    return !codexFormulas[formula];
  });

  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const lastCol = Math.max(sheet.getLastColumn(), iaCol);
  const rowRange = sheet.getRange(2, 1, maxRows - 1, lastCol);

  const ruleGood = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulas.good)
    .setBackground(EMAIL_IA_CFG.colors.good)
    .setRanges([rowRange])
    .build();

  const ruleBad = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulas.bad)
    .setBackground(EMAIL_IA_CFG.colors.bad)
    .setRanges([rowRange])
    .build();

  const ruleFixed = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulas.fixed)
    .setBackground(EMAIL_IA_CFG.colors.fixed)
    .setRanges([rowRange])
    .build();

  sheet.setConditionalFormatRules(kept.concat([ruleGood, ruleBad, ruleFixed]));
}

function hasCodexColorRules_(sheet, iaCol) {
  const colLetter = columnToLetter_(iaCol);
  const expected = [
    `=$${colLetter}2="${EMAIL_IA_CFG.statusGood}"`,
    `=$${colLetter}2="${EMAIL_IA_CFG.statusBad}"`,
    `=$${colLetter}2="${EMAIL_IA_CFG.statusFixed}"`
  ].map((f) => String(f).toLowerCase().replace(/\s+/g, ""));

  const found = {};
  sheet.getConditionalFormatRules().forEach((rule) => {
    const cond = rule.getBooleanCondition();
    if (!cond) return;
    if (String(cond.getCriteriaType()) !== "CUSTOM_FORMULA") return;
    const values = cond.getCriteriaValues() || [];
    if (!values.length) return;
    const f = String(values[0] || "").toLowerCase().replace(/\s+/g, "");
    found[f] = true;
  });

  return expected.every((f) => !!found[f]);
}

function columnToLetter_(col) {
  let out = "";
  let c = Number(col || 0);
  while (c > 0) {
    const rem = (c - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    c = Math.floor((c - 1) / 26);
  }
  return out || "A";
}
