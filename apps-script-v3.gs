// ═══════════════════════════════════════════════════════════
//  ARBOLADO ZAPOPAN — Apps Script v3
//  Maneja: PDF a Drive + Dictamen + Reinspección + Búsqueda por folio
// ═══════════════════════════════════════════════════════════

var CONFIG = {
  carpeta:       "Arbolado_Zapopan",
  nombreSheet:   "Concentrado_Arbolado",
  hojaTab:       "Dictámenes",
  hojaReinsp:    "Reinspecciones"
};

// ═══════════════════════════════════════════════════════════
//  ENTRADA PRINCIPAL
// ═══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var tipo  = datos.tipo;

    if (tipo === "ping")          return ok({msg: "Conectado — Arbolado Zapopan v3"});
    if (tipo === "pdf")           { guardarPDF(datos); return ok({msg: "PDF guardado"}); }
    if (tipo === "dictamen")      { actualizarSheets(datos); return ok({msg: "Dictamen guardado"}); }
    if (tipo === "reinspeccion")  { guardarReinspeccion(datos); return ok({msg: "Reinspección guardada"}); }

    return ok({msg: "Tipo no reconocido: " + tipo});
  } catch (err) {
    return fail(err.message);
  }
}

function doGet(e) {
  // Búsqueda de folio para la app de reinspección
  if (e.parameter.tipo === "buscar" && e.parameter.folio) {
    return buscarFolioGet(e.parameter.folio);
  }
  return ok({msg: "Apps Script activo — Arbolado Zapopan v3"});
}

// ═══════════════════════════════════════════════════════════
//  BUSCAR FOLIO (GET — para la app de reinspección)
// ═══════════════════════════════════════════════════════════
function buscarFolioGet(folio) {
  try {
    var ss   = obtenerSheet();
    var hoja = ss.getSheetByName(CONFIG.hojaTab);
    if (!hoja) return ok({ok: false, msg: "Hoja no encontrada"});

    var valores = hoja.getDataRange().getValues();
    var arboles = [];
    var supervisor = "";

    // Buscar todas las filas con ese folio (col 0)
    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      if (String(fila[0]).trim().toUpperCase() === folio.toUpperCase()) {
        if (!supervisor && fila[3]) supervisor = fila[3];
        arboles.push({
          nombre:     fila[5]  || "",
          especie:    fila[6]  || "",
          cientifico: fila[7]  || "",
          altura:     fila[8]  || "",
          dap:        fila[9]  || "",
          estado:     fila[11] || "",
          fase:       fila[12] || "",
          sitio:      fila[19] || "",
          riesgo:     fila[26] || "",
          manejo:     (fila[21]||"") + " " + (fila[22]||"")
        });
      }
    }

    if (!arboles.length) {
      return ContentService
        .createTextOutput(JSON.stringify({ok: false, msg: "Folio no encontrado"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ok: true, supervisor: supervisor, arboles: arboles}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return fail(err.message);
  }
}

// ═══════════════════════════════════════════════════════════
//  GUARDAR REINSPECCIÓN EN SHEETS
// ═══════════════════════════════════════════════════════════
function guardarReinspeccion(datos) {
  var ss   = obtenerSheet();
  var hoja = ss.getSheetByName(CONFIG.hojaReinsp);
  if (!hoja) hoja = ss.insertSheet(CONFIG.hojaReinsp);

  var encabezados = [
    "Folio Reinsp.", "Folio Original", "Fecha", "Supervisor/a", "Motivo",
    "Árbol", "Especie", "N. Científico", "Altura nueva", "DAP nuevo",
    "Estado nuevo", "Fase nueva", "Copa", "Fuste", "Sistema radicular",
    "Riesgo nuevo", "Manejo nuevo", "Notas árbol",
    "Conclusión", "Acción final", "Fecha registro"
  ];

  if (hoja.getLastRow() === 0) {
    hoja.appendRow(encabezados);
    var r = hoja.getRange(1, 1, 1, encabezados.length);
    r.setBackground("#1B4332").setFontColor("white").setFontWeight("bold");
    hoja.setFrozenRows(1);
  }

  var arboles = datos.arboles || [];
  arboles.forEach(function(a) {
    var fila = [
      datos.folio_reinsp  || "",
      datos.folio_orig    || "",
      datos.fecha         || "",
      datos.supervisor    || "",
      datos.motivo        || "",
      a.nombre            || "",
      a.especie           || "",
      a.cientifico        || "",
      a.altura            || "",
      a.dap               || "",
      a.estado            || "",
      a.fase              || "",
      (a.ch_copa  || []).join(" | "),
      (a.ch_fuste || []).join(" | "),
      (a.ch_raiz  || []).join(" | "),
      a.riesgo            || "",
      a.manejo            || "",
      a.notas             || "",
      datos.conclusion    || "",
      datos.accion_final  || "",
      new Date().toLocaleString("es-MX")
    ];
    hoja.appendRow(fila);
    var numFila = hoja.getLastRow();
    var color = a.riesgo === "alto" ? "#FDEAEA" : a.riesgo === "medio" ? "#FFF9E6" : "#F0FFF4";
    hoja.getRange(numFila, 1, 1, fila.length).setBackground(color);
  });

  hoja.autoResizeColumns(1, encabezados.length);
  Logger.log("Reinspección guardada: " + datos.folio_reinsp);
}

// ═══════════════════════════════════════════════════════════
//  GUARDAR PDF EN DRIVE (sin cambios)
// ═══════════════════════════════════════════════════════════
function guardarPDF(datos) {
  var nombre  = datos.nombre;
  var b64     = datos.data;
  var carpeta = obtenerCarpeta();
  var bytes   = Utilities.base64Decode(b64);
  var blob    = Utilities.newBlob(bytes, MimeType.PDF, nombre);
  var existentes = carpeta.getFilesByName(nombre);
  while (existentes.hasNext()) existentes.next().setTrashed(true);
  var archivo = carpeta.createFile(blob);
  Logger.log("PDF guardado: " + nombre);
  return archivo;
}

// ═══════════════════════════════════════════════════════════
//  ACTUALIZAR SHEETS — DICTAMEN (sin cambios)
// ═══════════════════════════════════════════════════════════
function actualizarSheets(datos) {
  var ss   = obtenerSheet();
  var hoja = ss.getSheetByName(CONFIG.hojaTab) || ss.insertSheet(CONFIG.hojaTab);

  var encabezados = [
    "Folio","Fecha","Hora","Supervisor/a","Programa",
    "Árbol","Especie","Nombre científico","Altura (m)","DAP (cm)",
    "Fase fenológica","Estado físico","Ciclo biológico","Ramificación",
    "Defectos raíz","Defectos fuste","Defectos copa","Plagas/enfermedades",
    "Afectaciones / Diana","Sitio","Requerimientos operativos",
    "Poda (Art.)","Derribo (Art.)","NAE Derribo","NAE Trasplante",
    "Clasificación riesgo","Puntaje","Notas campo","Notas resultado",
    "Observaciones generales","Última actualización"
  ];

  if (hoja.getLastRow() === 0) {
    hoja.appendRow(encabezados);
    var rango = hoja.getRange(1, 1, 1, encabezados.length);
    rango.setBackground("#F47920").setFontColor("white").setFontWeight("bold").setFontSize(9);
    hoja.setFrozenRows(1);
    hoja.setRowHeight(1, 24);
  }

  var fila = [
    datos.folio||"", datos.fecha||"", datos.hora||"", datos.supervisor||"", datos.programa||"",
    datos.nombre||"", datos.especie||"", datos.cientifico||"",
    datos.altura||"", datos.dap||"", datos.fase||"", datos.estado||"",
    datos.ciclo||"", datos.ram||"",
    (datos.ch_raiz||[]).join(" | "), (datos.ch_fuste||[]).join(" | "),
    (datos.ch_copa||[]).join(" | "), (datos.ch_plagas||[]).join(" | "),
    (datos.ch_diana||[]).join(" | "), datos.sitio||"",
    (datos.ch_req||[]).join(" | "), (datos.ch_poda||[]).join(" | "),
    (datos.ch_derribo||[]).join(" | "), (datos.ch_nae_der||[]).join(" | "),
    (datos.ch_nae_tras||[]).join(" | "),
    datos.riesgo==="alto"?"ALTO":"No Aplica", datos._pts||0,
    datos.notas||"", datos.notas_r||"", datos.obs||"",
    new Date().toLocaleString("es-MX")
  ];

  var clave = (datos.folio||"") + "|" + (datos.nombre||"");
  var valores = hoja.getDataRange().getValues();
  var filaExistente = -1;
  for (var i = 1; i < valores.length; i++) {
    if ((valores[i][0]||"") + "|" + (valores[i][5]||"") === clave) {
      filaExistente = i + 1; break;
    }
  }

  if (filaExistente > 0) {
    hoja.getRange(filaExistente, 1, 1, fila.length).setValues([fila]);
    colorearFila(hoja, filaExistente, datos.riesgo);
  } else {
    hoja.appendRow(fila);
    colorearFila(hoja, hoja.getLastRow(), datos.riesgo);
  }
  hoja.autoResizeColumns(1, encabezados.length);
}

function colorearFila(hoja, numFila, riesgo) {
  var color = riesgo === "alto" ? "#FDEAEA" : "#F9FFF9";
  hoja.getRange(numFila, 1, 1, 31).setBackground(color);
}

// ═══════════════════════════════════════════════════════════
//  UTILIDADES
// ═══════════════════════════════════════════════════════════
function obtenerCarpeta() {
  var iter = DriveApp.getFoldersByName(CONFIG.carpeta);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(CONFIG.carpeta);
}

function obtenerSheet() {
  var carpeta = obtenerCarpeta();
  var iter    = DriveApp.getFilesByName(CONFIG.nombreSheet);
  if (iter.hasNext()) return SpreadsheetApp.open(iter.next());
  var ss = SpreadsheetApp.create(CONFIG.nombreSheet);
  DriveApp.getFileById(ss.getId()).moveTo(carpeta);
  return ss;
}

function ok(data) {
  var payload = {ok: true};
  if (data) for (var k in data) payload[k] = data[k];
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function fail(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok: false, error: msg})).setMimeType(ContentService.MimeType.JSON);
}
