/**
 * ============================================================
 *  VICUS - LABORATORIO  |  Google Apps Script (PROFESIONAL V4)
 *  Hospital Natalio Burd
 * ============================================================
 */

const CONFIG_LABORATORIO = {
  folderPDF:     '1wIsTZr3Fn_IcNVFK62qIW5Vp9F0QYt6z', // Informes Técnicos Semanales
  sheetId:       '1M6GAFAifSBEUp3yrM_EXFUv3PYAPwjmx',  // Hoja de Cálculo Semanal
  alertasFolder: '1oqG7iuAAxdgflTpg3i686k_FRUPBLBoS'  // Informes Técnicos (generados por usuario)
};

const SENSORES = [
  { id: '3027200', k: '9VCYTLGCB2Y8NZ28', n: 'Laboratorio Principal - Sensor 1', eq: 'Briket NHC12246', field: 'field1', subtype: 'reactivos' },
  { id: '3027200', k: '9VCYTLGCB2Y8NZ28', n: 'Laboratorio Principal - Sensor 2', eq: 'Inelro NHC12288', field: 'field2', subtype: 'reactivos' }
];

// =====================================================================
// FUNCIÓN PRINCIPAL DEL ACTIVADOR SEMANAL
// =====================================================================
function iniciarReporteSemanal() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty("CURRENT_PHASE");
  ejecutarReporteSemanal();
}

function ejecutarReporteSemanal() {
  const properties = PropertiesService.getScriptProperties();
  const phase = properties.getProperty("CURRENT_PHASE") || "pdfs";

  const hoy           = new Date();
  const haceSieteDias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  const fechaEmision  = fmtFecha(hoy).split(' ')[0];
  const rangoTexto    = fmtFecha(haceSieteDias).split(' ')[0] + " - " + fechaEmision;

  if (phase === "pdfs") {
    console.log("Iniciando Fase 1: Generación de PDFs Laboratorio...");
    SENSORES.forEach(s => {
      try {
        const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
        if (!data || !data.feeds || data.feeds.length === 0) return;

        const trazabilidad = "AUTO-SEM-LAB-" + Utilities.formatDate(hoy, "GMT-3", "yyyyMMdd") + "-" + s.id + "-" + s.field;
        const analizada    = analizarDatos(data.feeds, s.field, s.subtype);
        const conectividad = analizarConectividad(data.feeds, s.field);
        const grafico      = generarGraficoCurva(data.feeds, s.field, s.n);

        const pdfBlob = generarPDFOficial(s, fechaEmision, rangoTexto, trazabilidad, analizada, conectividad, grafico);

        const carpeta = DriveApp.getFolderById(CONFIG_LABORATORIO.folderPDF);
        const file    = carpeta.createFile(pdfBlob);
        file.setName("Informe_Oficial_" + s.n.replace(/ /g, "_") + "_" + trazabilidad + ".pdf");
        console.log(" -> PDF generado para: " + s.n);
      } catch (e) {
        console.error("  Error en " + s.n + ": " + e.message);
      }
    });

    properties.setProperty("CURRENT_PHASE", "sheet");
    crearTriggerDeContinuacion();
    console.log("Fase PDFs completada. Generando planilla en 1 minuto...");
  } else {
    console.log("Iniciando Fase 2: Generación de Planilla Consolidada Laboratorio...");
    const feedsPorSensor = [];
    SENSORES.forEach(s => {
      try {
        const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
        if (data && data.feeds) feedsPorSensor.push({ sensor: s, feeds: data.feeds });
      } catch (e) {
        console.error("  Error obteniendo feeds para planilla, sensor " + s.n + ": " + e.message);
      }
    });

    try {
      if (feedsPorSensor.length > 0) {
        generarSheetSemanal(feedsPorSensor, rangoTexto, hoy);
        console.log(" -> Planilla consolidada creada exitosamente.");
      }
    } catch (e) {
      console.error("  Error al generar Sheet semanal: " + e.message);
    }

    properties.deleteProperty("CURRENT_PHASE");
    eliminarTriggersDeContinuacion();
    console.log("¡Reporte Semanal de Laboratorio completado exitosamente!");
  }
}

function crearTriggerDeContinuacion() {
  eliminarTriggersDeContinuacion();
  ScriptApp.newTrigger("ejecutarReporteSemanal").timeBased().after(60000).create();
}

function eliminarTriggersDeContinuacion() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "ejecutarReporteSemanal") ScriptApp.deleteTrigger(t);
  });
}

// =====================================================================
// GENERACIÓN DE PDF OFICIAL
// =====================================================================
function generarPDFOficial(sensor, fecha, rango, trazabilidad, analizada, conectividad, grafico) {
  const doc  = DocumentApp.create('Temp_Reporte_' + sensor.n);
  const body = doc.getBody();
  body.setMarginLeft(20).setMarginRight(20).setMarginTop(20).setMarginBottom(20);
  const anchoMax = 555;

  // Encabezado con logo
  const logo   = buscarLogoEnDrive("logo_rih.jpg");
  const header = doc.addHeader();
  if (logo) {
    const hp = header.getParagraphs()[0];
    hp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    hp.appendInlineImage(logo).setWidth(anchoMax).setHeight(60);
    hp.setSpacingBefore(0).setSpacingAfter(0);
  } else {
    const hp = header.getParagraphs()[0];
    hp.setText("").setSpacingBefore(0).setSpacingAfter(0);
  }
  header.appendHorizontalRule();

  // Título
  const t1 = body.appendParagraph("INFORME TÉCNICO DE CADENA DE FRÍO - LABORATORIO");
  t1.setFontSize(14).setBold(true).setForegroundColor("#00384d").setSpacingAfter(4);

  body.appendParagraph("Según Disposición ANMAT 10.872/2020")
    .setFontSize(9).setItalic(true).setSpacingAfter(10);

  body.appendParagraph("Dispositivo: " + sensor.n + " (2°C a 8°C)")
    .setBold(true).setFontSize(11).setSpacingBefore(0).setSpacingAfter(2);
  body.appendParagraph("Equipo/Artefacto: " + sensor.eq)
    .setItalic(true).setFontSize(10).setSpacingAfter(2);
  body.appendParagraph("Período: " + rango)
    .setBold(true).setFontSize(10).setSpacingAfter(2);
  body.appendParagraph("Emisión: " + fecha + " | Trazabilidad: " + trazabilidad)
    .setFontSize(8).setSpacingAfter(10);

  // Gráfico
  body.appendParagraph("CURVA TÉRMICA SEMANAL").setBold(true).setFontSize(10).setSpacingBefore(4).setSpacingAfter(4);
  const pChart = body.appendParagraph("");
  pChart.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pChart.appendInlineImage(grafico).setWidth(anchoMax).setHeight(300);

  // Tabla de alertas
  body.appendParagraph("\n⚠️ ALERTAS Y RECUPERACIONES (2°C - 8°C)").setBold(true).setFontSize(10).setSpacingAfter(4);
  const tablaAlertas = [["Fecha y Hora", "Valor", "Estado", "Duración", "Pico Registrado"]];
  if (analizada.alertasFilas.length > 0) {
    analizada.alertasFilas.forEach(f => tablaAlertas.push([f.h, f.v, f.e, f.d, f.p || "--"]));
  } else {
    tablaAlertas.push(["-", "-", "✅ Sin eventos fuera de rango", "-", "-"]);
  }
  estilizarTabla(body.appendTable(tablaAlertas));

  // Eventos de conectividad
  body.appendParagraph("\n📡 EVENTOS DETECTADOS (>10 min sin datos)").setBold(true).setFontSize(10).setSpacingAfter(4);
  const tablaWifi = [["Inicio", "Fin", "Tipo de Corte", "T. Antes", "T. Desp.", "Duración"]];
  if (conectividad.filas.length > 0) {
    conectividad.filas.forEach(f => tablaWifi.push([f.inicio, f.fin, f.tipo, f.antes, f.despues, f.duracion]));
  } else {
    tablaWifi.push(["-", "-", "✅ Sin interrupciones significativas", "-", "-", "-"]);
  }
  estilizarTabla(body.appendTable(tablaWifi));

  // Análisis y recomendaciones
  body.appendParagraph("\nANÁLISIS TÉCNICO:").setBold(true).setFontSize(10);
  if (analizada.textoAnalisis) body.appendParagraph(analizada.textoAnalisis).setFontSize(9).setItalic(true);
  if (conectividad.analisis && conectividad.analisis !== "Sin problemas de conectividad. Monitoreo continuo confirmado.") {
    body.appendParagraph("\nConectividad:").setBold(true).setFontSize(9);
    body.appendParagraph(conectividad.analisis).setFontSize(9).setItalic(true);
  }

  body.appendParagraph("\nRECOMENDACIONES:").setBold(true).setFontSize(10).setForegroundColor("#00384d");
  body.appendParagraph(analizada.textoRecom + "\n• " + conectividad.recom).setFontSize(9);

  const pNota = body.appendParagraph("\n⚙️ NOTA TÉCNICA:");
  pNota.setBold(true).setFontSize(9).setForegroundColor("#475569");
  body.appendParagraph(analizada.notaTecnica).setFontSize(8).setItalic(true).setForegroundColor("#475569");

  const pResp = body.appendParagraph("\n⚠️ RESPONSABILIDAD:");
  pResp.setBold(true).setFontSize(9).setForegroundColor("#475569");
  body.appendParagraph(analizada.notaResponsabilidad).setFontSize(8).setItalic(true).setForegroundColor("#475569");

  // Pie de página
  const footer = doc.addFooter();
  footer.insertHorizontalRule(0);
  const fp = footer.getParagraphs()[0];
  const logoF = buscarLogoEnDrive("footer.jpg");
  if (logoF) {
    fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    fp.appendInlineImage(logoF).setWidth(anchoMax).setHeight(50);
    fp.setSpacingBefore(0).setSpacingAfter(0);
  } else {
    fp.setText("").setSpacingBefore(0).setSpacingAfter(0);
  }

  doc.saveAndClose();
  Utilities.sleep(2000);

  let pdf    = null;
  const docId = doc.getId();
  for (let i = 0; i < 3; i++) {
    try {
      pdf = DriveApp.getFileById(docId).getAs('application/pdf');
      break;
    } catch (e) {
      console.warn("Intento " + (i + 1) + " de generar PDF falló. Reintentando...");
      Utilities.sleep(3000);
    }
  }
  DriveApp.getFileById(docId).setTrashed(true);
  if (!pdf) throw new Error("No se pudo generar el PDF por error de servidor en Google Drive.");
  return pdf;
}

// =====================================================================
// ANÁLISIS DE CONECTIVIDAD
// =====================================================================
function analizarConectividad(feeds, field) {
  let filas        = [];
  let totalMinutos = 0;
  const minOk      = 2.0;
  const maxOk      = 8.0;

  for (let i = 1; i < feeds.length; i++) {
    const d1   = new Date(feeds[i - 1].created_at);
    const d2   = new Date(feeds[i].created_at);
    const diff = (d2 - d1) / 60000;
    if (diff > 10) {
      const v1   = parseFloat(feeds[i - 1][field]);
      const v2   = parseFloat(feeds[i][field]);
      const tipo = (v2 > maxOk || v2 < minOk) ? 'Corte Energía' : 'Corte WiFi';
      filas.push({
        inicio:   fmtFecha(d1).replace(/\/\d{4}/, ""),
        fin:      fmtFecha(d2).replace(/\/\d{4}/, ""),
        tipo:     tipo,
        antes:    isNaN(v1) ? "--" : v1.toFixed(2) + "°C",
        despues:  isNaN(v2) ? "--" : v2.toFixed(2) + "°C",
        duracion: formatDur(diff)
      });
      totalMinutos += diff;
    }
  }

  return {
    filas:    filas,
    analisis: filas.length > 0
      ? "Se detectaron " + filas.length + " " + (filas.length === 1 ? "interrupción" : "interrupciones") + " de datos.\n• Tiempo total sin datos: " + formatDur(totalMinutos) + "\n• Durante los cortes no se puede garantizar el control de la cadena de frío."
      : "Sin problemas de conectividad. Monitoreo continuo confirmado.",
    recom: filas.length > 0
      ? "Verificar el estado del router y la conexión a internet.\n• Revisar la distancia entre el sensor y el punto de acceso WiFi.\n• Considerar registro manual de temperatura durante los períodos sin datos.\n• Evaluar instalación de UPS para el equipo de red."
      : "Mantener el equipo de red en condiciones óptimas para asegurar monitoreo continuo."
  };
}

// =====================================================================
// ANÁLISIS DE DATOS TÉRMICOS
// =====================================================================
function analizarDatos(feeds, field, subtype) {
  const minOk = 2.0;
  const maxOk = 8.0;

  const alertasFilas = [];
  let enAlerta    = false;
  let inicioAlerta = null;
  let valorAlerta  = null;
  let tipoAlerta   = null;
  let picoValor    = null;
  let picoTs       = null;
  let stats        = [];

  const valores = feeds
    .map(f => ({ ts: new Date(f.created_at), val: parseFloat(f[field]) }))
    .filter(f => !isNaN(f.val) && f.val !== -127);

  valores.forEach(punto => {
    const fuera = punto.val < minOk || punto.val > maxOk;
    const tipo  = punto.val > maxOk ? "ALTA" : punto.val < minOk ? "BAJA" : null;

    if (fuera && !enAlerta) {
      enAlerta     = true;
      inicioAlerta = punto.ts;
      valorAlerta  = punto.val;
      tipoAlerta   = tipo;
      picoValor    = punto.val;
      picoTs       = punto.ts;
    } else if (fuera && enAlerta) {
      if (tipoAlerta === "ALTA" && punto.val > picoValor) { picoValor = punto.val; picoTs = punto.ts; }
      if (tipoAlerta === "BAJA" && punto.val < picoValor) { picoValor = punto.val; picoTs = punto.ts; }
    } else if (!fuera && enAlerta) {
      const dur     = (punto.ts - inicioAlerta) / 60000;
      const picoStr = picoValor !== null ? picoValor.toFixed(1) + "°C (" + fmtFecha(picoTs) + ")" : "--";
      alertasFilas.push({
        h: fmtFecha(inicioAlerta),
        v: valorAlerta.toFixed(1) + "°C",
        e: tipoAlerta === "ALTA" ? "⚠️ Alerta Alta (>8°C)" : "⚠️ Alerta Baja (<2°C)",
        d: formatDur(dur),
        p: picoStr
      });
      stats.push({ s: tipoAlerta === "ALTA" ? "Alta" : "Baja", d: dur });
      enAlerta  = false;
      picoValor = null;
      picoTs    = null;
    }
  });

  if (enAlerta && inicioAlerta) {
    const ultimo  = valores[valores.length - 1].ts;
    const dur     = (ultimo - inicioAlerta) / 60000;
    const picoStr = picoValor !== null ? picoValor.toFixed(1) + "°C (" + fmtFecha(picoTs) + ")" : "--";
    alertasFilas.push({
      h: fmtFecha(inicioAlerta),
      v: valorAlerta.toFixed(1) + "°C",
      e: (tipoAlerta === "ALTA" ? "⚠️ Alerta Alta (>8°C)" : "⚠️ Alerta Baja (<2°C)") + " (en curso)",
      d: formatDur(dur),
      p: picoStr
    });
    stats.push({ s: tipoAlerta === "ALTA" ? "Alta" : "Baja", d: dur });
  }

  const tieneAltas = stats.some(s => s.s === 'Alta');
  const tieneBajas = stats.some(s => s.s === 'Baja');
  const durTotal   = stats.reduce((a, b) => a + b.d, 0);

  let textoAnalisis = "Estabilidad térmica confirmada. Sin desvíos en el período.";
  let textoRecom    = "• Continuar monitoreo habitual.\n• Realizar mantenimiento preventivo según calendario.\n• Verificar calibración del sensor periódicamente.";

  if (stats.length > 0) {
    textoAnalisis = "Se detectaron " + stats.length + " " + (stats.length === 1 ? "desvío térmico" : "desvíos térmicos") + " (duración total: " + formatDur(durTotal) + ").\n";
    if (tieneAltas) textoAnalisis += "• Temperatura ALTA (>8°C): riesgo de degradación de reactivos termolábiles.\n";
    if (tieneBajas) textoAnalisis += "• Temperatura BAJA (<2°C): riesgo de congelación de reactivos refrigerados.\n";
    if (durTotal < 30)       textoAnalisis += "Los desvíos fueron breves. Se recomienda monitorear las próximas horas.";
    else if (durTotal < 120) textoAnalisis += "Desvíos de moderada duración. Evaluar posible afectación de reactivos.";
    else                     textoAnalisis += "Desvíos prolongados. Requiere evaluación técnica inmediata según Disposición ANMAT 10.872/2020.";

    textoRecom = "";
    if (tieneAltas) {
      textoRecom += "• Temperatura ALTA detectada: verificar sistema de refrigeración y sellado de puertas.\n";
      textoRecom += "  → Los reactivos refrigerados pueden degradarse irreversiblemente por encima de 8°C.\n";
      textoRecom += "• Controlar termostato y estado del compresor.\n";
      textoRecom += "• Evaluar la aptitud de los reactivos afectados según protocolo vigente.\n";
      textoRecom += "  → La Disposición ANMAT 10.872/2020 exige evaluación documentada ante toda ruptura de cadena de frío.\n";
    }
    if (tieneBajas) {
      textoRecom += "• Temperatura BAJA detectada: revisar configuración del termostato.\n";
      textoRecom += "  → La congelación puede desnaturalizar enzimas y reactivos sensibles.\n";
      textoRecom += "• Verificar que no haya contacto directo de reactivos con el evaporador.\n";
    }
    textoRecom += "• Documentar el evento en el registro de incidencias de cadena de frío.\n";
    textoRecom += "  → La normativa exige trazabilidad completa para auditorías sanitarias.\n";
  }

  const notaTecnica = "Ante cualquier desvío térmico o falla del equipo, la intervención correctiva debe ser realizada por personal técnico calificado (Técnico en Refrigeración matriculado o servicio técnico autorizado por el fabricante). Toda intervención debe quedar documentada con fecha, descripción y firma del responsable, conforme a la Disposición ANMAT 10.872/2020 y las Buenas Prácticas de Almacenamiento (Resolución ANMAT 368/2000).";
  const notaResponsabilidad = "La responsabilidad del cumplimiento de las condiciones de conservación y de la cadena de frío en el sector Laboratorio recae sobre el Director Técnico de Laboratorio, conforme a la Disposición ANMAT 10.872/2020. Ante cualquier incidente, el Director Técnico debe ser notificado de forma inmediata.";

  return { alertasFilas, textoAnalisis, textoRecom, notaTecnica, notaResponsabilidad };
}

// =====================================================================
// GRÁFICO DE CURVA TÉRMICA
// =====================================================================
function generarGraficoCurva(feeds, field, nombre) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Tiempo")
    .addColumn(Charts.ColumnType.NUMBER, "°C");

  let vals = feeds.map(f => parseFloat(f[field])).filter(v => !isNaN(v) && v !== -127);
  if (vals.length === 0) {
    dataTable.addRow(["Sin datos", 5]);
    return Charts.newLineChart().setDataTable(dataTable).setDimensions(2200, 520).setColors(["#3b82f6"]).build().getAs('image/png');
  }

  let minVal = Math.min(...vals);
  let maxVal = Math.max(...vals);
  let yMin   = minVal - 0.5;
  let yMax   = maxVal + 0.5;

  // Detectar gaps reales en los datos crudos
  const UMBRAL_GAP_MS = 10 * 60 * 1000;
  const gapsReales    = [];
  for (let i = 1; i < feeds.length; i++) {
    const t1 = new Date(feeds[i - 1].created_at).getTime();
    const t2 = new Date(feeds[i].created_at).getTime();
    if ((t2 - t1) > UMBRAL_GAP_MS) gapsReales.push({ desde: t1, hasta: t2 });
  }

  const numPuntos   = 800;
  const step        = Math.max(1, Math.floor(feeds.length / numPuntos));
  const muestreados = [];
  for (let i = 0; i < feeds.length; i += step) muestreados.push(feeds[i]);

  let ultimoTsValido = null;
  for (let i = 0; i < muestreados.length; i++) {
    const f    = muestreados[i];
    const val  = parseFloat(f[field]);
    const date = new Date(f.created_at);
    if (isNaN(date.getTime())) continue;
    const esInvalido = isNaN(val) || val === -127;
    if (ultimoTsValido !== null) {
      const hayGap = gapsReales.some(g => g.desde >= ultimoTsValido && g.hasta <= date.getTime());
      if (hayGap) dataTable.addRow([fmtFecha(date).replace(/\/\d{4}/, ""), null]);
    }
    if (esInvalido) {
      dataTable.addRow([fmtFecha(date).replace(/\/\d{4}/, ""), null]);
    } else {
      dataTable.addRow([fmtFecha(date).replace(/\/\d{4}/, ""), val]);
      ultimoTsValido = date.getTime();
    }
  }

  return Charts.newLineChart()
    .setDataTable(dataTable)
    .setDimensions(2200, 520)
    .setColors(["#3b82f6"])
    .setOption("areaOpacity", 0.1)
    .setOption("lineWidth", 1.5)
    .setOption("vAxis", {
      gridlines: { count: 8, color: '#cbd5e1' },
      viewWindow: { min: yMin, max: yMax },
      format: '#.0°C',
      textStyle: { fontSize: 14, color: '#000000', bold: true },
      textPosition: 'out'
    })
    .setOption("hAxis", {
      slantedText: true,
      slantedTextAngle: 45,
      textStyle: { fontSize: 12, color: '#000000', bold: true },
      gridlines: { color: 'none' },
      showTextEvery: 60
    })
    .setOption("chartArea", { width: '94%', height: '70%', left: '4%', right: '1%', top: '4%' })
    .setOption("legend", { position: 'none' })
    .setOption("backgroundColor", "white")
    .build().getAs('image/png');
}

// =====================================================================
// FUNCIONES DE UTILIDAD
// =====================================================================
function estilizarTabla(t) {
  const r0 = t.getRow(0);
  for (let i = 0; i < r0.getNumCells(); i++) r0.getCell(i).setBackgroundColor("#f1f5f9").setBold(true).setFontSize(9);
  for (let i = 1; i < t.getNumRows(); i++) {
    for (let j = 0; j < t.getRow(i).getNumCells(); j++) t.getRow(i).getCell(j).setFontSize(8);
  }
}

function formatDur(m) {
  return m < 60 ? Math.round(m) + "m" : Math.floor(m / 60) + "h " + Math.round(m % 60) + "m";
}

/**
 * Obtiene TODOS los feeds de los últimos N días usando paginación hacia atrás.
 */
function fetchThingSpeakDataCompleto(id, key, dias) {
  const ahora  = new Date();
  const inicio = new Date(ahora.getTime() - dias * 24 * 60 * 60 * 1000);

  let todosLosFeeds = [];
  let fechaHasta    = new Date(ahora);
  let intentos      = 0;
  const MAX_INTENTOS = 10;

  while (intentos < MAX_INTENTOS) {
    const startStr = Utilities.formatDate(inicio,     "GMT-3", "yyyy-MM-dd'T'HH:mm:ss");
    const endStr   = Utilities.formatDate(fechaHasta, "GMT-3", "yyyy-MM-dd'T'HH:mm:ss");

    const url  = "https://api.thingspeak.com/channels/" + id + "/feeds.json?api_key=" + key + "&start=" + startStr + "-03:00&end=" + endStr + "-03:00&results=8000";
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());

    if (!data.feeds || data.feeds.length === 0) break;

    const primerTs = todosLosFeeds.length > 0 ? new Date(todosLosFeeds[0].created_at).getTime() : Infinity;
    const nuevos   = data.feeds.filter(f => new Date(f.created_at).getTime() < primerTs);
    todosLosFeeds  = nuevos.concat(todosLosFeeds);

    if (data.feeds.length < 8000) break;

    fechaHasta = new Date(new Date(data.feeds[0].created_at).getTime() - 60000);
    if (fechaHasta <= inicio) break;

    intentos++;
    Utilities.sleep(500);
  }

  const inicioMs = inicio.getTime();
  todosLosFeeds = todosLosFeeds.filter(f => new Date(f.created_at).getTime() >= inicioMs);
  return { feeds: todosLosFeeds };
}

function buscarLogoEnDrive(n) {
  const f = DriveApp.getFilesByName(n);
  return f.hasNext() ? f.next().getBlob() : null;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'guardarPDF') {
      const bytes = Utilities.base64Decode(body.pdfData);
      const blob  = Utilities.newBlob(bytes, 'application/pdf', body.filename);
      const file  = DriveApp.getFolderById(CONFIG_LABORATORIO.alertasFolder).createFile(blob);
      return ContentService.createTextOutput(JSON.stringify({ result: true, url: file.getUrl() })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ result: false, error: "Acción no reconocida" })).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ result: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Vicus Laboratorio Online.");
}

function fmtFecha(date, conSegundos) {
  const d = new Date(date.getTime() - 3 * 60 * 60 * 1000);
  const p = n => String(n).padStart(2, '0');
  const base = p(d.getUTCDate()) + "/" + p(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear() + " " + p(d.getUTCHours()) + ":" + p(d.getUTCMinutes());
  return conSegundos ? base + ':' + p(d.getUTCSeconds()) : base;
}

// =====================================================================
// GENERACIÓN DE PLANILLA SEMANAL (GOOGLE SHEETS)
// =====================================================================
function abrirOCrearSheet(sheetFolderId, nombreBase) {
  const carpeta = DriveApp.getFolderById(sheetFolderId);
  const nombreSheet = "VICUS_" + nombreBase;

  // Buscar si ya existe un Sheet en esa carpeta
  const archivos = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (archivos.hasNext()) {
    const f = archivos.next();
    if (f.getName() === nombreSheet) {
      return SpreadsheetApp.openById(f.getId());
    }
  }

  // No existe — crear uno nuevo dentro de la carpeta
  const nuevaSS = SpreadsheetApp.create(nombreSheet);
  const file = DriveApp.getFileById(nuevaSS.getId());
  carpeta.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return nuevaSS;
}

function generarSheetSemanal(feedsPorSensor, rangoTexto, fechaHoy) {
  const ss         = abrirOCrearSheet(CONFIG_LABORATORIO.sheetId, "Laboratorio_Semanal");
  const nombreHoja = "Semana " + Utilities.formatDate(fechaHoy, "GMT-3", "dd-MM-yyyy");
  const hojaExistente = ss.getSheetByName(nombreHoja);
  if (hojaExistente) ss.deleteSheet(hojaExistente);
  const hoja = ss.insertSheet(nombreHoja);

  hoja.getRange("A1").setValue("REGISTRO SEMANAL DE TEMPERATURAS - LABORATORIO")
      .setFontSize(13).setFontWeight("bold").setFontColor("#00384d");
  hoja.getRange("A2").setValue("Período: " + rangoTexto).setFontSize(10).setFontStyle("italic");
  hoja.getRange("A3").setValue("Generado: " + fmtFecha(fechaHoy)).setFontSize(9).setFontColor("#64748b");

  // Encabezados dinámicos
  const encabezados = ["Fecha / Hora"];
  feedsPorSensor.forEach(fs => {
    encabezados.push(fs.sensor.n + "\n(" + fs.sensor.eq + ")");
  });

  const filaEncabezado = 5;
  const rangoEnc = hoja.getRange(filaEncabezado, 1, 1, encabezados.length);
  rangoEnc.setValues([encabezados]);
  rangoEnc.setBackground("#00384d").setFontColor("#ffffff").setFontWeight("bold")
          .setFontSize(10).setWrap(true).setVerticalAlignment("middle")
          .setHorizontalAlignment("center");
  hoja.setRowHeight(filaEncabezado, 45);

  // Unificar timestamps agrupando por minuto
  const mapaTemp = {};
  feedsPorSensor.forEach((fs, idx) => {
    fs.feeds.forEach(feed => {
      const val = parseFloat(feed[fs.sensor.field]);
      if (isNaN(val) || val === -127) return;
      const d     = new Date(feed.created_at);
      const clave = fmtFecha(d);
      if (!mapaTemp[clave]) mapaTemp[clave] = { fecha: clave, valores: {} };
      if (mapaTemp[clave].valores[idx] !== undefined) {
        mapaTemp[clave].valores[idx] = (mapaTemp[clave].valores[idx] + val) / 2;
      } else {
        mapaTemp[clave].valores[idx] = val;
      }
    });
  });

  const claves = Object.keys(mapaTemp).sort((a, b) => {
    const toDate = s => {
      const [fecha, hora] = s.split(' ');
      const [dd, mm, yy]  = fecha.split('/');
      return new Date(yy + "-" + mm + "-" + dd + "T" + hora + ":00");
    };
    return toDate(a) - toDate(b);
  });

  const filas = claves.map(clave => {
    const entrada = mapaTemp[clave];
    const fila    = [entrada.fecha];
    feedsPorSensor.forEach((_, idx) => {
      const v = entrada.valores[idx];
      fila.push(v !== undefined ? Math.round(v * 100) / 100 : "");
    });
    return fila;
  });

  if (filas.length > 0) {
    const filaInicio = filaEncabezado + 1;
    
    // Asegurar que la hoja tenga suficientes filas para los datos + el resumen
    const requiredRows = filaInicio + filas.length + 10;
    if (hoja.getMaxRows() < requiredRows) {
      hoja.insertRowsAfter(hoja.getMaxRows(), requiredRows - hoja.getMaxRows());
    }

    hoja.getRange(filaInicio, 1, filas.length, encabezados.length).setValues(filas);

    // Formato condicional: rojo si fuera de rango (2°C - 8°C) — igual para todos los sensores de Laboratorio
    const todasLasReglas = [];
    feedsPorSensor.forEach((fs, idx) => {
      const col      = idx + 2;
      const rangoCol = hoja.getRange(filaInicio, col, filas.length, 1);

      todasLasReglas.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(8.0)
        .setBackground("#fecaca").setFontColor("#dc2626")
        .setRanges([rangoCol]).build());

      todasLasReglas.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(2.0)
        .setBackground("#bfdbfe").setFontColor("#1d4ed8")
        .setRanges([rangoCol]).build());
    });
    hoja.setConditionalFormatRules(todasLasReglas);

    hoja.setColumnWidth(1, 140);
    feedsPorSensor.forEach((_, idx) => hoja.setColumnWidth(idx + 2, 130));

    // Optimización del fondo cebra: una sola llamada a setBackgrounds
    const colorMatrix = [];
    for (let i = 0; i < filas.length; i++) {
      const color = i % 2 === 0 ? "#f8fafc" : "#ffffff";
      colorMatrix.push(Array(encabezados.length).fill(color));
    }
    hoja.getRange(filaInicio, 1, filas.length, encabezados.length).setBackgrounds(colorMatrix);

    hoja.getRange(filaInicio, 2, filas.length, feedsPorSensor.length)
        .setHorizontalAlignment("center").setNumberFormat("0.00");
  }

  // Resumen estadístico
  const filaResumen = filaEncabezado + filas.length + 2;
  hoja.getRange(filaResumen, 1).setValue("RESUMEN ESTADÍSTICO")
      .setFontWeight("bold").setFontColor("#00384d").setFontSize(10);

  const etiquetas = ["Mínimo (°C)", "Máximo (°C)", "Promedio (°C)", "Lecturas totales"];
  etiquetas.forEach((etiq, i) => {
    hoja.getRange(filaResumen + 1 + i, 1).setValue(etiq).setFontWeight("bold");
  });

  feedsPorSensor.forEach((fs, idx) => {
    const col    = idx + 2;
    const valores = filas
      .map(f => f[col - 1])
      .filter(v => v !== "" && !isNaN(v))
      .map(Number);

    if (valores.length > 0) {
      const minVal = Math.min(...valores);
      const maxVal = Math.max(...valores);
      const avg    = Math.round((valores.reduce((a, b) => a + b, 0) / valores.length) * 100) / 100;
      hoja.getRange(filaResumen + 1, col).setValue(minVal);
      hoja.getRange(filaResumen + 2, col).setValue(maxVal);
      hoja.getRange(filaResumen + 3, col).setValue(avg);
      hoja.getRange(filaResumen + 4, col).setValue(valores.length);
    } else {
      hoja.getRange(filaResumen + 1, col).setValue("--");
      hoja.getRange(filaResumen + 2, col).setValue("--");
      hoja.getRange(filaResumen + 3, col).setValue("--");
      hoja.getRange(filaResumen + 4, col).setValue(0);
    }
  });

  hoja.getRange(filaResumen + 1, 1, 4, encabezados.length)
      .setBackground("#f1f5f9").setBorder(true, true, true, true, true, true);
  hoja.setFrozenRows(filaEncabezado);

  console.log("Sheet semanal generado: " + nombreHoja + " (" + filas.length + " registros)");
}

/** Convierte número de columna a letra (1→A, 2→B, 27→AA, etc.) */
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col    = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Función de prueba — ejecutar desde el editor de Apps Script
 * para generar el Sheet semanal sin esperar el trigger automático.
 */
function probarSheetSemanal() {
  const hoy           = new Date();
  const haceSieteDias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  const rangoTexto    = Utilities.formatDate(haceSieteDias, "GMT-3", "dd/MM/yyyy") + " - " + Utilities.formatDate(hoy, "GMT-3", "dd/MM/yyyy");

  const feedsPorSensor = [];
  SENSORES.forEach(s => {
    try {
      const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
      if (!data || !data.feeds || data.feeds.length === 0) return;
      feedsPorSensor.push({ sensor: s, feeds: data.feeds });
    } catch (e) {
      console.error("Error cargando sensor " + s.n + ": " + e.message);
    }
  });

  if (feedsPorSensor.length === 0) {
    console.error("No se obtuvieron datos de ningún sensor.");
    return;
  }
  generarSheetSemanal(feedsPorSensor, rangoTexto, hoy);
}
