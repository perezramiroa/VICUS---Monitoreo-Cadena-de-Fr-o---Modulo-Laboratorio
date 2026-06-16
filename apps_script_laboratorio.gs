/**
 * ============================================================
 *  VICUS - LABORATORIO  |  Google Apps Script (PROFESIONAL V4)
 *  Hospital Natalio Burd
 * ============================================================
 */

const CONFIG_LABORATORIO = {
  folderPDF: '1wIsTZr3Fn_IcNVFK62qIW5Vp9F0QYt6z', // Informes Técnicos Semanales
  sheetId: '1M6GAFAifSBEUp3yrM_EXFUv3PYAPwjmx', // Hoja de Cálculo Semanal
  alertasFolder: '1oqG7iuAAxdgflTpg3i686k_FRUPBLBoS' // Informes Técnicos (generados por usuario)
};

const SENSORES = [
  { id: '3027200', k: '9VCYTLGCB2Y8NZ28', n: 'Laboratorio Principal', eq: 'Heladera Laboratorio', field: 'field1', subtype: 'reactivos' }
];

function iniciarReporteSemanal() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty("CURRENT_PHASE");
  ejecutarReporteSemanal();
}

function ejecutarReporteSemanal() {
  const properties = PropertiesService.getScriptProperties();
  const phase = properties.getProperty("CURRENT_PHASE") || "pdfs";

  const hoy = new Date();
  const haceSieteDias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  const fechaEmision = fmtFecha(hoy).split(' ')[0];
  const rangoTexto = fmtFecha(haceSieteDias).split(' ')[0] + " - " + fechaEmision;

  if (phase === "pdfs") {
    console.log("Iniciando Fase 1: Generación de PDFs Laboratorio...");
    SENSORES.forEach(s => {
      try {
        const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
        if (!data || !data.feeds || data.feeds.length === 0) return;

        const trazabilidad = "AUTO-SEM-LAB-" + Utilities.formatDate(hoy, "GMT-3", "yyyyMMdd") + "-" + s.id + "-" + s.field;
        const analizada = analizarDatos(data.feeds, s.field, s.subtype);
        const conectividad = analizarConectividad(data.feeds, s.field, s.subtype);
        const grafico = generarGraficoCurva(data.feeds, s.field, s.n, s.subtype);

        const pdfBlob = generarPDFOficial(s, fechaEmision, rangoTexto, trazabilidad, analizada, conectividad, grafico);
        
        const carpeta = DriveApp.getFolderById(CONFIG_LABORATORIO.folderPDF);
        const file = carpeta.createFile(pdfBlob);
        file.setName("Informe_Oficial_" + s.n.replace(/ /g,"_") + "_" + trazabilidad + ".pdf");
        console.log(" -> PDF generado para: " + s.n);
      } catch (e) {
        console.error("  Error en " + s.n + ": " + e.message);
      }
    });

    properties.setProperty("CURRENT_PHASE", "sheet");
    crearTriggerDeContinuacion();
  } else {
    console.log("Iniciando Fase 2: Generación de Planilla Consolidada Laboratorio...");
    const feedsPorSensor = [];
    SENSORES.forEach(s => {
      try {
        const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
        if (data && data.feeds) {
          feedsPorSensor.push({ sensor: s, feeds: data.feeds });
        }
      } catch (e) {}
    });

    try {
      if (feedsPorSensor.length > 0) {
        generarSheetSemanal(feedsPorSensor, rangoTexto, hoy);
      }
    } catch (e) {
      console.error("  Error al generar Sheet semanal: " + e.message);
    }

    properties.deleteProperty("CURRENT_PHASE");
    eliminarTriggersDeContinuacion();
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

function generarPDFOficial(sensor, fecha, rango, trazabilidad, analizada, conectividad, grafico) {
  const doc = DocumentApp.create('Temp_Reporte_' + sensor.n);
  const body = doc.getBody();
  body.setMarginLeft(20).setMarginRight(20).setMarginTop(20).setMarginBottom(20);
  const anchoMax = 555;

  const logo = buscarLogoEnDrive("logo_rih.jpg");
  const header = doc.addHeader();
  if (logo) {
    const hp = header.appendParagraph("");
    hp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    hp.appendInlineImage(logo).setWidth(anchoMax).setHeight(60);
  }
  header.appendHorizontalRule();

  const t1 = body.appendParagraph("INFORME TÉCNICO DE CADENA DE FRÍO\nLABORATORIO - REACTIVOS");
  t1.setFontSize(14).setBold(true).setForegroundColor("#f97316").setSpacingAfter(4);

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

  body.appendParagraph("CURVA TÉRMICA SEMANAL").setBold(true).setFontSize(10).setSpacingBefore(4).setSpacingAfter(4);
  const pChart = body.appendParagraph("");
  pChart.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pChart.appendInlineImage(grafico).setWidth(anchoMax).setHeight(300);

  body.appendParagraph("\n⚠️ ALERTAS Y RECUPERACIONES").setBold(true).setFontSize(10).setSpacingAfter(4);
  const tablaAlertas = [["Fecha y Hora", "Valor", "Estado", "Duración", "Pico Registrado"]];
  if (analizada.alertasFilas.length > 0) {
    analizada.alertasFilas.forEach(f => tablaAlertas.push([f.h, f.v, f.e, f.d, f.p || "--"]));
  } else {
    tablaAlertas.push(["-", "-", "✅ Sin eventos fuera de rango", "-", "-"]);
  }
  estilizarTabla(body.appendTable(tablaAlertas));

  body.appendParagraph("\n📡 EVENTOS DETECTADOS (>10 min sin datos)").setBold(true).setFontSize(10).setSpacingAfter(4);
  const tablaWifi = [["Inicio", "Fin", "Tipo de Corte", "T. Antes", "T. Desp.", "Duración"]];
  if (conectividad.filas.length > 0) {
    conectividad.filas.forEach(f => tablaWifi.push([f.inicio, f.fin, f.tipo, f.antes, f.despues, f.duracion]));
  } else {
    tablaWifi.push(["-", "-", "✅ Sin interrupciones", "-", "-", "-"]);
  }
  estilizarTabla(body.appendTable(tablaWifi));

  body.appendParagraph("\nANÁLISIS TÉCNICO:").setBold(true).setFontSize(10);
  if (analizada.textoAnalisis) body.appendParagraph(analizada.textoAnalisis).setFontSize(9).setItalic(true);
  
  if (conectividad.analisis && conectividad.analisis !== "Sin problemas de conectividad.") {
    body.appendParagraph("\nConectividad:").setBold(true).setFontSize(9);
    body.appendParagraph(conectividad.analisis).setFontSize(9).setItalic(true);
  }

  body.appendParagraph("\nRECOMENDACIONES:").setBold(true).setFontSize(10).setForegroundColor("#f97316");
  body.appendParagraph(analizada.textoRecom + "\n• " + conectividad.recom).setFontSize(9);

  const pNota = body.appendParagraph("\n⚙️ NOTA TÉCNICA:");
  pNota.setBold(true).setFontSize(9).setForegroundColor("#475569");
  body.appendParagraph(analizada.notaTecnica).setFontSize(8).setItalic(true).setForegroundColor("#475569");

  const pResp = body.appendParagraph("\n⚠️ RESPONSABILIDAD:");
  pResp.setBold(true).setFontSize(9).setForegroundColor("#475569");
  body.appendParagraph(analizada.notaResponsabilidad).setFontSize(8).setItalic(true).setForegroundColor("#475569");

  const footer = doc.addFooter();
  footer.appendHorizontalRule();
  const logoF = buscarLogoEnDrive("footer.jpg");
  if (logoF) {
    const fp = footer.appendParagraph("");
    fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    fp.appendInlineImage(logoF).setWidth(anchoMax).setHeight(50);
  }

  doc.saveAndClose();
  Utilities.sleep(2000);
  
  let pdf = null;
  const docId = doc.getId();
  for (let i = 0; i < 3; i++) {
    try {
      pdf = DriveApp.getFileById(docId).getAs('application/pdf');
      break;
    } catch (e) {
      Utilities.sleep(3000);
    }
  }
  DriveApp.getFileById(docId).setTrashed(true);
  return pdf;
}

function analizarConectividad(feeds, field, subtype) {
  let filas = [];
  let totalMinutos = 0;
  const maxOk = 8.0;
  const minOk = 2.0;
  
  for (let i = 1; i < feeds.length; i++) {
    const d1 = new Date(feeds[i-1].created_at);
    const d2 = new Date(feeds[i].created_at);
    const diff = (d2 - d1) / 60000;
    if (diff > 10) {
      const v1 = parseFloat(feeds[i-1][field]);
      const v2 = parseFloat(feeds[i][field]);
      const tipo = (v2 > maxOk || v2 < minOk) ? 'Corte Energía/Temperatura' : 'Corte WiFi';
      filas.push({
        inicio: fmtFecha(d1).replace(/\/\d{4}/, ""),
        fin: fmtFecha(d2).replace(/\/\d{4}/, ""),
        tipo: tipo,
        antes: isNaN(v1) ? "--" : v1.toFixed(1) + "°C",
        despues: isNaN(v2) ? "--" : v2.toFixed(1) + "°C",
        duracion: formatDur(diff)
      });
      totalMinutos += diff;
    }
  }
  return {
    filas: filas,
    analisis: filas.length > 0 ? \`Detectadas \${filas.length} interrupciones.\` : "Sin problemas de conectividad.",
    recom: filas.length > 0 ? "Verificar red y posibles fallas de energía." : "Mantener condiciones actuales."
  };
}

function analizarDatos(feeds, field, subtype) {
  const minOk = 2.0;
  const maxOk = 8.0;

  const alertasFilas = [];
  let enAlerta = false;
  let inicioAlerta = null;
  let valorAlerta = null;
  let tipoAlerta = null;
  let picoValor = null;
  let picoTs = null;
  let stats = [];

  const valores = feeds
    .map(f => ({ ts: new Date(f.created_at), val: parseFloat(f[field]) }))
    .filter(f => !isNaN(f.val) && f.val !== -127);

  valores.forEach(punto => {
    const fuera = punto.val < minOk || punto.val > maxOk;
    const tipo = punto.val > maxOk ? "ALTA" : (punto.val < minOk) ? "BAJA" : null;

    if (fuera && !enAlerta && tipo) {
      enAlerta = true;
      inicioAlerta = punto.ts;
      valorAlerta = punto.val;
      tipoAlerta = tipo;
      picoValor = punto.val;
      picoTs = punto.ts;
    } else if (fuera && enAlerta) {
      if (tipoAlerta === "ALTA" && punto.val > picoValor) { picoValor = punto.val; picoTs = punto.ts; }
      if (tipoAlerta === "BAJA" && punto.val < picoValor) { picoValor = punto.val; picoTs = punto.ts; }
    } else if (!fuera && enAlerta) {
      const dur = (punto.ts - inicioAlerta) / 60000;
      const picoStr = picoValor !== null ? picoValor.toFixed(1) + "°C" : "--";
      alertasFilas.push({
        h: fmtFecha(inicioAlerta),
        v: valorAlerta.toFixed(1) + "°C",
        e: tipoAlerta === "ALTA" ? \`⚠️ Alta (>\${maxOk}°C)\` : \`⚠️ Baja (<\${minOk}°C)\`,
        d: formatDur(dur),
        p: picoStr
      });
      stats.push({ s: tipoAlerta === "ALTA" ? "Alta" : "Baja", d: dur });
      enAlerta = false;
      picoValor = null;
    }
  });

  if (enAlerta && inicioAlerta) {
    const ultimo = valores[valores.length - 1].ts;
    const dur = (ultimo - inicioAlerta) / 60000;
    alertasFilas.push({
      h: fmtFecha(inicioAlerta),
      v: valorAlerta.toFixed(1) + "°C",
      e: "⚠️ Alerta en curso",
      d: formatDur(dur),
      p: "--"
    });
    stats.push({ s: tipoAlerta === "ALTA" ? "Alta" : "Baja", d: dur });
  }

  let textoAnalisis = "Estabilidad térmica confirmada. Sin desvíos de temperatura.";
  let textoRecom = "• Continuar monitoreo habitual.";
  if (stats.length > 0) {
    textoAnalisis = \`Detectados \${stats.length} desvíos termicos.\n• Reactivos de diagnostico fuera de rango (\${minOk} a \${maxOk}°C).\`;
    textoRecom = "• Identificar reactivos afectados y evaluar descarte si superan tiempo maximo o se congelaron.\n• Revisar equipo y contactar servicio tecnico.";
  }

  const notaTecnica = "Toda intervencion debe quedar documentada con fecha, descripcion y firma del responsable, conforme a la Disposicion ANMAT 10.872/2020 y BPA.";
  const notaResponsabilidad = "La responsabilidad del cumplimiento de la cadena de frio recae sobre el Servicio de Laboratorio y su Director Tecnico.";

  return { alertasFilas, textoAnalisis, textoRecom, notaTecnica, notaResponsabilidad };
}

function generarGraficoCurva(feeds, field, nombre, subtype) {
  const dataTable = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Tiempo").addColumn(Charts.ColumnType.NUMBER, "°C");
  let vals = feeds.map(f => parseFloat(f[field])).filter(v => !isNaN(v) && v !== -127);
  if (vals.length === 0) {
    dataTable.addRow(["Sin datos", 0]);
    return Charts.newLineChart().setDataTable(dataTable).setDimensions(1200, 400).build().getAs('image/png');
  }

  let yMin = 0;
  let yMax = 15;

  const numPuntos = 500;
  const step = Math.max(1, Math.floor(feeds.length / numPuntos));

  for (let i = 0; i < feeds.length; i += step) {
    const f = feeds[i];
    const val = parseFloat(f[field]);
    if (!isNaN(val) && val !== -127) {
      dataTable.addRow([fmtFecha(new Date(f.created_at)).replace(/\/\d{4}/, ""), val]);
    }
  }

  return Charts.newLineChart()
    .setDataTable(dataTable)
    .setDimensions(2200, 520)
    .setColors(["#f97316"])
    .setOption("vAxis", { viewWindow: { min: yMin, max: yMax } })
    .build().getAs('image/png');
}

function estilizarTabla(t) {
  const r0 = t.getRow(0);
  for(let i=0; i<r0.getNumCells(); i++) r0.getCell(i).setBackgroundColor("#f1f5f9").setBold(true).setFontSize(9);
  for(let i=1; i<t.getNumRows(); i++) {
    for(let j=0; j<t.getRow(i).getNumCells(); j++) t.getRow(i).getCell(j).setFontSize(8);
  }
}

function formatDur(m) {
  return m < 60 ? Math.round(m) + "m" : Math.floor(m/60) + "h " + Math.round(m%60) + "m";
}

function fetchThingSpeakDataCompleto(id, key, dias) {
  const res = UrlFetchApp.fetch(\`https://api.thingspeak.com/channels/\${id}/feeds.json?api_key=\${key}&minutes=\${dias*1440}&results=8000\`);
  return { feeds: JSON.parse(res.getContentText()).feeds || [] };
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
      const blob = Utilities.newBlob(bytes, 'application/pdf', body.filename);
      const file = DriveApp.getFolderById(CONFIG_LABORATORIO.alertasFolder).createFile(blob);
      return ContentService.createTextOutput(JSON.stringify({result: true, url: file.getUrl()})).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({result: false, error: "Acción no reconocida"})).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({result: false, error: e.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) { return ContentService.createTextOutput("Vicus Laboratorio Online."); }

function fmtFecha(date) {
  const d = new Date(date.getTime() - 3 * 60 * 60 * 1000);
  const p = n => String(n).padStart(2, '0');
  return \`\${p(d.getUTCDate())}/\${p(d.getUTCMonth()+1)}/\${d.getUTCFullYear()} \${p(d.getUTCHours())}:\${p(d.getUTCMinutes())}\`;
}

function generarSheetSemanal(feedsPorSensor, rangoTexto, fechaHoy) {
  const ss = SpreadsheetApp.openById(CONFIG_LABORATORIO.sheetId);
  const nombreHoja = "Semana " + Utilities.formatDate(fechaHoy, "GMT-3", "dd-MM-yyyy");
  const hojaExistente = ss.getSheetByName(nombreHoja);
  if (hojaExistente) ss.deleteSheet(hojaExistente);
  const hoja = ss.insertSheet(nombreHoja);

  hoja.getRange("A1").setValue("REGISTRO SEMANAL DE TEMPERATURAS - LABORATORIO").setFontWeight("bold").setFontColor("#f97316");
}
