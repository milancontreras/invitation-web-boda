/**
 * Invitación de boda · conecta la página web con esta hoja de Google.
 *
 * 1. En la hoja: Extensiones > Apps Script. Borra lo que haya, pega este código y guarda.
 * 2. Recarga la hoja. Aparece el menú "💌 Boda"; la primera vez la hoja se prepara sola
 *    (si no, usa "💌 Boda > Preparar hoja" y acepta los permisos).
 * 3. En Apps Script: Implementar > Nueva implementación > Aplicación web.
 *    Ejecutar como: Yo · Quién tiene acceso: Cualquier persona. Copia la URL que termina en /exec
 *    y pégala en API_URL dentro de index.html.
 */

const HOJA = 'Invitados';
const RESUMEN = 'Resumen';
const CONFIG = 'Config';
const COL = { codigo: 1, invitacion: 2, pases: 3, telefono: 4, link: 5, enviar: 6, respuesta: 7, personas: 8, mensaje: 9, fecha: 10 };
const ENCABEZADOS = ['Código', 'Invitación para', 'Pases', 'WhatsApp (593…)', 'Link personal', 'Enviar', 'Respuesta', 'Personas que asisten', 'Mensaje', 'Fecha de respuesta'];
const LETRAS = 'abcdefghjkmnpqrstuvwxyz23456789';
const VINO = '#3B1219', MARFIL = '#F3E6BE';

/**
 * Separador de argumentos de las fórmulas según el idioma de la hoja.
 * En español (es_ES, es_EC…) Sheets usa ";" y en inglés ",". Se prueba una vez y se guarda.
 */
function separador() {
  const props = PropertiesService.getDocumentProperties();
  let sep = props.getProperty('SEP');
  if (sep) return sep;
  const ss = SpreadsheetApp.getActive();
  const celda = (ss.getSheetByName(CONFIG) || ss.getSheets()[0]).getRange('Z1');
  celda.setFormula('=SUM(1,1)');
  SpreadsheetApp.flush();
  sep = celda.getValue() === 2 ? ',' : ';';
  celda.clearContent();
  props.setProperty('SEP', sep);
  return sep;
}

/* ---------------- Menú y preparación ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('💌 Boda')
    .addItem('Preparar hoja', 'prepararHoja')
    .addItem('Completar códigos y links', 'completarFilas')
    .addToUi();
  if (!SpreadsheetApp.getActive().getSheetByName(CONFIG)) prepararHoja();
}

function onEdit(e) {
  if (!e || e.range.getSheet().getName() !== HOJA) return;
  if (e.range.getRow() < 2) return;
  if (e.range.getColumn() <= COL.pases && e.range.getLastColumn() >= COL.invitacion) completarFilas();
}

function prepararHoja() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(HOJA); }

  sh.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS])
    .setFontWeight('bold').setBackground(VINO).setFontColor(MARFIL);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  [90, 220, 60, 140, 320, 80, 110, 110, 280, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('D2:D').setNumberFormat('@');
  sh.getRange('G2:G').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Pendiente', 'Asistirá', 'No asistirá'], true).build());
  const filas = sh.getRange('A2:J');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G2="Asistirá"').setBackground('#DCE3C2').setRanges([filas]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G2="No asistirá"').setBackground('#F1D2CA').setRanges([filas]).build()
  ]);

  let cf = ss.getSheetByName(CONFIG);
  if (!cf) {
    cf = ss.insertSheet(CONFIG);
    cf.getRange('A1:B3').setValues([
      ['Ajuste', 'Valor'],
      ['Dirección de la página', 'https://tu-boda.netlify.app/'],
      ['Mensaje para WhatsApp', '¡Hola! Con mucho cariño te compartimos nuestra invitación de boda:']
    ]);
    cf.getRange('A1:B1').setFontWeight('bold').setBackground(VINO).setFontColor(MARFIL);
    cf.setColumnWidth(1, 220);
    cf.setColumnWidth(2, 480);
  }

  let rs = ss.getSheetByName(RESUMEN);
  if (!rs) {
    rs = ss.insertSheet(RESUMEN, 0);
    const S = separador();
    rs.getRange('A1:B8').setValues([
      ['Resumen', 'Total'],
      ['Invitaciones', '=COUNTA(Invitados!B2:B)'],
      ['Pases entregados', '=SUM(Invitados!C2:C)'],
      ['Invitaciones que confirmaron', `=COUNTIF(Invitados!G2:G${S}"Asistirá")`],
      ['Personas confirmadas', '=SUM(Invitados!H2:H)'],
      ['Invitaciones que no asistirán', `=COUNTIF(Invitados!G2:G${S}"No asistirá")`],
      ['Invitaciones sin responder', '=B2-B4-B6'],
      ['Pases sin responder', `=B3-SUMIF(Invitados!G2:G${S}"Asistirá"${S}Invitados!C2:C)-SUMIF(Invitados!G2:G${S}"No asistirá"${S}Invitados!C2:C)`]
    ]);
    rs.getRange('A1:B1').setFontWeight('bold').setBackground(VINO).setFontColor(MARFIL);
    rs.getRange('B2:B8').setFontWeight('bold').setFontSize(13).setFontColor(VINO);
    rs.setColumnWidth(1, 260);
  }
  completarFilas();
}

/** Pone código, link, botón de WhatsApp y estado "Pendiente" a cada invitado nuevo. */
function completarFilas() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA);
  const ultima = sh.getLastRow();
  if (ultima < 2) return;
  const n = ultima - 1;
  const datos = sh.getRange(2, 1, n, COL.respuesta).getValues();
  const usados = new Set(datos.map(f => String(f[0]).trim().toLowerCase()).filter(Boolean));

  const S = separador();
  const codigos = [], formulas = [], estados = [];
  datos.forEach((f, i) => {
    const r = i + 2;
    const hay = String(f[COL.invitacion - 1]).trim() !== '';
    let codigo = String(f[COL.codigo - 1]).trim();
    if (hay && !codigo) { codigo = nuevoCodigo(usados); usados.add(codigo); }
    codigos.push([codigo]);
    formulas.push(hay ? [
      `=IF(A${r}=""${S}""${S}Config!$B$2&"?i="&A${r})`,
      `=IF(OR(A${r}=""${S}D${r}="")${S}""${S}HYPERLINK("https://wa.me/"&REGEXREPLACE(TO_TEXT(D${r})${S}"[^0-9]"${S}"")&"?text="&ENCODEURL(Config!$B$3&" "&E${r})${S}"Enviar"))`
    ] : ['', '']);
    estados.push([hay && !f[COL.respuesta - 1] ? 'Pendiente' : f[COL.respuesta - 1]]);
  });
  sh.getRange(2, COL.codigo, n, 1).setValues(codigos);
  sh.getRange(2, COL.link, n, 2).setFormulas(formulas);
  sh.getRange(2, COL.respuesta, n, 1).setValues(estados);
}

function nuevoCodigo(usados) {
  let c;
  do {
    c = '';
    for (let i = 0; i < 6; i++) c += LETRAS[Math.floor(Math.random() * LETRAS.length)];
  } while (usados.has(c));
  return c;
}

/* ---------------- Lo que usa la página web ---------------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'rsvp') return json(guardarRespuesta(p));
    return json(buscarInvitado(p.i));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    return json(guardarRespuesta(JSON.parse((e.postData && e.postData.contents) || '{}')));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function encontrarFila(codigo) {
  codigo = String(codigo || '').trim().toLowerCase();
  if (!codigo) return null;
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA);
  const ultima = sh.getLastRow();
  if (ultima < 2) return null;
  const codigos = sh.getRange(2, COL.codigo, ultima - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim().toLowerCase() === codigo) return i + 2;
  }
  return null;
}

function buscarInvitado(codigo) {
  const fila = encontrarFila(codigo);
  if (!fila) return { ok: false, error: 'no_encontrado' };
  const v = SpreadsheetApp.getActive().getSheetByName(HOJA).getRange(fila, 1, 1, COL.fecha).getValues()[0];
  return {
    ok: true,
    invitacion: String(v[COL.invitacion - 1]),
    pases: Number(v[COL.pases - 1]) || 1,
    respuesta: v[COL.respuesta - 1] || 'Pendiente',
    personas: Number(v[COL.personas - 1]) || 0,
    mensaje: String(v[COL.mensaje - 1] || '')
  };
}

function guardarRespuesta(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fila = encontrarFila(p.i);
    if (!fila) return { ok: false, error: 'no_encontrado' };
    const sh = SpreadsheetApp.getActive().getSheetByName(HOJA);
    const pases = Number(sh.getRange(fila, COL.pases).getValue()) || 1;
    const asiste = p.asiste === 'si';
    const personas = asiste ? Math.max(1, Math.min(pases, parseInt(p.personas, 10) || 1)) : 0;
    let mensaje = String(p.mensaje || '').slice(0, 500);
    if (/^[=+\-@]/.test(mensaje)) mensaje = "'" + mensaje; // evita que un mensaje se convierta en fórmula
    const respuesta = asiste ? 'Asistirá' : 'No asistirá';
    sh.getRange(fila, COL.respuesta, 1, 4).setValues([[respuesta, personas, mensaje, new Date()]]);
    return { ok: true, respuesta: respuesta, personas: personas };
  } finally {
    lock.releaseLock();
  }
}
