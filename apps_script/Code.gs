/**
 * Backend del Relevamiento Cinzano — Google Apps Script.
 *
 * Recibe los envíos del formulario (POST), los guarda como filas en la pestaña
 * "Respuestas" de la planilla, y guarda las fotos en una carpeta de Drive.
 * También responde GET ?action=list con los clientes ya relevados (para que el
 * formulario los marque con ✓ en cualquier celular).
 *
 * INSTALACIÓN (una sola vez):
 *  1. Crear una planilla nueva en Google Sheets (sheets.new), ponerle nombre
 *     "Relevamiento Cinzano".
 *  2. Menú Extensiones → Apps Script. Borrar el contenido y pegar este archivo.
 *  3. Botón "Implementar" → "Nueva implementación" → tipo "Aplicación web":
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier persona
 *     → Implementar. Autorizar los permisos cuando lo pida.
 *  4. Copiar la URL que termina en /exec y pasársela a Claude para que la
 *     incruste en el formulario.
 *
 * Si más adelante se cambia este código, hay que hacer Implementar → Gestionar
 * implementaciones → editar (lápiz) → Versión: Nueva → Implementar, para que la
 * MISMA URL /exec sirva la versión nueva.
 */

var TZ = 'America/Argentina/Buenos_Aires';
var HOJA = 'Respuestas';
var CARPETA_FOTOS = 'Relevamiento Cinzano - Fotos';

// ID de la planilla. Se usa cuando el script es INDEPENDIENTE (creado desde
// script.google.com). Si el script está pegado dentro de la planilla (Extensiones →
// Apps Script), se puede dejar en '' y usa la planilla que lo contiene.
var ID_PLANILLA = '1SUVSBW1lIQk2rWIZthLrqJyrDlUKmAtOaM2WucatLoU';

// Clave del equipo: tiene que ser LA MISMA que "clave" en config.json del formulario.
// Sirve para que solo el formulario pueda escribir en la planilla (la URL /exec es
// pública porque los vendedores no tienen cuenta de Google). Con '' no se valida.
var CLAVE = '__PONER_CLAVE__';

// Orden de las columnas de artículos. El payload trae el id de cada uno, así que
// los valores se ubican por id (el orden de acá solo define las columnas).
var ARTICULOS = [
  { id: 'rosso',   nombre: 'Cinzano Rosso' },
  { id: 'bianco',  nombre: 'Cinzano Bianco' },
  { id: 'segundo', nombre: 'Cinzano Segundo' },
  { id: 'spritz',  nombre: 'Cinzano Spritz' }
];

function encabezado_() {
  var head = ['Fecha', 'Hora', 'Vendedor cód.', 'Vendedor', 'Cliente cód.', 'Cliente', 'Dirección'];
  for (var i = 0; i < ARTICULOS.length; i++) {
    head.push(ARTICULOS[i].nombre, 'Precio ' + ARTICULOS[i].nombre.replace('Cinzano ', ''));
  }
  head.push('Tiene algún Cinzano', '¿Nos compró?', 'Observaciones', 'Foto', 'Recibido', 'Id');
  return head;
}
function colId_() { return encabezado_().length; }          // última columna
function colRecibido_() { return encabezado_().length - 1; }

function planilla_() {
  return ID_PLANILLA ? SpreadsheetApp.openById(ID_PLANILLA) : SpreadsheetApp.getActiveSpreadsheet();
}

// Cada formulario guarda en SU PROPIO archivo de Google Sheets. El id se guarda en las
// propiedades del script la primera vez y se reusa siempre; si el archivo no existe
// todavía, se crea solo (no hay que crearlo a mano ni hardcodear ningún id).
// Sin `formulario` (el caso de Cinzano) usa la planilla original.
function planillaDe_(formId, nombreArchivo) {
  if (!formId) return planilla_();
  var props = PropertiesService.getScriptProperties();
  var clave = 'planilla_' + String(formId);
  var id = props.getProperty(clave);
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (err) { props.deleteProperty(clave); }   // la borraron: se crea de nuevo
  }
  var ss = SpreadsheetApp.create(nombreArchivo || ('Relevamiento ' + formId));
  props.setProperty(clave, ss.getId());
  return ss;
}

// Lo que se sabe de cada formulario, para poder pasarle el link al usuario.
function planillasConocidas_() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var out = [];
  for (var k in props) {
    if (k.indexOf('planilla_') !== 0) continue;
    var info = { formulario: k.slice(9), id: props[k] };
    try {
      var ss = SpreadsheetApp.openById(props[k]);
      info.nombre = ss.getName();
      info.url = ss.getUrl();
      info.hojas = ss.getSheets().map(function (h) { return h.getName(); });
    } catch (err) { info.error = 'no se pudo abrir'; }
    out.push(info);
  }
  return out;
}

// Devuelve la pestaña, creándola con su encabezado y formatos la primera vez.
// Sirve para cualquier formulario: el encabezado y el archivo se los pasa quien la llama.
function hojaDe_(nombre, head, ss) {
  ss = ss || planilla_();
  var h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  if (h.getLastRow() === 0) {
    h.appendRow(head);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, head.length).setFontWeight('bold');
    // Fecha y Recibido se escriben como Date reales: el formato lo fija la columna,
    // así no dependen de la configuración regional de la planilla.
    h.getRange(2, 1, h.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy');
    // Hora como TEXTO: si no, Sheets interpreta "16:43" como valor de hora y al leerlo
    // por API vuelve como fecha 1899-12-30 desfasada por el huso (bajar_respuestas.js
    // igual lo recupera, pero así queda limpio de entrada).
    h.getRange(2, 2, h.getMaxRows() - 1, 1).setNumberFormat('@');
    h.getRange(2, head.length - 1, h.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    h.setColumnWidth(6, 220);
    h.setColumnWidth(7, 240);
    // en un archivo nuevo, la "Hoja 1" que crea Google queda vacía y molestando
    var hojas = ss.getSheets();
    for (var i = 0; i < hojas.length; i++) {
      var otra = hojas[i];
      if (otra.getSheetId() !== h.getSheetId() && otra.getLastRow() === 0 && hojas.length > 1) {
        try { ss.deleteSheet(otra); } catch (err) {}
      }
    }
  }
  return h;
}

function hoja_() { return hojaDe_(HOJA, encabezado_()); }

// ---------------- formularios genéricos ----------------
// El payload trae la definición de las columnas, así un formulario nuevo NO necesita
// tocar este código ni reimplementar: manda `hoja` y `campos` [{etiqueta, tipo, valor}].
var PREFIJO = ['Fecha', 'Hora', 'Vendedor cód.', 'Vendedor', 'Cliente cód.', 'Cliente', 'Dirección'];
var SUFIJO = ['Observaciones', 'Foto', 'Recibido', 'Id'];

function encabezadoGen_(campos) {
  var head = PREFIJO.slice();
  for (var i = 0; i < campos.length; i++) head.push(String(campos[i].etiqueta || ('Campo ' + (i + 1))));
  return head.concat(SUFIJO);
}

// Nombre de pestaña seguro (Sheets no acepta algunos caracteres)
function nombreHoja_(s) {
  var n = String(s || 'Respuestas').replace(/[\[\]\*\/\\\?:]/g, ' ').trim().slice(0, 90);
  return n || 'Respuestas';
}

function carpeta_(nombre) {
  var it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}
function carpetaFotos_() { return carpeta_(CARPETA_FOTOS); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ¿ya se insertó este envío? (la cola offline puede reenviar algo que sí llegó)
// La columna Id es siempre la última del encabezado de esa pestaña.
function yaExiste_(h, id, colId) {
  var n = h.getLastRow() - 1;
  if (n <= 0) return false;
  var ids = h.getRange(2, colId || h.getLastColumn(), n, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {          // desde el final: los reenvíos son recientes
    if (String(ids[i][0]) === id) return true;
  }
  return false;
}

// ===================== RESUMEN (hoja "Resumen", ex "Hoja 1") =====================
// Cartera por vendedor sacada del padrón de Gescom (rutas de preventa con día de visita).
// Es una foto: para refrescarla, regenerar esta tabla y volver a correr armarResumen().
var UNIVERSO_FECHA = '14/08/2026';
var UNIVERSO = [
  [1, 'ROCIO MORE', 'JUAN', 214], [2, 'MAGALI BADER', 'JUAN', 227], [3, 'SOFIA AQUINO', 'JUAN', 231],
  [4, 'SAMANTA MACHADO', 'JUAN', 231], [5, 'VALENTINA GOMEZ', 'JUAN', 202], [6, 'SILVIA BOUTET', 'JUAN', 223],
  [7, 'ANGELA ECHEVERRIA', 'JUAN', 212], [8, 'AILEN BLANCO', 'JUAN', 217], [9, 'PILAR FERNANDEZ', 'JUAN', 205],
  [10, 'IVANA CASTILLO', 'JUAN', 212], [11, 'LORENA ESCOBAR', 'JUAN', 220], [12, 'FIAMA JARA', 'JUAN', 212],
  [24, 'KAF DANIEL SANCHEZ', 'LUCIA', 83], [25, 'KAF ROMINA MERELLO', 'LUCIA', 80], [26, 'Kaf Gaby Peralta', 'LUCIA', 80],
  [27, 'Kaf Maria Jose Minio', 'LUCIA', 81], [28, 'VENDEDOR 28 KAF', 'LUCIA', 154],
  [31, 'MARIA VICTORIA ZERDA', 'FLOR', 210], [32, 'CAROLINA ECHEVERRIA', 'FLOR', 227], [33, 'SOFIA CATIVIELA', 'FLOR', 208],
  [34, 'NARA GONZALEZ', 'FLOR', 213], [35, 'AGUSTINA COCA', 'FLOR', 230], [37, 'GIMENA ESCUDERO', 'FLOR', 202],
  [38, 'NAHUEL BAETA', 'FLOR', 221], [44, 'BELEN ROMERO', 'FLOR', 220], [45, 'NATALIA POLITO', 'FLOR', 243]
];

var ROJO = '#b32330', CREMA = '#faf7f2', BORDE = '#e5ddd5', GRIS = '#8a817c';

/**
 * Arma el tablero de seguimiento en la primera hoja. Todo con FÓRMULAS: se actualiza
 * solo a medida que entran relevamientos, no hay que volver a correrlo.
 * Correr desde el editor: elegir armarResumen y apretar Ejecutar.
 */
function armarResumen() {
  var ss = ID_PLANILLA ? SpreadsheetApp.openById(ID_PLANILLA) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Resumen') || ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (!sh) sh = ss.insertSheet('Resumen', 0);
  sh.setName('Resumen');
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
  sh.clear();
  sh.clearFormats();
  var filtros = sh.getFilter();
  if (filtros) filtros.remove();

  hoja_();                                       // asegura que exista la pestaña Respuestas
  var R = "'" + HOJA + "'!";                     // prefijo de referencia a la hoja de datos
  var totalCartera = 0;
  for (var i = 0; i < UNIVERSO.length; i++) totalCartera += UNIVERSO[i][3];

  var relevados = 'IFERROR(COUNTA(UNIQUE(FILTER(' + R + 'E2:E,' + R + 'E2:E<>""))),0)';
  var pintadas = [];   // [rango, color de fondo, negrita, color de letra]

  function banda(fila, texto) {
    sh.getRange(fila, 1).setValue(texto);
    pintadas.push([fila + ':' + fila, ROJO, true, '#ffffff']);
  }
  function encabezados(fila, cols) {
    sh.getRange(fila, 1, 1, cols.length).setValues([cols]);
    pintadas.push([fila + ':' + fila, CREMA, true, '#2b2320']);
  }

  // ---------- título ----------
  sh.getRange('A1').setValue('RELEVAMIENTO CINZANO — seguimiento');
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor(ROJO);
  sh.getRange('A2').setValue('Último relevamiento cargado');
  sh.getRange('B2').setFormula('=IFERROR(TEXT(MAX(' + R + 'T2:T),"dd/MM/yyyy HH:mm"),"todavía no hay datos")');
  sh.getRange('D2').setValue('Cartera del padrón al ' + UNIVERSO_FECHA);
  sh.getRange('A2:F2').setFontColor(GRIS);

  // ---------- avance ----------
  banda(4, 'AVANCE DEL CENSO');
  var avance = [
    ['Clientes relevados', '=' + relevados, 'de ' + totalCartera + ' de cartera', '=IFERROR(B5/' + totalCartera + ',0)'],
    ['Faltan por relevar', '=' + totalCartera + '-B5', '', '=IFERROR(1-D5,1)'],
    ['Relevamientos cargados', '=COUNTA(' + R + 'E2:E)', 'incluye recargas del mismo cliente', ''],
    ['Vendedores que arrancaron', '=IFERROR(COUNTA(UNIQUE(FILTER(' + R + 'C2:C,' + R + 'C2:C<>""))),0)', 'de ' + UNIVERSO.length, ''],
    ['Relevados hoy', '=IFERROR(SUMPRODUCT((INT(' + R + 'A2:A)=TODAY())*(' + R + 'A2:A<>"")),0)', '', ''],
    ['Relevados últimos 7 días', '=IFERROR(SUMPRODUCT((INT(' + R + 'A2:A)>=TODAY()-6)*(' + R + 'A2:A<>"")),0)', '', '']
  ];
  sh.getRange(5, 1, avance.length, 4).setValues(avance);
  sh.getRange('D5:D6').setNumberFormat('0.0%');

  // ---------- hallazgos ----------
  banda(12, 'QUÉ SE ENCONTRÓ EN LA CALLE');
  var hallazgos = [
    ['Clientes con algún Cinzano', '=COUNTIF(' + R + 'P2:P,"SI")', '=IFERROR(B13/' + relevados + ',0)'],
    ['Clientes sin nada de Cinzano', '=COUNTIF(' + R + 'P2:P,"NO")', '=IFERROR(B14/' + relevados + ',0)'],
    ['Clientes que nos compraron', '=COUNTIF(' + R + 'Q2:Q,"SI")', '=IFERROR(B15/' + relevados + ',0)'],
    ['Clientes con foto', '=COUNTIF(' + R + 'S2:S,"https://*")', '=IFERROR(B16/' + relevados + ',0)']
  ];
  sh.getRange(13, 1, hallazgos.length, 3).setValues(hallazgos);
  sh.getRange('C13:C16').setNumberFormat('0.0%');

  // ---------- artículos ----------
  banda(18, 'PRESENCIA Y PRECIOS DE GÓNDOLA POR ARTÍCULO');
  encabezados(19, ['Artículo', 'Clientes que lo tienen', '% de los relevados', 'Precio promedio', 'Mínimo', 'Máximo']);
  var arts = [];
  for (var a = 0; a < ARTICULOS.length; a++) {
    var colTiene = String.fromCharCode(72 + a * 2);      // H, J, L, N
    var colPrecio = String.fromCharCode(73 + a * 2);     // I, K, M, O
    var f = 20 + a;
    arts.push([
      ARTICULOS[a].nombre,
      '=COUNTIF(' + R + colTiene + '2:' + colTiene + ',"SI")',
      '=IFERROR(B' + f + '/' + relevados + ',0)',
      '=IFERROR(AVERAGEIF(' + R + colPrecio + '2:' + colPrecio + ',">0"),"—")',
      '=IFERROR(MIN(FILTER(' + R + colPrecio + '2:' + colPrecio + ',' + R + colPrecio + '2:' + colPrecio + '>0)),"—")',
      '=IFERROR(MAX(' + R + colPrecio + '2:' + colPrecio + '),"—")'
    ]);
  }
  sh.getRange(20, 1, arts.length, 6).setValues(arts);
  sh.getRange(20, 3, arts.length, 1).setNumberFormat('0.0%');
  sh.getRange(20, 4, arts.length, 3).setNumberFormat('$#,##0.00');

  // ---------- por vendedor ----------
  var f0 = 20 + ARTICULOS.length + 2;
  banda(f0, 'AVANCE POR VENDEDOR');
  encabezados(f0 + 1, ['Vendedor', 'Equipo', 'Cartera', 'Relevados', '% avance', 'Con Cinzano', 'Nos compró']);
  var filas = [];
  for (var v = 0; v < UNIVERSO.length; v++) {
    var u = UNIVERSO[v], fila = f0 + 2 + v;
    filas.push([
      u[1], u[2], u[3],
      '=IFERROR(COUNTA(UNIQUE(FILTER(' + R + 'E2:E,' + R + 'C2:C=' + u[0] + '))),0)',
      '=IFERROR(D' + fila + '/C' + fila + ',0)',
      '=COUNTIFS(' + R + 'C2:C,' + u[0] + ',' + R + 'P2:P,"SI")',
      '=COUNTIFS(' + R + 'C2:C,' + u[0] + ',' + R + 'Q2:Q,"SI")'
    ]);
  }
  var fTot = f0 + 2 + UNIVERSO.length;
  filas.push(['TOTAL', '', totalCartera,
    '=SUM(D' + (f0 + 2) + ':D' + (fTot - 1) + ')',
    '=IFERROR(D' + fTot + '/C' + fTot + ',0)',
    '=SUM(F' + (f0 + 2) + ':F' + (fTot - 1) + ')',
    '=SUM(G' + (f0 + 2) + ':G' + (fTot - 1) + ')']);
  sh.getRange(f0 + 2, 1, filas.length, 7).setValues(filas);
  sh.getRange(f0 + 2, 5, filas.length, 1).setNumberFormat('0.0%');
  pintadas.push([fTot + ':' + fTot, CREMA, true, '#2b2320']);

  // ---------- comentarios ----------
  var fc = fTot + 2;
  banda(fc, 'COMENTARIOS DE LOS VENDEDORES (los más nuevos primero)');
  encabezados(fc + 1, ['Fecha', 'Vendedor', 'Cliente', 'Comentario']);
  sh.getRange(fc + 2, 1).setFormula(
    '=IFERROR(QUERY(' + R + 'A2:R,"select Col1, Col4, Col6, Col18 where Col18 is not null and Col18 <> \'\' ' +
    'order by Col1 desc limit 100",0),"Todavía no hay comentarios cargados")');
  sh.getRange(fc + 2, 1, 200, 1).setNumberFormat('dd/MM/yyyy');

  // ---------- formato general ----------
  for (var p = 0; p < pintadas.length; p++) {
    var rr = sh.getRange(pintadas[p][0]);
    rr.setBackground(pintadas[p][1]).setFontWeight(pintadas[p][2] ? 'bold' : 'normal').setFontColor(pintadas[p][3]);
  }
  var anchos = [300, 130, 210, 130, 110, 130, 130];
  for (var c = 0; c < anchos.length; c++) sh.setColumnWidth(c + 1, anchos[c]);
  sh.getRange(1, 1, sh.getMaxRows(), 7).setVerticalAlignment('middle');
  sh.setFrozenRows(2);
  if (sh.getMaxColumns() > 7) sh.deleteColumns(8, sh.getMaxColumns() - 7);
  SpreadsheetApp.flush();
  return 'Resumen armado (' + (fc + 2) + ' filas)';
}

function doPost(e) {
  try {
    var p;
    try { p = JSON.parse(e.postData.contents); }
    catch (err) { return json_({ ok: false, error: 'json inválido' }); }
    if (CLAVE && String((p && p.clave) || '') !== CLAVE) return json_({ ok: false, error: 'clave inválida' });

    // Descarga de todo lo relevado (la usa bajar_respuestas.js para armar el Excel).
    // Va por POST para no pasar la clave en la URL. `hoja` elige el formulario.
    if (p && p.accion === 'export') {
      var nomb = nombreHoja_(p.hoja || HOJA);
      var ssx = planillaDe_(p.formulario, p.archivo);
      var hx = ssx.getSheetByName(nomb);
      if (!hx) return json_({ ok: true, encabezado: [], filas: [], vacia: true });
      var nCols = hx.getLastColumn();
      var enc = hx.getRange(1, 1, 1, nCols).getValues()[0];
      var nf = hx.getLastRow() - 1;
      var filas = [];
      if (nf > 0) {
        // getValues (no getDisplayValues): los números viajan como números y no
        // dependen de la configuración regional de la planilla. Las dos columnas de
        // fecha se formatean acá, con la zona horaria del script.
        var vals = hx.getRange(2, 1, nf, nCols).getValues();
        var iRec = nCols - 2;                    // ..., Recibido, Id
        for (var k = 0; k < vals.length; k++) {
          var f = vals[k];
          if (f[0] instanceof Date) f[0] = Utilities.formatDate(f[0], TZ, 'dd/MM/yyyy');
          if (f[iRec] instanceof Date) f[iRec] = Utilities.formatDate(f[iRec], TZ, 'dd/MM/yyyy HH:mm');
          filas.push(f);
        }
      }
      return json_({ ok: true, encabezado: enc, filas: filas });
    }

    if (!p || !p.clienteCod || !p.vendedorCod) return json_({ ok: false, error: 'payload incompleto' });

    // Formulario genérico: el payload define las columnas y la pestaña destino.
    if (p.campos) return altaGenerica_(p);

    var id = String(p._id || ((p._k || '') + '|' + (p.fecha || '')));
    var cache = CacheService.getScriptCache();
    if (cache.get('id:' + id)) return json_({ ok: true, dup: true });

    // Fecha del RELEVAMIENTO (la manda el form): con cola offline puede ser de otro día
    // que el de recepción. Si viene rara, se cae a la hora del servidor.
    var recibido = new Date();
    var fechaRel = p.fecha ? new Date(p.fecha) : recibido;
    if (isNaN(fechaRel.getTime())) fechaRel = recibido;

    // La foto se sube FUERA del lock (tarda segundos y no necesita serializarse).
    // Si falla, el relevamiento se guarda igual: los datos importan más que la foto.
    var fotoUrl = '';
    if (p.foto && String(p.foto).indexOf('data:image') === 0) {
      try {
        var b64 = String(p.foto).split(',')[1];
        var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg',
          'cli' + p.clienteCod + '_' + Utilities.formatDate(fechaRel, TZ, 'yyyyMMdd_HHmmss') + '_' + id.slice(0, 8) + '.jpg');
        fotoUrl = carpetaFotos_().createFile(blob).getUrl();
      } catch (err) {
        fotoUrl = 'ERROR al subir la foto: ' + err;
      }
    }

    // Valores de artículos por id (no por posición)
    var porId = {};
    var arts = p.articulos || [];
    for (var i = 0; i < arts.length; i++) {
      if (arts[i] && arts[i].id) porId[arts[i].id] = arts[i];
    }

    var fila = [
      fechaRel,
      Utilities.formatDate(fechaRel, TZ, 'HH:mm'),
      p.vendedorCod, p.vendedor || '',
      p.clienteCod, p.cliente || '', p.direccion || ''
    ];
    var alguno = 'NO';
    for (var j = 0; j < ARTICULOS.length; j++) {
      var a = porId[ARTICULOS[j].id] || arts[j] || {};
      if (a.tiene === 'SI') alguno = 'SI';
      fila.push(a.tiene === 'SI' ? 'SI' : 'NO',
        (a.precio === '' || a.precio == null) ? '' : Number(a.precio));
    }
    fila.push(alguno, p.nosCompro === 'SI' ? 'SI' : 'NO', p.obs || '', fotoUrl, recibido, id);

    // Solo el appendRow necesita el lock. tryLock (no waitLock) para poder responder
    // JSON siempre: una excepción devolvería una página HTML que el form no entiende.
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return json_({ ok: false, error: 'servidor ocupado', reintentar: true });
    try {
      var h = hoja_();
      if (yaExiste_(h, id, colId_())) {
        cache.put('id:' + id, '1', 21600);
        return json_({ ok: true, dup: true });
      }
      h.appendRow(fila);
      cache.put('id:' + id, '1', 21600);   // 6 h
      return json_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    // reintentar: el form lo trata como transitorio (y tras varios intentos lo aparta
    // de la cola, así un error permanente no traba los envíos que están detrás).
    return json_({ ok: false, error: String(err), reintentar: true });
  }
}

// Alta de cualquier formulario definido por el payload (hoja + campos).
// Mismas garantías que el camino de Cinzano: idempotencia por _id, fecha del
// relevamiento, foto fuera del lock y lock sólo alrededor del appendRow.
function altaGenerica_(p) {
  var id = String(p._id || ((p._k || '') + '|' + (p.fecha || '')));
  var cache = CacheService.getScriptCache();
  if (cache.get('id:' + id)) return json_({ ok: true, dup: true });

  var campos = p.campos || [];
  var head = encabezadoGen_(campos);
  var nomb = nombreHoja_(p.hoja || HOJA);

  var recibido = new Date();
  var fechaRel = p.fecha ? new Date(p.fecha) : recibido;
  if (isNaN(fechaRel.getTime())) fechaRel = recibido;

  var fotoUrl = '';
  if (p.foto && String(p.foto).indexOf('data:image') === 0) {
    try {
      var b64 = String(p.foto).split(',')[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg',
        'cli' + p.clienteCod + '_' + Utilities.formatDate(fechaRel, TZ, 'yyyyMMdd_HHmmss') + '_' + id.slice(0, 8) + '.jpg');
      fotoUrl = carpeta_(p.carpetaFotos || (nomb + ' - Fotos')).createFile(blob).getUrl();
    } catch (err) {
      fotoUrl = 'ERROR al subir la foto: ' + err;
    }
  }

  var fila = [
    fechaRel,
    Utilities.formatDate(fechaRel, TZ, 'HH:mm'),
    p.vendedorCod, p.vendedor || '',
    p.clienteCod, p.cliente || '', p.direccion || ''
  ];
  for (var i = 0; i < campos.length; i++) {
    var c = campos[i], v = c.valor;
    if (v === '' || v === null || v === undefined) fila.push('');
    else if (c.tipo === 'monto' || c.tipo === 'numero') fila.push(Number(v));
    else fila.push(String(v));
  }
  fila.push(p.obs || '', fotoUrl, recibido, id);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_({ ok: false, error: 'servidor ocupado', reintentar: true });
  try {
    var ss = planillaDe_(p.formulario, p.archivo);
    var h = hojaDe_(nomb, head, ss);
    // Si el formulario cambió de columnas, no mezclar datos viejos con nuevos.
    var actual = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].join('|');
    if (actual !== head.join('|')) {
      return json_({ ok: false, error: 'las columnas de la pestaña "' + nomb + '" no coinciden con el formulario' });
    }
    if (yaExiste_(h, id, head.length)) {
      cache.put('id:' + id, '1', 21600);
      return json_({ ok: true, dup: true });
    }
    h.appendRow(fila);
    cache.put('id:' + id, '1', 21600);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var accion = (e && e.parameter && e.parameter.action) || '';
    if (accion === 'list') {
      // ?f=<pestaña> y ?form=<formulario> (el formulario define en qué archivo buscar)
      var nombreF = nombreHoja_((e.parameter && e.parameter.f) || HOJA);
      var formF = (e.parameter && e.parameter.form) || '';
      var h = planillaDe_(formF).getSheetByName(nombreF);
      if (!h) return json_({ ok: true, rows: [] });
      var n = h.getLastRow() - 1;
      var rows = [];
      if (n > 0) {
        var vals = h.getRange(2, 1, n, 5).getValues();   // Fecha, Hora, Vend cód., Vend, Cliente cód.
        for (var i = 0; i < vals.length; i++) {
          var f = vals[i][0];
          var ddmm = (f instanceof Date && !isNaN(f.getTime()))
            ? Utilities.formatDate(f, TZ, 'dd/MM')
            : String(f).slice(0, 5);
          rows.push([Number(vals[i][2]), Number(vals[i][4]), ddmm]);
        }
      }
      return json_({ ok: true, rows: rows });
    }
    if (accion === 'ping') {
      // `generico: true` le dice a los formularios nuevos que esta versión ya sabe
      // guardar cualquier formulario en su propia pestaña. Sin esto, un formulario
      // genérico apuntando a una versión vieja escribiría filas vacías en Respuestas.
      return json_({
        ok: true, ping: true, generico: true, archivos: true, version: 'generico-2',
        filas: Math.max(0, hoja_().getLastRow() - 1)
      });
    }
    // Links de los archivos que fue creando cada formulario
    if (accion === 'planillas') {
      return json_({ ok: true, planillas: planillasConocidas_(), principal: planilla_().getUrl() });
    }
    return json_({ ok: true, info: 'backend relevamiento cinzano' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
