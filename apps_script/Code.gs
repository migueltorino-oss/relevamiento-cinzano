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

function hoja_() {
  var ss = ID_PLANILLA ? SpreadsheetApp.openById(ID_PLANILLA) : SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName(HOJA);
  if (!h) h = ss.insertSheet(HOJA);
  if (h.getLastRow() === 0) {
    var head = encabezado_();
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
    h.getRange(2, colRecibido_(), h.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    h.setColumnWidth(6, 220);
    h.setColumnWidth(7, 240);
  }
  return h;
}

function carpetaFotos_() {
  var it = DriveApp.getFoldersByName(CARPETA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FOTOS);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ¿ya se insertó este envío? (la cola offline puede reenviar algo que sí llegó)
function yaExiste_(h, id) {
  var n = h.getLastRow() - 1;
  if (n <= 0) return false;
  var ids = h.getRange(2, colId_(), n, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {          // desde el final: los reenvíos son recientes
    if (String(ids[i][0]) === id) return true;
  }
  return false;
}

function doPost(e) {
  try {
    var p;
    try { p = JSON.parse(e.postData.contents); }
    catch (err) { return json_({ ok: false, error: 'json inválido' }); }
    if (CLAVE && String((p && p.clave) || '') !== CLAVE) return json_({ ok: false, error: 'clave inválida' });

    // Descarga de todo lo relevado (la usa bajar_respuestas.js para armar el Excel).
    // Va por POST para no pasar la clave en la URL.
    if (p && p.accion === 'export') {
      var hx = hoja_();
      var enc = encabezado_();
      var nf = hx.getLastRow() - 1;
      var filas = [];
      if (nf > 0) {
        // getValues (no getDisplayValues): los números viajan como números y no
        // dependen de la configuración regional de la planilla. Las dos columnas de
        // fecha se formatean acá, con la zona horaria del script.
        var vals = hx.getRange(2, 1, nf, enc.length).getValues();
        var iRec = colRecibido_() - 1;
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
      if (yaExiste_(h, id)) {
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

function doGet(e) {
  try {
    var accion = (e && e.parameter && e.parameter.action) || '';
    if (accion === 'list') {
      var h = hoja_();
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
      return json_({ ok: true, ping: true, filas: Math.max(0, hoja_().getLastRow() - 1) });
    }
    return json_({ ok: true, info: 'backend relevamiento cinzano' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
