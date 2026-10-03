/* =====================================================================
   Sivarcito™ — Backend en Google Sheets (Google Apps Script)
   ---------------------------------------------------------------------
   Instalación / actualización:
   1. En la hoja de cálculo: Extensiones ▸ Apps Script. Pega este archivo en Code.gs.
   2. Ejecuta la función "configurar" (botón ▶) y acepta los permisos
      (Hojas, Drive para las fotos de documentos y conexión externa para Wompi).
      Crea/repara las hojas Productos, Reservas, Clientes y Config y los
      disparadores automáticos (borrado de fotos y limpieza de pagos vencidos).
   3. Si vienes de la versión anterior, ejecuta UNA vez "actualizarTarifas"
      (o menú Sivarcito ▸ Aplicar tarifas y contacto nuevos).
   4. Credenciales de Wompi (NUNCA en la hoja ni en la página):
      Configuración del proyecto (engranaje) ▸ Propiedades del script ▸ agrega
        WOMPI_CLIENT_ID      = App ID del aplicativo de Wompi
        WOMPI_CLIENT_SECRET  = API Secret del aplicativo de Wompi
      Sin estas propiedades la reserva se registra igual, pero sin cobro en línea.
   5. (Opcional, recomendado) Servicios ▸ + ▸ Drive API: así las fotos de
      documentos se borran definitivamente en vez de ir a la papelera.
   6. Implementar ▸ Gestionar implementaciones ▸ editar ▸ Nueva versión
      (la URL /exec no cambia). Ejecutar como: Yo · Acceso: Cualquier persona.
   Los cambios en las HOJAS se ven en la página al instante, sin reimplementar.
   ===================================================================== */

// ID de tu Google Sheets (la parte entre /d/ y /edit de su URL).
// Solo hace falta si el script NO se creó desde Extensiones ▸ Apps Script de la hoja.
const SHEET_ID = '1pI1tmnAUWsreXTa_GV4ZxpeTdTgFIBcutmtkv596Clg';
function libro() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
}

const HOJAS = {
  productos: {
    nombre: 'Productos',
    enc: ['ID', 'Ocultar', 'Categoría', 'Producto', 'Descripción', 'Talla', 'Precio', 'Cantidad', 'Etiqueta', 'Fotos'],
    anchos: [50, 65, 120, 300, 320, 70, 75, 80, 90, 320]
  },
  // Las columnas nuevas van al final para no desordenar las reservas ya existentes.
  reservas: {
    nombre: 'Reservas',
    enc: ['Registrado', 'Nombre', 'Teléfono', 'Email', 'Inicio', 'Fin', 'Días', 'Total', 'Notas', 'Estado',
          'Reserva', 'Tipo documento', 'N.º documento', 'Foto documento', 'Enlace de pago', 'Transacción Wompi', 'Punto de entrega'],
    anchos: [140, 180, 110, 190, 95, 95, 50, 70, 260, 120, 110, 130, 130, 240, 240, 160, 230]
  },
  clientes: {
    nombre: 'Clientes',
    enc: ['Teléfono', 'Nombre', 'Email', 'Tipo documento', 'N.º documento', 'Primera reserva', 'Última reserva', 'Reservas'],
    anchos: [120, 200, 200, 130, 130, 130, 130, 80]
  },
  config: {
    nombre: 'Config',
    enc: ['Clave', 'Valor', 'Nota'],
    anchos: [130, 160, 460]
  }
};

const CONFIG_INICIAL = [
  ['precioDia', 20, 'Precio por día en alquileres cortos (menos de "diasSemana" días)'],
  ['precioDiaSemana', 7, 'Precio por día cuando el alquiler es de "diasSemana" días o más (7 × $7 = $49)'],
  ['diasSemana', 7, 'Días mínimos para aplicar el precio semanal'],
  ['precioMes', 180, 'Precio por cada bloque de "diasMes" días. Nunca se cobra más que esto por un bloque'],
  ['diasMes', 30, 'Días que cuenta un mes de alquiler'],
  ['telefono', '+503 6272-5022', 'Teléfono para llamadas que aparece en la página'],
  ['whatsapp', '50362725022', 'Número de WhatsApp sin + ni espacios'],
  ['puntosEntrega', 'Parque Cuscatlán | Plaza León (frente a Las Terrazas)', 'Puntos de entrega y devolución que el cliente puede elegir, separados por |'],
  ['urlSitio', '', 'Dirección pública de la página (https://…). Wompi regresa aquí al cliente después de pagar'],
  ['scooterMantenimiento', 'SI', 'SI = la página muestra "En mantenimiento" y no acepta reservas. NO = alquiler activo'],
  ['minutosPago', 30, 'Minutos que se apartan las fechas mientras el cliente paga en Wompi'],
  ['wompiPruebas', 'NO', 'SI = acepta pagos de prueba de Wompi como pagados. Déjalo en NO cuando cobres de verdad']
];
// Claves que la página puede leer. Nada más de Config sale del servidor.
const CONFIG_PUBLICA = ['precioDia', 'precioDiaSemana', 'diasSemana', 'precioMes', 'diasMes', 'telefono', 'whatsapp', 'puntosEntrega', 'scooterMantenimiento'];

// [ID, Ocultar, Categoría, Producto, Descripción, Talla, Precio, Cantidad, Etiqueta, Fotos]
// Filas con el mismo Producto + Categoría se muestran como UNA tarjeta con botones de talla.
const CATALOGO_INICIAL = [
  [0, false, 'Ropa', 'Camisa sublimada blanca, piel de durazno', 'Camisa blanca sublimada en tela piel de durazno.', '0 - L', 8.99, '', '', '0.jpeg\n1.jpeg\n2.jpeg'],
  [1, false, 'Ropa', 'Camisa sublimada blanca, piel de durazno', '', 'XL', 9.99, '', '', ''],
  [2, false, 'Ropa', 'Camisa sublimada blanca, piel de durazno', '', 'XXL', 11.99, '', '', ''],
  [3, false, 'Ropa', 'Camisa sublimada blanca, piel de durazno', '', 'XXXL', 14.99, '', '', ''],
  [4, false, 'Ropa', 'Camisa negra, algodón estampado DTF', 'Camisa negra de algodón con estampado DTF.', 'XL', 15.99, '', '', '4 (1).jpeg\n4 (2).jpeg'],
  [5, false, 'Ropa', 'Camisa negra de niño, algodón estampado DTF', 'Camisa negra de algodón para niño con estampado DTF.', '', 14.99, '', '', '5 (1).jpeg\n5 (2).jpeg'],
  [6, false, 'Personalizados', 'Termos personalizados', 'Termo personalizado con el diseño que elijas.', '', 14.99, '', '', '6.jpeg'],
  [7, false, 'Personalizados', 'Combo: termo, taza y camisa', 'Termo, taza y camisa personalizados en un solo combo.', '', 27, '', 'Combo', '7.jpeg'],
  [8, false, 'Personalizados', 'Mouse pad cuadrado', 'Mouse pad cuadrado con diseño estampado.', '', 8, 1, '', '8 (1).jpeg\n8 (2).jpeg'],
  [9, false, 'Personalizados', 'Bolsa ecológica', 'Bolsa ecológica reutilizable con diseño estampado.', '', 9.99, '', '', '9 (1).jpeg\n9 (2).jpeg\n9 (3).jpeg\n9 (4).jpeg\n9 (5).jpeg\n9 (6).jpeg'],
  [10, false, 'Ropa', 'Hoodie con dos dibujos', 'Hoodie con diseño adelante y atrás.', '', 23.99, '', '', '10 (1).jpeg\n10 (2).jpeg'],
  [11, false, 'Audio', 'Audífonos inalámbricos On-Ear BALL, negro', 'Audífonos inalámbricos On-Ear modelo BALL.', '', 14.99, 0, '', '11.jpeg'],
  [12, false, 'Audio', 'Audífonos inalámbricos On-Ear BALL, púrpura', 'Audífonos inalámbricos On-Ear modelo BALL.', '', 14.99, 1, '', '12.jpeg'],
  [13, false, 'Audio', 'Audífonos inalámbricos On-Ear BALL, beige', 'Audífonos inalámbricos On-Ear modelo BALL.', '', 14.99, 1, '', '13.jpeg'],
  [14, false, 'Gorras', 'Gorra Baseball camuflaje NY, negro', 'Gorra tipo baseball con diseño camuflaje y logo NY.', '', 16.99, 1, '', '14.jpeg'],
  [15, false, 'Gorras', 'Gorra Baseball plana NY, gris/negro', 'Gorra de visera plana con logo NY en gris y negro.', '', 15.99, 1, '', '15.jpeg'],
  [16, false, 'Gorras', 'Gorra Baseball L*, marrón', 'Gorra tipo baseball color marrón.', '', 16.99, 1, '', '16.jpeg'],
  [17, false, 'Ropa', 'Camisa estilo Preppy blanco/vino', 'Camisa estilo Preppy en blanco con vino.', 'M', 17.99, 1, '', '17.jpeg'],
  [18, false, 'Accesorios', 'Llaveros amantes del café', 'Llavero temático para los amantes del café.', '', 1.25, 15, '', '18.jpeg'],
  [19, false, 'Accesorios', 'Bolso de bandolera multifuncional', 'Bolso cruzado multifuncional. Colores: negro, azul y gris.', '', 7, 3, '', '19.jpeg'],
  [20, false, 'Ramos', 'Ramo 1', 'Ramo decorativo para regalo.', '', 7, 1, '', '20.jpeg'],
  [21, false, 'Ramos', 'Ramo 2', 'Ramo decorativo para regalo.', '', 10, 1, '', '21.jpeg'],
  [22, false, 'Ramos', 'Ramo 3', 'Ramo decorativo para regalo.', '', 7, 1, '', '22.jpeg'],
  [23, false, 'Ramos', 'Ramo 4', 'Ramo decorativo para regalo.', '', 7, 1, '', '23.jpeg'],
  [24, false, 'Ramos', 'Llavero flor (c/u)', 'Llavero con flor, precio por unidad.', '', 2.5, 4, '', '24.jpeg'],
  [25, false, 'Audio', 'Audífonos inalámbricos On-Ear, negro', 'Audífonos inalámbricos On-Ear.', '', 14.99, 1, '', '25.jpeg'],
  [26, false, 'Lociones', 'Lociones de bolsillo', 'Loción en spray de bolsillo. Varios aromas para dama y caballero; pregúntanos por WhatsApp cuáles hay disponibles.', '', 7, '', '', fotosNumeradas(26, 12)],
  [27, false, 'Accesorios', 'Pulseras de paracord', 'Pulsera tejida de paracord con broche. Varios colores disponibles.', '', 7.6, '', '', fotosNumeradas(27, 5)]
];

// "26 (1).jpeg", "26 (2).jpeg"… una por línea, como en la columna Fotos
function fotosNumeradas(id, n) {
  return Array.from({ length: n }, (_, i) => id + ' (' + (i + 1) + ').jpeg').join('\n');
}

const CATEGORIAS = ['Ropa', 'Gorras', 'Personalizados', 'Audio', 'Accesorios', 'Ramos', 'Lociones'];

// Estados de una reserva. "Pendiente" queda por compatibilidad con reservas viejas
// y para las que no pudieron generar enlace de pago (se cobran a mano).
const ESTADOS = ['Pago pendiente', 'Pagada', 'Confirmada', 'Entregada', 'Devuelta', 'Cancelada', 'Expirada', 'Pendiente'];
const LIBERAN_FECHAS = ['Cancelada', 'Expirada'];
const BORRAN_FOTO = ['Devuelta', 'Cancelada', 'Expirada'];   // al pasar a uno de estos se elimina la foto del documento
const TIPOS_DOC = ['DUI', 'Pasaporte', 'Licencia de conducir'];
const HORAS_EXPIRA_PAGO = 24;   // sin pago en este tiempo: Expirada y se borra la foto

/* =====================================================================
   Menú dentro de Google Sheets
   ===================================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sivarcito')
    .addItem('Crear / reparar hojas', 'configurar')
    .addItem('Aplicar tarifas y contacto nuevos', 'actualizarTarifas')
    .addItem('Agregar productos nuevos del catálogo', 'agregarProductosNuevos')
    .addToUi();
}

// Agrega a la hoja Productos las filas de CATALOGO_INICIAL cuyo ID todavía no existe.
// No toca los productos que ya tienes (ni sus precios o cantidades).
function agregarProductosNuevos() {
  const sh = libro().getSheetByName(HOJAS.productos.nombre);
  const ids = filas(HOJAS.productos.nombre).map(r => String(r['ID']).trim());
  const nuevos = CATALOGO_INICIAL.filter(r => ids.indexOf(String(r[0])) < 0);
  if (nuevos.length) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevos.length, nuevos[0].length).setValues(nuevos);
    configurar();   // casillas, categorías y formato de precio en las filas nuevas
  }
  return nuevos.length + ' producto(s) agregado(s)';
}

/* =====================================================================
   Creación automática de hojas, encabezados y disparadores
   ===================================================================== */
function configurar() {
  const ss = libro();
  Object.keys(HOJAS).forEach(k => prepararHoja(ss, HOJAS[k]));

  const prod = ss.getSheetByName(HOJAS.productos.nombre);
  if (prod.getLastRow() < 2) {
    prod.getRange(2, 1, CATALOGO_INICIAL.length, CATALOGO_INICIAL[0].length).setValues(CATALOGO_INICIAL);
  }
  const n = HOJAS.productos.enc.length;
  const filasProd = prod.getMaxRows() - 1;
  prod.getRange(2, 2, filasProd).insertCheckboxes();
  prod.getRange(2, 3, filasProd).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CATEGORIAS, true).setAllowInvalid(true).build());
  prod.getRange(2, 7, filasProd).setNumberFormat('$0.00');
  prod.getRange(2, 1, filasProd, n).setVerticalAlignment('top').setWrap(true);
  prod.getRange(1, 2).setNote('Marca la casilla para ocultar el producto de la página sin borrarlo.');
  prod.getRange(1, 8).setNote('Vacío = no se muestra la cantidad.\n0 o "AGOTADO" = aparece Agotado.');
  prod.getRange(1, 10).setNote('Nombre del archivo tal cual está en la carpeta productos/ de GitHub (ej. 26.jpeg). Varias fotos: una por línea.');
  prod.getRange(1, 4).setNote('Filas con el mismo Producto y Categoría se juntan en una tarjeta con botones de talla.');

  // Config: agrega las claves que falten sin tocar los valores que ya tengas
  const cfg = ss.getSheetByName(HOJAS.config.nombre);
  const claves = cfg.getLastRow() > 1 ? cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  CONFIG_INICIAL.filter(r => claves.indexOf(r[0]) < 0).forEach(r => cfg.appendRow(r));

  const res = ss.getSheetByName(HOJAS.reservas.nombre);
  const col = (h) => HOJAS.reservas.enc.indexOf(h) + 1;
  const filasRes = res.getMaxRows() - 1;
  res.getRange(2, col('Estado'), filasRes).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(true).build());
  res.getRange(2, col('Total'), filasRes).setNumberFormat('$0.00');
  res.getRange(1, col('Estado')).setNote(
    'Pago pendiente → el cliente aún no paga (sus fechas se liberan al pasar los minutos de "minutosPago" en Config; a las ' + HORAS_EXPIRA_PAGO + ' h pasa a Expirada).\n' +
    'Pagada → Wompi confirmó el pago.\n' +
    'Devuelta / Cancelada / Expirada → la foto del documento se borra de Drive automáticamente.');
  res.getRange(1, col('Foto documento')).setNote('Enlace privado a la foto en tu Drive. Solo tú puedes abrirla.');

  // Borra la hoja vacía que Google crea por defecto
  ss.getSheets().forEach(sh => {
    const esNuestra = Object.keys(HOJAS).some(k => HOJAS[k].nombre === sh.getName());
    if (!esNuestra && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  carpetaDocumentos();
  instalarDisparadores(ss);
  ss.setActiveSheet(prod);
  return 'Listo';
}

function prepararHoja(ss, def) {
  let sh = ss.getSheetByName(def.nombre);
  if (!sh) sh = ss.insertSheet(def.nombre);
  const actual = sh.getRange(1, 1, 1, def.enc.length).getValues()[0];
  if (actual.join('|') !== def.enc.join('|')) {
    sh.getRange(1, 1, 1, def.enc.length).setValues([def.enc]);
  }
  sh.getRange(1, 1, 1, def.enc.length)
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f4c81');
  sh.setFrozenRows(1);
  def.anchos.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  return sh;
}

function asegurarHojas() {
  const ss = libro();
  const falta = Object.keys(HOJAS).some(k => !ss.getSheetByName(HOJAS[k].nombre));
  if (falta) configurar();
}

function instalarDisparadores(ss) {
  const nuestros = ['alEditarHoja', 'limpieza'];
  ScriptApp.getProjectTriggers()
    .filter(t => nuestros.indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  // onEdit "instalable": el simple no tiene permiso para borrar archivos de Drive
  ScriptApp.newTrigger('alEditarHoja').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('limpieza').timeBased().everyHours(1).create();
}

// Pasa la hoja Config a las tarifas y el contacto actuales (se ejecuta una vez).
function actualizarTarifas() {
  const cfg = libro().getSheetByName(HOJAS.config.nombre);
  const forzar = ['precioDia', 'precioDiaSemana', 'diasSemana', 'precioMes', 'diasMes', 'telefono', 'whatsapp', 'puntosEntrega'];
  const datos = cfg.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    const clave = String(datos[i][0]).trim();
    if (clave === 'precioSemana' || clave === 'deposito') { cfg.deleteRow(i + 1); continue; }   // ya no se usan
    const nuevo = CONFIG_INICIAL.find(r => r[0] === clave);
    if (nuevo && forzar.indexOf(clave) >= 0) cfg.getRange(i + 1, 2, 1, 2).setValues([[nuevo[1], nuevo[2]]]);
  }
  configurar();   // agrega las claves que falten
  return 'Tarifas actualizadas';
}

/* =====================================================================
   API para la página web
   GET  (JSONP): productos, disponibilidad, estadoReserva
   POST (JSON):  reservar (con foto del documento) y avisos (webhook) de Wompi
   ===================================================================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    asegurarHojas();
    switch (p.action) {
      case 'productos':      out = { ok: true, productos: leerProductos() }; break;
      case 'disponibilidad': out = { ok: true, ocupadas: leerOcupadas(), config: configPublica() }; break;
      case 'estadoReserva':  out = estadoReserva(p.reserva, p.idTransaccion); break;
      default:               out = { ok: true, mensaje: 'API Sivarcito activa' };
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  const json = JSON.stringify(out);
  if (p.callback && /^[\w$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    asegurarHojas();
    if (body.action === 'reservar') out = reservar(body);
    else if (campo(body, 'IdTransaccion')) out = webhookWompi(body);
    else out = { ok: false, error: 'Acción no válida.' };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function filas(nombreHoja) {
  const sh = libro().getSheetByName(nombreHoja);
  const v = sh.getDataRange().getValues();
  const enc = v.shift() || [];
  return v.map((r, i) => {
    const o = { _fila: i + 2 };
    enc.forEach((h, k) => o[h] = r[k]);
    return o;
  });
}

// Escribe una fila nueva respetando el orden real de los encabezados de la hoja
function agregarFila(nombreHoja, obj) {
  const sh = libro().getSheetByName(nombreHoja);
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(enc.map(h => (h in obj ? obj[h] : '')));
  return sh.getLastRow();
}

function escribirCampos(nombreHoja, fila, obj) {
  const sh = libro().getSheetByName(nombreHoja);
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(obj).forEach(k => {
    const c = enc.indexOf(k);
    if (c >= 0) sh.getRange(fila, c + 1).setValue(obj[k]);
  });
}

function leerProductos() {
  const grupos = new Map();
  filas(HOJAS.productos.nombre).forEach(r => {
    const nombre = String(r['Producto'] || '').trim();
    if (!nombre || r['Ocultar'] === true) return;
    const cat = String(r['Categoría'] || 'Otros').trim();
    const key = cat + '|' + nombre.toLowerCase();
    let g = grupos.get(key);
    if (!g) {
      g = { id: String(r['ID']), nombre: nombre, categoria: cat, desc: '', etiqueta: '',
            img: icono(cat + ' ' + nombre), fotos: [], variantes: [] };
      grupos.set(key, g);
    }
    if (!g.desc && r['Descripción']) g.desc = String(r['Descripción']).trim();
    if (!g.etiqueta && r['Etiqueta']) g.etiqueta = String(r['Etiqueta']).trim();
    listaFotos(r['Fotos']).forEach(f => { if (g.fotos.indexOf(f) < 0) g.fotos.push(f); });
    g.variantes.push({ t: String(r['Talla'] || '').trim(), precio: Number(r['Precio']) || 0, stock: aStock(r['Cantidad']) });
  });

  return Array.from(grupos.values()).map(g => {
    const v = g.variantes;
    const p = { id: g.id, nombre: g.nombre, categoria: g.categoria, desc: g.desc,
                etiqueta: g.etiqueta, img: g.img, fotos: g.fotos, precio: v[0].precio };
    if (v.length > 1) {
      p.tallas = v.map(x => ({ t: x.t || 'Única', precio: x.precio, stock: x.stock }));
      p.stock = v.every(x => x.stock === 0) ? 0 : null;
    } else {
      p.stock = v[0].stock;
      if (v[0].t) p.talla = v[0].t;
    }
    return p;
  });
}

function listaFotos(celda) {
  return String(celda || '').split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
}

function aStock(c) {
  if (c === '' || c === null || c === undefined) return null;
  if (/agotado/i.test(String(c))) return 0;
  const n = Number(c);
  return isNaN(n) ? null : n;
}

function icono(texto) {
  const t = texto.toLowerCase();
  if (/aud[ií]fono|earbud/.test(t)) return 'audifonos';
  if (/gorra/.test(t)) return 'gorra';
  if (/camisa|hoodie|sudadera/.test(t)) return 'camisa';
  if (/mouse/.test(t)) return 'mouse';
  if (/bolso|mochila/.test(t)) return 'mochila';
  return '';
}

function leerConfig() {
  const c = {};
  filas(HOJAS.config.nombre).forEach(r => { if (r['Clave']) c[String(r['Clave']).trim()] = r['Valor']; });
  return c;
}

// Opción para quien no conoce los puntos fijos: el lugar se acuerda después por WhatsApp
const PUNTO_OTRO = 'Otro punto (coordinar por WhatsApp)';

// Si la clave no existe todavía en Config, se asume que SÍ está en mantenimiento
function enMantenimiento(cfg) {
  const v = cfg.scooterMantenimiento;
  return !/^no$/i.test(String(v === undefined || v === '' ? 'SI' : v).trim());
}

function puntosEntrega(cfg) {
  return String(cfg.puntosEntrega || '').split(/[|\n]+/).map(s => s.trim()).filter(Boolean);
}

function configPublica() {
  const c = leerConfig(), out = {};
  CONFIG_PUBLICA.forEach(k => { if (c[k] !== undefined && c[k] !== '') out[k] = c[k]; });
  out.puntosEntrega = puntosEntrega(c);
  out.scooterMantenimiento = enMantenimiento(c) ? 'SI' : 'NO';
  out.pagosEnLinea = !!credencialesWompi();
  return out;
}

/* =====================================================================
   Tarifas — el ÚNICO lugar donde se decide cuánto se cobra
   ===================================================================== */
function tarifas(cfg) {
  const num = (v, d) => (isNaN(Number(v)) || v === '' ? d : Number(v));
  return {
    dia: num(cfg.precioDia, 20),
    diaSemana: num(cfg.precioDiaSemana, 7),
    diasSemana: num(cfg.diasSemana, 7) || 7,
    mes: num(cfg.precioMes, 180),
    diasMes: num(cfg.diasMes, 30) || 30
  };
}

// 1–6 días: $20/día · 7 días o más: $7/día · cada 30 días: $180 (tope por bloque).
// Ej.: 3 días = $60 · 7 días = $49 · 20 días = $140 · 30 días = $180 · 33 días = $201.
function precioAlquiler(dias, t) {
  if (!dias) return 0;
  const tarifaDia = dias >= t.diasSemana ? t.diaSemana : t.dia;
  const meses = Math.floor(dias / t.diasMes), resto = dias % t.diasMes;
  const total = meses * t.mes + Math.min(resto * tarifaDia, t.mes);
  return Math.round(total * 100) / 100;
}

/* =====================================================================
   Reservas
   ===================================================================== */
function fechaISO(v) {
  if (v instanceof Date) return Utilities.formatDate(v, libro().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function sumarDia(isoStr) {
  const [y, m, d] = isoStr.split('-').map(Number);
  const f = new Date(Date.UTC(y, m - 1, d + 1));
  return f.toISOString().slice(0, 10);
}

function pagoVencido(r, cfg) {
  if (String(r['Estado']).trim() !== 'Pago pendiente') return false;
  const reg = r['Registrado'] instanceof Date ? r['Registrado'].getTime() : 0;
  const minutos = Number(cfg.minutosPago) || 30;
  return reg > 0 && Date.now() - reg > minutos * 60000;
}

function leerOcupadas(excluirReserva) {
  const cfg = leerConfig();
  const set = {};
  filas(HOJAS.reservas.nombre).forEach(r => {
    if (LIBERAN_FECHAS.indexOf(String(r['Estado']).trim()) >= 0) return;
    if (pagoVencido(r, cfg)) return;
    if (excluirReserva && String(r['Reserva']) === excluirReserva) return;
    const ini = fechaISO(r['Inicio']), fin = fechaISO(r['Fin']);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ini) || !/^\d{4}-\d{2}-\d{2}$/.test(fin)) return;
    for (let d = ini, n = 0; d <= fin && n < 400; d = sumarDia(d), n++) set[d] = true;
  });
  return Object.keys(set).sort();
}

// Evita que un texto se interprete como fórmula en la hoja
function limpio(s, max) {
  s = String(s || '').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function soloDigitos(s) { return String(s || '').replace(/\D/g, ''); }

function nuevoIdReserva() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += c[Math.floor(Math.random() * c.length)];
  return 'SV-' + s;
}

function validarDocumento(tipo, numero) {
  if (TIPOS_DOC.indexOf(tipo) < 0) return 'Elige el tipo de documento.';
  const n = String(numero || '').trim().toUpperCase();
  if (tipo === 'DUI') {
    if (!/^\d{8}-?\d$/.test(n)) return 'El DUI debe tener 9 dígitos (ej. 01234567-8).';
  } else if (!/^[A-Z0-9-]{5,20}$/.test(n)) {
    return 'Número de documento inválido.';
  }
  return '';
}

function reservar(p) {
  const reISO = /^\d{4}-\d{2}-\d{2}$/;
  const nombre = String(p.nombre || '').trim(), tel = soloDigitos(p.telefono);
  if (nombre.length < 3) return { ok: false, error: 'Escribe tu nombre completo.' };
  if (tel.length < 8) return { ok: false, error: 'Escribe un teléfono válido.' };
  const errDoc = validarDocumento(p.tipoDoc, p.numDoc);
  if (errDoc) return { ok: false, error: errDoc };
  if (!p.foto) return { ok: false, error: 'Falta la foto del documento.' };
  const cfg = leerConfig();
  if (enMantenimiento(cfg)) return { ok: false, error: 'El scooter está en mantenimiento. Por ahora no hay reservas.' };
  let punto = puntosEntrega(cfg).find(x => x === String(p.punto || '').trim());
  if (!punto && p.punto === PUNTO_OTRO) {
    const sugerido = String(p.puntoOtro || '').trim().slice(0, 120);
    punto = PUNTO_OTRO + (sugerido ? ': ' + sugerido : '');
  }
  if (!punto) return { ok: false, error: 'Elige un punto de entrega.' };
  if (p.acepta !== true) return { ok: false, error: 'Debes aceptar las políticas de alquiler y privacidad.' };
  if (!reISO.test(p.inicio || '') || !reISO.test(p.fin || '')) return { ok: false, error: 'Fechas inválidas.' };
  if (p.fin < p.inicio) return { ok: false, error: 'La fecha de fin es anterior al inicio.' };
  const hoy = Utilities.formatDate(new Date(), libro().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (p.inicio < hoy) return { ok: false, error: 'La fecha de inicio ya pasó.' };

  const pagosEnLinea = !!credencialesWompi();
  const id = nuevoIdReserva();
  let dias = 0, total = 0, foto;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ocupadas = leerOcupadas();
    for (let d = p.inicio; d <= p.fin; d = sumarDia(d)) {
      if (ocupadas.indexOf(d) >= 0) return { ok: false, error: 'Ese rango incluye días ya reservados.' };
      if (++dias > 90) return { ok: false, error: 'El máximo es 90 días.' };
    }
    total = precioAlquiler(dias, tarifas(cfg));   // el precio SIEMPRE lo calcula el servidor
    foto = guardarFotoDocumento(p.foto, id, nombre);

    agregarFila(HOJAS.reservas.nombre, {
      'Registrado': new Date(), 'Reserva': id,
      'Nombre': limpio(nombre, 100), 'Teléfono': limpio(p.telefono, 30), 'Email': limpio(p.email, 120),
      'Tipo documento': p.tipoDoc, 'N.º documento': limpio(String(p.numDoc).toUpperCase(), 30),
      'Foto documento': foto.getUrl(),
      'Punto de entrega': limpio(punto, 180),
      'Inicio': p.inicio, 'Fin': p.fin, 'Días': dias, 'Total': total, 'Notas': limpio(p.notas, 500),
      'Estado': pagosEnLinea ? 'Pago pendiente' : 'Pendiente'
    });
    registrarCliente(p, nombre);
  } finally {
    lock.releaseLock();
  }

  const r = { ok: true, reserva: id, dias: dias, total: total, pagoUrl: null };
  if (!pagosEnLinea) {
    r.mensaje = 'Reserva registrada. Te contactaremos por WhatsApp para confirmarla y coordinar el pago.';
    return r;
  }
  try {
    const enlace = crearEnlacePago(id, total, dias, cfg);
    escribirCampos(HOJAS.reservas.nombre, buscarReserva(id)._fila, { 'Enlace de pago': enlace.url });
    r.pagoUrl = enlace.url;
    r.minutosPago = Number(cfg.minutosPago) || 30;
    r.mensaje = 'Tus fechas quedan apartadas ' + r.minutosPago + ' minutos mientras completas el pago.';
  } catch (err) {
    // Sin enlace: queda como Pendiente para cobrar a mano y no se vence sola
    const fila = buscarReserva(id)._fila;
    escribirCampos(HOJAS.reservas.nombre, fila, { 'Estado': 'Pendiente', 'Enlace de pago': 'Error: ' + String(err.message || err).slice(0, 200) });
    r.mensaje = 'Reserva registrada, pero no se pudo generar el enlace de pago. Te contactaremos por WhatsApp.';
  }
  return r;
}

function buscarReserva(id) {
  if (!id) return null;
  return filas(HOJAS.reservas.nombre).find(r => String(r['Reserva']) === String(id)) || null;
}

function estadoReserva(id, idTransaccion) {
  let r = buscarReserva(id);
  if (!r) return { ok: false, error: 'No encontramos esa reserva.' };
  if (idTransaccion && String(r['Estado']).trim() !== 'Pagada' && credencialesWompi()) {
    try { verificarTransaccion(id, idTransaccion); } catch (err) { /* el webhook lo reintenta */ }
    r = buscarReserva(id);
  }
  const estado = String(r['Estado']).trim();
  return {
    ok: true, reserva: id, estado: estado,
    inicio: fechaISO(r['Inicio']), fin: fechaISO(r['Fin']), dias: r['Días'], total: r['Total'],
    punto: String(r['Punto de entrega'] || ''),
    pagoUrl: estado === 'Pago pendiente' && /^https:/.test(String(r['Enlace de pago'])) ? r['Enlace de pago'] : null
  };
}

// Cada cliente queda registrado una vez (por teléfono) y se actualiza en cada reserva
function registrarCliente(p, nombre) {
  const tel = soloDigitos(p.telefono);
  const ahora = new Date();
  const existente = filas(HOJAS.clientes.nombre).find(r => soloDigitos(r['Teléfono']) === tel);
  const datos = {
    'Nombre': limpio(nombre, 100), 'Email': limpio(p.email, 120),
    'Tipo documento': p.tipoDoc, 'N.º documento': limpio(String(p.numDoc).toUpperCase(), 30), 'Última reserva': ahora
  };
  if (existente) {
    datos['Reservas'] = (Number(existente['Reservas']) || 0) + 1;
    escribirCampos(HOJAS.clientes.nombre, existente._fila, datos);
  } else {
    datos['Teléfono'] = limpio(p.telefono, 30);
    datos['Primera reserva'] = ahora;
    datos['Reservas'] = 1;
    agregarFila(HOJAS.clientes.nombre, datos);
  }
}

/* =====================================================================
   Fotos de documentos en Drive (carpeta privada del dueño)
   ===================================================================== */
function carpetaDocumentos() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CARPETA_DOCUMENTOS');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* se borró: se crea otra */ }
  }
  const f = DriveApp.createFolder('Sivarcito · Documentos de clientes (privado)');
  props.setProperty('CARPETA_DOCUMENTOS', f.getId());
  return f;
}

function guardarFotoDocumento(dataUrl, reserva, nombre) {
  const m = /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('La foto del documento no es una imagen válida.');
  if (m[2].length > 8 * 1024 * 1024) throw new Error('La foto es demasiado grande.');
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const nombreArchivo = reserva + ' - ' + String(nombre).replace(/[^\wÁÉÍÓÚÑáéíóúñ ]/g, '').slice(0, 40) + '.' + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nombreArchivo);
  const archivo = carpetaDocumentos().createFile(blob);
  try { archivo.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (e) { /* ya es privado por defecto */ }
  return archivo;
}

function idDeDrive(url) {
  const m = /\/d\/([-\w]{20,})|[?&]id=([-\w]{20,})/.exec(String(url || ''));
  return m ? (m[1] || m[2]) : '';
}

function borrarArchivo(id) {
  try {
    if (typeof Drive !== 'undefined') { Drive.Files.remove(id); return; }   // definitivo (servicio avanzado)
  } catch (e) { /* sigue con la papelera */ }
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* ya no existía */ }
}

function eliminarFotoDeFila(fila) {
  const sh = libro().getSheetByName(HOJAS.reservas.nombre);
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const c = enc.indexOf('Foto documento') + 1;
  if (!c) return;
  const celda = sh.getRange(fila, c);
  const id = idDeDrive(celda.getValue());
  if (!id) return;
  borrarArchivo(id);
  celda.setValue('Eliminada el ' + Utilities.formatDate(new Date(), libro().getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm'));
}

// Disparador instalable: al cambiar a mano el Estado a Devuelta / Cancelada / Expirada
function alEditarHoja(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== HOJAS.reservas.nombre) return;
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cEstado = enc.indexOf('Estado') + 1;
  if (!cEstado || e.range.getColumn() > cEstado || e.range.getLastColumn() < cEstado) return;
  for (let f = Math.max(2, e.range.getRow()); f <= e.range.getLastRow(); f++) {
    const estado = String(sh.getRange(f, cEstado).getValue()).trim();
    if (BORRAN_FOTO.indexOf(estado) >= 0) eliminarFotoDeFila(f);
  }
}

// Cada hora: los pagos que nunca se completaron pasan a Expirada y se borra su foto
function limpieza() {
  const limite = Date.now() - HORAS_EXPIRA_PAGO * 3600000;
  filas(HOJAS.reservas.nombre).forEach(r => {
    if (String(r['Estado']).trim() !== 'Pago pendiente') return;
    const reg = r['Registrado'] instanceof Date ? r['Registrado'].getTime() : 0;
    if (!reg || reg > limite) return;
    escribirCampos(HOJAS.reservas.nombre, r._fila, { 'Estado': 'Expirada' });
    eliminarFotoDeFila(r._fila);
  });
}

/* =====================================================================
   Wompi (El Salvador) — enlaces de pago
   El monto sale de precioAlquiler(); la página nunca lo envía.
   Revisa los nombres de campos con la documentación de tu panel de Wompi.
   ===================================================================== */
const WOMPI = { token: 'https://id.wompi.sv/connect/token', api: 'https://api.wompi.sv' };

function credencialesWompi() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('WOMPI_CLIENT_ID'), secret = p.getProperty('WOMPI_CLIENT_SECRET');
  return id && secret ? { id: id, secret: secret } : null;
}

function tokenWompi() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get('wompi_token');
  if (guardado) return guardado;
  const c = credencialesWompi();
  if (!c) throw new Error('Wompi no está configurado.');
  const r = UrlFetchApp.fetch(WOMPI.token, {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'client_credentials', client_id: c.id, client_secret: c.secret, audience: 'wompi_api' }
  });
  if (r.getResponseCode() !== 200) throw new Error('Wompi rechazó las credenciales (' + r.getResponseCode() + ').');
  const j = JSON.parse(r.getContentText());
  cache.put('wompi_token', j.access_token, Math.max(60, Math.min(21600, (Number(j.expires_in) || 3600) - 120)));
  return j.access_token;
}

function apiWompi(metodo, ruta, cuerpo) {
  const o = { method: metodo, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + tokenWompi() } };
  if (cuerpo) { o.contentType = 'application/json'; o.payload = JSON.stringify(cuerpo); }
  const r = UrlFetchApp.fetch(WOMPI.api + ruta, o);
  const code = r.getResponseCode(), txt = r.getContentText();
  if (code < 200 || code >= 300) throw new Error('Wompi ' + code + ': ' + txt.slice(0, 300));
  return txt ? JSON.parse(txt) : {};
}

// Busca una propiedad sin importar mayúsculas (Wompi usa PascalCase en el webhook y camelCase en la API)
function campo(obj, nombre) {
  if (!obj || typeof obj !== 'object') return undefined;
  if (nombre in obj) return obj[nombre];
  const k = Object.keys(obj).find(x => x.toLowerCase() === nombre.toLowerCase());
  return k ? obj[k] : undefined;
}

function crearEnlacePago(id, total, dias, cfg) {
  const configuracion = { urlWebhook: ScriptApp.getService().getUrl(), notificarTransaccionCliente: true };
  const sitio = String(cfg.urlSitio || '').trim();
  if (/^https:\/\//.test(sitio)) {
    configuracion.urlRedirect = sitio.replace(/#.*$/, '') + (sitio.indexOf('?') >= 0 ? '&' : '?') + 'reserva=' + encodeURIComponent(id);
  }
  const r = apiWompi('post', '/EnlacePago', {
    identificadorEnlaceComercio: id,
    monto: total,
    nombreProducto: 'Alquiler Scooter Xiaomi 6 · ' + dias + (dias === 1 ? ' día' : ' días') + ' · ' + id,
    formaPago: { permitirTarjetaCreditoDebido: true, permitirPagoConPuntoAgricola: false, permitirPagoEnCuotasAgricola: false },
    configuracion: configuracion,
    limitesDeUso: { cantidadMaximaPagosExitosos: 1 }
  });
  const url = campo(r, 'urlEnlace');
  if (!url) throw new Error('Wompi no devolvió el enlace de pago.');
  return { url: url, id: campo(r, 'idEnlace') };
}

// Aviso de Wompi: no se confía en su contenido; se consulta la transacción a la API.
function webhookWompi(body) {
  const enlace = campo(body, 'EnlacePago') || {};
  const reserva = campo(enlace, 'IdentificadorEnlaceComercio');
  const idTx = campo(body, 'IdTransaccion');
  if (!reserva || !idTx) return { ok: false };
  verificarTransaccion(String(reserva), String(idTx));
  return { ok: true };
}

function verificarTransaccion(reserva, idTx) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = buscarReserva(reserva);
    if (!r) return false;
    if (String(r['Estado']).trim() === 'Pagada') return true;
    // Una transacción solo puede pagar UNA reserva
    const usada = filas(HOJAS.reservas.nombre).some(x => String(x['Transacción Wompi']) === String(idTx));
    if (usada) return false;

    const t = apiWompi('get', '/TransaccionCompra/' + encodeURIComponent(idTx));
    const aprobada = campo(t, 'esAprobada') === true || /^exitosa/i.test(String(campo(t, 'resultadoTransaccion') || ''));
    if (!aprobada) return false;
    const monto = Number(campo(t, 'monto'));
    if (!(Math.abs(monto - Number(r['Total'])) < 0.01)) return false;   // el monto debe ser el que calculó el servidor
    const ident = campo(t, 'identificadorEnlaceComercio') || campo(campo(t, 'enlacePago') || {}, 'identificadorEnlaceComercio');
    if (ident && String(ident) !== reserva) return false;

    const esReal = campo(t, 'esReal');
    const pruebas = /^s[ií]$/i.test(String(leerConfig().wompiPruebas || '').trim());
    if (esReal === false && !pruebas) {
      escribirCampos(HOJAS.reservas.nombre, r._fila, { 'Notas': limpio((r['Notas'] ? r['Notas'] + ' · ' : '') + 'Pago de PRUEBA ignorado (' + idTx + ')', 500) });
      return false;
    }

    const cambios = { 'Estado': 'Pagada', 'Transacción Wompi': String(idTx) };
    const anterior = String(r['Estado']).trim();
    if (anterior !== 'Pago pendiente' && anterior !== 'Pendiente') {
      cambios['Notas'] = limpio((r['Notas'] ? r['Notas'] + ' · ' : '') + '⚠ Pagó cuando estaba "' + anterior + '": revisa fechas y pide de nuevo el documento', 500);
    } else {
      const choque = leerOcupadas(reserva);
      for (let d = fechaISO(r['Inicio']); d <= fechaISO(r['Fin']); d = sumarDia(d)) {
        if (choque.indexOf(d) >= 0) { cambios['Notas'] = limpio((r['Notas'] ? r['Notas'] + ' · ' : '') + '⚠ Pagó tarde y sus fechas chocan con otra reserva', 500); break; }
      }
    }
    escribirCampos(HOJAS.reservas.nombre, r._fila, cambios);
    return true;
  } finally {
    lock.releaseLock();
  }
}
