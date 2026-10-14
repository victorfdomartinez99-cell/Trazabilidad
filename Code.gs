// ============================================================================
// SISTEMA DE TRAZABILIDAD: ADUANAS Y ALMACÉN (WH)
// BACKEND GOOGLE APPS SCRIPT
// ============================================================================

function doGet(e) {
  var htmlOutput = HtmlService.createTemplateFromFile('MenuMovil').evaluate()
    .setTitle('Trazabilidad Aduanas & Almacén')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

function doPost(e) {
  var respuesta;
  try {
    var contenido = e && e.postData && e.postData.contents || '';
    if (!contenido || contenido.length > 1000000) throw new Error('Solicitud no válida.');
    var solicitud = JSON.parse(contenido);
    var secreto = PropertiesService.getScriptProperties().getProperty('VERCEL_API_SECRET');
    if (!secreto || solicitud.secret !== secreto) throw new Error('Solicitud no autorizada.');
    if (typeof solicitud.method !== 'string' || !Array.isArray(solicitud.args) || solicitud.args.length > 8) {
      throw new Error('Solicitud no válida.');
    }

    var metodos = {
      obtenerEstadoBloqueoLogin: obtenerEstadoBloqueoLogin,
      iniciarSesion: iniciarSesion,
      cerrarSesion: cerrarSesion,
      listarAuditoria: listarAuditoria,
      repararEstructuraHojas: repararEstructuraHojas,
      migrarCarrosAValores: migrarCarrosAValores,
      listarUsuarios: listarUsuarios,
      crearUsuario: crearUsuario,
      eliminarUsuarioPermanentemente: eliminarUsuarioPermanentemente,
      agregarClasificaciones: agregarClasificaciones,
      listarPartesAdmin: listarPartesAdmin,
      actualizarParte: actualizarParte,
      eliminarParte: eliminarParte,
      registrarMovimientoUnificado: registrarMovimientoUnificado,
      listarRegistros: listarRegistros,
      listarHistorial: listarHistorial,
      obtenerResumenHXH: obtenerResumenHXH,
      actualizarRegistro: actualizarRegistro,
      obtenerResumenVisual: obtenerResumenVisual,
      obtenerResumenPorDia: obtenerResumenPorDia
    };
    if (!Object.prototype.hasOwnProperty.call(metodos, solicitud.method)) {
      throw new Error('La operación solicitada no está habilitada.');
    }
    var metodo = metodos[solicitud.method];
    respuesta = { ok: true, result: metodo.apply(null, solicitud.args) };
  } catch (error) {
    var mensaje = error && error.message ? error.message : 'Error interno del servidor.';
    console.error('Error en API de Vercel: ' + mensaje);
    respuesta = { ok: false, error: mensaje };
  }
  return ContentService.createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Trazabilidad Móvil')
    .addItem('📱 Abrir Sistema Completo', 'abrirMenuMovil')
    .addSeparator()
    .addItem('Registrar Entrada (Aduanas)', 'abrirEntradasAduanas')
    .addItem('Registrar Salida (Aduanas)', 'abrirSalidasAduanas')
    .addSeparator()
    .addItem('Registrar Entrada (Almacén)', 'abrirEntradasAlmacen')
    .addItem('Registrar Salida (Almacén)', 'abrirSalidasAlmacen')
    .addToUi();
}

function abrirMenuMovil() { mostrarMenu_('Control de Trazabilidad'); }
function abrirEntradasAduanas() { mostrarMenu_('Entradas (Aduanas)'); }
function abrirSalidasAduanas() { mostrarMenu_('Salidas (Aduanas)'); }
function abrirEntradasAlmacen() { mostrarMenu_('Entradas (Almacén)'); }
function abrirSalidasAlmacen() { mostrarMenu_('Salidas (Almacén)'); }

function mostrarMenu_(titulo) {
  var html = HtmlService.createTemplateFromFile('MenuMovil').evaluate()
    .setWidth(480).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, titulo);
}

// Hojas ocultas que almacenan credenciales y sesiones pendientes.
var HOJA_USUARIOS_ = 'Usuarios';
var ROL_DEV_ = 'DEV';
var ROLES_ADMIN_ = ['DEV', 'ADMAD', 'ADMWH', 'ADMQM'];
var OPERACIONES_POR_DEPARTAMENTO_ = {
  UAD: ['aduanas_entrada', 'aduanas_salida'],
  UWH: ['almacen_entrada', 'almacen_salida'],
  UQM: ['calidad_entrada', 'calidad_salida']
};
var OPERACIONES_TODAS_ = ['aduanas_entrada', 'aduanas_salida', 'almacen_entrada', 'almacen_salida', 'calidad_entrada', 'calidad_salida'];
var HOJA_AUDITORIA_ = 'Auditoria Administrativa';
var OPERACION_HEADERS_ = ['ID', 'Fecha', 'Escaneo', 'Rolling', 'Numero de parte', 'No. carro/caja', 'Cantidad', 'Linea', 'Empleado', 'Turno', 'Nombre', 'Carro', 'Corte', 'Comentarios', 'Hora registro'];
var DURACION_SESION_SEGUNDOS_ = 28800;
var MAX_INTENTOS_LOGIN_ = 5;
var DURACION_BLOQUEO_LOGIN_MS_ = 3 * 60 * 1000;
var PROPIEDAD_GENERACION_CACHE_CARROS_ = 'CARRO_CACHE_GENERACION_V1';
var RESUMEN_CACHE_VERSION_ = 'resumen_fecha_v1';
var RESUMEN_CACHE_TTL_SEGUNDOS_ = 21600;
var RESUMEN_CACHE_BLOQUE_FILAS_ = 500;
var RESUMEN_CACHE_MAX_BLOQUE_BYTES_ = 70000;

function nuevoIdRegistro_() { return Utilities.getUuid(); }

function asegurarEstructuraOperaciones_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  OPERACIONES_TODAS_.forEach(function(operacion) {
    var hoja = ss.getSheetByName(nombreHojaOperacion_(operacion));
    if (!hoja) hoja = ss.insertSheet(nombreHojaOperacion_(operacion));
    if (hoja.getLastRow() === 0) hoja.getRange(1, 1, 1, OPERACION_HEADERS_.length).setValues([OPERACION_HEADERS_]);
    else if (String(hoja.getRange(1, 1).getValue()).trim().toUpperCase() !== 'ID') {
      hoja.insertColumnBefore(1);
      hoja.getRange(1, 1).setValue('ID');
    }
    hoja.getRange(1, 1, 1, OPERACION_HEADERS_.length).setValues([OPERACION_HEADERS_]);
    if (hoja.getLastRow() > 1) {
      var ids = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
      var cambios = false;
      ids.forEach(function(fila) {
        if (!String(fila[0] || '').trim()) { fila[0] = nuevoIdRegistro_(); cambios = true; }
      });
      if (cambios) hoja.getRange(2, 1, ids.length, 1).setValues(ids);
    }
    hoja.hideColumns(1);
  });
}

function asegurarAuditoria_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_AUDITORIA_) || ss.insertSheet(HOJA_AUDITORIA_);
  if (hoja.getLastRow() === 0) hoja.appendRow(['Fecha', 'Usuario', 'Rol', 'Accion', 'Entidad', 'Detalle']);
  hoja.hideSheet();
  return hoja;
}

function registrarAuditoria_(usuario, accion, entidad, detalle) {
  var hoja = asegurarAuditoria_();
  hoja.appendRow([new Date(), textoLiteralHoja_(usuario && usuario.empleado || 'SISTEMA'), textoLiteralHoja_(usuario && usuario.rol || 'SISTEMA'), textoLiteralHoja_(accion), textoLiteralHoja_(entidad), textoLiteralHoja_(String(detalle || '').slice(0, 1000))]);
}

function listarAuditoria(token, filtros) {
  var usuario = requireSession_(token);
  if (!esDEV_(usuario)) throw new Error('Solo DEV puede consultar la auditoría administrativa.');
  var hoja = asegurarAuditoria_();
  filtros = filtros || {};
  var pagina = Math.max(Number(filtros.pagina) || 1, 1);
  var porPagina = Math.min(Math.max(Number(filtros.porPagina) || 30, 1), 100);
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 6).getValues() : [];
  var fecha = String(filtros.fecha || '').trim();
  var empleadoFiltro = String(filtros.empleado || '').trim().toUpperCase();
  var departamentoFiltro = String(filtros.departamento || '').trim().toUpperCase();
  var lista = filas.filter(function(fila) {
    var departamento = departamentoDeRol_(fila[2]);
    return (!fecha || formatearFecha(fila[0]) === fecha) &&
      (!empleadoFiltro || String(fila[1]).toUpperCase().indexOf(empleadoFiltro) >= 0) &&
      (!departamentoFiltro || departamento === departamentoFiltro);
  }).reverse().map(function(fila) {
    return { fecha: fila[0], usuario: fila[1], rol: fila[2], departamento: departamentoDeRol_(fila[2]) || '', accion: fila[3], entidad: fila[4], detalle: fila[5] };
  });
  return { items: lista.slice((pagina - 1) * porPagina, pagina * porPagina), total: lista.length, pagina: pagina, porPagina: porPagina, paginas: Math.max(1, Math.ceil(lista.length / porPagina)) };
}

function esUsuarioQM_(usuario) {
  var rol = String(usuario && usuario.rol || '').trim().toUpperCase();
  var departamento = String(usuario && usuario.departamento || '').trim().toUpperCase();
  return departamento === 'UQM' || rol === 'ADMQM';
}

function validarComentarios_(usuario, comentarios, operacion) {
  var texto = String(comentarios || '').trim();
  var esOperacionCalidad = String(operacion || '').indexOf('calidad_') === 0;
  var puedeComentar = esUsuarioQM_(usuario) || (esDEV_(usuario) && esOperacionCalidad);
  if (texto && !puedeComentar) throw new Error('Solo usuarios de Calidad o DEV en operaciones de Calidad pueden agregar o editar comentarios.');
  return puedeComentar ? texto : '';
}

function resolverFilaOperacion_(operacion, referencia) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
  if (!hoja) throw new Error('Registro no encontrado.');
  var valor = referencia && typeof referencia === 'object' ? (referencia.id || referencia.registroId || referencia.fila) : referencia;
  var id = String(valor || '').trim();
  if (id && hoja.getLastRow() > 1) {
    var ids = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
    var indiceId = ids.findIndex(function(fila) { return String(fila[0]).trim() === id; });
    if (indiceId >= 0) return { hoja: hoja, fila: indiceId + 2, id: id };
  }
  var fila = Number(valor);
  if (!isFinite(fila) || fila < 2 || fila > hoja.getLastRow()) throw new Error('Registro no encontrado.');
  var idFila = String(hoja.getRange(fila, 1).getValue() || '').trim();
  if (!idFila) {
    idFila = nuevoIdRegistro_();
    hoja.getRange(fila, 1).setValue(idFila);
  }
  return { hoja: hoja, fila: fila, id: idFila };
}

function esDEV_(usuario) { return usuario && String(usuario.rol || '').trim().toUpperCase() === ROL_DEV_; }
function esAdministrador_(usuario) { return usuario && ROLES_ADMIN_.indexOf(String(usuario.rol || '').trim().toUpperCase()) >= 0; }
function operacionesDeUsuario_(usuario) {
  if (esDEV_(usuario)) return OPERACIONES_TODAS_.slice();
  return (OPERACIONES_POR_DEPARTAMENTO_[String(usuario.departamento || '').trim().toUpperCase()] || []).slice();
}

function ejecutarConBloqueo_(accion) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Hay otra operación en proceso. Inténtalo nuevamente.');
  try {
    return accion();
  } finally {
    lock.releaseLock();
  }
}

function configurarSeguridadInicial() {
  asegurarHojasSeguridad_();
  return 'Seguridad inicializada.';
}

function repararEstructuraHojas(token) {
  var usuario = requireSession_(token);
  if (!esDEV_(usuario)) throw new Error('Solo DEV puede reparar la estructura de hojas.');
  return ejecutarConBloqueo_(function() {
    asegurarHojasSeguridad_();
    return 'Estructura de hojas reparada correctamente.';
  });
}

function escribirFilasContiguas_(hoja, filaInicial, columna, valoresPorFila) {
  var inicioLote = -1;
  var lote = [];
  function escribirLote_() {
    if (!lote.length) return;
    hoja.getRange(filaInicial + inicioLote, columna, lote.length, lote[0].length).setValues(lote);
    inicioLote = -1;
    lote = [];
  }
  valoresPorFila.forEach(function(valores, indice) {
    if (valores === null) {
      escribirLote_();
      return;
    }
    if (inicioLote < 0) inicioLote = indice;
    lote.push(valores);
  });
  escribirLote_();
}

function asegurarHojasSeguridad_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usuarios = ss.getSheetByName(HOJA_USUARIOS_) || ss.insertSheet(HOJA_USUARIOS_);
  if (usuarios.getLastRow() === 0) usuarios.appendRow(['Empleado', 'Nombre', 'Rol', 'Departamento', 'Hash', 'Salt', 'Estado', 'Intentos', 'Fecha bloqueo', 'Creado por', 'Reservado', 'CambioRequerido', 'Fecha del registro']);
  if (usuarios.getLastColumn() < 13) usuarios.getRange(1, 13).setValue('Fecha del registro');
  var filas = usuarios.getLastRow() > 1 ? usuarios.getRange(2, 1, usuarios.getLastRow() - 1, 13).getValues() : [];
  var indiceDev = filas.findIndex(function(fila) { return String(fila[2]).toUpperCase() === ROL_DEV_; });
  var valoresHashSalt = [];
  var valoresReservados = [];
  var valoresReservadosDev = [];
  var fechasFaltantes = [];
  filas.forEach(function(fila, indice) {
    valoresHashSalt[indice] = null;
    valoresReservados[indice] = null;
    valoresReservadosDev[indice] = null;
    fechasFaltantes[indice] = null;
    if (String(fila[2]).toUpperCase() === 'DEV') {
      if (indice === indiceDev) valoresReservadosDev[indice] = [''];
      return;
    }
    var empleado = String(fila[0] || '').trim();
    if (!empleado) return;
    var salt = String(fila[5] || '');
    valoresHashSalt[indice] = [hashPassword_(empleado, salt), salt];
    valoresReservados[indice] = ['', 'NO'];
    if (!fila[12]) fechasFaltantes[indice] = [new Date()];
  });
  escribirFilasContiguas_(usuarios, 2, 5, valoresHashSalt);
  escribirFilasContiguas_(usuarios, 2, 11, valoresReservados);
  escribirFilasContiguas_(usuarios, 2, 11, valoresReservadosDev);
  escribirFilasContiguas_(usuarios, 2, 13, fechasFaltantes);
  ['Entradas de calidad', 'Salidas de calidad'].forEach(function(nombre) {
    var hojaCalidad = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    if (hojaCalidad.getLastRow() === 0) hojaCalidad.appendRow(['Fecha', 'Escaneo', 'Rolling', 'Numero de parte', 'Numero de carro', 'Cantidad', 'Linea', 'Empleado', 'Nombre', 'Auxiliar', 'Carro', 'Corte', 'Comentarios']);
    asegurarEstatusCalidad_(hojaCalidad);
  });
  asegurarEstructuraOperaciones_();
  ['Entradas de calidad', 'Salidas de calidad'].forEach(function(nombre) {
    reestructurarRegistrosCalidad_(ss.getSheetByName(nombre));
  });
  asegurarAuditoria_();
  usuarios.hideSheet();
}

function recalcularCarrosComoValores_(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return 0;
  var cantidad = hoja.getLastRow() - 1;
  var registros = hoja.getRange(2, 2, cantidad, 12).getValues();
  var zonaHoraria = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'GMT';
  var normalizar = function(valor) {
    if (valor instanceof Date) return Utilities.formatDate(valor, zonaHoraria, 'yyyy-MM-dd');
    return String(valor == null ? '' : valor).trim();
  };
  var grupos = {};
  var esCalidad = hoja.getName().indexOf('calidad') >= 0;
  var carros = registros.map(function(actual, indice) {
    if (!normalizar(actual[0]) || !normalizar(actual[3]) || !normalizar(actual[6]) || !normalizar(actual[8]) || (!esCalidad && !normalizar(actual[11]))) return [''];
    var componentes = [actual[0], actual[3], actual[6], actual[8]];
    if (!esCalidad) componentes.push(actual[11]);
    var clave = componentes.map(normalizar).join('|');
    grupos[clave] = (grupos[clave] || 0) + 1;
    return ['Carro ' + grupos[clave]];
  });
  hoja.getRange(2, 12, cantidad, 1).setValues(carros);
  return cantidad;
}

function calcularCarroNuevo_(hoja, fila) {
  var actual = hoja.getRange(fila, 2, 1, 12).getValues()[0];
  var zonaHoraria = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'GMT';
  var normalizar = function(valor) {
    if (valor instanceof Date) return Utilities.formatDate(valor, zonaHoraria, 'yyyy-MM-dd');
    return String(valor == null ? '' : valor).trim();
  };
  var esCalidad = hoja.getName().indexOf('calidad') >= 0;
  if (!normalizar(actual[0]) || !normalizar(actual[3]) || !normalizar(actual[6]) || !normalizar(actual[8]) || (!esCalidad && !normalizar(actual[11]))) return '';
  var previos = fila > 2 ? hoja.getRange(2, 2, fila - 2, 12).getValues() : [];
  var total = 1;
  previos.forEach(function(registro) {
    var coincide = normalizar(registro[0]) === normalizar(actual[0]) &&
      normalizar(registro[3]) === normalizar(actual[3]) &&
      normalizar(registro[6]) === normalizar(actual[6]) &&
      normalizar(registro[8]) === normalizar(actual[8]);
    if (!esCalidad) coincide = coincide && normalizar(registro[11]) === normalizar(actual[11]);
    if (coincide) total++;
  });
  return 'Carro ' + total;
}

function normalizarTextoCache_(valor) {
  return String(valor == null ? '' : valor).replace(/[^a-zA-Z0-9_-]/g, '_');
}

function fechaLogicaCarroCache_(datos) {
  var opciones = [
    datos && datos.fecha,
    datos && datos.fechaOperacion,
    datos && datos.fechaRegistro,
    datos && datos.fechaLogica
  ];
  for (var i = 0; i < opciones.length; i++) {
    var valor = String(opciones[i] == null ? '' : opciones[i]).trim();
    if (valor) return valor;
  }

  var zonaHoraria = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'GMT';
  return Utilities.formatDate(new Date(), zonaHoraria, 'yyyy-MM-dd');
}

function claveCarroCache_(operacion, token, datos, generacion) {
  var nombre = String(operacion || '').trim();
  if (!nombre || nombre.indexOf('calidad_') === 0) return null;
  var fecha = fechaLogicaCarroCache_(datos);
  var escaneo = String((datos && datos.escaneo) || '').trim().toUpperCase();
  var segmentosEscaneo = escaneo.split('-');
  var parte = segmentosEscaneo.length >= 3 ? String(segmentosEscaneo[2] || '').trim() : '';
  if (!parte) parte = String((datos && datos.parte) || '').trim();
  var linea = String((datos && datos.linea) || '').trim();
  var turno = String((datos && datos.turno) || '').trim();
  var corte = String((datos && datos.corte) || '').trim();
  var base = [nombre, fecha, parte, linea, turno, corte].join('|');
  var version = generacion;
  if (version == null) version = PropertiesService.getScriptProperties().getProperty(PROPIEDAD_GENERACION_CACHE_CARROS_) || '';
  if (version) base = String(version) + '|' + base;
  return 'carro_cache_' + normalizarTextoCache_(base);
}

function registrarClaveCarroCache_(operacion, token, clave) {
  if (!clave) return;
}

function limpiarCarroCacheSesion_(token) {
  return;
}

function leerUltimoCarroCache_(operacion, token, datos) {
  var clave = claveCarroCache_(operacion, token, datos);
  if (!clave) return null;
  var valor = CacheService.getScriptCache().get(clave);
  if (!valor) return null;
  var numero = Number(String(valor).trim());
  return isFinite(numero) ? numero : null;
}

function guardarUltimoCarroCache_(operacion, token, numero, datos) {
  var clave = claveCarroCache_(operacion, token, datos);
  if (!clave) return;
  registrarClaveCarroCache_(operacion, token, clave);
  CacheService.getScriptCache().put(clave, String(Math.max(1, Number(numero) || 1)), 28800);
}

function obtenerCarroAsignadoDesdeCacheOHoja_(hoja, operacion, token, nuevaFila, datos) {
  var nombreHoja = String(hoja && hoja.getName ? hoja.getName() : '').toLowerCase();
  if (nombreHoja.indexOf('calidad') >= 0) return calcularCarroNuevo_(hoja, nuevaFila);

  var ultimoCarro = leerUltimoCarroCache_(operacion, token, datos);
  if (ultimoCarro !== null) {
    var siguiente = ultimoCarro + 1;
    guardarUltimoCarroCache_(operacion, token, siguiente, datos);
    return 'Carro ' + siguiente;
  }

  var carroCalculado = calcularCarroNuevo_(hoja, nuevaFila);
  var numero = Number(String(carroCalculado || '').replace(/\D+/g, ''));
  if (!isFinite(numero) || numero < 1) numero = 1;
  guardarUltimoCarroCache_(operacion, token, numero, datos);
  return carroCalculado || ('Carro ' + numero);
}

function asegurarEstatusCalidad_(hoja) {
  if (!hoja) return;
  hoja.getRange(1, 16).setValue('Estatus');
}

function reestructurarRegistrosCalidad_(hoja) {
  return;
}

function migrarCarrosAValoresUnaVez_() {
  var propiedades = PropertiesService.getScriptProperties();
  if (propiedades.getProperty('CARRO_VALOR_V3') === 'SI') return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  OPERACIONES_TODAS_.forEach(function(operacion) {
    recalcularCarrosComoValores_(ss.getSheetByName(nombreHojaOperacion_(operacion)));
  });
  propiedades.setProperty('CARRO_VALOR_V3', 'SI');
}

function migrarCarrosAValores(token) {
  var usuario = requireSession_(token);
  if (!esDEV_(usuario)) throw new Error('Solo el usuario DEV puede recalcular la columna de carros.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Hay otro registro en proceso. Espera un momento e inténtalo de nuevo.');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultado = [];
    OPERACIONES_TODAS_.forEach(function(operacion) {
      var hoja = ss.getSheetByName(nombreHojaOperacion_(operacion));
      var cantidad = recalcularCarrosComoValores_(hoja);
      resultado.push(nombreHojaOperacion_(operacion) + ': ' + cantidad + ' registros');
    });
    SpreadsheetApp.flush();

    var propiedades = PropertiesService.getScriptProperties();
    var generacion = String((Number(propiedades.getProperty(PROPIEDAD_GENERACION_CACHE_CARROS_)) || 0) + 1);
    propiedades.setProperty(PROPIEDAD_GENERACION_CACHE_CARROS_, generacion);
    sincronizarCacheCarrosConHojas_(ss, generacion);
    propiedades.setProperty('CARRO_VALOR_V3', 'SI');
    return 'Migración completada. ' + resultado.join(' | ');
  } finally {
    lock.releaseLock();
  }
}

function sincronizarCacheCarrosConHojas_(ss, generacion) {
  var cache = CacheService.getScriptCache();
  var zonaHoraria = ss.getSpreadsheetTimeZone() || 'GMT';
  var contadores = {};

  OPERACIONES_TODAS_.forEach(function(operacion) {
    if (operacion.indexOf('calidad_') === 0) return;
    var hoja = ss.getSheetByName(nombreHojaOperacion_(operacion));
    if (!hoja || hoja.getLastRow() < 2) return;

    var filas = hoja.getRange(2, 2, hoja.getLastRow() - 1, 12).getValues();
    filas.forEach(function(fila) {
      var fecha = fila[0] instanceof Date
        ? Utilities.formatDate(fila[0], zonaHoraria, 'yyyy-MM-dd')
        : String(fila[0] == null ? '' : fila[0]).trim();
      var numeroCarro = Number(String(fila[10] || '').replace(/\D+/g, ''));
      var datos = {
        fecha: fecha,
        escaneo: String(fila[1] || ''),
        parte: String(fila[3] || ''),
        linea: fila[6],
        turno: fila[8],
        corte: fila[11]
      };
      if (!fecha || !datos.parte || !datos.linea || !datos.turno || !datos.corte || !isFinite(numeroCarro) || numeroCarro < 1) return;

      var clave = claveCarroCache_(operacion, '', datos, generacion);
      if (clave) contadores[clave] = String(numeroCarro);
    });
  });

  var claves = Object.keys(contadores);
  for (var inicio = 0; inicio < claves.length; inicio += 100) {
    var lote = {};
    claves.slice(inicio, inicio + 100).forEach(function(clave) {
      lote[clave] = contadores[clave];
    });
    cache.putAll(lote, 28800);
  }
  return claves.length;
}

function obtenerEstadoBloqueoLogin(empleado) {
  var id = String(empleado || '').trim();
  if (!id) return { bloqueado: false, restanteMs: 0, bloqueadaHasta: '' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_USUARIOS_);
  if (!hoja || hoja.getLastRow() <= 1) return { bloqueado: false, restanteMs: 0, bloqueadaHasta: '' };
  var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 12).getValues();
  var fila = filas.find(function(item) { return String(item[0]).trim() === id; });
  if (!fila) return { bloqueado: false, restanteMs: 0, bloqueadaHasta: '' };
  var bloqueadaHasta = fila[8] ? new Date(fila[8]).getTime() : 0;
  var ahora = Date.now();
  var restanteMs = isFinite(bloqueadaHasta) && bloqueadaHasta > ahora ? Math.max(0, bloqueadaHasta - ahora) : 0;
  return { bloqueado: !!restanteMs, restanteMs: restanteMs, bloqueadaHasta: restanteMs ? new Date(bloqueadaHasta).toISOString() : '' };
}

function iniciarSesion(empleado, password) {
  return ejecutarConBloqueo_(function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_USUARIOS_);
  if (!hoja || hoja.getLastRow() <= 1) {
    asegurarHojasSeguridad_();
    hoja = ss.getSheetByName(HOJA_USUARIOS_);
  }
  if (!hoja) throw new Error('Empleado o contraseña incorrectos.');
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 12).getValues() : [];
  var indice = filas.findIndex(function(fila) { return String(fila[0]).trim() === String(empleado).trim(); });
  if (indice < 0) throw new Error('Empleado o contraseña incorrectos.');
  var fila = filas[indice];
  var real = indice + 2;
  if (String(fila[6] || 'ACTIVO') !== 'ACTIVO') {
    throw new Error('Usuario inactivo.');
  }
  var ahora = Date.now();
  var bloqueadaHasta = fila[8] ? new Date(fila[8]).getTime() : 0;
  var intentos = Math.max(0, Number(fila[7]) || 0);
  if (isFinite(bloqueadaHasta) && bloqueadaHasta > ahora) {
    throw new Error('Cuenta bloqueada temporalmente. Inténtalo más tarde.');
  }
  if (fila[8] && isFinite(bloqueadaHasta) && bloqueadaHasta <= ahora) {
    intentos = 0;
  }
  if (hashPassword_(password, fila[5]) !== String(fila[4])) {
    intentos++;
    var nuevoBloqueo = intentos >= MAX_INTENTOS_LOGIN_ ? new Date(ahora + DURACION_BLOQUEO_LOGIN_MS_) : '';
    hoja.getRange(real, 8, 1, 2).setValues([[intentos, nuevoBloqueo]]);
    throw new Error('Empleado o contraseña incorrectos.');
  }
  hoja.getRange(real, 8, 1, 2).setValues([[0, '']]);
  var cambioRequerido = String(fila[11] || 'NO').toUpperCase() === 'SI';
  var usuario = {
    empleado: String(fila[0]),
    nombre: String(fila[1]),
    rol: String(fila[2]),
    departamento: String(fila[3]),
    requiereCambio: false
  };
  var token = Utilities.getUuid();
  var datosSesion = JSON.stringify(usuario);
  if (esDEV_(usuario)) {
    PropertiesService.getScriptProperties().setProperty('sesion_dev_' + token, datosSesion);
  } else {
    CacheService.getScriptCache().put('sesion_' + token, datosSesion, DURACION_SESION_SEGUNDOS_);
  }
  return { token: token, usuario: usuario, catalogos: catalogosParaSesion_(usuario) };
  });
}

function cerrarSesion(token) {
  if (token) {
    CacheService.getScriptCache().remove('sesion_' + token);
    PropertiesService.getScriptProperties().deleteProperty('sesion_dev_' + token);
    limpiarCarroCacheSesion_(token);
  }
  return true;
}

function requireSession_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(token)) {
    throw new Error('Sesión expirada. Inicia sesión nuevamente.');
  }
  var raw = CacheService.getScriptCache().get('sesion_' + token);
  if (!raw && token) raw = PropertiesService.getScriptProperties().getProperty('sesion_dev_' + token);
  if (!raw) throw new Error('Sesión expirada. Inicia sesión nuevamente.');
  var usuario = JSON.parse(raw);
  if (!esDEV_(usuario)) CacheService.getScriptCache().put('sesion_' + token, raw, DURACION_SESION_SEGUNDOS_);
  return usuario;
}

function hashPassword_(password, salt) {
  var texto = salt ? String(salt) + '|' + String(password || '') : String(password || '');
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function(byte) { var value = byte < 0 ? byte + 256 : byte; return ('0' + value.toString(16)).slice(-2); }).join('');
}

function limpiarTextoSeguro_(valor, maxLength) {
  var texto = String(valor == null ? '' : valor)
    .replace(/[\u0000-\u001F\u007F]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (maxLength && texto.length > maxLength) texto = texto.slice(0, maxLength);
  return texto;
}

function textoLiteralHoja_(valor) {
  var texto = String(valor == null ? '' : valor);
  return /^[\t\r\n ]*[=+\-@]/.test(texto) ? "'" + texto : texto;
}

function limpiarIdSeguro_(valor, maxLength) {
  var texto = String(valor == null ? '' : valor).replace(/[^A-Za-z0-9_.@_\-/]/g, '').trim();
  if (maxLength && texto.length > maxLength) texto = texto.slice(0, maxLength);
  return texto;
}

function limpiarCodigoEscaneo_(valor, obligatorio) {
  var texto = String(valor == null ? '' : valor).replace(/[\u0000-\u001F\u007F]/g, '').trim().toUpperCase();
  if (!texto && !obligatorio) return '';
  if (!texto || texto.length > 80 || /^[=+@]/.test(texto)) throw new Error('El código de etiqueta no es válido.');
  return texto;
}

function limpiarNumeroParte_(valor) {
  var texto = limpiarTextoSeguro_(valor, 40).toUpperCase();
  if (!texto || texto.length > 40 || /^[=+\-@]/.test(texto)) throw new Error('El número de parte no es válido.');
  return texto;
}

function leerMatrizEmpleados_() {
  var catalogos = { empleados: {}, empleadosPorDepartamento: {} };
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de empleados');
  if (!hoja || hoja.getLastRow() <= 1) return catalogos;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues().forEach(function(fila) {
    var id = String(fila[0]).replace(/\.0$/, '').trim();
    if (!id) return;
    var nombre = String(fila[1]).trim();
    var departamento = String(fila[2] || '').trim().toUpperCase();
    catalogos.empleados[id] = nombre;
    if (!catalogos.empleadosPorDepartamento[departamento]) catalogos.empleadosPorDepartamento[departamento] = {};
    catalogos.empleadosPorDepartamento[departamento][id] = nombre;
  });
  return catalogos;
}

function obtenerCatalogosBase_() {
  var cache = CacheService.getScriptCache();
  var cacheado = cache.get('catalogos_base');
  if (cacheado) {
    var catalogosCacheados = JSON.parse(cacheado);
    if (!catalogosCacheados.empleadosPorDepartamento) {
      var matrizEmpleadosCacheada = leerMatrizEmpleados_();
      catalogosCacheados.empleados = matrizEmpleadosCacheada.empleados;
      catalogosCacheados.empleadosPorDepartamento = matrizEmpleadosCacheada.empleadosPorDepartamento;
      cache.put('catalogos_base', JSON.stringify(catalogosCacheados), 60);
    }
    return catalogosCacheados;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var catalogos = {
    empleados: {}, partes: {}, fechaAduanas: '', fechaAlmacen: '',
    turnoActual: 'A', corteActual: '1', empleadosPorDepartamento: {}
  };

  try {
    var matrizEmpleados = leerMatrizEmpleados_();
    catalogos.empleados = matrizEmpleados.empleados;
    catalogos.empleadosPorDepartamento = matrizEmpleados.empleadosPorDepartamento;

    var hojaFG = ss.getSheetByName('Matriz de Finish God') || ss.getSheetByName('Matriz de Finish Goods');
    if (hojaFG && hojaFG.getLastRow() > 1) {
      var valsFG = hojaFG.getRange(2, 1, hojaFG.getLastRow() - 1, 2).getValues();
      valsFG.forEach(function(fila) {
        var noParte = normalizarParteServidor_(fila[0]);
        if (noParte) catalogos.partes[noParte] = String(fila[1]).trim();
      });
    }

    var resumenAduanas = ss.getSheetByName('Resumen aduanas');
    if (resumenAduanas) {
      var metadatosAduanas = resumenAduanas.getRange('C1:L1').getValues()[0];
      catalogos.fechaAduanas = formatearFecha(metadatosAduanas[0]);
      catalogos.turnoActual = String(metadatosAduanas[4] || 'A').trim();
      catalogos.corteActual = String(metadatosAduanas[9] || '1').trim();
    }

    var resumenWH = ss.getSheetByName('Resumen WH');
    if (resumenWH) catalogos.fechaAlmacen = formatearFecha(resumenWH.getRange('C1').getValue());
  } catch (err) {
    Logger.log('Error cargando catálogos: ' + err.message);
  }
  cache.put('catalogos_base', JSON.stringify(catalogos), 60);
  return catalogos;
}

function normalizarParteServidor_(valor) {
  return String(valor || '').replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ').replace(/\s*\/\s*/g, '/').toUpperCase();
}

function catalogosParaSesion_(usuario) {
  var catalogos = obtenerCatalogosBase_();
  var departamento = departamentoDeRol_(usuario.rol);
  if (usuario.rol !== 'DEV' && departamento) {
    catalogos.empleados = catalogos.empleadosPorDepartamento[departamento] || {};
  }
  delete catalogos.empleadosPorDepartamento;
  return catalogos;
}

function tipoParteHistorial_(parte) {
  var tipos = obtenerCatalogosBase_().partes;
  return String(tipos[normalizarParteServidor_(parte)] || '').trim().toUpperCase();
}

function obtenerCatalogosIniciales(token) {
  var usuario = requireSession_(token);
  var catalogos = obtenerCatalogosBase_();
  delete catalogos.empleadosPorDepartamento;
  catalogos.usuario = usuario;
  return catalogos;
}

function puedeUsarOperacion_(usuario, operacion) {
  return operacionesDeUsuario_(usuario).indexOf(String(operacion || '').trim()) >= 0;
}

function departamentoDeRol_(rol) {
  var valor = String(rol || '').trim().toUpperCase();
  if (valor === 'UAD' || valor === 'ADMAD') return 'UAD';
  if (valor === 'UWH' || valor === 'ADMWH') return 'UWH';
  if (valor === 'UQM' || valor === 'ADMQM') return 'UQM';
  return 'TODOS';
}

function puedeAdministrarUsuarios_(usuario) { return esAdministrador_(usuario); }
function puedeConsultarMatrizPartes_(usuario) {
  return esAdministrador_(usuario) || ['UAD', 'UWH', 'UQM'].indexOf(String(usuario && usuario.rol || '').trim().toUpperCase()) >= 0;
}

function puedeGestionarUsuario_(solicitante, objetivo) {
  if (esDEV_(solicitante)) return objetivo[2] !== ROL_DEV_;
  return esAdministrador_(solicitante) && objetivo[2] !== ROL_DEV_ && objetivo[3] === departamentoDeRol_(solicitante.rol);
}

function rolAdministradorDepartamento_(departamento) {
  return departamento === 'UAD' ? 'ADMAD' : departamento === 'UWH' ? 'ADMWH' : departamento === 'UQM' ? 'ADMQM' : '';
}

function puedeGestionarOperacion_(usuario, operacion) {
  return esAdministrador_(usuario) && (esDEV_(usuario) || puedeUsarOperacion_(usuario, operacion));
}

function puedeEditarRegistro_(usuario, operacion, cambios) {
  if (esAdministrador_(usuario)) return true;
  if (!puedeUsarOperacion_(usuario, operacion)) return false;
  var claves = Object.keys(cambios || {});
  if (!claves.length) return false;
  var permitidos = ['cantidad', 'carroCaja'];
  return claves.every(function(clave) { return permitidos.indexOf(clave) >= 0; });
}

function puedeRecibirSolicitudEliminacion_(usuario, operacion) {
  var rol = String(usuario.rol || '').trim().toUpperCase();
  if (rol === 'DEV') return true;
  var departamentosPorOperacion = {
    aduanas: 'ADMAD',
    almacen: 'ADMWH',
    calidad: 'ADMQM'
  };
  var prefijo = String(operacion || '').trim().toLowerCase().split('_')[0];
  return departamentosPorOperacion[prefijo] === rol;
}

function leerRegistroParaEliminar_(operacion, fila) {
  if (!nombreHojaOperacion_(operacion)) throw new Error('Operación no válida.');
  var referencia = resolverFilaOperacion_(operacion, fila);
  var hoja = referencia.hoja;
  var numeroFila = referencia.fila;
  var valores = hoja.getRange(numeroFila, 2, 1, 13).getValues()[0];
  return {
    hoja: hoja,
    fila: numeroFila,
    id: referencia.id,
    fecha: formatearFecha(valores[0]),
    empleado: String(valores[7] || '').trim(),
    nombre: String(valores[8] || '').trim(),
    parte: String(valores[3] || '').trim(),
    cantidad: valores[5],
    carroCaja: String(valores[4] || '').trim(),
    carro: String(valores[10] || '').trim(),
    corte: String(valores[11] || '').trim()
  };
}

function buscarUsuario(token, empleado) {
  var solicitante = requireSession_(token);
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_USUARIOS_);
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 10).getValues() : [];
  var fila = filas.find(function(item) { return String(item[0]).trim() === String(empleado).trim(); });
  if (!fila) return null;
  if (solicitante.rol !== 'DEV' && fila[3] !== departamentoDeRol_(solicitante.rol)) throw new Error('No puedes consultar usuarios de otro departamento.');
  return { empleado: fila[0], nombre: fila[1], rol: fila[2], departamento: fila[3], estado: fila[6], intentos: fila[7] };
}

function listarUsuarios(token, pagina, porPagina) {
  var usuario = requireSession_(token);
  if (!puedeAdministrarUsuarios_(usuario)) throw new Error('No tienes permiso para administrar usuarios.');
  var cache = CacheService.getScriptCache();
  var claveCache = 'usuarios_admin_' + String(usuario.rol || '').toUpperCase();
  var cacheado = cache.get(claveCache);
  var lista;
  if (cacheado) {
    lista = JSON.parse(cacheado);
  } else {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_USUARIOS_);
    var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 12).getValues() : [];
    lista = filas.filter(function(fila) { return usuario.rol === 'DEV' || fila[3] === departamentoDeRol_(usuario.rol); })
      .map(function(fila) { return { empleado: fila[0], nombre: fila[1], rol: fila[2], departamento: fila[3], estado: fila[6], intentos: fila[7], cambioRequerido: false }; });
    cache.put(claveCache, JSON.stringify(lista), 60);
  }
  if (String(pagina).toUpperCase() === 'TODAS') {
    return { items: lista, total: lista.length, pagina: 1, porPagina: lista.length, paginas: 0 };
  }
  pagina = Math.max(Number(pagina) || 1, 1); porPagina = Math.min(Math.max(Number(porPagina) || 10, 1), 50);
  return { items: lista.slice((pagina - 1) * porPagina, pagina * porPagina), total: lista.length, pagina: pagina, porPagina: porPagina, paginas: Math.max(1, Math.ceil(lista.length / porPagina)) };
}

function invalidarCacheUsuarios_() {
  var cache = CacheService.getScriptCache();
  ['DEV', 'ADMAD', 'ADMWH', 'ADMQM'].forEach(function(rol) { cache.remove('usuarios_admin_' + rol); });
}

function crearUsuario(token, datos) {
  var creador = requireSession_(token);
  if (!puedeAdministrarUsuarios_(creador)) throw new Error('No tienes permiso para administrar usuarios.');
  var empleado = limpiarIdSeguro_(datos && datos.empleado, 20);
  var nombre = limpiarTextoSeguro_(datos && datos.nombre, 80);
  var rolSeleccionado = limpiarIdSeguro_(datos && datos.rol, 10).toUpperCase();
  var departamento = limpiarIdSeguro_(datos && datos.departamento, 10).toUpperCase();
  if (!empleado || !nombre) throw new Error('Empleado y nombre son obligatorios.');
  var rolesValidos = ['U', 'ADM'];
  if (!rolesValidos.includes(rolSeleccionado) || !['UAD', 'UWH', 'UQM'].includes(departamento)) throw new Error('Rol o departamento no válido.');
  var rol = rolSeleccionado === 'U' ? departamento : 'ADM' + departamento.slice(1);
  if (rol === 'ADMAD' && departamento !== 'UAD' || rol === 'ADMWH' && departamento !== 'UWH' || rol === 'ADMQM' && departamento !== 'UQM') throw new Error('El administrador debe pertenecer a su departamento.');
  if (creador.rol !== 'DEV' && (rolSeleccionado !== 'U' || departamento !== departamentoDeRol_(creador.rol))) throw new Error('Solo puedes crear usuarios estándar de tu departamento.');
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_USUARIOS_);
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 12).getValues() : [];
  if (filas.some(function(fila) { return String(fila[0]) === empleado; })) throw new Error('Ese empleado ya existe.');
  var salt = Utilities.getUuid();
  var password = empleado;
  hoja.appendRow([empleado, textoLiteralHoja_(nombre), rol, departamento, hashPassword_(password, salt), salt, 'ACTIVO', 0, '', creador.empleado, '', 'NO', new Date()]);
  registrarEmpleadoEnMatriz_(empleado, nombre, departamento);
  registrarAuditoria_(creador, 'CREAR_USUARIO', empleado, 'Rol: ' + rol + ', departamento: ' + departamento);
  invalidarCacheUsuarios_();
  return true;
}

function registrarEmpleadoEnMatriz_(empleado, nombre, departamento) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName('Matriz de empleados');
  if (!hoja) hoja = ss.insertSheet('Matriz de empleados');
  if (hoja.getLastRow() === 0) hoja.appendRow(['Numero de empleado', 'Nombre', 'Departamento']);
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues() : [];
  var id = String(empleado).trim();
  var indice = filas.findIndex(function(fila) { return String(fila[0]).trim() === id; });
  if (indice >= 0) hoja.getRange(indice + 2, 1, 1, 3).setValues([[id, textoLiteralHoja_(String(nombre).trim()), departamento]]);
  else hoja.appendRow([id, textoLiteralHoja_(String(nombre).trim()), departamento]);
  CacheService.getScriptCache().remove('catalogos_base');
}

function eliminarEmpleadoDeMatriz_(empleado) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de empleados');
  if (!hoja || hoja.getLastRow() < 2) return;
  var filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, 3).getValues();
  for (var i = filas.length - 1; i >= 0; i--) {
    if (String(filas[i][0]).replace(/\.0$/, '').trim() === String(empleado).trim()) hoja.deleteRow(i + 2);
  }
  CacheService.getScriptCache().remove('catalogos_base');
}

function eliminarUsuarioPermanentemente(token, empleado) {
  var solicitante = requireSession_(token);
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_USUARIOS_);
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 10).getValues() : [];
  var indice = filas.findIndex(function(fila) { return String(fila[0]).trim() === String(empleado).trim(); });
  if (indice < 0) throw new Error('Usuario no encontrado.');
  var objetivo = filas[indice];
  if (esDEV_(solicitante) && String(objetivo[0]).trim() === String(solicitante.empleado).trim()) {
    throw new Error('El usuario DEV no puede eliminar su propia cuenta.');
  }
  if (solicitante.rol !== 'DEV' && !(['ADMAD', 'ADMWH', 'ADMQM'].includes(solicitante.rol) && objetivo[2] !== 'DEV' && objetivo[3] === departamentoDeRol_(solicitante.rol))) {
    throw new Error('Solo puedes borrar usuarios estándar de tu departamento.');
  }
  if (objetivo[2] === 'DEV' && solicitante.rol !== 'DEV') throw new Error('No puedes borrar al usuario DEV.');
  hoja.deleteRow(indice + 2);
  eliminarEmpleadoDeMatriz_(empleado);
  registrarAuditoria_(solicitante, 'ELIMINAR_USUARIO', empleado, 'Usuario eliminado permanentemente.');
  invalidarCacheUsuarios_();
  return true;
}

function agregarClasificaciones(token, clasificaciones) {
  return ejecutarConBloqueo_(function() {
    return agregarClasificacionesSinBloqueo_(token, clasificaciones);
  });
}

function agregarClasificacionesSinBloqueo_(token, clasificaciones) {
  var usuario = requireSession_(token);
  if (!puedeAdministrarUsuarios_(usuario)) throw new Error('Solo administradores y DEV pueden agregar números de parte.');
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de Finish God') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de Finish Goods');
  if (!hoja) throw new Error('No se encontró la matriz de partes.');
  var entradas = (clasificaciones || []).filter(function(item) { return item && String(item.parte).trim() && String(item.tipo).trim(); }).slice(0, 2);
  if (!entradas.length) throw new Error('Agrega al menos una clasificación.');
  var ultima = hoja.getLastRow();
  var existentes = ultima > 1 ? hoja.getRange(2, 1, ultima - 1, 2).getValues() : [];
  entradas.forEach(function(item) {
    var parte = limpiarNumeroParte_(item.parte);
    var tipo = limpiarTextoSeguro_(item.tipo, 40);
    if (!parte || !tipo) return;
    var indice = existentes.findIndex(function(fila) { return String(fila[0]).trim().toUpperCase() === parte; });
    if (indice >= 0) throw new Error('El número de parte ya existe.');
    hoja.appendRow([parte, textoLiteralHoja_(tipo)]);
    existentes.push([parte, tipo]);
    registrarAuditoria_(usuario, 'CREAR_PARTE', parte, 'Tipo: ' + tipo);
  });
  CacheService.getScriptCache().remove('catalogos_base');
  return obtenerCatalogosBase_().partes;
}

function listarPartesAdmin(token, filtro, pagina, porPagina, busqueda) {
  var usuario = requireSession_(token);
  if (!puedeConsultarMatrizPartes_(usuario)) throw new Error('No tienes permiso para consultar la matriz.');
  var partes = obtenerCatalogosBase_().partes;
  var texto = String(busqueda || '').trim().toUpperCase();
  var lista = Object.keys(partes).sort().filter(function(parte) { return (!filtro || filtro === 'Todos' || partes[parte] === filtro) && (!texto || parte.toUpperCase().indexOf(texto) >= 0); })
    .map(function(parte) { return { parte: parte, tipo: partes[parte] }; });
  if (String(pagina).toUpperCase() === 'TODAS') {
    return { items: lista, total: lista.length, pagina: 1, porPagina: lista.length, paginas: 0 };
  }
  pagina = Math.max(Number(pagina) || 1, 1); porPagina = Math.min(Math.max(Number(porPagina) || 30, 1), 100);
  return { items: lista.slice((pagina - 1) * porPagina, pagina * porPagina), total: lista.length, pagina: pagina, porPagina: porPagina, paginas: Math.max(1, Math.ceil(lista.length / porPagina)) };
}

function actualizarParte(token, parteAnterior, parte, tipo) {
  return ejecutarConBloqueo_(function() {
    return actualizarParteSinBloqueo_(token, parteAnterior, parte, tipo);
  });
}

function actualizarParteSinBloqueo_(token, parteAnterior, parte, tipo) {
  var usuario = requireSession_(token);
  if (!puedeAdministrarUsuarios_(usuario)) throw new Error('Solo administradores y DEV pueden modificar partes.');
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de Finish God') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de Finish Goods');
  if (!hoja) throw new Error('No se encontró la matriz de partes.');
  var anterior = limpiarNumeroParte_(parteAnterior);
  var nuevo = limpiarNumeroParte_(parte);
  var nuevoTipo = limpiarTextoSeguro_(tipo, 40);
  if (!anterior || !nuevo || !nuevoTipo) throw new Error('Completa el número de parte y el tipo.');
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues() : [];
  var indice = filas.findIndex(function(fila) { return String(fila[0]).trim().toUpperCase() === anterior; });
  if (indice < 0) throw new Error('Parte no encontrada.');
  var parteActual = String(filas[indice][0]).trim().toUpperCase();
  var tipoActual = String(filas[indice][1]).trim();
  if (parteActual === nuevo && tipoActual === nuevoTipo) return { mensaje: 'No se han generado cambios' };
  var duplicada = filas.some(function(fila, i) { return i !== indice && String(fila[0]).trim().toUpperCase() === nuevo; });
  if (duplicada) throw new Error('Ya existe ese número de parte.');
  hoja.getRange(indice + 2, 1, 1, 2).setValues([[nuevo, textoLiteralHoja_(nuevoTipo)]]);
  CacheService.getScriptCache().remove('catalogos_base');
  registrarAuditoria_(usuario, 'EDITAR_PARTE', nuevo, 'Anterior: ' + anterior + ', tipo: ' + nuevoTipo);
  return { mensaje: 'Cambios realizados y guardados correctamente.' };
}

function eliminarParte(token, parte) {
  return ejecutarConBloqueo_(function() {
    return eliminarParteSinBloqueo_(token, parte);
  });
}

function eliminarParteSinBloqueo_(token, parte) {
  var usuario = requireSession_(token);
  if (!puedeAdministrarUsuarios_(usuario)) throw new Error('Solo administradores y DEV pueden borrar partes.');
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de Finish God') || SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matriz de Finish Goods');
  if (!hoja) throw new Error('No se encontró la matriz de partes.');
  var filas = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues() : [];
  var indice = filas.findIndex(function(fila) { return String(fila[0]).trim().toUpperCase() === String(parte).trim().toUpperCase(); });
  if (indice < 0) throw new Error('Parte no encontrada.');
  hoja.deleteRow(indice + 2);
  CacheService.getScriptCache().remove('catalogos_base');
  registrarAuditoria_(usuario, 'ELIMINAR_PARTE', parte, 'Parte eliminada.');
  return true;
}

function registrarMovimientoUnificado(datos, token) {
  var usuario = requireSession_(token);
  if (!puedeUsarOperacion_(usuario, datos.tipoOperacion)) throw new Error('Tu usuario no tiene permiso para esta operación.');
  datos.escaneo = limpiarCodigoEscaneo_(datos.escaneo, true);
  datos.escaneo2 = limpiarCodigoEscaneo_(datos.escaneo2, datos.esLateral === 'Si');
  var escaneoPrincipal = analizarEscaneoRegistro_(datos.escaneo);
  datos.linea = escaneoPrincipal.linea;
  datos.turno = String(datos.turno || '').trim().toUpperCase();
  if (!/^[A-D]$/.test(datos.turno)) throw new Error('Selecciona un turno válido (A, B, C o D).');
  if (datos.esLateral === 'Si') {
    var escaneoSecundario = analizarEscaneoRegistro_(datos.escaneo2);
    if (escaneoPrincipal.cantidad !== escaneoSecundario.cantidad) {
      throw new Error('Los dos escaneos laterales deben tener la misma cantidad.');
    }
  }
  validarCantidadPositiva_(datos.cantidad);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Hay otro registro en proceso. Espera un momento e inténtalo de nuevo.');
  try {
    return registrarMovimientoConLock_(datos, usuario, token);
  } finally {
    lock.releaseLock();
  }
}

function validarCantidadPositiva_(cantidad) {
  var valor = Number(cantidad);
  if (!isFinite(valor) || valor <= 0 || Math.floor(valor) !== valor) throw new Error('La cantidad debe ser un número entero positivo.');
  return valor;
}

function analizarEscaneoRegistro_(escaneo) {
  var segmentos = String(escaneo || '').trim().split('-').map(function(segmento) { return segmento.trim(); });
  if (segmentos.length < 7 || segmentos.slice(0, 7).some(function(segmento) { return !segmento; })) {
    throw new Error('El escaneo debe contener al menos siete valores separados por guiones.');
  }
  var cantidad = validarCantidadPositiva_(segmentos[4]);
  return { rolling: segmentos[1], parte: segmentos[2], cantidad: cantidad, linea: segmentos[5] };
}

function validarNumeroCarroCaja_(valor) {
  var texto = String(valor || '').trim();
  if (/^Caja$/i.test(texto)) return 'Caja';
  if (/^(?:[1-9]\d{0,3}|10000)$/.test(texto)) return texto;
  throw new Error('El No. carro/caja debe ser Caja o un número del 1 al 10000.');
}

function registrarMovimientoConLock_(datos, usuario, token) {
  datos.escaneo = limpiarCodigoEscaneo_(datos.escaneo, true);
  datos.escaneo2 = limpiarCodigoEscaneo_(datos.escaneo2, datos.esLateral === 'Si');
  datos.numEmpleado = validarEmpleadoRegistro_(limpiarIdSeguro_(datos.numEmpleado, 20), usuario);
  datos.numCarroCaja = validarNumeroCarroCaja_(limpiarIdSeguro_(datos.numCarroCaja, 20));
  datos.linea = limpiarTextoSeguro_(datos.linea, 20);
  datos.turno = limpiarTextoSeguro_(datos.turno, 10).toUpperCase();
  datos.corte = limpiarTextoSeguro_(datos.corte, 20);
  datos.fecha = limpiarTextoSeguro_(datos.fecha, 20);
  datos.comentarios = limpiarTextoSeguro_(datos.comentarios, 200);
  if (String(datos.tipoOperacion || '').indexOf('calidad_') === 0) datos.corte = '';
  datos.comentarios = validarComentarios_(usuario, datos.comentarios, datos.tipoOperacion);
  datos.usuarioSesion = usuario.empleado;
  datos.token = token || datos.token || '';
  var mapaHojas = {
    aduanas_entrada: 'Entradas de aduanas',
    aduanas_salida: 'Salidas de aduanas',
    almacen_entrada: 'Entradas de almacén',
    almacen_salida: 'Salidas de almacén',
    calidad_entrada: 'Entradas de calidad',
    calidad_salida: 'Salidas de calidad'
  };
  var nombreHoja = mapaHojas[datos.tipoOperacion] || datos.nombreHoja;
  if (!nombreHoja) throw new Error('Tipo de operación no reconocido.');
  var idSolicitud = String(datos.idSolicitud || '').trim();
  if (idSolicitud && !/^[A-Za-z0-9_-]{1,60}$/.test(idSolicitud)) throw new Error('El identificador del movimiento no es válido.');

  if (datos.esLateral === 'Si') {
    validarLateralesServidor_(datos.escaneo, datos.escaneo2);
    var datos1 = Object.assign({}, datos, { escaneo: datos.escaneo, idSolicitud: idSolicitud ? idSolicitud + '-1' : '' });
    var datos2 = Object.assign({}, datos, { escaneo: datos.escaneo2.trim(), idSolicitud: idSolicitud ? idSolicitud + '-2' : '' });
    var detallesLaterales = [guardarEnHojaOptimizada(datos1, nombreHoja, token), guardarEnHojaOptimizada(datos2, nombreHoja, token)];
    SpreadsheetApp.flush();
    var avisoCalidadLateral = null;
    if (datos.tipoOperacion === 'aduanas_entrada' || datos.tipoOperacion === 'aduanas_salida') {
      var resultadosCalidadLaterales = detallesLaterales.map(function(detalle) {
        var validado = sincronizarLiberacionQMConAduanas_(datos.tipoOperacion, detalle.fila);
        return !validado && detalle.yaExistia ? validarLiberacionQMParaRegistro_(datos.tipoOperacion, detalle.fila) : validado;
      });
      var validadoLateral = resultadosCalidadLaterales.every(function(validado) { return validado; });
      avisoCalidadLateral = {
        mensaje: validadoLateral ? 'Material validado por Calidad' : 'Material no registrado por el departamento de Calidad',
        validado: validadoLateral
      };
    } else if (datos.tipoOperacion === 'almacen_entrada' || datos.tipoOperacion === 'almacen_salida') {
      var resultadosCalidadWHLaterales = detallesLaterales.map(function(detalle) {
        return validarLiberacionQMParaRegistro_(datos.tipoOperacion, detalle.fila);
      });
      var validadoWHLateral = resultadosCalidadWHLaterales.every(function(validado) { return validado; });
      avisoCalidadLateral = {
        mensaje: validadoWHLateral ? 'Material validado por Calidad' : 'Material no registrado por el departamento de Calidad',
        validado: validadoWHLateral
      };
    }
    return {
      exito: true,
      mensaje: '¡Ambas piezas laterales registradas con éxito!',
      avisoCalidad: avisoCalidadLateral,
      detalles: detallesLaterales
    };
  }

  var detalle = guardarEnHojaOptimizada(datos, nombreHoja, token);
  SpreadsheetApp.flush();
  var avisoCalidad = null;
  if (datos.tipoOperacion === 'aduanas_entrada' || datos.tipoOperacion === 'aduanas_salida') {
    var validado = sincronizarLiberacionQMConAduanas_(datos.tipoOperacion, detalle.fila);
    if (!validado && detalle.yaExistia) validado = validarLiberacionQMParaRegistro_(datos.tipoOperacion, detalle.fila);
    avisoCalidad = {
      mensaje: validado ? 'Material validado por Calidad' : 'Material no registrado por el departamento de Calidad',
      validado: validado
    };
  } else if (datos.tipoOperacion === 'almacen_entrada' || datos.tipoOperacion === 'almacen_salida') {
    var validadoWH = validarLiberacionQMParaRegistro_(datos.tipoOperacion, detalle.fila);
    avisoCalidad = {
      mensaje: validadoWH ? 'Material validado por Calidad' : 'Material no registrado por el departamento de Calidad',
      validado: validadoWH
    };
  }
  return { exito: true, mensaje: "¡Registro guardado con éxito en '" + nombreHoja + "'!", avisoCalidad: avisoCalidad, detalles: [detalle] };
}

function validarLateralesServidor_(escaneo, escaneo2) {
  var partes = obtenerCatalogosBase_().partes;
  var extraerTipo = function(valor) {
    var segmentos = String(valor || '').trim().split('-');
    return segmentos.length >= 3 ? partes[normalizarParteServidor_(segmentos[2])] : '';
  };
  var tipo1 = extraerTipo(escaneo);
  var tipo2 = extraerTipo(escaneo2);
  if (tipo1 === tipo2 && (tipo1 === 'B/C L' || tipo1 === 'B/C R')) throw new Error('Se registraron dos veces ' + tipo1 + '. Escanea el lado faltante.');
  if (!['B/C L', 'B/C R'].includes(tipo1) || !['B/C L', 'B/C R'].includes(tipo2) || tipo1 === tipo2) throw new Error('Para un lateral debes escanear una pieza B/C L y una pieza B/C R.');
}

function validarEmpleadoRegistro_(empleado, usuario) {
  var id = String(empleado || '').trim();
  if (!id) throw new Error('Selecciona el empleado que realizará el registro.');
  var catalogos = obtenerCatalogosBase_();
  if (!catalogos.empleados[id] && id !== usuario.empleado) throw new Error('El empleado seleccionado no existe en la matriz de empleados.');
  return id;
}

function guardarEnHojaOptimizada(datos, nombreHoja, token) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) throw new Error("No se encontró la pestaña '" + nombreHoja + "'.");

  var id = String(datos.idSolicitud || '').trim() || nuevoIdRegistro_();
  if (datos.idSolicitud && hoja.getLastRow() > 1) {
    var coincidenciaId = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).matchCase(true).findNext();
    if (coincidenciaId) {
      var filaExistente = coincidenciaId.getRow();
      var existente = hoja.getRange(filaExistente, 2, 1, 14).getValues()[0];
      var coincideSolicitud = (!datos.fecha || formatearFecha(existente[0]) === String(datos.fecha).trim()) &&
        String(existente[1] || '').trim().toUpperCase() === String(datos.escaneo || '').trim().toUpperCase() &&
        String(existente[4] || '').trim() === String(datos.numCarroCaja || '').trim() &&
        Number(existente[5]) === Number(datos.cantidad) &&
        String(existente[6] || '').trim() === String(datos.linea || '').trim() &&
        String(existente[7] || '').trim() === String(datos.numEmpleado || '').trim() &&
        String(existente[8] || '').trim() === String(datos.turno || '').trim() &&
        String(existente[11] || '').trim() === String(datos.corte || '').trim() &&
        String(existente[12] || '').trim() === String(datos.comentarios || '').trim();
      if (!coincideSolicitud) throw new Error('Este movimiento ya se registró con datos distintos. Consulta el historial antes de volver a intentarlo.');
      return { id: id, fila: filaExistente, escaneo: String(existente[1] || ''), rolling: String(existente[2] || ''), noParte: String(existente[3] || ''), cantidad: existente[5], linea: existente[6], carro: String(existente[10] || ''), carroEtiqueta: String(existente[4] || ''), yaExistia: true };
    }
  }

  var nuevaFila = hoja.getLastRow() + 1;
  var esCalidad = nombreHoja.indexOf('calidad') >= 0;
  if (esCalidad) asegurarEstatusCalidad_(hoja);
  var esAlmacen = nombreHoja.toLowerCase().indexOf('almacén') !== -1;
  var fechaSeleccionada = String(datos.fecha || '').trim();
  var fechaRegistro;
  if (fechaSeleccionada) {
    fechaRegistro = new Date(fechaSeleccionada + 'T00:00:00');
  } else {
    var celdaFecha = hoja.getRange(nuevaFila, 2);
    celdaFecha.setFormula(esAlmacen ? "='Resumen WH'!C1" : "='Resumen aduanas'!$C$1");
    SpreadsheetApp.flush();
    fechaRegistro = celdaFecha.getValue();
  }

  var escaneo = String(datos.escaneo || '').trim().toUpperCase();
  var datosEscaneo = analizarEscaneoRegistro_(escaneo);
  var rolling = datosEscaneo.rolling;
  var parte = datosEscaneo.parte;
  var rollingFormula = rolling.replace(/"/g, '""');
  var parteFormula = parte.replace(/"/g, '""');
  var filaValores = [
    escaneo,
    '=IFERROR(REGEXEXTRACT(C' + nuevaFila + ', "^[^-]+-([^-]+)"), "' + rollingFormula + '")',
    '=IFERROR(REGEXEXTRACT(C' + nuevaFila + ', "^[^-]+-[^-]+-([^-]+)"), "' + parteFormula + '")',
    textoLiteralHoja_(String(datos.numCarroCaja || '')),
    Number(datos.cantidad), textoLiteralHoja_(String(datos.linea || '')), textoLiteralHoja_(String(datos.numEmpleado || '')), textoLiteralHoja_(String(datos.turno || '')),
    '=VLOOKUP(I' + nuevaFila + ', \'Matriz de empleados\'!A:B, 2, FALSE)',
    '',
    textoLiteralHoja_(String(datos.corte || '')),
    textoLiteralHoja_(String(datos.comentarios || '')),
    textoLiteralHoja_(String(datos.horaRegistro || ''))
  ];
  hoja.getRange(nuevaFila, 1, 1, filaValores.length + 2).setValues([[id, fechaRegistro].concat(filaValores)]);
  var carroAsignado = esCalidad ? calcularCarroNuevo_(hoja, nuevaFila) : obtenerCarroAsignadoDesdeCacheOHoja_(hoja, datos.tipoOperacion, token, nuevaFila, datos);
  hoja.getRange(nuevaFila, 12).setValue(textoLiteralHoja_(String(carroAsignado || '')));
  if (esCalidad) hoja.getRange(nuevaFila, 16).setValue('Pendiente');
  return { id: id, fila: nuevaFila, escaneo: escaneo, rolling: rolling, noParte: parte, cantidad: datos.cantidad, linea: datos.linea, carro: carroAsignado, carroEtiqueta: datos.numCarroCaja };
}

function sincronizarLiberacionQMConAduanas_(operacionAduanas, filaAduanas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaAduanas = ss.getSheetByName(nombreHojaOperacion_(operacionAduanas));
  var operacionQM = operacionAduanas === 'aduanas_entrada' ? 'calidad_entrada' : 'calidad_salida';
  var hojaQM = ss.getSheetByName(nombreHojaOperacion_(operacionQM));
  if (!hojaAduanas || !hojaQM || filaAduanas < 2 || hojaQM.getLastRow() < 2) return false;

  asegurarEstatusCalidad_(hojaQM);
  var registroAduanas = hojaAduanas.getRange(filaAduanas, 2, 1, 14).getValues()[0];
  var fecha = formatearFecha(registroAduanas[0]);
  var turno = String(registroAduanas[8] || '').trim();
  var parte = normalizarParteServidor_(registroAduanas[3]);
  var cantidad = Number(registroAduanas[5]) || 0;
  var linea = String(registroAduanas[6] || '').trim();
  if (!fecha || !turno || !parte || !linea || !cantidad) return false;

  var cantidadQM = hojaQM.getLastRow() - 1;
  var filasQM = hojaQM.getRange(2, 2, cantidadQM, 15).getValues();
  var candidatos = [];
  filasQM.forEach(function(fila, indice) {
    if (String(fila[14] || '').trim().toUpperCase() !== 'PENDIENTE') return;
    if (formatearFecha(fila[0]) !== fecha ||
        String(fila[8] || '').trim() !== turno ||
        normalizarParteServidor_(fila[3]) !== parte ||
        String(fila[6] || '').trim() !== linea ||
        (Number(fila[5]) || 0) !== cantidad) return;
    var minutos = minutosHoraHistorial_(fila[13]);
    candidatos.push({ fila: indice + 2, hora: minutos === null ? Number.MAX_SAFE_INTEGER : minutos });
  });
  if (!candidatos.length) return false;

  candidatos.sort(function(a, b) { return a.hora - b.hora || a.fila - b.fila; });
  var filaQM = candidatos[0].fila;
  hojaQM.getRange(filaQM, 6).setValue(textoLiteralHoja_(String(registroAduanas[4] || '')));
  hojaQM.getRange(filaQM, 12).setValue(textoLiteralHoja_(String(registroAduanas[10] || '')));
  hojaQM.getRange(filaQM, 13).setValue(textoLiteralHoja_(String(registroAduanas[11] || '')));
  hojaQM.getRange(filaQM, 16).setValue('Corregido');
  return true;
}

function validarLiberacionQMParaRegistro_(operacionAlmacen, filaAlmacen) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaAlmacen = ss.getSheetByName(nombreHojaOperacion_(operacionAlmacen));
  var operacionQM = String(operacionAlmacen || '').slice(-8) === '_entrada' ? 'calidad_entrada' : 'calidad_salida';
  var hojaQM = ss.getSheetByName(nombreHojaOperacion_(operacionQM));
  if (!hojaAlmacen || !hojaQM || filaAlmacen < 2 || hojaQM.getLastRow() < 2) return false;

  asegurarEstatusCalidad_(hojaQM);
  var registroAlmacen = hojaAlmacen.getRange(filaAlmacen, 2, 1, 14).getValues()[0];
  var fecha = formatearFecha(registroAlmacen[0]);
  var turno = String(registroAlmacen[8] || '').trim();
  var parte = normalizarParteServidor_(registroAlmacen[3]);
  var cantidad = Number(registroAlmacen[5]) || 0;
  var linea = String(registroAlmacen[6] || '').trim();
  if (!fecha || !turno || !parte || !linea || !cantidad) return false;

  var filasQM = hojaQM.getRange(2, 2, hojaQM.getLastRow() - 1, 15).getValues();
  return filasQM.some(function(fila) {
    return String(fila[14] || '').trim().toUpperCase() === 'CORREGIDO' &&
      formatearFecha(fila[0]) === fecha &&
      String(fila[8] || '').trim() === turno &&
      normalizarParteServidor_(fila[3]) === parte &&
      String(fila[6] || '').trim() === linea &&
      (Number(fila[5]) || 0) === cantidad;
  });
}

function registrarEntradaAduanas(datos, token) { datos.tipoOperacion = 'aduanas_entrada'; return registrarMovimientoUnificado(datos, token).mensaje; }
function registrarSalidaAduanas(datos, token) { datos.tipoOperacion = 'aduanas_salida'; return registrarMovimientoUnificado(datos, token).mensaje; }
function registrarEntradaAlmacen(datos, token) { datos.tipoOperacion = 'almacen_entrada'; return registrarMovimientoUnificado(datos, token).mensaje; }
function registrarSalidaAlmacen(datos, token) { datos.tipoOperacion = 'almacen_salida'; return registrarMovimientoUnificado(datos, token).mensaje; }

function nombreHojaOperacion_(operacion) {
  return { aduanas_entrada: 'Entradas de aduanas', aduanas_salida: 'Salidas de aduanas', almacen_entrada: 'Entradas de almacén', almacen_salida: 'Salidas de almacén', calidad_entrada: 'Entradas de calidad', calidad_salida: 'Salidas de calidad' }[operacion];
}

function listarRegistros(token, operacion, limite) {
  var usuario = requireSession_(token);
  if (!puedeUsarOperacion_(usuario, operacion)) throw new Error('No tienes permiso para consultar esta operación.');
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
  var inicio = Math.max(2, hoja.getLastRow() - Math.min(Number(limite) || 50, 100) + 1);
  var cantidad = hoja.getLastRow() >= inicio ? hoja.getLastRow() - inicio + 1 : 0;
  if (!cantidad) return [];
  var ids = hoja.getRange(inicio, 1, cantidad, 1).getValues();
  return hoja.getRange(inicio, 2, cantidad, 13).getValues().map(function(fila, indice) {
    return { id: ids[indice][0], fila: inicio + indice, fecha: fila[0], escaneo: fila[1], rolling: fila[2], parte: fila[3], carroCaja: fila[4], cantidad: fila[5], linea: fila[6], empleado: fila[7], turno: fila[8], registra: fila[9], carro: fila[10], corte: fila[11], comentarios: esUsuarioQM_(usuario) ? fila[12] : '' };
  }).reverse();
}

function operacionesHistorial_(usuario) {
  return operacionesDeUsuario_(usuario);
}

function claveHistorial_(fila) {
  return [
    formatearFecha(fila[0]),
    String(fila[8] || '').trim(),
    String(fila[6] || '').trim(),
    normalizarParteServidor_(fila[3]),
    tipoParteHistorial_(fila[3]),
    String(fila[5] || '').trim(),
    String(fila[4] || '').trim(),
    String(fila[10] || '').trim(),
    String(fila[11] || '').trim()
  ].join('|');
}

function obtenerFilasHistorialEnBloquesDesdeAbajo_(hoja, fecha, turno, corte, incluirEstadoQM) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  var inicioRango = 2;
  var finRango = ultimaFila;
  if (fecha) {
    var bloqueFecha = leerBloqueFechaDesdeAbajo_(hoja, fecha, ultimaFila);
    if (bloqueFecha.fin < bloqueFecha.inicio) return [];
    inicioRango = bloqueFecha.inicio;
    finRango = bloqueFecha.fin;
  }
  var columnas = incluirEstadoQM ? 16 : 15;
  var registros = [];
  for (var fin = finRango; fin >= inicioRango; fin -= RESUMEN_CACHE_BLOQUE_FILAS_) {
    var inicio = Math.max(inicioRango, fin - RESUMEN_CACHE_BLOQUE_FILAS_ + 1);
    var filas = hoja.getRange(inicio, 1, fin - inicio + 1, columnas).getValues();
    for (var indice = filas.length - 1; indice >= 0; indice--) {
      var fila = filas[indice];
      var valores = fila.slice(1, 15);
      if (fecha && formatearFecha(valores[0]) !== String(fecha).trim()) continue;
      if (turno && String(valores[8] || '').trim() !== turno) continue;
      if (corte && String(valores[11] || '').trim() !== corte) continue;
      registros.push({
        id: fila[0],
        fila: inicio + indice,
        valores: valores,
        estadoQM: incluirEstadoQM ? fila[15] : ''
      });
    }
  }
  return registros;
}

function listarHistorial(token, filtros) {
  var usuario = requireSession_(token);
  filtros = filtros || {};
  var partesCatalogo = obtenerCatalogosBase_().partes;
  var fecha = String(filtros.fecha || '').trim();
  var turno = String(filtros.turno || '').trim();
  var corte = String(filtros.corte || '').trim();
  var tipoMovimiento = String(filtros.tipoMovimiento || 'entrada').trim().toLowerCase();
  if (tipoMovimiento !== 'entrada' && tipoMovimiento !== 'salida') throw new Error('El tipo de movimiento no es válido.');
  var parteFiltro = String(filtros.parte || '').trim().toUpperCase();
  var departamentoFiltro = String(filtros.departamento || '').trim().toUpperCase();
  var operaciones = operacionesHistorial_(usuario);
  operaciones = operaciones.filter(function(operacion) {
    return operacion.endsWith('_' + tipoMovimiento);
  });
  if (departamentoFiltro && usuario.rol === 'DEV') {
    operaciones = operaciones.filter(function(operacion) {
      return departamentoDeOperacion_(operacion) === departamentoFiltro;
    });
  }
  var validacionesQM = {};
  ['calidad_' + tipoMovimiento].forEach(function(operacion) {
    var hojaQM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
    if (!hojaQM || hojaQM.getLastRow() < 2) return;
    var filasQM = obtenerFilasHistorialEnBloquesDesdeAbajo_(hojaQM, fecha, turno, corte, true);
    filasQM.forEach(function(registro) {
      if (String(registro.estadoQM || '').trim().toUpperCase() !== 'CORREGIDO') return;
      validacionesQM[claveHistorial_(registro.valores)] = { comentarios: registro.valores[12] || '', registro: registro.valores[9] || '' };
    });
  });
  var registros = [];
  operaciones.forEach(function(operacion) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
    if (!hoja || hoja.getLastRow() < 2) return;
    var filasHistorial = obtenerFilasHistorialEnBloquesDesdeAbajo_(hoja, fecha, turno, corte, false);
    filasHistorial.forEach(function(registro) {
      var fila = registro.valores;
      if (parteFiltro && normalizarParteServidor_(fila[3]).toUpperCase().indexOf(parteFiltro) < 0) return;
      var esQM = operacion.indexOf('calidad_') === 0;
      var validacionQM = validacionesQM[claveHistorial_(fila)];
      if (esQM && usuario.rol === 'DEV' && validacionQM) return;
      registros.push({
        id: registro.id, fila: registro.fila,
        operacion: operacion,
        operacionNombre: operacion === 'aduanas_entrada' ? 'Entrada Aduanas' : operacion === 'aduanas_salida' ? 'Salida Aduanas' : operacion === 'almacen_entrada' ? 'Entrada Almacén' : operacion === 'almacen_salida' ? 'Salida Almacén' : operacion === 'calidad_entrada' ? 'Entrada Calidad' : 'Salida Calidad',
        fecha: formatearFecha(fila[0]), linea: fila[6], parte: normalizarParteServidor_(fila[3]), tipo: partesCatalogo[normalizarParteServidor_(fila[3])] || 'No registrado',
        rolling: fila[2], hora: formatearHoraHistorial_(fila[13]), departamento: departamentoDeOperacion_(operacion), carroCaja: fila[4], carro: fila[10], corte: fila[11], turno: fila[8], comentarios: esUsuarioQM_(usuario) ? (fila[12] || (validacionQM && validacionQM.comentarios) || '') : ((validacionQM && validacionQM.comentarios) || ''), registro: fila[9] || (validacionQM && validacionQM.registro) || '', cantidad: fila[5],
        validadoQM: esQM ? null : !!validacionQM
      });
    });
  });
  registros.sort(function(a, b) { return String(b.fecha).localeCompare(String(a.fecha)) || Number(b.fila) - Number(a.fila); });
  return { items: registros, total: registros.length, pagina: 1, porPagina: registros.length, paginas: 1 };
}

function departamentoDeOperacion_(operacion) {
  var valor = String(operacion || '').toLowerCase();
  if (valor.indexOf('aduanas_') === 0) return 'UAD';
  if (valor.indexOf('almacen_') === 0) return 'UWH';
  if (valor.indexOf('calidad_') === 0) return 'UQM';
  return '';
}

function formatearHoraHistorial_(valor) {
  if (!valor) return '';
  if (valor instanceof Date) return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'HH:mm:ss');
  return String(valor).trim();
}

function minutosHoraHistorial_(valor) {
  var texto = formatearHoraHistorial_(valor);
  var partes = texto.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!partes) return null;
  var horas = Number(partes[1]);
  var minutos = Number(partes[2]);
  if (horas === 24 && minutos === 0) return 1440;
  if (horas > 23 || minutos > 59) return null;
  return horas * 60 + minutos;
}

function obtenerResumenHXH(token, filtros) {
  var usuario = requireSession_(token);
  filtros = filtros || {};
  var fecha = String(filtros.fecha || '').trim();
  var turno = String(filtros.turno || '').trim();
  var departamento = String(filtros.departamento || '').trim().toUpperCase();
  var inicio = minutosHoraHistorial_(filtros.horaInicio);
  var final = minutosHoraHistorial_(filtros.horaFinal);
  if (!fecha || !turno || inicio === null || final === null) throw new Error('Completa fecha, turno, hora de inicio y hora final.');
  if (inicio > final) throw new Error('La hora de inicio no puede ser mayor que la hora final.');
  if (!['UAD', 'UWH', 'UQM'].includes(departamento)) throw new Error('Selecciona un departamento válido.');
  var registros = [];
  var partesCatalogo = obtenerCatalogosBase_().partes;
  operacionesHistorial_(usuario).forEach(function(operacion) {
    if (departamentoDeOperacion_(operacion) !== departamento) return;
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
    if (!hoja || hoja.getLastRow() < 2) return;
    obtenerFilasOperacionFiltradas_(hoja, fecha, turno, '').filas.forEach(function(fila) {
      if (formatearFecha(fila[0]) !== fecha || String(fila[8] || '').trim() !== turno) return;
      var hora = minutosHoraHistorial_(fila[13]);
      if (hora === null || hora < inicio || hora > final) return;
      var linea = String(fila[6] || '').trim();
      var parte = normalizarParteServidor_(String(fila[3] || ''));
      if (!linea || !parte) return;
      registros.push({
        operacion: operacion,
        linea: linea,
        parte: parte,
        tipo: partesCatalogo[parte] || 'No registrado',
        cantidad: Number(fila[5]) || 0,
        carroCaja: String(fila[4] || '').trim(),
        hora: formatearHoraHistorial_(fila[13])
      });
    });
  });
  var grupos = {};
  registros.forEach(function(registro) {
    var esLateral = registro.tipo === 'B/C L' || registro.tipo === 'B/C R';
    var clave;
    if (esLateral) {
      clave = registro.linea;
      if (!grupos[clave]) grupos[clave] = { linea: registro.linea, parteL: '', parteR: '', cantidad: 0, laterales: true };
      if (registro.tipo === 'B/C L') grupos[clave].parteL = registro.parte;
      if (registro.tipo === 'B/C R') grupos[clave].parteR = registro.parte;
      grupos[clave].cantidad += registro.cantidad;
    } else {
      clave = [registro.linea, registro.parte].join('|');
      if (!grupos[clave]) grupos[clave] = { linea: registro.linea, parte: registro.parte, tipo: registro.tipo, cantidad: 0, laterales: false };
      grupos[clave].cantidad += registro.cantidad;
    }
  });
  var items = Object.keys(grupos).map(function(clave) {
    var grupo = grupos[clave];
    if (grupo.laterales) {
      grupo.parte = grupo.parteL && grupo.parteR ? grupo.parteL + ' / ' + grupo.parteR : (grupo.parteL || grupo.parteR);
      grupo.tipo = 'B/C L-R';
      grupo.cantidad = grupo.cantidad / 2;
    }
    return { linea: grupo.linea, parte: grupo.parte, tipo: grupo.tipo, cantidad: grupo.cantidad };
  }).sort(function(a, b) {
    var numeroA = Number(a.linea);
    var numeroB = Number(b.linea);
    if (!isNaN(numeroA) && !isNaN(numeroB) && numeroA !== numeroB) return numeroA - numeroB;
    return String(a.parte).localeCompare(String(b.parte));
  });
  return { fecha: fecha, turno: turno, departamento: departamento, horaInicio: filtros.horaInicio, horaFinal: filtros.horaFinal, items: items, total: items.reduce(function(total, item) { return total + item.cantidad; }, 0) };
}

function eliminarRegistro(token, operacion, fila) {
  return ejecutarConBloqueo_(function() {
    return eliminarRegistroSinBloqueo_(token, operacion, fila);
  });
}

function eliminarRegistroSinBloqueo_(token, operacion, fila) {
  var usuario = requireSession_(token);
  if (!puedeGestionarOperacion_(usuario, operacion)) throw new Error('No tienes permiso para borrar este registro.');
  var registro = leerRegistroParaEliminar_(operacion, fila);
  registro.hoja.deleteRow(registro.fila);
  actualizarCacheResumenTrasCambio_(registro.hoja, registro.fecha);
  registrarAuditoria_(usuario, 'ELIMINAR_REGISTRO', registro.id, 'Operación: ' + operacion + ', fila histórica: ' + registro.fila);
  return true;
}

function actualizarRegistro(token, operacion, fila, cambios) {
  return ejecutarConBloqueo_(function() {
    return actualizarRegistroSinBloqueo_(token, operacion, fila, cambios);
  });
}

function actualizarRegistroSinBloqueo_(token, operacion, fila, cambios) {
  var usuario = requireSession_(token);
  var esAdmin = esAdministrador_(usuario);
  if (!esAdmin && !puedeEditarRegistro_(usuario, operacion, cambios)) {
    throw new Error('Solo puedes modificar cantidad y número de carro/caja.');
  }
  if (!esAdmin && !puedeUsarOperacion_(usuario, operacion)) {
    throw new Error('No puedes modificar registros de este departamento.');
  }
  if (esAdmin && !puedeGestionarOperacion_(usuario, operacion)) {
    throw new Error('No puedes modificar registros de este departamento.');
  }
  var referencia = resolverFilaOperacion_(operacion, fila);
  var hoja = referencia.hoja;
  fila = referencia.fila;
  var fechaCacheResumen = formatearFecha(hoja.getRange(Number(fila), 2).getValue());
  if (!esDEV_(usuario)) {
    var fechaRegistro = fechaCacheResumen;
    if (fechaRegistro !== formatearFecha(new Date())) throw new Error('Solo puedes editar registros de la fecha actual.');
  }
  if (Object.prototype.hasOwnProperty.call(cambios, 'cantidad')) validarCantidadPositiva_(cambios.cantidad);
  if (Object.prototype.hasOwnProperty.call(cambios, 'carroCaja')) cambios.carroCaja = validarNumeroCarroCaja_(cambios.carroCaja);
  if (Object.prototype.hasOwnProperty.call(cambios, 'linea')) {
    var linea = String(cambios.linea || '').trim();
    if (!/^(?:[1-9]|1[0-4])$/.test(linea) && linea !== '8003 RWK') throw new Error('Selecciona una línea válida.');
    cambios.linea = linea;
  }
  if (Object.prototype.hasOwnProperty.call(cambios, 'corte') && !['1', '2', '3', 'Acumulado'].includes(String(cambios.corte))) throw new Error('El corte no es válido.');
  if (Object.prototype.hasOwnProperty.call(cambios, 'escaneo')) cambios.escaneo = limpiarCodigoEscaneo_(cambios.escaneo, true);
  if (Object.prototype.hasOwnProperty.call(cambios, 'comentarios')) cambios.comentarios = validarComentarios_(usuario, cambios.comentarios, operacion);
  var valores = hoja.getRange(Number(fila), 3, 1, 11).getValues()[0];
  ['escaneo', 'rolling', 'parte', 'carroCaja', 'cantidad', 'linea', 'empleado', 'turno', 'registra', 'carro', 'corte'].forEach(function(clave, i) {
    if (Object.prototype.hasOwnProperty.call(cambios, clave)) {
      var valor = cambios[clave];
      if (['escaneo', 'carroCaja', 'linea', 'empleado', 'turno', 'registra', 'corte'].indexOf(clave) >= 0) {
        valor = textoLiteralHoja_(String(valor == null ? '' : valor));
      }
      valores[i] = valor;
    }
  });
  hoja.getRange(Number(fila), 3, 1, 11).setValues([valores]);
  if (Object.prototype.hasOwnProperty.call(cambios, 'comentarios')) hoja.getRange(Number(fila), 14).setValue(textoLiteralHoja_(String(cambios.comentarios == null ? '' : cambios.comentarios)));
  actualizarCacheResumenTrasCambio_(hoja, fechaCacheResumen);
  registrarAuditoria_(usuario, 'EDITAR_REGISTRO', referencia.id, 'Operación: ' + operacion + ', campos: ' + Object.keys(cambios).join(', '));
  return true;
}

function obtenerResumenConciliacion(token) {
  requireSession_(token);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resultado = { fecha: '', turno: '', corte: '', totalAduanas: 0, totalAlmacen: 0, discrepancias: 0, detalles: [] };
  try {
    var aduanas = ss.getSheetByName('Resumen aduanas');
    var almacen = ss.getSheetByName('Resumen WH');
    if (!aduanas) return resultado;
    var metadatosAduanas = aduanas.getRange('C1:L1').getValues()[0];
    resultado.fecha = formatearFecha(metadatosAduanas[0]);
    resultado.turno = String(metadatosAduanas[4] || '');
    resultado.corte = String(metadatosAduanas[9] || '');
    var filas = Math.min(aduanas.getLastRow(), 200);
    if (filas < 3) return resultado;
    var datosAduanas = aduanas.getRange(3, 1, filas - 2, 19).getValues();
    var datosAlmacen = almacen ? almacen.getRange(3, 1, filas - 2, 15).getValues() : [];
    datosAduanas.forEach(function(fila, i) {
      var totalAduana = Number(fila[17]) || 0;
      var totalWH = datosAlmacen[i] ? Number(datosAlmacen[i][13]) || 0 : 0;
      if (fila[2] || totalAduana || totalWH) {
        var error = totalAduana !== totalWH || String(fila[18] || 'OK').trim() === 'ERROR';
        resultado.totalAduanas += totalAduana;
        resultado.totalAlmacen += totalWH;
        if (error) resultado.discrepancias++;
        resultado.detalles.push({ linea: fila[0], parte: fila[2], aduanas: totalAduana, almacen: totalWH, diferencia: totalAduana - totalWH, estado: error ? 'ERROR' : 'OK' });
      }
    });
  } catch (err) { Logger.log('Error en conciliación: ' + err.message); }
  return resultado;
}

function obtenerResumenVisual(token, filtros) {
  var usuario = requireSession_(token);
  filtros = filtros || {};
  var fecha = String(filtros.fecha || '').trim();
  var turno = String(filtros.turno || '').trim();
  var corte = String(filtros.corte || 'Acumulado').trim();
  var tipoMovimiento = String(filtros.tipoMovimiento || 'entrada').trim().toLowerCase();
  if (tipoMovimiento !== 'entrada' && tipoMovimiento !== 'salida') throw new Error('El tipo de movimiento no es válido.');
  var operacionesPermitidas = operacionesDeUsuario_(usuario);
  var operacionSeleccionada = String(filtros.operacion || '').trim();
  var grupoSeleccionado = operacionSeleccionada && operacionesPermitidas.indexOf(operacionSeleccionada) >= 0
    ? operacionSeleccionada.split('_')[0]
    : '';
  var operaciones = grupoSeleccionado
    ? operacionesPermitidas.filter(function(operacion) { return operacion.indexOf(grupoSeleccionado + '_') === 0; })
    : operacionesPermitidas;
  var departamentoFiltro = String(filtros.departamento || '').trim().toUpperCase();
  if (grupoSeleccionado === 'calidad') {
    operaciones = operaciones.concat(['aduanas_entrada', 'aduanas_salida', 'almacen_entrada', 'almacen_salida']);
  } else if (grupoSeleccionado === 'aduanas') {
    operaciones = operaciones.concat(['almacen_entrada', 'almacen_salida']);
  } else if (grupoSeleccionado === 'almacen') {
    operaciones = operaciones.concat(['aduanas_entrada', 'aduanas_salida']);
  }
  operaciones = operaciones.filter(function(operacion) {
    return operacion.endsWith('_' + tipoMovimiento);
  });
  var operacionesTotales = operaciones.slice();
  if (esDEV_(usuario) && departamentoFiltro) {
    operaciones = operaciones.filter(function(operacion) {
      return departamentoDeOperacion_(operacion) === departamentoFiltro;
    });
  }
  var operacionesAduanas = operacionesTotales.filter(function(op) { return op === 'aduanas_entrada' || op === 'aduanas_salida'; });
  var operacionesAlmacen = operacionesTotales.filter(function(op) { return op === 'almacen_entrada' || op === 'almacen_salida'; });
  var operacionesCalidad = operacionesTotales.filter(function(op) { return op === 'calidad_entrada' || op === 'calidad_salida'; });
  var operacionesAduanasVista = operaciones.filter(function(op) { return op === 'aduanas_entrada' || op === 'aduanas_salida'; });
  var operacionesAlmacenVista = operaciones.filter(function(op) { return op === 'almacen_entrada' || op === 'almacen_salida'; });
  var operacionesCalidadVista = operaciones.filter(function(op) { return op === 'calidad_entrada' || op === 'calidad_salida'; });
  var aduanas = agregarMovimientosResumen_(operacionesAduanas, fecha, turno, corte);
  var almacen = agregarMovimientosResumen_(operacionesAlmacen, fecha, turno, corte);
  var calidad = agregarMovimientosResumen_(operacionesCalidad, fecha, turno, corte);
  var aduanasVista = agregarMovimientosResumen_(operacionesAduanasVista, fecha, turno, corte);
  var almacenVista = agregarMovimientosResumen_(operacionesAlmacenVista, fecha, turno, corte);
  var calidadVista = agregarMovimientosResumen_(operacionesCalidadVista, fecha, turno, corte);
  var partes = obtenerCatalogosBase_().partes;
  var llaves = {};
  Object.keys(aduanas).concat(Object.keys(almacen)).concat(Object.keys(calidad)).forEach(function(llave) { llaves[llave] = true; });
  var filas = Object.keys(llaves).map(function(llave) {
    var datosA = aduanas[llave] || { linea: '', parte: '', carros: {}, carroEtiquetas: {}, total: 0 };
    var datosW = almacen[llave] || { carros: {}, carroEtiquetas: {}, total: 0 };
    var datosQ = calidad[llave] || { carros: {}, carroEtiquetas: {}, total: 0 };
    var carros = [];
    var carroEtiquetas = [];
    var datosAVista = aduanasVista[llave] || { carros: {}, carroEtiquetas: {} };
    var datosWVista = almacenVista[llave] || { carros: {}, carroEtiquetas: {} };
    var datosQVista = calidadVista[llave] || { carros: {}, carroEtiquetas: {} };
    var datosCarros = grupoSeleccionado === 'almacen' ? datosWVista : grupoSeleccionado === 'calidad' ? datosQVista : datosAVista;
    for (var i = 1; i <= 50; i++) {
      var totalCarro = datosCarros.carros[i] || 0;
      carros.push(totalCarro);
      carroEtiquetas.push(datosCarros.carroEtiquetas[i] || '');
    }
    var parte = datosA.parte || datosW.parte || datosQ.parte;
    var esLateral = parte === 'B/C L / B/C R';
    var totalAduanas = datosA.total || 0;
    var totalAlmacen = datosW.total || 0;
    var totalCalidad = datosQ.total || 0;
    var esResumenCalidad = grupoSeleccionado === 'calidad';
    var diferencia = esResumenCalidad ? 0 : totalAduanas - totalAlmacen - totalCalidad;
    var estado = esResumenCalidad ? 'OK' : (totalAduanas + totalAlmacen + totalCalidad === 0 || totalAduanas === totalAlmacen + totalCalidad ? 'OK' : 'ERROR');
    return { linea: datosA.linea || datosW.linea || datosQ.linea, parte: parte, tipo: esLateral ? 'B/C L / B/C R' : partes[normalizarParteServidor_(parte)] || 'No registrado', carros: carros, carroEtiquetas: carroEtiquetas, totalAduanas: totalAduanas, totalAlmacen: totalAlmacen, totalCalidad: totalCalidad, diferencia: diferencia, estado: estado };
  }).filter(function(fila) { return fila.totalAduanas || fila.totalAlmacen || fila.totalCalidad; }).sort(function(a, b) { return Number(a.linea) - Number(b.linea) || a.parte.localeCompare(b.parte); });
  return { fecha: fecha, turno: turno, corte: corte, tipoMovimiento: tipoMovimiento, grupoOperacion: grupoSeleccionado, departamento: departamentoFiltro, filas: filas, totalAduanas: filas.reduce(function(s, f) { return s + f.totalAduanas; }, 0), totalAlmacen: filas.reduce(function(s, f) { return s + f.totalAlmacen; }, 0), discrepancias: filas.filter(function(f) { return f.estado === 'ERROR'; }).length };
}

function obtenerResumenDeHoja_(nombreHoja, fecha, corte) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 3) return null;
  var fechaHoja = formatearFecha(hoja.getRange('C1').getValue());
  if (fecha && fechaHoja && fecha !== fechaHoja) return null;
  var esAduanas = nombreHoja === 'Resumen aduanas';
  var corteHoja = String(hoja.getRange('L1').getValue() || '').trim();
  if (corte && corteHoja && corteHoja !== corte) return null;
  var ancho = esAduanas ? 19 : 15;
  var filasHoja = hoja.getRange(3, 1, hoja.getLastRow() - 2, ancho).getValues();
  var partes = obtenerCatalogosBase_().partes;
  var filas = filasHoja.map(function(fila) {
    var parte = normalizarParteServidor_(fila[2]);
    var total = Number(fila[esAduanas ? 17 : 13]) || 0;
    var carros = [];
    for (var i = 3; i <= (esAduanas ? 16 : 12); i++) carros.push(Number(fila[i]) || 0);
    return { linea: fila[0], parte: parte, tipo: partes[parte] || 'No registrado', carros: carros, totalAduanas: esAduanas ? total : 0, totalAlmacen: esAduanas ? 0 : total, diferencia: 0, estado: 'OK' };
  }).filter(function(fila) { return fila.parte || fila.totalAduanas || fila.totalAlmacen; });
  return { fecha: fecha || fechaHoja, corte: corte, filas: filas, totalAduanas: filas.reduce(function(s, f) { return s + f.totalAduanas; }, 0), totalAlmacen: filas.reduce(function(s, f) { return s + f.totalAlmacen; }, 0), discrepancias: 0 };
}

function claveCacheResumenFecha_(hoja, fecha) {
  var idHoja = Utilities.base64EncodeWebSafe(hoja.getName()).replace(/=+$/g, '');
  var idLibro = hoja.getParent().getId();
  return [RESUMEN_CACHE_VERSION_, idLibro, idHoja, String(fecha).replace(/[^0-9]/g, '')].join('_');
}

function leerCacheResumenFecha_(hoja, fecha) {
  try {
    var cache = CacheService.getScriptCache();
    var prefijo = claveCacheResumenFecha_(hoja, fecha);
    var textoMeta = cache.get(prefijo + '_meta');
    if (!textoMeta) return null;
    var meta = JSON.parse(textoMeta);
    if (!meta || !meta.version || !Array.isArray(meta.chunks) || Number(meta.count) < 0) return null;
    var filas = [];
    for (var i = 0; i < meta.chunks.length; i++) {
      var textoBloque = cache.get(prefijo + '_' + meta.version + '_' + i);
      if (!textoBloque) return null;
      var bloque = JSON.parse(textoBloque);
      if (!Array.isArray(bloque)) return null;
      filas = filas.concat(bloque);
    }
    if (filas.length !== meta.count) return null;
    return { meta: meta, filas: filas };
  } catch (error) {
    Logger.log('No se pudo leer la caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + error.message);
    return null;
  }
}

function guardarCacheResumenFecha_(hoja, fecha, filas, ultimaFila, ultimoId) {
  var cache;
  var prefijo;
  var version;
  var bloques = [];
  var metaAnterior = null;
  var meta;
  try {
    cache = CacheService.getScriptCache();
    prefijo = claveCacheResumenFecha_(hoja, fecha);
    version = Utilities.getUuid();
    var bloqueActual = [];
    var bytesActuales = 2;
    filas.forEach(function(fila) {
      var bytesFila = Utilities.newBlob(JSON.stringify(fila)).getBytes().length + 1;
      if (bloqueActual.length && bytesActuales + bytesFila > RESUMEN_CACHE_MAX_BLOQUE_BYTES_) {
        bloques.push(JSON.stringify(bloqueActual));
        bloqueActual = [];
        bytesActuales = 2;
      }
      bloqueActual.push(fila);
      bytesActuales += bytesFila;
    });
    if (bloqueActual.length || !bloques.length) bloques.push(JSON.stringify(bloqueActual));
    meta = {
      version: version,
      chunks: bloques.map(function(_, indice) { return indice; }),
      count: filas.length,
      ultimaFila: ultimaFila,
      ultimoId: ultimoId
    };
    try {
      var textoMetaAnterior = cache.get(prefijo + '_meta');
      if (textoMetaAnterior) metaAnterior = JSON.parse(textoMetaAnterior);
    } catch (errorMetaAnterior) {
      Logger.log('No se pudo leer la versión previa de caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + errorMetaAnterior.message);
    }
  } catch (errorMeta) {
    Logger.log('No se pudo preparar la caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + errorMeta.message);
    return false;
  }
  try {
    bloques.forEach(function(bloque, indice) {
      cache.put(prefijo + '_' + version + '_' + indice, bloque, RESUMEN_CACHE_TTL_SEGUNDOS_);
    });
    cache.put(prefijo + '_meta', JSON.stringify(meta), RESUMEN_CACHE_TTL_SEGUNDOS_);
  } catch (error) {
    Logger.log('No se pudo guardar la caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + error.message);
    return false;
  }
  if (metaAnterior && metaAnterior.version && Array.isArray(metaAnterior.chunks)) {
    try {
      metaAnterior.chunks.forEach(function(indice) {
        cache.remove(prefijo + '_' + metaAnterior.version + '_' + indice);
      });
    } catch (errorLimpieza) {
      Logger.log('No se pudieron limpiar versiones anteriores de caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + errorLimpieza.message);
    }
  }
  return true;
}

function invalidarCacheResumenFecha_(hoja, fecha) {
  if (!fecha) return;
  try {
    CacheService.getScriptCache().remove(claveCacheResumenFecha_(hoja, fecha) + '_meta');
  } catch (error) {
    Logger.log('No se pudo invalidar la caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + error.message);
  }
}

function leerBloqueFechaDesdeAbajo_(hoja, fecha, ultimaFila) {
  var encontrado = false;
  var inicioFecha = -1;
  var finFecha = -1;
  for (var fin = ultimaFila; fin >= 2; fin -= RESUMEN_CACHE_BLOQUE_FILAS_) {
    var inicio = Math.max(2, fin - RESUMEN_CACHE_BLOQUE_FILAS_ + 1);
    var fechas = hoja.getRange(inicio, 2, fin - inicio + 1, 1).getValues();
    for (var indice = fechas.length - 1; indice >= 0; indice--) {
      var filaHoja = inicio + indice;
      var fechaFila = formatearFecha(fechas[indice][0]);
      if (fechaFila === fecha) {
        if (!encontrado) finFecha = filaHoja;
        inicioFecha = filaHoja;
        encontrado = true;
      } else if (encontrado) {
        return { inicio: inicioFecha, fin: finFecha };
      }
    }
  }
  return encontrado ? { inicio: inicioFecha, fin: finFecha } : { inicio: 2, fin: 1 };
}

function construirCacheResumenFecha_(hoja, fecha) {
  var ultimaFila = hoja.getLastRow();
  var filasCache = [];
  if (ultimaFila >= 2) {
    var bloque = leerBloqueFechaDesdeAbajo_(hoja, fecha, ultimaFila);
    if (bloque.fin >= bloque.inicio) {
      var filasHoja = hoja.getRange(bloque.inicio, 1, bloque.fin - bloque.inicio + 1, 15).getValues();
      filasHoja.forEach(function(fila) {
        var valores = fila.slice(1);
        valores[0] = formatearFecha(valores[0]);
        filasCache.push({ id: String(fila[0] || ''), valores: valores });
      });
    }
  }
  var ultimoId = ultimaFila >= 2 ? String(hoja.getRange(ultimaFila, 1).getValue() || '') : '';
  return { filas: filasCache, ultimaFila: ultimaFila, ultimoId: ultimoId };
}

function obtenerFilasResumenCacheadas_(hoja, fecha, turno, corte) {
  var cacheado = leerCacheResumenFecha_(hoja, fecha);
  if (!cacheado) {
    var construido = construirCacheResumenFecha_(hoja, fecha);
    if (!guardarCacheResumenFecha_(hoja, fecha, construido.filas, construido.ultimaFila, construido.ultimoId)) {
      invalidarCacheResumenFecha_(hoja, fecha);
    }
    cacheado = { meta: { ultimaFila: construido.ultimaFila, ultimoId: construido.ultimoId }, filas: construido.filas };
  } else {
    var ultimaFilaActual = hoja.getLastRow();
    var ultimoIdActual = ultimaFilaActual >= 2 ? String(hoja.getRange(ultimaFilaActual, 1).getValue() || '') : '';
    if (ultimaFilaActual !== Number(cacheado.meta.ultimaFila) || ultimoIdActual !== String(cacheado.meta.ultimoId || '')) {
      var filasNuevas = [];
      var hayIdBase = !!cacheado.meta.ultimoId;
      var encontrado = !hayIdBase;
      for (var fin = ultimaFilaActual; fin >= 2 && (!hayIdBase || !encontrado); fin -= RESUMEN_CACHE_BLOQUE_FILAS_) {
        var inicio = Math.max(2, fin - RESUMEN_CACHE_BLOQUE_FILAS_ + 1);
        var filasHoja = hoja.getRange(inicio, 1, fin - inicio + 1, 15).getValues();
        for (var indice = filasHoja.length - 1; indice >= 0; indice--) {
          var fila = filasHoja[indice];
          var id = String(fila[0] || '');
          if (hayIdBase && id === String(cacheado.meta.ultimoId || '')) {
            encontrado = true;
            break;
          }
          if (id) {
            var valores = fila.slice(1);
            valores[0] = formatearFecha(valores[0]);
            filasNuevas.push({ id: id, valores: valores });
          }
        }
      }
      if (encontrado) {
        filasNuevas.reverse().forEach(function(fila) {
          if (formatearFecha(fila.valores[0]) === fecha) cacheado.filas.push(fila);
        });
        cacheado.meta.ultimaFila = ultimaFilaActual;
        cacheado.meta.ultimoId = ultimoIdActual;
        if (!guardarCacheResumenFecha_(hoja, fecha, cacheado.filas, ultimaFilaActual, ultimoIdActual)) {
          invalidarCacheResumenFecha_(hoja, fecha);
        }
      } else {
        var reconstruido = construirCacheResumenFecha_(hoja, fecha);
        cacheado = { meta: { ultimaFila: reconstruido.ultimaFila, ultimoId: reconstruido.ultimoId }, filas: reconstruido.filas };
        if (!guardarCacheResumenFecha_(hoja, fecha, reconstruido.filas, reconstruido.ultimaFila, reconstruido.ultimoId)) {
          invalidarCacheResumenFecha_(hoja, fecha);
        }
      }
    }
  }
  var filas = [];
  cacheado.filas.forEach(function(fila) {
    var valores = fila.valores;
    if (turno && String(valores[8] || '').trim() !== turno) return;
    if (corte && String(valores[11] || '').trim() !== corte) return;
    filas.push(valores);
  });
  return filas;
}

function actualizarCacheResumenTrasCambio_(hoja, fecha) {
  if (!fecha) return;
  try {
    var cache = CacheService.getScriptCache();
    var claveMeta = claveCacheResumenFecha_(hoja, fecha) + '_meta';
    if (!cache.get(claveMeta)) return;
    var reconstruido = construirCacheResumenFecha_(hoja, fecha);
    if (!guardarCacheResumenFecha_(hoja, fecha, reconstruido.filas, reconstruido.ultimaFila, reconstruido.ultimoId)) {
      cache.remove(claveMeta);
    }
  } catch (error) {
    Logger.log('No se pudo actualizar la caché del resumen para ' + hoja.getName() + ' (' + fecha + '): ' + error.message);
    invalidarCacheResumenFecha_(hoja, fecha);
  }
}

function obtenerResumenPorDia(token, fecha, turno) {
  var usuario = requireSession_(token);
  if (!puedeAdministrarUsuarios_(usuario)) throw new Error('Solo administradores y DEV pueden consultar el resumen por dia.');
  turno = String(turno || '').trim();
  var operaciones = operacionesDeUsuario_(usuario);
  var partes = obtenerCatalogosBase_().partes;
  var grupos = {};
  operaciones.forEach(function(operacion) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
    if (!hoja || hoja.getLastRow() < 2) return;
    var filas = fecha
      ? obtenerFilasResumenCacheadas_(hoja, String(fecha).trim(), turno, '')
      : obtenerFilasOperacionFiltradas_(hoja, fecha, turno, '').filas;
    filas.forEach(function(fila) {
      if (fecha && formatearFecha(fila[0]) !== String(fecha).trim()) return;
      if (turno && String(fila[8] || '').trim() !== turno) return;
      var linea = String(fila[6] || '').trim();
      var parte = normalizarParteServidor_(fila[3]);
      var corte = String(fila[11] || '').trim();
      if (!linea || !parte || !['1', '2', '3', 'Acumulado'].includes(corte)) return;
      var llave = linea + '|' + parte;
      if (!grupos[llave]) grupos[llave] = { linea: linea, parte: parte, cortes: { '1': 0, '2': 0, '3': 0, Acumulado: 0 } };
      var cantidad = Number(fila[5]) || 0;
      grupos[llave].cortes[corte] += operacion.indexOf('_salida') >= 0 ? -cantidad : cantidad;
    });
  });
  var filasResultado = Object.keys(grupos).map(function(llave) {
    var fila = grupos[llave];
    var produccion = fila.cortes['1'] + fila.cortes['2'] + fila.cortes['3'] + fila.cortes.Acumulado;
    return { linea: fila.linea, parte: fila.parte, tipo: partes[normalizarParteServidor_(fila.parte)] || 'No registrado', corte1: fila.cortes['1'], corte2: fila.cortes['2'], corte3: fila.cortes['3'], acumulado: fila.cortes.Acumulado, produccionTotal: produccion };
  }).sort(function(a, b) { return Number(a.linea) - Number(b.linea) || a.parte.localeCompare(b.parte); });
  return { fecha: String(fecha || ''), turno: turno, filas: filasResultado, produccionTotalDia: filasResultado.reduce(function(total, fila) { return total + fila.produccionTotal; }, 0) };
}

function agregarMovimientosResumen_(operaciones, fecha, turno, corte) {
  var resultado = {};
  operaciones.forEach(function(operacion) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHojaOperacion_(operacion));
    if (!hoja || hoja.getLastRow() < 2) return;
    var filas = fecha
      ? obtenerFilasResumenCacheadas_(hoja, fecha, turno, corte)
      : obtenerFilasOperacionFiltradas_(hoja, fecha, turno, corte).filas;
    filas.forEach(function(fila) {
      var fechaFila = formatearFecha(fila[0]);
      if (fecha && fechaFila !== fecha) return;
      if (turno && String(fila[8] || '').trim() !== turno) return;
      if (corte && String(fila[11]).trim() !== corte) return;
      var linea = String(fila[6] || '').trim();
      var parte = normalizarParteServidor_(fila[3]);
      if (!linea || !parte) return;
      var carroTexto = String(fila[10] || '').replace(/[^0-9]/g, '');
      var carro = Number(carroTexto);
      var esLateral = parte === 'B/C L' || parte === 'B/C R';
      var llave = linea + '|' + (esLateral ? 'B/C LR' : parte);
        if (!resultado[llave]) resultado[llave] = { linea: linea, parte: esLateral ? 'B/C L / B/C R' : parte, carros: {}, carroEtiquetas: {}, total: 0 };
      var cantidad = Number(fila[5]) || 0;
      resultado[llave].total += cantidad;
        if (carro >= 1 && carro <= 50) {
          resultado[llave].carros[carro] = (resultado[llave].carros[carro] || 0) + cantidad;
          if (!resultado[llave].carroEtiquetas[carro]) resultado[llave].carroEtiquetas[carro] = String(fila[4] || '').trim();
        }
    });
  });
  return resultado;
}

function obtenerFilasOperacionFiltradas_(hoja, fecha, turno, corte) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { inicio: 2, filas: [] };

  var inicio = 2;
  var fin = ultimaFila;
  var fechaBuscada = String(fecha || '').trim();
  var turnoBuscado = String(turno || '').trim();
  var corteBuscado = String(corte || '').trim();

  if (fechaBuscada) {
    var bloqueFecha = leerBloqueFechaDesdeAbajo_(hoja, fechaBuscada, ultimaFila);
    if (bloqueFecha.fin < bloqueFecha.inicio) return { inicio: 2, filas: [] };
    inicio = bloqueFecha.inicio;
    fin = bloqueFecha.fin;
  }

  if (turnoBuscado && corteBuscado) {
    var inicioFiltros = inicio;
    var filtrosTurnoCorte = hoja.getRange(inicio, 10, fin - inicio + 1, 4).getValues();
    var primerTurno = -1;
    var ultimoTurno = -1;
    filtrosTurnoCorte.forEach(function(fila, indice) {
      if (String(fila[0] || '').trim() === turnoBuscado) {
        if (primerTurno < 0) primerTurno = indice;
        ultimoTurno = indice;
      }
    });
    if (primerTurno < 0) return { inicio: inicio, filas: [] };
    inicio = inicioFiltros + primerTurno;
    fin = inicioFiltros + ultimoTurno;
    var primerCorte = -1;
    var ultimoCorte = -1;
    for (var indiceFiltro = inicio - inicioFiltros; indiceFiltro <= fin - inicioFiltros; indiceFiltro++) {
      if (String(filtrosTurnoCorte[indiceFiltro][3] || '').trim() === corteBuscado) {
        if (primerCorte < 0) primerCorte = indiceFiltro;
        ultimoCorte = indiceFiltro;
      }
    }
    if (primerCorte < 0) return { inicio: inicio, filas: [] };
    inicio = inicioFiltros + primerCorte;
    fin = inicioFiltros + ultimoCorte;
  } else if (turnoBuscado) {
    var turnos = hoja.getRange(inicio, 10, fin - inicio + 1, 1).getValues();
    var primerTurno = -1;
    var ultimoTurno = -1;
    turnos.forEach(function(fila, indice) {
      if (String(fila[0] || '').trim() === turnoBuscado) {
        if (primerTurno < 0) primerTurno = indice;
        ultimoTurno = indice;
      }
    });
    if (primerTurno < 0) return { inicio: inicio, filas: [] };
    inicio += primerTurno;
    fin = inicio + (ultimoTurno - primerTurno);
  } else if (corteBuscado) {
    var cortes = hoja.getRange(inicio, 13, fin - inicio + 1, 1).getValues();
    var primerCorte = -1;
    var ultimoCorte = -1;
    cortes.forEach(function(fila, indice) {
      if (String(fila[0] || '').trim() === corteBuscado) {
        if (primerCorte < 0) primerCorte = indice;
        ultimoCorte = indice;
      }
    });
    if (primerCorte < 0) return { inicio: inicio, filas: [] };
    inicio += primerCorte;
    fin = inicio + (ultimoCorte - primerCorte);
  }

  var cantidadFilas = fin - inicio + 1;
  var columnasDisponibles = hoja.getMaxColumns();
  var columnasALeer = Math.min(14, Math.max(0, columnasDisponibles - 1));
  var filas = columnasALeer ? hoja.getRange(inicio, 2, cantidadFilas, columnasALeer).getValues() : [];
  if (columnasALeer < 14) {
    filas.forEach(function(fila) {
      while (fila.length < 14) fila.push('');
    });
  }
  return { inicio: inicio, filas: filas };
}

function formatearFecha(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (typeof val === 'number') return Utilities.formatDate(new Date(Math.round((val - 25569) * 86400 * 1000)), 'UTC', 'yyyy-MM-dd');
  return String(val);
}
