/**
 * Finanzas Consultorio – Dr. Villamizar
 * Servidor (Google Apps Script). Va pegado en Extensiones > Apps Script de la hoja de Google.
 * Primera vez: ejecuta la función instalar() desde el editor.
 */

const TZ = 'America/Lima';
const BIOPSIA_CU = 80;               // Valor por defecto; el real se edita en la pestaña Ajustes
const FORMAS = ['Efectivo', 'Transferencia', 'Yape', 'Tarjeta', 'Depósito'];

const H = {
  Usuarios:       ['id', 'nombre', 'rol', 'pin', 'activo'],
  Catalogo:       ['id', 'categoria', 'servicio', 'precio', 'tipo', 'pap', 'biopsia_cu', 'activo', 'pide_detalle'],
  Presupuesto:    ['id', 'dia', 'concepto', 'monto', 'categoria', 'desde', 'hasta', 'estimado', 'activo'],
  Movimientos:    ['id', 'fecha', 'mes', 'usuario', 'tipo', 'bolsa', 'concepto', 'monto', 'forma_pago',
                   'paciente', 'proc_id', 'pres_id', 'origen_id', 'estado', 'nota', 'anulado', 'comprobante', 'ref'],
  Procedimientos: ['id', 'fecha', 'tipo', 'tipo_doc', 'num_doc', 'nombres', 'apellidos', 'celular',
                   'precio_total', 'presupuesto_gastos', 'forma_cobro', 'nota', 'usuario', 'anulado', 'fecha_proc'],
  Pacientes:      ['num_doc', 'tipo_doc', 'nombres', 'apellidos', 'celular', 'actualizado'],
  Correcciones:   ['id', 'fecha', 'usuario', 'mov_id', 'detalle', 'motivo', 'estado'],
  Ajustes:        ['ajuste', 'valor', 'nota'],
  Plantillas:     ['procedimiento', 'concepto', 'monto_sugerido', 'orden', 'activo'],
  Equivalencias:  ['codigo', 'servicio_id', 'precio_desde', 'descripcion']
};
// Columnas que deben guardarse como texto (para no perder ceros ni convertir a fecha)
const TEXTO = {
  Usuarios: ['id', 'pin'], Catalogo: ['id'], Presupuesto: ['id', 'desde', 'hasta'],
  Movimientos: ['id', 'mes', 'proc_id', 'pres_id', 'origen_id', 'ref'],
  Equivalencias: ['codigo', 'servicio_id'],
  Procedimientos: ['id', 'num_doc', 'celular', 'fecha_proc'], Pacientes: ['num_doc', 'celular'],
  Correcciones: ['id', 'mov_id']
};

// ================= Web app =================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Consultorio – Finanzas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setFaviconUrl('https://danielvillamizar.com/consultorio/icon-192.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================= Utilidades =================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(n) {
  const s = ss_().getSheetByName(n);
  if (!s) throw new Error('Falta la hoja "' + n + '". Ejecuta instalar() en el editor.');
  return s;
}
// Memoria por ejecución (cada pestaña se lee una sola vez por acción) y caché de configuración (5 min)
const MEM_ = {};
const CONF_ = ['Usuarios', 'Catalogo', 'Presupuesto', 'Plantillas', 'Ajustes', 'Equivalencias'];
function valores_(n) {
  if (MEM_[n]) return MEM_[n];
  let v = null;
  const cache = CacheService.getScriptCache();
  if (CONF_.indexOf(n) >= 0) {
    const c = cache.get('hoja_' + n);
    if (c) { try { v = JSON.parse(c); } catch (e) { v = null; } }
  }
  if (!v) {
    v = sh_(n).getDataRange().getValues();
    if (CONF_.indexOf(n) >= 0) {
      v = v.map(function (r) { return r.map(function (x) { return x instanceof Date ? Utilities.formatDate(x, TZ, n === 'Presupuesto' ? 'yyyy-MM' : 'yyyy-MM-dd') : x; }); });
      try { cache.put('hoja_' + n, JSON.stringify(v), 300); } catch (e) {}
    }
  }
  MEM_[n] = v;
  return v;
}
function olvidar_(n) {
  delete MEM_[n];
  if (CONF_.indexOf(n) >= 0) CacheService.getScriptCache().remove('hoja_' + n);
}
function limpiarCache() { CONF_.forEach(function (n) { olvidar_(n); }); }
// Al editar la hoja a mano, la app toma los cambios de inmediato
function onEdit(e) { try { const n = e.range.getSheet().getName(); if (CONF_.indexOf(n) >= 0) olvidar_(n); } catch (err) {} }
function rows_(n) {
  const v = valores_(n).slice();
  const h = v.shift() || [];
  const out = [];
  v.forEach(function (r, i) {
    if (r[0] === '' || r[0] === null) return;
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    out.push(o);
  });
  return out;
}
function add_(n, o) {
  olvidar_(n);
  sh_(n).appendRow(H[n].map(function (k) { return (o[k] === undefined || o[k] === null) ? '' : o[k]; }));
}
function set_(n, row, changes) {
  olvidar_(n);
  const s = sh_(n);
  Object.keys(changes).forEach(function (k) {
    const c = H[n].indexOf(k);
    if (c >= 0) s.getRange(row, c + 1).setValue(changes[k]);
  });
}
function id_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }
function yes_(v) { return v === true || String(v).trim().toUpperCase() === 'SI'; }
function mm_(v) { return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM') : String(v || '').trim(); }
function mesActual_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM'); }
function fmt_(d, p) { return (d instanceof Date) ? Utilities.formatDate(d, TZ, p || 'dd/MM/yyyy HH:mm') : String(d || ''); }
function difMeses_(a, b) {
  const x = a.split('-').map(Number), y = b.split('-').map(Number);
  return (y[0] - x[0]) * 12 + (y[1] - x[1]);
}
function lock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}
function suma_(arr) { return num_(arr.reduce(function (s, m) { return s + num_(m.monto); }, 0)); }
function ajuste_(k, def) {
  try {
    const r = rows_('Ajustes').find(function (x) { return String(x.ajuste).trim() === k; });
    return r && String(r.valor) !== '' ? num_(r.valor) : def;
  } catch (e) { return def; }
}
function biopsiaCU_() { return ajuste_('biopsia_cuello_uterino', BIOPSIA_CU); }
const COMP = ['Factura', 'Boleta', 'Recibo', 'Código de operación', 'Sin comprobante'];
function comprobante_(tipo, numero) {
  tipo = String(tipo || '').trim(); numero = String(numero || '').trim();
  if (!tipo && !numero) return '';
  if (tipo && COMP.indexOf(tipo) < 0) throw new Error('Tipo de comprobante no válido.');
  if (tipo === 'Sin comprobante') return 'Sin comprobante';
  return ((tipo || 'Comprobante') + (numero ? ' ' + numero : '')).trim();
}
function fechaOk_(f) {
  f = String(f || '').trim();
  if (f && !/^\d{4}-\d{2}-\d{2}$/.test(f)) throw new Error('Revisa la fecha de la cirugía.');
  return f;
}
function hoyIso_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function fechaPago_(f) {
  f = String(f || '').trim() || hoyIso_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f)) throw new Error('Revisa la fecha de pago.');
  if (f > hoyIso_()) throw new Error('La fecha de pago no puede ser futura.');
  return { iso: f, date: new Date(f + 'T12:00:00-05:00'), mes: f.slice(0, 7) };
}
function plantillas_() {
  const out = {};
  try {
    rows_('Plantillas').filter(function (p) { return String(p.activo).toUpperCase() !== 'NO'; })
      .sort(function (a, b) { return num_(a.orden) - num_(b.orden); })
      .forEach(function (p) {
        const k = String(p.procedimiento).trim();
        (out[k] = out[k] || []).push({ concepto: String(p.concepto).trim(), monto: String(p.monto_sugerido) === '' ? '' : num_(p.monto_sugerido) });
      });
  } catch (e) {}
  return out;
}
function fechaTxt_(f) { f = mm_(f); return /^\d{4}-\d{2}-\d{2}/.test(f) ? f.slice(8, 10) + '/' + f.slice(5, 7) + '/' + f.slice(0, 4) : ''; }
function forma_(f) { if (FORMAS.indexOf(f) < 0) throw new Error('Elige la forma de pago.'); return f; }

// ================= Acceso con PIN =================
function usuarios() {
  return rows_('Usuarios')
    .filter(function (u) { return String(u.activo).toUpperCase() !== 'NO'; })
    .map(function (u) { return { id: String(u.id), nombre: u.nombre }; });
}
function chk_(a) {
  if (!a || !a.id) throw new Error('Vuelve a ingresar con tu PIN.');
  const cache = CacheService.getScriptCache();
  const key = 'fallos_' + a.id;
  const fallos = Number(cache.get(key) || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos. Espera 10 minutos.');
  const u = rows_('Usuarios').find(function (x) {
    return String(x.id) === String(a.id) && String(x.pin).trim() === String(a.pin).trim() &&
      String(x.activo).toUpperCase() !== 'NO';
  });
  if (!u) { cache.put(key, String(fallos + 1), 600); throw new Error('PIN incorrecto.'); }
  return u;
}
function cambiarPin(a, actual, nuevo) {
  const u = chk_({ id: a && a.id, pin: actual });
  nuevo = String(nuevo || '').trim();
  if (!/^\d{6}$/.test(nuevo)) throw new Error('El PIN nuevo debe tener 6 números.');
  if (nuevo === String(actual).trim()) throw new Error('El PIN nuevo debe ser distinto al actual.');
  set_('Usuarios', u._row, { pin: nuevo });
  return { ok: true, pin: nuevo };
}
function adm_(a) {
  const u = chk_(a);
  if (String(u.rol).toLowerCase() !== 'admin') throw new Error('Solo el doctor puede hacer esto.');
  return u;
}
function login(id, pin) {
  const u = chk_({ id: id, pin: pin });
  return { id: String(u.id), nombre: u.nombre, rol: String(u.rol).toLowerCase(), pin: String(pin) };
}

// ================= Datos iniciales para la app =================
function init(a) {
  const u = chk_(a);
  const movs = rows_('Movimientos');
  const cat = rows_('Catalogo').filter(function (c) { return yes_(c.activo); }).map(function (c) {
    return { id: String(c.id), categoria: c.categoria, servicio: c.servicio, precio: num_(c.precio),
      tipo: c.tipo, pap: num_(c.pap), biopsia: yes_(c.biopsia_cu), detalle: yes_(c.pide_detalle) };
  });
  const mes = mesActual_();
  const pres = presMes_(mes, movs.filter(function (m) { return mm_(m.mes) === mes && !yes_(m.anulado); }));
  const cnt = {};
  movs.forEach(function (m) {
    if (m.tipo === 'ingreso' && m.bolsa === 'consultorio' && !yes_(m.anulado)) cnt[m.concepto] = (cnt[m.concepto] || 0) + 1;
  });
  const top = Object.keys(cnt).sort(function (x, y) { return cnt[y] - cnt[x]; }).slice(0, 4);
  return { user: { id: String(u.id), nombre: u.nombre, rol: String(u.rol).toLowerCase() },
    cat: cat, pres: pres, top: top, procs: procs_(movs), formas: FORMAS, comps: COMP, biopsiaCU: biopsiaCU_(), mes: mes, plantillas: plantillas_(), hoy: hoyIso_() };
}

function arranque(a, mes) {
  const ini = init(a);
  const r = resumen_(mes || ini.mes);
  delete r.all; delete r.mv; delete r.procs;
  return { init: ini, resumen: r };
}

// ================= Procedimientos =================
function procs_(movs) {
  movs = movs || rows_('Movimientos');
  const act = movs.filter(function (m) { return !yes_(m.anulado); });
  return rows_('Procedimientos').filter(function (p) { return !yes_(p.anulado); }).map(function (p) {
    const id = String(p.id);
    const mine = act.filter(function (m) { return String(m.proc_id) === id; });
    const cobrado = suma_(mine.filter(function (m) { return m.tipo === 'ingreso'; }));
    const gastos = suma_(mine.filter(function (m) { return m.tipo === 'gasto'; }));
    const patologia = suma_(mine.filter(function (m) { return m.tipo === 'patologia'; }));
    const pend = mine.filter(function (m) { return m.tipo === 'previsto' && m.estado === 'pendiente'; });
    const porFacturar = suma_(pend);
    const precio = num_(p.precio_total);
    return { id: id, fecha: fmt_(p.fecha, 'dd/MM/yyyy'), mesCreado: fmt_(p.fecha, 'yyyy-MM'), tipo: p.tipo,
      tipo_doc: p.tipo_doc, num_doc: String(p.num_doc), nombre: (p.nombres + ' ' + p.apellidos).trim(),
      celular: String(p.celular), precio: precio, presupuesto: num_(p.presupuesto_gastos),
      forma_cobro: p.forma_cobro, nota: p.nota, cobrado: cobrado, gastos: num_(gastos + patologia),
      patologia: patologia, porFacturar: porFacturar, saldo: Math.max(0, num_(precio - cobrado)), ganancia: num_(cobrado - gastos - patologia),
      gananciaEst: num_(precio - gastos - patologia - porFacturar), pendientes: pend.length, faltanDatos: !String(p.num_doc || '').trim(),
      nombres: p.nombres, apellidos: p.apellidos, fecha_proc: fechaTxt_(p.fecha_proc), fecha_proc_iso: mm_(p.fecha_proc).slice(0, 10) };
  }).reverse();
}

function valPac_(d) {
  const t = d.tipo_doc;
  const n = String(d.num_doc || '').trim().toUpperCase();
  if (['DNI', 'CE', 'PAS'].indexOf(t) < 0) throw new Error('Elige el tipo de documento.');
  if (t === 'DNI' && !/^\d{8}$/.test(n)) throw new Error('El DNI debe tener 8 dígitos.');
  if (t === 'CE' && !/^[A-Z0-9]{9,12}$/.test(n)) throw new Error('El carnet de extranjería debe tener al menos 9 caracteres.');
  if (t === 'PAS' && !/^[A-Z0-9]{6,12}$/.test(n)) throw new Error('El pasaporte debe tener entre 6 y 12 letras o números.');
  const nom = String(d.nombres || '').trim(), ape = String(d.apellidos || '').trim();
  if (!nom || !ape) throw new Error('Escribe nombres y apellidos.');
  const cel = String(d.celular || '').replace(/[\s-]/g, '');
  if (cel.charAt(0) === '+') { if (!/^\+\d{8,15}$/.test(cel)) throw new Error('Revisa el celular con código de país.'); }
  else if (!/^9\d{8}$/.test(cel)) throw new Error('El celular debe tener 9 dígitos y empezar con 9.');
  return { tipo_doc: t, num_doc: n, nombres: titulo_(nom), apellidos: titulo_(ape), celular: cel };
}
function titulo_(s) { return s.toLowerCase().replace(/(^|\s)(\S)/g, function (m, a, b) { return a + b.toUpperCase(); }); }

function buscarPaciente(a, num) {
  chk_(a);
  const n = String(num || '').trim().toUpperCase();
  const p = rows_('Pacientes').find(function (x) { return String(x.num_doc).toUpperCase() === n; });
  return p ? { tipo_doc: p.tipo_doc, num_doc: String(p.num_doc), nombres: p.nombres, apellidos: p.apellidos, celular: String(p.celular) } : null;
}
function upsertPac_(p) {
  const r = rows_('Pacientes').find(function (x) { return String(x.num_doc).toUpperCase() === p.num_doc; });
  const o = Object.assign({}, p, { actualizado: new Date() });
  if (r) set_('Pacientes', r._row, o); else add_('Pacientes', o);
}

function crearProcedimiento(a, d) {
  const u = chk_(a);
  const pac = valPac_(d);
  const cat = rows_('Catalogo').find(function (x) { return String(x.id) === String(d.cat_id) && x.tipo === 'procedimiento'; });
  if (!cat) throw new Error('Elige el tipo de procedimiento.');
  const detalle = String(d.detalle || '').trim();
  if (yes_(cat.pide_detalle) && !detalle) throw new Error('Especifica la cirugía o el procedimiento.');
  const tipoTxt = cat.servicio + (detalle ? ' · ' + detalle : '');
  const precio = num_(d.precio), pago = num_(d.pago);
  let pres = num_(d.presupuesto);
  if (!(precio > 0)) throw new Error('Ingresa el precio total.');
  if (pres < 0) throw new Error('El presupuesto de gastos no puede ser negativo.');
  if (pago < 0 || pago > precio) throw new Error('El pago no puede ser mayor al precio total.');
  if (pago > 0) forma_(d.forma);
  const fproc = fechaOk_(d.fecha_proc);
  const gastos = (d.gastos || []).filter(function (g) { return String(g.concepto || '').trim(); }).map(function (g) {
    const o = { concepto: String(g.concepto).trim(), monto: num_(g.monto), pagado: !!g.pagado };
    if (o.pagado) {
      if (!(o.monto > 0)) throw new Error('Escribe el monto de "' + o.concepto + '" para marcarlo como pagado.');
      o.forma = forma_(g.forma);
      o.f = fechaPago_(g.fecha || fproc);
      o.comp = comprobante_(g.comp_tipo, g.comp_num);
    }
    return o;
  });
  if (gastos.length && !(pres > 0)) pres = suma_(gastos);
  return lock_(function () {
    upsertPac_(pac);
    const id = id_(), now = new Date();
    add_('Procedimientos', Object.assign({ id: id, fecha: now, tipo: tipoTxt }, pac, {
      precio_total: precio, presupuesto_gastos: pres,
      forma_cobro: d.forma_cobro === 'adelantos' ? 'Con adelantos' : 'Pago único',
      nota: d.nota || '', usuario: u.nombre, anulado: 'NO', fecha_proc: fproc }));
    if (pago > 0) {
      add_('Movimientos', { id: id_(), fecha: now, mes: mesActual_(), usuario: u.nombre, tipo: 'ingreso',
        bolsa: 'procedimiento', concepto: tipoTxt + ' – ' + (pago >= precio ? 'pago completo' : 'adelanto'),
        monto: pago, forma_pago: d.forma, paciente: pac.nombres + ' ' + pac.apellidos, proc_id: id, anulado: 'NO' });
    }
    const nombre = pac.nombres + ' ' + pac.apellidos;
    gastos.forEach(function (g) {
      const pid = id_();
      add_('Movimientos', { id: pid, fecha: now, mes: mesActual_(), usuario: u.nombre, tipo: 'previsto', bolsa: 'procedimiento',
        concepto: g.concepto, monto: g.monto, paciente: nombre, proc_id: id, estado: g.pagado ? 'facturado' : 'pendiente', anulado: 'NO' });
      if (g.pagado) {
        add_('Movimientos', { id: id_(), fecha: g.f.date, mes: g.f.mes, usuario: u.nombre, tipo: 'gasto', bolsa: 'procedimiento',
          concepto: g.concepto, monto: g.monto, forma_pago: g.forma, paciente: nombre, proc_id: id, origen_id: pid,
          anulado: 'NO', comprobante: g.comp });
      }
    });
    return { ok: true, id: id };
  });
}

function detalleProc(a, id) {
  chk_(a);
  const movs = rows_('Movimientos');
  const p = procs_(movs).find(function (x) { return x.id === String(id); });
  if (!p) throw new Error('Procedimiento no encontrado.');
  p.movs = movs.filter(function (m) { return String(m.proc_id) === String(id); }).map(movOut_).reverse();
  return p;
}

function fechaProc(a, id, fecha) {
  chk_(a);
  const f = fechaOk_(fecha);
  const p = rows_('Procedimientos').find(function (x) { return String(x.id) === String(id); });
  if (!p) throw new Error('Procedimiento no encontrado.');
  set_('Procedimientos', p._row, { fecha_proc: f });
  return { ok: true };
}
function anotarPorFacturar(a, procId, concepto, monto) {
  const u = chk_(a);
  concepto = String(concepto || '').trim(); monto = num_(monto);
  if (!concepto) throw new Error('Escribe qué factura falta llegar.');
  if (monto < 0) throw new Error('El monto no puede ser negativo.');
  return lock_(function () {
    const p = procs_().find(function (x) { return x.id === String(procId); });
    if (!p) throw new Error('Procedimiento no encontrado.');
    add_('Movimientos', { id: id_(), fecha: new Date(), mes: mesActual_(), usuario: u.nombre, tipo: 'previsto', bolsa: 'procedimiento',
      concepto: concepto, monto: monto, paciente: p.nombre, proc_id: p.id, estado: 'pendiente', anulado: 'NO' });
    return { ok: true };
  });
}
function registrarFactura(a, prevId, d) {
  const u = chk_(a);
  const monto = num_(d.monto);
  if (!(monto > 0)) throw new Error('Ingresa el monto real.');
  forma_(d.forma);
  const comp = comprobante_(d.comp_tipo, d.comp_num);
  const f = fechaPago_(d.fecha);
  return lock_(function () {
    const pv = rows_('Movimientos').find(function (m) { return String(m.id) === String(prevId) && m.tipo === 'previsto'; });
    if (!pv || yes_(pv.anulado)) throw new Error('Factura por llegar no encontrada.');
    if (pv.estado !== 'pendiente') throw new Error('Esta factura ya fue registrada.');
    add_('Movimientos', { id: id_(), fecha: f.date, mes: f.mes, usuario: u.nombre, tipo: 'gasto', bolsa: 'procedimiento',
      concepto: pv.concepto, monto: monto, forma_pago: d.forma, paciente: pv.paciente, proc_id: String(pv.proc_id),
      origen_id: String(pv.id), nota: d.nota || '', anulado: 'NO', comprobante: comp });
    set_('Movimientos', pv._row, { estado: 'facturado' });
    return { ok: true };
  });
}
function pagarPendientes(a, procId, items) {
  const u = chk_(a);
  items = items || [];
  const prep = items.map(function (it) {
    const o = { id: String(it.id), monto: num_(it.monto), pagado: !!it.pagado };
    if (o.monto < 0) throw new Error('Revisa los montos.');
    if (o.pagado) {
      if (!(o.monto > 0)) throw new Error('Escribe el monto de cada concepto marcado como pagado.');
      o.forma = forma_(it.forma); o.f = fechaPago_(it.fecha); o.comp = comprobante_(it.comp_tipo, it.comp_num);
    }
    return o;
  });
  return lock_(function () {
    const all = rows_('Movimientos');
    let n = 0;
    prep.forEach(function (o) {
      const pv = all.find(function (m) { return String(m.id) === o.id && m.tipo === 'previsto' && String(m.proc_id) === String(procId); });
      if (!pv || yes_(pv.anulado) || pv.estado !== 'pendiente') return;
      if (o.pagado) {
        add_('Movimientos', { id: id_(), fecha: o.f.date, mes: o.f.mes, usuario: u.nombre, tipo: 'gasto', bolsa: 'procedimiento',
          concepto: pv.concepto, monto: o.monto, forma_pago: o.forma, paciente: pv.paciente, proc_id: String(pv.proc_id),
          origen_id: String(pv.id), anulado: 'NO', comprobante: o.comp });
        set_('Movimientos', pv._row, { estado: 'facturado' });
        n++;
      } else if (o.monto !== num_(pv.monto)) {
        set_('Movimientos', pv._row, { monto: o.monto });
      }
    });
    return { ok: true, pagados: n };
  });
}
function setComprobante(a, movId, tipo, numero) {
  chk_(a);
  const comp = comprobante_(tipo, numero);
  const m = rows_('Movimientos').find(function (x) { return String(x.id) === String(movId); });
  if (!m || m.tipo !== 'gasto') throw new Error('Gasto no encontrado.');
  set_('Movimientos', m._row, { comprobante: comp });
  return { ok: true };
}
function porFacturar(a) {
  chk_(a);
  const movs = rows_('Movimientos');
  const procs = procs_(movs);
  return movs.filter(function (m) { return m.tipo === 'previsto' && m.estado === 'pendiente' && !yes_(m.anulado); }).map(function (m) {
    const o = movOut_(m);
    const p = procs.find(function (x) { return x.id === String(m.proc_id); });
    o.proc = p ? p.tipo + ' · ' + p.nombre + (p.fecha_proc ? ' (' + p.fecha_proc + ')' : '') : '';
    return o;
  }).reverse();
}
function anularProc(a, id) {
  adm_(a);
  return lock_(function () {
    const p = rows_('Procedimientos').find(function (x) { return String(x.id) === String(id); });
    if (!p) throw new Error('Procedimiento no encontrado.');
    set_('Procedimientos', p._row, { anulado: 'SI' });
    rows_('Movimientos').forEach(function (m) { if (String(m.proc_id) === String(id)) set_('Movimientos', m._row, { anulado: 'SI' }); });
    return { ok: true };
  });
}

// ================= Registros =================
function registrarIngreso(a, d) {
  const u = chk_(a);
  const monto = num_(d.monto);
  if (!(monto > 0)) throw new Error('Ingresa un monto mayor a 0.');
  forma_(d.forma);
  return lock_(function () {
    const now = new Date(), mes = mesActual_(), id = id_();
    const base = { fecha: now, mes: mes, usuario: u.nombre, forma_pago: d.forma, nota: d.nota || '', anulado: 'NO' };
    if (d.proc_id) {
      const p = procs_().find(function (x) { return x.id === String(d.proc_id); });
      if (!p) throw new Error('Procedimiento no encontrado.');
      if (monto > p.saldo + 0.001) throw new Error('El monto supera el saldo pendiente (S/ ' + p.saldo + ').');
      add_('Movimientos', Object.assign({}, base, { id: id, tipo: 'ingreso', bolsa: 'procedimiento',
        concepto: p.tipo + ' – ' + (p.saldo - monto <= 0.001 ? 'pago final' : 'abono'), monto: monto,
        paciente: p.nombre, proc_id: p.id }));
      return { ok: true, neto: monto };
    }
    const c = rows_('Catalogo').find(function (x) { return String(x.id) === String(d.cat_id); });
    if (!c) throw new Error('Elige el servicio.');
    const pac = String(d.paciente || '').trim();
    add_('Movimientos', Object.assign({}, base, { id: id, tipo: 'ingreso', bolsa: 'consultorio',
      concepto: c.servicio, monto: monto, paciente: pac }));
    let desc = 0;
    function pat(concepto, m) {
      desc += m;
      add_('Movimientos', { id: id_(), fecha: now, mes: mes, usuario: u.nombre, tipo: 'patologia', bolsa: 'consultorio',
        concepto: concepto, monto: m, paciente: pac, origen_id: id, estado: 'pendiente', anulado: 'NO' });
    }
    if (num_(c.pap) > 0) pat('PAP – ' + c.servicio, num_(c.pap));
    if (d.biopsia && yes_(c.biopsia_cu)) pat('Biopsia de cuello uterino – ' + c.servicio, biopsiaCU_());
    return { ok: true, neto: num_(monto - desc) };
  });
}

function registrarGasto(a, d) {
  const u = chk_(a);
  const monto = num_(d.monto);
  if (!(monto > 0)) throw new Error('Ingresa un monto mayor a 0.');
  const conc = String(d.concepto || '').trim(), nota = d.nota || '';
  return lock_(function () {
    const now = new Date(), mes = mesActual_();
    if (d.destino === 'procedimiento') {
      const p = procs_().find(function (x) { return x.id === String(d.proc_id); });
      if (!p) throw new Error('Elige el procedimiento.');
      if (d.biopsia) {
        add_('Movimientos', { id: id_(), fecha: now, mes: mes, usuario: u.nombre, tipo: 'patologia', bolsa: 'procedimiento',
          concepto: 'Biopsia de pieza quirúrgica – ' + p.tipo + (conc ? ' (' + conc + ')' : ''), monto: monto,
          paciente: p.nombre, proc_id: p.id, estado: 'pendiente', nota: nota, anulado: 'NO' });
        return { ok: true };
      }
      if (!conc) throw new Error('Escribe el concepto del gasto.');
      forma_(d.forma);
      add_('Movimientos', { id: id_(), fecha: now, mes: mes, usuario: u.nombre, tipo: 'gasto', bolsa: 'procedimiento',
        concepto: conc, monto: monto, forma_pago: d.forma, paciente: p.nombre, proc_id: p.id, nota: nota, anulado: 'NO',
        comprobante: comprobante_(d.comp_tipo, d.comp_num) });
      return { ok: true };
    }
    forma_(d.forma);
    let concepto = conc, pres = '';
    if (d.pres_id) {
      const p = rows_('Presupuesto').find(function (x) { return String(x.id) === String(d.pres_id); });
      if (!p) throw new Error('Pago del presupuesto no encontrado.');
      pres = String(p.id);
      concepto = p.concepto + (conc ? ' – ' + conc : '');
    } else if (!conc) throw new Error('Escribe el concepto del gasto.');
    add_('Movimientos', { id: id_(), fecha: now, mes: mes, usuario: u.nombre, tipo: 'gasto', bolsa: 'consultorio',
      concepto: concepto, monto: monto, forma_pago: d.forma, pres_id: pres, nota: nota, anulado: 'NO',
      comprobante: comprobante_(d.comp_tipo, d.comp_num) });
    return { ok: true };
  });
}

function transferir(a, monto, nota) {
  const u = adm_(a);
  monto = num_(monto);
  if (!(monto > 0)) throw new Error('Ingresa un monto mayor a 0.');
  return lock_(function () {
    add_('Movimientos', { id: id_(), fecha: new Date(), mes: mesActual_(), usuario: u.nombre, tipo: 'transferencia',
      bolsa: 'consultorio', concepto: 'Paso de procedimientos a consultorio', monto: monto, nota: nota || '', anulado: 'NO' });
    return { ok: true };
  });
}

// ================= Movimientos, anular y correcciones =================
function movOut_(m) {
  return { id: String(m.id), fecha: fmt_(m.fecha, 'dd/MM HH:mm'), tipo: m.tipo, bolsa: m.bolsa, concepto: m.concepto,
    monto: num_(m.monto), forma: m.forma_pago, paciente: m.paciente, usuario: m.usuario, estado: m.estado,
    nota: m.nota, anulado: yes_(m.anulado), origen: String(m.origen_id || ''), proc_id: String(m.proc_id || ''),
    comprobante: m.comprobante || '' };
}
function movimientos(a, mes) {
  chk_(a);
  mes = mes || mesActual_();
  return rows_('Movimientos').filter(function (m) { return mm_(m.mes) === mes; }).map(movOut_).reverse();
}
function anular(a, id) {
  adm_(a);
  return lock_(function () {
    let n = 0;
    const all = rows_('Movimientos');
    all.forEach(function (m) {
      const tgt = all.find(function (x) { return String(x.id) === String(id); });
      const mismoRef = tgt && tgt.ref && tgt.tipo === 'ingreso' && m.tipo === 'patologia' && String(m.ref) === String(tgt.ref);
      if (String(m.id) === String(id) || String(m.origen_id) === String(id) || mismoRef) {
        set_('Movimientos', m._row, { anulado: 'SI' }); n++;
        if (String(m.id) === String(id) && m.tipo === 'gasto' && m.origen_id) {
          const pv = all.find(function (x) { return String(x.id) === String(m.origen_id) && x.tipo === 'previsto'; });
          if (pv && !yes_(pv.anulado)) set_('Movimientos', pv._row, { estado: 'pendiente' });
        }
      }
    });
    if (!n) throw new Error('Registro no encontrado.');
    rows_('Correcciones').forEach(function (c) {
      if (String(c.mov_id) === String(id) && c.estado === 'pendiente') set_('Correcciones', c._row, { estado: 'resuelta' });
    });
    return { ok: true };
  });
}
function pedirCorreccion(a, id, motivo) {
  const u = chk_(a);
  motivo = String(motivo || '').trim();
  if (!motivo) throw new Error('Escribe qué hay que corregir.');
  const m = rows_('Movimientos').find(function (x) { return String(x.id) === String(id); });
  if (!m) throw new Error('Registro no encontrado.');
  add_('Correcciones', { id: id_(), fecha: new Date(), usuario: u.nombre, mov_id: String(id),
    detalle: m.concepto + ' · S/ ' + num_(m.monto) + ' · ' + fmt_(m.fecha, 'dd/MM HH:mm'), motivo: motivo, estado: 'pendiente' });
  return { ok: true };
}
function correcciones(a) {
  adm_(a);
  return rows_('Correcciones').filter(function (c) { return c.estado === 'pendiente'; }).map(function (c) {
    return { id: String(c.id), fecha: fmt_(c.fecha, 'dd/MM HH:mm'), usuario: c.usuario, mov_id: String(c.mov_id), detalle: c.detalle, motivo: c.motivo };
  }).reverse();
}
function resolverCorreccion(a, id, anularMov) {
  adm_(a);
  const c = rows_('Correcciones').find(function (x) { return String(x.id) === String(id); });
  if (!c) throw new Error('Corrección no encontrada.');
  if (anularMov) anular(a, c.mov_id);
  set_('Correcciones', c._row, { estado: 'resuelta' });
  return { ok: true };
}

// ================= Importar caja de Medicloud =================
const MC_FORMAS = { 'Efectivo': 'Efectivo', 'Transferencia': 'Transferencia', 'Tarjeta': 'Tarjeta', 'Depósito': 'Depósito',
  'Cheque': 'Depósito', 'Seguro/Med': 'Transferencia' };
function equivalencias_() {
  const out = {};
  try {
    rows_('Equivalencias').forEach(function (e) {
      const k = String(e.codigo).trim().toUpperCase();
      (out[k] = out[k] || []).push({ id: String(e.servicio_id).trim(), desde: num_(e.precio_desde) });
    });
  } catch (err) {}
  Object.keys(out).forEach(function (k) { out[k].sort(function (a, b) { return b.desde - a.desde; }); });
  return out;
}
function equiv_(eq, codigo, valor) {
  const l = eq[String(codigo || '').trim().toUpperCase()];
  if (!l) return '';
  const e = l.find(function (x) { return valor >= x.desde; }) || l[l.length - 1];
  return e.id;
}
function mcFecha_(f) {
  const m = String(f || '').match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}):(\d{2})/);
  if (!m) return null;
  return { date: new Date(m[1] + 'T' + m[2] + ':' + m[3] + ':00-05:00'), mes: m[1].slice(0, 7), txt: m[1].slice(8, 10) + '/' + m[1].slice(5, 7) + ' ' + m[2] + ':' + m[3] };
}
// tx: [{id, fecha, paciente, estado, forma, total, pagos:{Efectivo:..}, lineas:[{codigo, servicio, valor}]}]
function tokens_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z\s]/g, ' ')
    .split(/\s+/).filter(function (w) { return w.length > 2; });
}
function mismaPersona_(a, b) {
  const ta = tokens_(a), tb = tokens_(b);
  if (!ta.length || !tb.length) return false;
  const comunes = ta.filter(function (w) { return tb.indexOf(w) >= 0; }).length;
  return comunes >= Math.min(2, ta.length, tb.length);
}
function analizar_(tx) {
  const movs = rows_('Movimientos');
  const ya = {};
  movs.forEach(function (m) { if (m.ref && !yes_(m.anulado)) ya[String(m.ref)] = true; });
  const eq = equivalencias_();
  const cat = rows_('Catalogo');
  const catById = {};
  cat.forEach(function (c) { catById[String(c.id)] = c; });
  const procs = procs_(movs);
  return (tx || []).map(function (t) {
    const o = { id: String(t.id), paciente: String(t.paciente || '').trim(), total: num_(t.total) };
    const f = mcFecha_(t.fecha);
    o.fechaTxt = f ? f.txt : '';
    const pagos = [];
    Object.keys(t.pagos || {}).forEach(function (k) {
      const v = num_(t.pagos[k]);
      if (v > 0) pagos.push({ forma: MC_FORMAS[k] || 'Transferencia', monto: v });
    });
    o.pagos = pagos;
    o.lineas = (t.lineas || []).map(function (l) {
      const v = num_(l.valor);
      const sid = equiv_(eq, l.codigo, v);
      const c = sid ? catById[sid] : null;
      return { codigo: String(l.codigo || '').trim().toUpperCase(), servicio: String(l.servicio || '').trim(), valor: v,
        cat_id: c ? String(c.id) : '', cat_nombre: c ? c.servicio : '', proc: c ? c.tipo === 'procedimiento' : false };
    });
    if (ya[o.id]) o.estado = 'duplicado';
    else if (!f) o.estado = 'omitir';
    else if (String(t.estado || '').trim().toUpperCase() !== 'OK' || /pendiente/i.test(String(t.forma || '')) || !pagos.length) o.estado = 'omitir';
    else if (o.lineas.some(function (l) { return l.proc; })) {
      o.estado = 'procedimiento';
      const lp = o.lineas.find(function (l) { return l.proc; });
      o.procCat = lp.cat_id; o.procNombre = lp.cat_nombre;
      const cand = procs.find(function (p) { return p.saldo > 0 && String(p.tipo).indexOf(lp.cat_nombre) === 0 && mismaPersona_(p.nombre, o.paciente); });
      o.sugerido = cand ? cand.id : '';
    }
    else if (!o.lineas.length || o.lineas.some(function (l) { return !l.cat_id; })) o.estado = 'emparejar';
    else o.estado = 'listo';
    return o;
  });
}
function crearDesdeMedicloud_(u, o, f, cat) {
  if (!cat) return null;
  const partes = String(o.paciente || '').split(',');
  const apellidos = titulo_(String(partes[0] || '').trim()), nombres = titulo_(String(partes.slice(1).join(' ') || '').trim());
  const id = id_();
  const iso = Utilities.formatDate(f.date, TZ, 'yyyy-MM-dd');
  add_('Procedimientos', { id: id, fecha: f.date, tipo: cat.servicio, tipo_doc: '', num_doc: '', nombres: nombres || apellidos, apellidos: nombres ? apellidos : '',
    celular: '', precio_total: o.total, presupuesto_gastos: 0, forma_cobro: 'Pago único', nota: 'Creado desde Medicloud #' + o.id,
    usuario: u.nombre, anulado: 'NO', fecha_proc: iso });
  const nombre = ((nombres ? nombres + ' ' : '') + apellidos).trim();
  (plantillas_()[cat.servicio] || []).forEach(function (g) {
    add_('Movimientos', { id: id_(), fecha: new Date(), mes: mesActual_(), usuario: u.nombre, tipo: 'previsto', bolsa: 'procedimiento',
      concepto: g.concepto, monto: g.monto || 0, paciente: nombre, proc_id: id, estado: 'pendiente', anulado: 'NO' });
  });
  return { id: id, tipo: cat.servicio, nombre: nombre };
}
function completarPaciente(a, procId, d) {
  chk_(a);
  const pac = valPac_(d);
  return lock_(function () {
    const p = rows_('Procedimientos').find(function (x) { return String(x.id) === String(procId); });
    if (!p) throw new Error('Procedimiento no encontrado.');
    set_('Procedimientos', p._row, pac);
    upsertPac_(pac);
    return { ok: true };
  });
}
function previaImport(a, tx) {
  chk_(a);
  if (!tx || !tx.length) throw new Error('El archivo no tiene cobros.');
  return analizar_(tx);
}
// decisiones: { equiv: {CODIGO: cat_id|'omitir'}, procs: {txId: proc_id|'omitir'} }
function importarCaja(a, tx, decisiones) {
  const u = chk_(a);
  decisiones = decisiones || {};
  const nuevas = decisiones.equiv || {}, procsDec = decisiones.procs || {};
  return lock_(function () {
    // guardar equivalencias nuevas
    const eqActual = equivalencias_();
    Object.keys(nuevas).forEach(function (k) {
      const v = String(nuevas[k] || '');
      if (v && v !== 'omitir' && !eqActual[k.toUpperCase()]) add_('Equivalencias', { codigo: k.toUpperCase(), servicio_id: v, precio_desde: 0, descripcion: 'Agregado al importar' });
    });
    const an = analizar_(tx);
    const cat = rows_('Catalogo');
    const catById = {};
    cat.forEach(function (c) { catById[String(c.id)] = c; });
    const procs = procs_();
    const txById = {};
    (tx || []).forEach(function (t) { txById[String(t.id)] = t; });
    let n = 0, total = 0, omit = 0, dup = 0;
    an.forEach(function (o) {
      if (o.estado === 'duplicado') { dup++; return; }
      if (o.estado === 'omitir') { omit++; return; }
      const f = mcFecha_(txById[o.id].fecha);
      const formaPrincipal = o.pagos.slice().sort(function (x, y) { return y.monto - x.monto; })[0].forma;
      const base = { fecha: f.date, mes: f.mes, usuario: u.nombre + ' (Medicloud)', anulado: 'NO', ref: o.id };
      if (o.estado === 'procedimiento') {
        const pid = procsDec[o.id];
        let p = null;
        if (pid === 'crear') p = crearDesdeMedicloud_(u, o, f, catById[o.procCat]);
        else if (pid && pid !== 'omitir') p = procs.find(function (x) { return x.id === String(pid); });
        if (!p) { omit++; return; }
        o.pagos.forEach(function (pg) {
          add_('Movimientos', Object.assign({}, base, { id: id_(), tipo: 'ingreso', bolsa: 'procedimiento', concepto: p.tipo + ' – abono',
            monto: pg.monto, forma_pago: pg.forma, paciente: p.nombre, proc_id: p.id, nota: 'Medicloud #' + o.id }));
          total += pg.monto;
        });
        n++;
        return;
      }
      // servicios del consultorio
      const lineas = o.lineas.map(function (l) {
        const cid = l.cat_id || (nuevas[l.codigo] && nuevas[l.codigo] !== 'omitir' ? nuevas[l.codigo] : '');
        return Object.assign({}, l, { cat_id: cid });
      }).filter(function (l) { return l.cat_id; });
      if (!lineas.length) { omit++; return; }
      const partes = lineas.length === 1 ? o.pagos.map(function (pg) { return { l: lineas[0], monto: pg.monto, forma: pg.forma }; })
        : lineas.map(function (l) { return { l: l, monto: l.valor, forma: formaPrincipal }; });
      lineas.forEach(function (l) {
        const c = catById[l.cat_id];
        if (c && num_(c.pap) > 0) {
          add_('Movimientos', Object.assign({}, base, { id: id_(), tipo: 'patologia', bolsa: 'consultorio', concepto: 'PAP – ' + c.servicio,
            monto: num_(c.pap), paciente: o.paciente, estado: 'pendiente', nota: 'Medicloud #' + o.id }));
        }
      });
      partes.forEach(function (pt) {
        const c = catById[pt.l.cat_id];
        add_('Movimientos', Object.assign({}, base, { id: id_(), tipo: 'ingreso', bolsa: 'consultorio', concepto: c ? c.servicio : pt.l.servicio,
          monto: pt.monto, forma_pago: pt.forma, paciente: o.paciente, nota: 'Medicloud #' + o.id }));
        total += pt.monto;
      });
      n++;
    });
    return { ok: true, importados: n, total: num_(total), omitidos: omit, duplicados: dup };
  });
}

// ================= Presupuesto y resumen =================
function presMes_(mes, mv) {
  const actual = mesActual_();
  const dHoy = Number(Utilities.formatDate(new Date(), TZ, 'd'));
  return rows_('Presupuesto').filter(function (p) {
    const de = mm_(p.desde), ha = mm_(p.hasta);
    return yes_(p.activo) && (!de || de <= mes) && (!ha || mes <= ha);
  }).map(function (p) {
    const id = String(p.id), monto = num_(p.monto), dia = num_(p.dia);
    const pagado = suma_(mv.filter(function (m) { return m.tipo === 'gasto' && String(m.pres_id) === id; }));
    let estado = 'pendiente';
    if (pagado > 0 && (pagado >= monto || yes_(p.estimado))) estado = 'pagado';
    else if (pagado > 0) estado = 'parcial';
    else if (mes < actual || (mes === actual && dia < dHoy)) estado = 'vencido';
    else if (mes === actual && dia - dHoy <= 3) estado = 'proximo';
    const de = mm_(p.desde), ha = mm_(p.hasta);
    const cuota = (de && ha) ? (difMeses_(de, mes) + 1) + '/' + (difMeses_(de, ha) + 1) : '';
    return { id: id, dia: dia, concepto: p.concepto, categoria: p.categoria, monto: monto,
      estimado: yes_(p.estimado), pagado: pagado, estado: estado, cuota: cuota, hasta: ha };
  }).sort(function (x, y) { return x.dia - y.dia; });
}

function resumen_(mes) {
  const all = rows_('Movimientos');
  const mv = all.filter(function (m) { return mm_(m.mes) === mes && !yes_(m.anulado); });
  function S(f) { return suma_(mv.filter(f)); }
  const cIng = S(function (m) { return m.tipo === 'ingreso' && m.bolsa === 'consultorio'; });
  const cPat = S(function (m) { return m.tipo === 'patologia' && m.bolsa === 'consultorio'; });
  const cGas = S(function (m) { return m.tipo === 'gasto' && m.bolsa === 'consultorio'; });
  const tr = S(function (m) { return m.tipo === 'transferencia'; });
  const pIng = S(function (m) { return m.tipo === 'ingreso' && m.bolsa === 'procedimiento'; });
  const pGas = S(function (m) { return m.tipo === 'gasto' && m.bolsa === 'procedimiento'; });
  const pPat = S(function (m) { return m.tipo === 'patologia' && m.bolsa === 'procedimiento'; });
  const pres = presMes_(mes, mv);
  const meta = suma_(pres);
  const presPagado = num_(pres.reduce(function (s, p) { return s + p.pagado; }, 0));
  const cNeto = num_(cIng - cPat + tr);
  const procs = procs_(all);
  const porCobrar = procs.filter(function (p) { return p.saldo > 0; });
  const formas = {};
  mv.filter(function (m) { return m.tipo === 'ingreso'; }).forEach(function (m) {
    formas[m.forma_pago] = num_((formas[m.forma_pago] || 0) + num_(m.monto));
  });
  const patPend = suma_(all.filter(function (m) { return m.tipo === 'patologia' && !yes_(m.anulado) && m.estado !== 'pagado'; }));
  const prevPend = all.filter(function (m) { return m.tipo === 'previsto' && !yes_(m.anulado) && m.estado === 'pendiente'; });
  // Última importación de Medicloud (según la fecha del último cobro importado)
  let ultImp = null;
  const imp = all.filter(function (m) { return m.ref && !yes_(m.anulado) && m.tipo === 'ingreso' && m.fecha instanceof Date; });
  if (imp.length) {
    const ult = imp.reduce(function (x, m) { return m.fecha > x ? m.fecha : x; }, imp[0].fecha);
    const dia = Utilities.formatDate(ult, TZ, 'yyyy-MM-dd');
    const refs = {};
    imp.forEach(function (m) { if (Utilities.formatDate(m.fecha, TZ, 'yyyy-MM-dd') === dia) refs[String(m.ref)] = true; });
    ultImp = { dia: dia, txt: Utilities.formatDate(ult, TZ, 'dd/MM'), hora: Utilities.formatDate(ult, TZ, 'HH:mm'),
      n: Object.keys(refs).length, hoy: dia === hoyIso_() };
  }
  let dias = 0;
  if (mes === mesActual_()) {
    const y = Number(mes.slice(0, 4)), mo = Number(mes.slice(5, 7));
    dias = new Date(y, mo, 0).getDate() - Number(Utilities.formatDate(new Date(), TZ, 'd')) + 1;
  }
  return {
    mes: mes, dias: dias, all: all, mv: mv, procs: procs,
    consultorio: { ingresos: cIng, patologia: cPat, transferido: tr, neto: cNeto, gastos: cGas, meta: meta,
      presPagado: presPagado, porPagar: Math.max(0, num_(meta - presPagado)), faltante: Math.max(0, num_(meta - cNeto)) },
    procedimientos: { ingresos: pIng, gastos: pGas, patologia: pPat, transferido: tr, ganancia: num_(pIng - pGas - pPat - tr) },
    total: { ingresos: num_(cIng + pIng), egresos: num_(cGas + cPat + pGas + pPat), disponible: num_(cIng + pIng - cGas - cPat - pGas - pPat) },
    pres: pres, porCobrar: porCobrar, totalPorCobrar: num_(porCobrar.reduce(function (s, p) { return s + p.saldo; }, 0)),
    formas: formas, patPend: patPend, porFacturar: { n: prevPend.length, total: suma_(prevPend) }, ultImp: ultImp
  };
}

function resumen(a, mes) {
  chk_(a);
  const r = resumen_(mes || mesActual_());
  delete r.all; delete r.mv; delete r.procs;
  return r;
}

// ================= Patología =================
function patologia(a, mes) {
  chk_(a);
  mes = mes || mesActual_();
  const pat = rows_('Movimientos').filter(function (m) { return m.tipo === 'patologia' && !yes_(m.anulado); });
  const delMes = pat.filter(function (m) { return mm_(m.mes) === mes; });
  const pend = pat.filter(function (m) { return m.estado !== 'pagado'; });
  return {
    lista: delMes.map(movOut_).reverse(),
    consultorio: suma_(delMes.filter(function (m) { return m.bolsa === 'consultorio'; })),
    procedimientos: suma_(delMes.filter(function (m) { return m.bolsa === 'procedimiento'; })),
    total: suma_(delMes), pendientes: pend.map(movOut_), totalPend: suma_(pend)
  };
}
function pagarPatologia(a, ids, factura) {
  adm_(a);
  if (!ids || !ids.length) throw new Error('Marca al menos una línea.');
  const f = String(factura || '').trim();
  const hoy = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
  return lock_(function () {
    rows_('Movimientos').forEach(function (m) {
      if (m.tipo === 'patologia' && ids.indexOf(String(m.id)) >= 0) {
        set_('Movimientos', m._row, { estado: 'pagado',
          nota: (m.nota ? m.nota + ' · ' : '') + 'Pagado ' + hoy + (f ? ' · Factura ' + f : '') });
      }
    });
    return { ok: true };
  });
}

// ================= Informes =================
function nombreMes_(mes) {
  const n = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return n[Number(mes.slice(5, 7)) - 1] + ' ' + mes.slice(0, 4);
}
function datosInforme_(mes) {
  const r = resumen_(mes);
  const serv = {};
  r.mv.filter(function (m) { return m.tipo === 'ingreso' && m.bolsa === 'consultorio'; }).forEach(function (m) {
    serv[m.concepto] = serv[m.concepto] || { n: 0, t: 0 };
    serv[m.concepto].n++; serv[m.concepto].t = num_(serv[m.concepto].t + num_(m.monto));
  });
  const otrosGastos = r.mv.filter(function (m) { return m.tipo === 'gasto' && m.bolsa === 'consultorio' && !m.pres_id; });
  const procIds = {};
  r.mv.forEach(function (m) { if (m.proc_id) procIds[String(m.proc_id)] = true; });
  const procs = r.procs.filter(function (p) { return procIds[p.id] || p.mesCreado === mes; });
  const pat = r.mv.filter(function (m) { return m.tipo === 'patologia'; });
  return { r: r, serv: serv, otrosGastos: otrosGastos, procs: procs, pat: pat, nombre: nombreMes_(mes) };
}
function e_(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function s_(n) { return 'S/ ' + num_(n).toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
const TIPO_TXT = { ingreso: 'Ingreso', gasto: 'Gasto', patologia: 'Patología', transferencia: 'Paso a consultorio', previsto: 'Por facturar' };

function informePDF(a, mes) {
  chk_(a);
  mes = mes || mesActual_();
  const D = datosInforme_(mes), r = D.r, c = r.consultorio, p = r.procedimientos;
  function tabla(head, rowsArr) {
    return '<table><tr>' + head.map(function (h) { return '<th>' + e_(h) + '</th>'; }).join('') + '</tr>' +
      (rowsArr.length ? rowsArr.map(function (row) { return '<tr>' + row.map(function (v) { return '<td>' + e_(v) + '</td>'; }).join('') + '</tr>'; }).join('')
        : '<tr><td colspan="' + head.length + '">Sin registros</td></tr>') + '</table>';
  }
  let h = '<html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:10px;color:#222}' +
    'h1{font-size:16px;margin:0}h2{font-size:12px;margin:14px 0 4px;border-bottom:1px solid #999}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:6px}th,td{border:1px solid #ccc;padding:3px 4px;text-align:left}' +
    'th{background:#eef}.k td:first-child{width:60%}</style></head><body>';
  h += '<h1>Informe financiero – ' + e_(D.nombre) + '</h1><div>Consultorio Dr. Daniel Villamizar Hernández · generado ' +
    Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + '</div>';
  h += '<h2>Resumen</h2><table class="k">' +
    [['Ingresos consultorio', s_(c.ingresos)], ['Patología (consultorio)', '− ' + s_(c.patologia)],
     ['Pasado desde procedimientos', s_(c.transferido)], ['Neto consultorio', s_(c.neto)],
     ['Meta del presupuesto general', s_(c.meta)],
     ['Cumplimiento', c.meta ? Math.round(c.neto / c.meta * 100) + '%' : '-'],
     ['Faltante para la meta', s_(c.faltante)], ['Pagado del presupuesto', s_(c.presPagado)],
     ['Gastos consultorio (total)', s_(c.gastos)],
     ['Ingresos procedimientos', s_(p.ingresos)], ['Gastos procedimientos', s_(p.gastos)],
     ['Patología (procedimientos)', s_(p.patologia)], ['Ganancia procedimientos', s_(p.ganancia)],
     ['TOTAL INGRESOS', s_(r.total.ingresos)], ['TOTAL EGRESOS', s_(r.total.egresos)],
     ['DISPONIBLE DEL MES', s_(r.total.disponible)], ['Saldos por cobrar (procedimientos)', s_(r.totalPorCobrar)],
     ['Patología por pagar (todas las facturas)', s_(r.patPend)]]
      .map(function (x) { return '<tr><td>' + x[0] + '</td><td>' + x[1] + '</td></tr>'; }).join('') + '</table>';
  h += '<h2>Ingresos por forma de pago</h2>' + tabla(['Forma', 'Monto'], Object.keys(r.formas).map(function (k) { return [k, s_(r.formas[k])]; }));
  h += '<h2>Ingresos del consultorio por servicio</h2>' + tabla(['Servicio', 'Cantidad', 'Total'],
    Object.keys(D.serv).sort(function (x, y) { return D.serv[y].t - D.serv[x].t; }).map(function (k) { return [k, D.serv[k].n, s_(D.serv[k].t)]; }));
  h += '<h2>Presupuesto general</h2>' + tabla(['Día', 'Concepto', 'Cuota', 'Presupuestado', 'Pagado', 'Estado'],
    r.pres.map(function (x) { return [x.dia, x.concepto, x.cuota, s_(x.monto), s_(x.pagado), x.estado]; }));
  h += '<h2>Otros gastos del consultorio</h2>' + tabla(['Fecha', 'Concepto', 'Forma', 'Monto'],
    D.otrosGastos.map(function (m) { return [fmt_(m.fecha, 'dd/MM'), m.concepto, m.forma_pago, s_(m.monto)]; }));
  h += '<h2>Procedimientos</h2>' + tabla(['Fecha cirugía', 'Tipo', 'Paciente', 'Documento', 'Precio', 'Cobrado', 'Saldo', 'Presup. gastos', 'Gastos', 'Por facturar', 'Ganancia', 'Ganancia est.'],
    D.procs.map(function (x) { return [x.fecha_proc || x.fecha, x.tipo, x.nombre, x.tipo_doc + ' ' + x.num_doc, s_(x.precio), s_(x.cobrado), s_(x.saldo), s_(x.presupuesto), s_(x.gastos), s_(x.porFacturar), s_(x.ganancia), s_(x.gananciaEst)]; }));
  h += '<h2>Patología</h2>' + tabla(['Fecha', 'Concepto', 'Paciente', 'Bolsa', 'Estado', 'Monto'],
    D.pat.map(function (m) { return [fmt_(m.fecha, 'dd/MM'), m.concepto, m.paciente, m.bolsa, m.estado, s_(m.monto)]; }));
  const prev = r.all.filter(function (m) { return m.tipo === 'previsto' && !yes_(m.anulado) && m.estado === 'pendiente'; });
  h += '<h2>Facturas por llegar (estimado, aún no pagado)</h2>' + tabla(['Anotada', 'Concepto', 'Paciente', 'Estimado'],
    prev.map(function (m) { return [fmt_(m.fecha, 'dd/MM'), m.concepto, m.paciente, s_(m.monto)]; }));
  h += '<h2>Detalle de movimientos</h2>' + tabla(['Fecha', 'Tipo', 'Bolsa', 'Concepto', 'Paciente', 'Forma', 'Comprobante', 'Registró', 'Monto'],
    r.mv.filter(function (m) { return m.tipo !== 'previsto'; }).map(function (m) { return [fmt_(m.fecha, 'dd/MM HH:mm'), TIPO_TXT[m.tipo] || m.tipo, m.bolsa, m.concepto, m.paciente, m.forma_pago, m.comprobante, m.usuario, s_(m.monto)]; }));
  h += '</body></html>';
  const blob = HtmlService.createHtmlOutput(h).getAs(MimeType.PDF);
  return { name: 'Informe_Consultorio_' + mes + '.pdf', mime: 'application/pdf', b64: Utilities.base64Encode(blob.getBytes()) };
}

function informeExcel(a, mes) {
  chk_(a);
  mes = mes || mesActual_();
  const D = datosInforme_(mes), r = D.r, c = r.consultorio, p = r.procedimientos;
  const tmp = SpreadsheetApp.create('Informe ' + mes);
  function put(nombre, data, primera) {
    const s = primera ? tmp.getSheets()[0] : tmp.insertSheet(nombre);
    if (primera) s.setName(nombre);
    const w = data[0].length;
    s.getRange(1, 1, data.length, w).setValues(data.map(function (row) { while (row.length < w) row.push(''); return row; }));
    s.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#E6F1FB');
    s.setFrozenRows(1);
    s.autoResizeColumns(1, w);
  }
  put('Resumen', [['Concepto', 'Monto'],
    ['Ingresos consultorio', c.ingresos], ['Patología (consultorio)', -c.patologia], ['Pasado desde procedimientos', c.transferido],
    ['Neto consultorio', c.neto], ['Meta presupuesto general', c.meta], ['Faltante para la meta', c.faltante],
    ['Pagado del presupuesto', c.presPagado], ['Gastos consultorio', c.gastos],
    ['Ingresos procedimientos', p.ingresos], ['Gastos procedimientos', p.gastos], ['Patología (procedimientos)', p.patologia],
    ['Ganancia procedimientos', p.ganancia], ['TOTAL INGRESOS', r.total.ingresos], ['TOTAL EGRESOS', r.total.egresos],
    ['DISPONIBLE DEL MES', r.total.disponible], ['Saldos por cobrar', r.totalPorCobrar], ['Patología por pagar', r.patPend]], true);
  put('Movimientos', [['Fecha', 'Tipo', 'Bolsa', 'Concepto', 'Paciente', 'Forma de pago', 'Comprobante', 'Registró', 'Monto', 'Nota']]
    .concat(r.mv.filter(function (m) { return m.tipo !== 'previsto'; }).map(function (m) { return [fmt_(m.fecha), TIPO_TXT[m.tipo] || m.tipo, m.bolsa, m.concepto, m.paciente, m.forma_pago, m.comprobante, m.usuario, num_(m.monto), m.nota]; })));
  put('Servicios', [['Servicio', 'Cantidad', 'Total']].concat(Object.keys(D.serv).map(function (k) { return [k, D.serv[k].n, D.serv[k].t]; })));
  put('Presupuesto', [['Día', 'Concepto', 'Categoría', 'Cuota', 'Presupuestado', 'Pagado', 'Estado']]
    .concat(r.pres.map(function (x) { return [x.dia, x.concepto, x.categoria, x.cuota, x.monto, x.pagado, x.estado]; })));
  put('Procedimientos', [['Fecha cirugía', 'Tipo', 'Paciente', 'Documento', 'Celular', 'Precio', 'Cobrado', 'Saldo', 'Presup. gastos', 'Gastos', 'Por facturar', 'Ganancia', 'Ganancia estimada']]
    .concat(D.procs.map(function (x) { return [x.fecha_proc || x.fecha, x.tipo, x.nombre, x.tipo_doc + ' ' + x.num_doc, x.celular, x.precio, x.cobrado, x.saldo, x.presupuesto, x.gastos, x.porFacturar, x.ganancia, x.gananciaEst]; })));
  put('Patología', [['Fecha', 'Concepto', 'Paciente', 'Bolsa', 'Estado', 'Monto']]
    .concat(D.pat.map(function (m) { return [fmt_(m.fecha, 'dd/MM/yyyy'), m.concepto, m.paciente, m.bolsa, m.estado, num_(m.monto)]; })));
  SpreadsheetApp.flush();
  const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return { name: 'Informe_Consultorio_' + mes + '.xlsx', b64: Utilities.base64Encode(blob.getBytes()),
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
}

// ================= Reiniciar datos (ejecutar desde el editor, solo una vez) =================
function reiniciarDatosDePrueba() {
  const ss = ss_();
  const copia = DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' – respaldo pruebas ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  ['Movimientos', 'Procedimientos', 'Pacientes', 'Correcciones'].forEach(function (n) {
    const s = sh_(n);
    if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, Math.max(1, s.getLastColumn())).clearContent();
  });
  Logger.log('Datos de prueba borrados. Respaldo: ' + copia.getName());
}

// ================= Instalación (ejecutar una sola vez) =================
function instalar() {
  limpiarCache();
  const ss = ss_();
  Object.keys(H).forEach(function (n) {
    let s = ss.getSheetByName(n);
    if (!s) s = ss.insertSheet(n);
    if (s.getLastRow() === 0) {
      s.appendRow(H[n]);
      s.setFrozenRows(1);
      s.getRange(1, 1, 1, H[n].length).setFontWeight('bold').setBackground('#E6F1FB');
    } else {
      // Actualización: agrega al final las columnas nuevas que falten
      const head = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0].map(String);
      H[n].forEach(function (k) {
        if (head.indexOf(k) < 0) {
          head.push(k);
          s.getRange(1, head.length).setValue(k).setFontWeight('bold').setBackground('#E6F1FB');
        }
      });
    }
    (TEXTO[n] || []).forEach(function (k) {
      const c = H[n].indexOf(k) + 1;
      s.getRange(1, c, s.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (nm) {
    const d = ss.getSheetByName(nm);
    if (d && ss.getSheets().length > 1) ss.deleteSheet(d);
  });

  if (rows_('Usuarios').length === 0) {
    const pin = function () { return String(Math.floor(100000 + Math.random() * 900000)); };
    [['u1', 'Dr. Villamizar', 'admin'], ['u2', 'Kiara', 'asistente'], ['u3', 'Lucía', 'asistente'], ['u4', 'Asistente 3', 'asistente']]
      .forEach(function (x, i) { add_('Usuarios', { id: x[0], nombre: x[1], rol: x[2], pin: pin(), activo: i < 3 ? 'SI' : 'NO' }); });
  }

  if (rows_('Catalogo').length === 0) {
    const C = [
      ['Endometriosis', 'Consulta + Ecografía de endometriosis', 400, 0, 'NO'],
      ['Endometriosis', 'Ecografía de endometriosis (sin consulta)', 350, 0, 'NO'],
      ['Fertilidad', 'Consulta fertilidad + Ecografía transvaginal 5D', 300, 0, 'NO'],
      ['Fertilidad', 'Consulta fertilidad + Chequeo ginecológico', 500, 60, 'SI'],
      ['Fertilidad', 'Seguimiento folicular ecográfico', 0, 0, 'NO'],
      ['Gestantes', 'Consulta de embarazo + Ecografía', 250, 0, 'NO'],
      ['Gestantes', 'Inicio de control prenatal', 250, 0, 'NO'],
      ['Gestantes', 'Control prenatal mensual', 200, 0, 'NO'],
      ['Gestantes', 'Ecografía transvaginal obstétrica 5D', 150, 0, 'NO'],
      ['Gestantes', 'Ecografía obstétrica de control', 150, 0, 'NO'],
      ['Gestantes', 'Ecografía revelación de sexo 5D', 150, 0, 'NO'],
      ['Gestantes', 'Ecografía 5D caritas fetales', 250, 0, 'NO'],
      ['Ginecología', 'Consulta ginecológica', 150, 0, 'SI'],
      ['Ginecología', 'Consulta ginecológica + Ecografía transvaginal', 200, 0, 'SI'],
      ['Ginecología', 'Paquete ginecológico + Ecografía mamaria', 500, 60, 'SI'],
      ['Ginecología', 'Paquete ginecológico', 350, 60, 'SI'],
      ['Ginecología', 'Consulta + PAP + IVAA + Test de Schiller', 250, 60, 'SI'],
      ['Ginecología', 'Consulta + PAP + IVAA', 220, 60, 'SI'],
      ['Ginecología', 'Consulta + Eco transvaginal + Eco mamaria', 300, 0, 'NO'],
      ['Ginecología', 'Consulta + Ecografía mamaria', 250, 0, 'NO'],
      ['Ginecología', 'Consulta métodos anticonceptivos', 150, 0, 'NO'],
      ['Ginecología', 'Ecografía transvaginal ginecológica (sin consulta)', 100, 0, 'NO'],
      ['Ginecología', 'Ecografía mamaria (sin consulta)', 150, 0, 'NO'],
      ['Eco especializada', 'Ecografía genética 5D + consulta', 400, 0, 'NO'],
      ['Eco especializada', 'Ecografía genética 5D', 350, 0, 'NO'],
      ['Eco especializada', 'Ecografía morfológica 5D + consulta', 400, 0, 'NO'],
      ['Eco especializada', 'Ecografía morfológica 5D', 350, 0, 'NO'],
      ['Eco especializada', 'Ecografía de crecimiento fetal', 280, 0, 'NO']
    ];
    C.forEach(function (x, i) {
      add_('Catalogo', { id: 's' + (i + 1), categoria: x[0], servicio: x[1], precio: x[2], tipo: 'servicio', pap: x[3], biopsia_cu: x[4], activo: 'SI' });
    });
    [['HIFU', 'NO'], ['FRAXX', 'NO'], ['Parto vaginal', 'NO'], ['Cesárea', 'NO'], ['Cirugía laparoscópica', 'SI'], ['AMEU', 'NO'],
     ['Ginecología estética', 'SI'], ['Cirugía convencional', 'SI']].forEach(function (x, i) {
      add_('Catalogo', { id: 'p' + (i + 1), categoria: 'Procedimientos', servicio: x[0], precio: 0, tipo: 'procedimiento', pap: 0,
        biopsia_cu: 'NO', activo: 'SI', pide_detalle: x[1] });
    });
  }

  if (rows_('Presupuesto').length === 0) {
    const P = [
      [1, 'Internet', 180, 'Servicios', '', 'NO'],
      [2, 'Terreno', 1750, 'Compromisos', '', 'NO'],
      [5, 'Alquiler consultorio', 1600, 'Fijos', '', 'NO'],
      [7, 'Deuda Llermith', 2600, 'Deudas', '2027-12', 'NO'],
      [10, 'Deuda Sr. Zamora', 7600, 'Deudas', '2027-06', 'NO'],
      [20, 'Sueldo Villa', 5000, 'Sueldos', '', 'NO'],
      [27, 'Contadora', 120, 'Fijos', '', 'NO'],
      [27, 'SUNAT mensual', 70, 'Impuestos', '', 'SI'],
      [27, 'TDC Villa – deudas', 10000, 'Deudas', '', 'NO'],
      [27, 'Gastos de consultorio (materiales e insumos)', 500, 'Insumos', '', 'NO'],
      [29, 'Sueldo Kiara y Lucía', 4800, 'Sueldos', '', 'NO'],
      [29, 'SUNAT fraccionamiento', 691, 'Deudas', '2028-03', 'NO']
    ];
    P.forEach(function (x, i) {
      add_('Presupuesto', { id: 'b' + (i + 1), dia: x[0], concepto: x[1], monto: x[2], categoria: x[3],
        desde: x[4] ? '2026-11' : '', hasta: x[4], estimado: x[5], activo: 'SI' });
    });
  }
  // Actualización v4 del catálogo de procedimientos
  const cat = rows_('Catalogo');
  cat.forEach(function (c) {
    if (c.tipo !== 'procedimiento') return;
    if (c.servicio === 'Parto') set_('Catalogo', c._row, { servicio: 'Parto vaginal', pide_detalle: 'NO' });
    else if (c.servicio === 'Cirugía') set_('Catalogo', c._row, { servicio: 'Cirugía laparoscópica', pide_detalle: 'SI' });
    else if (c.servicio === 'Ginecología estética') set_('Catalogo', c._row, { pide_detalle: 'SI' });
    else if (!c.pide_detalle) set_('Catalogo', c._row, { pide_detalle: 'NO' });
  });
  if (!rows_('Catalogo').some(function (c) { return c.servicio === 'Cirugía convencional'; })) {
    let n = 1;
    while (cat.some(function (c) { return String(c.id) === 'p' + n; })) n++;
    add_('Catalogo', { id: 'p' + n, categoria: 'Procedimientos', servicio: 'Cirugía convencional', precio: 0, tipo: 'procedimiento',
      pap: 0, biopsia_cu: 'NO', activo: 'SI', pide_detalle: 'SI' });
  }
  [['s29', 'Gestantes', 'Control prenatal semanal', 100], ['s30', 'Ginecología', 'Control ginecológico', 50]].forEach(function (x) {
    if (!rows_('Catalogo').some(function (c) { return String(c.id) === x[0] || c.servicio === x[2]; })) {
      add_('Catalogo', { id: x[0], categoria: x[1], servicio: x[2], precio: x[3], tipo: 'servicio', pap: 0, biopsia_cu: 'NO', activo: 'SI', pide_detalle: 'NO' });
    }
  });
  if (rows_('Equivalencias').length === 0) {
    const fx = rows_('Catalogo').find(function (c) { return c.servicio === 'FRAXX'; });
    [['CONS', 's13'], ['CTV', 's14'], ['CPNN', 's8'], ['ECOTV', 's22'], ['EOBS', 's10'], ['ETVOBST', 's9'], ['CPS', 's29'], ['CG', 's30'],
     ['SF', 's5'], ['PG', 's16'], ['PGEM', 's15'], ['CPI', 's18'], ['CEE', 's19'], ['EM', 's23'], ['CF', 's3'], ['CEO', 's6'],
     ['ED', 's1'], ['EG', 's25', 0], ['EG', 's24', 400], ['MORFO', 's27', 0], ['MORFO', 's26', 400]].forEach(function (x) {
      add_('Equivalencias', { codigo: x[0], servicio_id: x[1], precio_desde: x[2] || 0, descripcion: 'Inicial' });
    });
    if (fx) add_('Equivalencias', { codigo: 'FRAX', servicio_id: String(fx.id), precio_desde: 0, descripcion: 'Procedimiento' });
  }
  const hf = rows_('Catalogo').find(function (c) { return c.servicio === 'HIFU'; });
  if (hf && !rows_('Equivalencias').some(function (e) { return String(e.codigo).toUpperCase() === 'HIFU'; })) {
    add_('Equivalencias', { codigo: 'HIFU', servicio_id: String(hf.id), precio_desde: 0, descripcion: 'Procedimiento' });
  }
  if (rows_('Plantillas').length === 0) {
    const CL = 'Clínica (sala de operaciones + hospitalización)';
    const P = {
      'Cesárea': [CL, 'Anestesiólogo', 'Pediatra', 'Cirujano principal', 'Cirujano ayudante', 'Instrumentista',
        'Asistente al cirujano principal', 'Técnico de apoyo', 'Alimentación', 'Curación', 'Medicamentos'],
      'Parto vaginal': [CL, 'Anestesiólogo', 'Pediatra', 'Cirujano principal', 'Obstetra especialista',
        'Asistente al cirujano principal', 'Técnico de apoyo', 'Alimentación', 'Curación', 'Medicamentos'],
      'Cirugía laparoscópica': [CL, 'Alquiler de torre de laparoscopia', 'Anestesiólogo', 'Cirujano principal', 'Cirujano ayudante',
        'Instrumentista', 'Asistente al cirujano principal', 'Técnico de apoyo', 'Alimentación', 'Curación', 'Medicamentos'],
      'Cirugía convencional': [CL, 'Anestesiólogo', 'Cirujano principal', 'Cirujano ayudante', 'Instrumentista',
        'Asistente al cirujano principal', 'Técnico de apoyo', 'Alimentación', 'Curación', 'Medicamentos'],
      'AMEU': ['Anestesiólogo', 'Cirujano principal', 'Asistente al cirujano principal', 'Técnico de sala', 'Clínica', 'Medicamentos'],
      'FRAXX': ['Cánula', 'Médico tratante', 'Anestesia local'],
      'HIFU': ['Alquiler de equipo', 'Médico tratante', 'Anestesia local']
    };
    Object.keys(P).forEach(function (proc) {
      P[proc].forEach(function (c, i) { add_('Plantillas', { procedimiento: proc, concepto: c, monto_sugerido: '', orden: i + 1, activo: 'SI' }); });
    });
  }
  if (rows_('Ajustes').length === 0) {
    add_('Ajustes', { ajuste: 'biopsia_cuello_uterino', valor: BIOPSIA_CU, nota: 'Descuento de Patología por biopsia de cuello uterino (S/)' });
  }
  limpiarCache();
  Logger.log('Listo. Revisa la hoja "Usuarios" para ver los PIN de cada persona.');
}
