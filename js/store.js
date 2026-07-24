/* =============================================
   STORE.JS - Capa de datos
   Proveedores / Equipos / Mantenimientos → conectados
   al backend real de Google Apps Script.
   Alertas / Correos de alerta → siguen en localStorage
   por ahora (ver nota al final del archivo).
   ============================================= */

const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbwS2SpNZD5usDRrUSNE_CMgRuagqRIgtlXWSjmn-g1czxCzpo9_0R0bI8pXtPuB6idhjg/exec';
const API_KEY = 'acerimallas-2026-x7k9m2'; // debe ser IDÉNTICA a la de Code.gs

// ---------------------------------------------------------------------
// Mapeo de filas crudas del Sheet -> objetos JS (según el orden de
// columnas definido en Code.gs)
// ---------------------------------------------------------------------
function mapProveedor(row) {
  return {
    id: row[0], fechaRegistro: row[1], razonSocial: row[2], nombreComercial: row[3],
    ruc: row[4], telefono1: row[5], telefono2: row[6], ciudad: row[7], provincia: row[8],
    correoRetenciones: row[9], direccion: row[10], representanteLegal: row[11],
    contacto1: row[12], contacto2: row[13], formaPago: row[14], entidadBancaria: row[15],
    tipoCuenta: row[16], numeroCuenta: row[17], titularCuenta: row[18],
    archivoRuc: row[19], archivoRepLegal: row[20], archivoNombramiento: row[21],
    certificadoBancario: row[22], area: row[23], estado: row[24],
    camposConError: row[25] ? String(row[25]).split(',').map(s => s.trim()).filter(Boolean) : [],
    linkToken: row[26], observacion: row[27] || ''
  };
}

function mapEquipo(row) {
  let datosEspecificos = {};
  try { datosEspecificos = row[7] ? JSON.parse(row[7]) : {}; } catch (e) { datosEspecificos = {}; }
  return {
    id: row[0], proveedorId: row[1], nombre: row[2], ubicacion: row[3], serie: row[4],
    proximoMantenimiento: row[5], tipo: row[6] || 'otro', datosEspecificos
  };
}

function mapMantenimiento(row) {
  return { id: row[0], equipoId: row[1], fecha: row[2], tipo: row[3], observacion: row[4], detalle: row[5], pdfUrl: row[6], origen: row[7] };
}

// ---------------------------------------------------------------------
// PROVEEDORES (backend real)
// ---------------------------------------------------------------------
async function getProveedores() {
  const res = await fetch(`${WEBAPP_URL}?accion=listarProveedores&apiKey=${encodeURIComponent(API_KEY)}`);
  const data = await res.json();
  if (!data.ok) { console.error(data.error); return []; }
  return data.proveedores.map(mapProveedor);
}

async function guardarProveedor(datos) {
  const res = await fetch(WEBAPP_URL, {
    method: 'POST',
    body: JSON.stringify({ accion: 'guardarProveedor', datos, apiKey: API_KEY })
  });
  return res.json();
}

async function actualizarEstadoProveedor(id, estado, camposConError, observacion) {
  const res = await fetch(WEBAPP_URL, {
    method: 'POST',
    body: JSON.stringify({ accion: 'actualizarEstadoProveedor', id, estado, camposConError, observacion, apiKey: API_KEY })
  });
  return res.json();
}

// ---------------------------------------------------------------------
// EQUIPOS (backend real)
// ---------------------------------------------------------------------
async function getEquipos(proveedorId) {
  const base = proveedorId ? `${WEBAPP_URL}?accion=listarEquipos&proveedorId=${proveedorId}` : `${WEBAPP_URL}?accion=listarEquipos`;
  const url = `${base}&apiKey=${encodeURIComponent(API_KEY)}`;
  const res = await fetch(url);
  const data = await res.json();
  if (!data.ok) { console.error(data.error); return []; }
  return data.equipos.map(mapEquipo);
}

async function agregarEquipo(datos) {
  const res = await fetch(WEBAPP_URL, {
    method: 'POST',
    body: JSON.stringify({ accion: 'agregarEquipo', datos, apiKey: API_KEY })
  });
  return res.json();
}

// ---------------------------------------------------------------------
// MANTENIMIENTOS (backend real)
// ---------------------------------------------------------------------
async function getMantenimientos(equipoId) {
  const base = equipoId ? `${WEBAPP_URL}?accion=listarMantenimientos&equipoId=${equipoId}` : `${WEBAPP_URL}?accion=listarMantenimientos`;
  const url = `${base}&apiKey=${encodeURIComponent(API_KEY)}`;
  const res = await fetch(url);
  const data = await res.json();
  if (!data.ok) { console.error(data.error); return []; }
  return data.mantenimientos.map(mapMantenimiento);
}

async function agregarMantenimiento(datos) {
  const res = await fetch(WEBAPP_URL, {
    method: 'POST',
    body: JSON.stringify({ accion: 'agregarMantenimiento', datos, apiKey: API_KEY })
  });
  return res.json();
}

// ---------------------------------------------------------------------
// Utilidad: convertir un <input type="file"> a base64 para mandarlo
// al backend (Code.gs espera { base64, nombre, tipo })
// ---------------------------------------------------------------------
function archivoABase64(file) {
  return new Promise((resolve, reject) => {
    if (!file) { resolve(null); return; }
    const reader = new FileReader();
    reader.onload = () => resolve({
      base64: reader.result.split(',')[1],
      nombre: file.name,
      tipo: file.type
    });
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ---------------------------------------------------------------------
// ALERTAS Y CORREOS DE ALERTA
// NOTA: Code.gs ya registra las alertas en la hoja "Log_Alertas" y ya
// envía los correos automáticos (nuevo proveedor, corrección, verificado).
// Lo que todavía sigue en localStorage es la LISTA de destinatarios que
// configuras en la pantalla de Configuración — Code.gs por ahora usa un
// correo fijo (CORREOS_EMPRESA_DEFAULT). Cuando quieras que la lista de
// "Configuración > Correos" controle a quién le llegan los correos reales,
// se agrega una hoja "Config" en el Sheet y una función en Code.gs para
// leerla — dilo y lo conectamos.
// ---------------------------------------------------------------------
const DB_KEYS = {
  alertas: 'acerimallas_alertas',
  correosAlerta: 'acerimallas_correos_alerta'
};

function getAlertas() { return JSON.parse(localStorage.getItem(DB_KEYS.alertas) || '[]'); }
function saveAlertas(list) { localStorage.setItem(DB_KEYS.alertas, JSON.stringify(list)); }
function agregarAlerta(mensaje, tipo) {
  const alertas = getAlertas();
  alertas.unshift({ id: Date.now(), mensaje, tipo, fecha: new Date().toISOString(), leida: false });
  saveAlertas(alertas);
}

function getCorreosAlerta() { return JSON.parse(localStorage.getItem(DB_KEYS.correosAlerta) || '[]'); }
function saveCorreosAlerta(list) { localStorage.setItem(DB_KEYS.correosAlerta, JSON.stringify(list)); }

async function enviarCorreoPruebaBackend(correos) {
  const res = await fetch(WEBAPP_URL, {
    method: 'POST',
    body: JSON.stringify({ accion: 'enviarCorreoPrueba', correos, apiKey: API_KEY })
  });
  return res.json();
}
