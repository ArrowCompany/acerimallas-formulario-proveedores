/* =====================================================================
   CODE.GS — Backend del Sistema de Proveedores y Mantenimientos
   Acerimallas S.A.

   PASOS PARA ACTIVARLO (5-10 min):
   1. Crea un Google Sheet nuevo con 4 hojas llamadas exactamente:
      "Proveedores", "Equipos", "Mantenimientos", "Log_Alertas"
   2. Copia el ID del Sheet (está en la URL, entre /d/ y /edit) y pégalo
      abajo en SHEET_ID.
   3. Crea una carpeta en Google Drive para los PDFs/fotos/firmas,
      copia su ID y pégalo en DRIVE_FOLDER_ID.
   4. Ve a Extensiones > Apps Script en tu Sheet, pega este código.
   5. Reemplaza CORREOS_EMPRESA con el/los correos que reciben alertas
      (o usa la hoja de Configuración de correos desde la app).
   6. Implementar > Nueva implementación > Tipo: Aplicación web.
      - Ejecutar como: Yo
      - Quién tiene acceso: Cualquier usuario
   7. Copia la URL que te da (termina en /exec) y pégala en
      WEBAPP_URL dentro de js/store.js (o donde hagas los fetch).
   ===================================================================== */

const SHEET_ID = '1WeNgnCUaE2Pp2H5LDNUOmMQ3lkx1YXlCmhg1eoJGpLE';
const DRIVE_FOLDER_ID = '1Wfx7XyVZ3TjUSKiL7kLYNEi9LhHdxJ7o';
const API_KEY = 'acerimallas-2026-x7k9m2'; // cámbiala por cualquier texto largo que tú elijas
const FORM_URL = 'https://arrowcompany.github.io/acerimallas-formulario-proveedores/registro_proveedor.html';
const CORREOS_EMPRESA_DEFAULT = ['arrowrelax@gmail.com']; // se puede sobreescribir desde la hoja de config

// Nombre para mostrar en correos y alertas: "Nombre comercial (Razón social)"
function nombreProveedor(nombreComercial, razonSocial) {
  if (nombreComercial && razonSocial) return `${nombreComercial} (${razonSocial})`;
  return nombreComercial || razonSocial || 'Proveedor';
}

// Etiquetas legibles de cada campo del formulario, para listarlas en el
// correo de corrección (mismo orden/textos que la ficha del proveedor)
const ETIQUETAS_CAMPOS = {
  tipoProveedor: 'Tipo de proveedor',
  razonSocial: 'Razón social',
  nombreComercial: 'Nombre comercial',
  ruc: 'RUC',
  archivoRuc: 'RUC actualizado / SRI (PDF)',
  telefono1: 'Teléfono 1',
  telefono2: 'Teléfono 2',
  ciudad: 'Ciudad',
  provincia: 'Provincia',
  correoRetenciones: 'Correo (retenciones)',
  direccion: 'Dirección',
  representanteLegal: 'Representante legal',
  archivoRepLegal: 'Documento representante legal (PDF)',
  archivoNombramiento: 'Nombramiento representante (PDF)',
  contacto1: 'Contacto adicional 1',
  contacto2: 'Contacto adicional 2',
  formaPago: 'Forma de pago',
  entidadBancaria: 'Entidad bancaria',
  tipoCuenta: 'Tipo de cuenta',
  numeroCuenta: 'Número de cuenta',
  titularCuenta: 'Titular de la cuenta',
  certificadoBancario: 'Certificado bancario (PDF)',
  area: 'Área de servicio'
};

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.apiKey !== API_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No autorizado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const accion = body.accion;

  let resultado;
  switch (accion) {
    case 'guardarProveedor':
      resultado = guardarProveedor(body.datos);
      break;
    case 'actualizarEstadoProveedor':
      resultado = actualizarEstadoProveedor(body.id, body.estado, body.camposConError, body.observacion);
      break;
    case 'agregarEquipo':
      resultado = agregarEquipo(body.datos);
      break;
    case 'eliminarEquipo':
      resultado = eliminarEquipo(body.equipoId);
      break;
    case 'recuperarEquipo':
      resultado = recuperarEquipo(body.equipoId);
      break;
    case 'actualizarEquipo':
      resultado = actualizarEquipo(body.equipoId, body.datos);
      break;
    case 'eliminarProveedor':
      resultado = eliminarProveedor(body.id);
      break;
    case 'recuperarProveedor':
      resultado = recuperarProveedor(body.id);
      break;
    case 'marcarAlertasLeidas':
      resultado = marcarAlertasLeidas();
      break;
    case 'agregarMantenimiento':
      resultado = agregarMantenimiento(body.datos);
      break;
    case 'actualizarProximoMantenimiento':
      resultado = actualizarProximoMantenimiento(body.equipoId, body.nuevaFecha);
      break;
    case 'enviarCorreoPrueba':
      resultado = enviarCorreoPrueba(body.correos);
      break;
    case 'actualizarProveedorPorToken':
      resultado = actualizarProveedorPorToken(body.token, body.datos);
      break;
    default:
      resultado = { ok: false, error: 'Acción no reconocida' };
  }

  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e.parameter.apiKey !== API_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No autorizado' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const accion = e.parameter.accion;
  let resultado;

  switch (accion) {
    case 'listarProveedores':
      resultado = listarProveedores();
      break;
    case 'listarEquipos':
      resultado = listarEquipos(e.parameter.proveedorId);
      break;
    case 'listarEquiposEliminados':
      resultado = listarEquiposEliminados();
      break;
    case 'listarProveedoresEliminados':
      resultado = listarProveedoresEliminados();
      break;
    case 'listarAlertas':
      resultado = listarAlertas();
      break;
    case 'obtenerArchivosProveedor':
      resultado = obtenerArchivosProveedor(e.parameter.id);
      break;
    case 'listarMantenimientos':
      resultado = listarMantenimientos(e.parameter.equipoId);
      break;
    case 'obtenerProveedorPorToken':
      resultado = obtenerProveedorPorToken(e.parameter.token);
      break;
    default:
      resultado = { ok: false, error: 'Acción no reconocida' };
  }

  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------
// PROVEEDORES
// ---------------------------------------------------------------------
function guardarProveedor(datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const id = Utilities.getUuid();
  const linkToken = Utilities.getUuid();

  // Sube cada archivo adjunto (viene en base64 desde el formulario) a Drive
  const archivos = {};
  ['archivoRuc', 'archivoRepLegal', 'archivoNombramiento', 'archivoCertBancario'].forEach(campo => {
    if (datos[campo] && datos[campo].base64) {
      archivos[campo] = subirArchivoADrive(datos[campo], `${campo}_${datos.razonSocial}`);
    }
  });

  sheet.appendRow([
    id, new Date(), datos.razonSocial, datos.nombreComercial, datos.ruc,
    datos.telefono1, datos.telefono2, datos.ciudad, datos.provincia,
    datos.correoRetenciones, datos.direccion, datos.representanteLegal,
    datos.contacto1, datos.contacto2, datos.formaPago, datos.entidadBancaria,
    datos.tipoCuenta, datos.numeroCuenta, datos.titularCuenta,
    archivos.archivoRuc || '', archivos.archivoRepLegal || '',
    archivos.archivoNombramiento || '', archivos.archivoCertBancario || '',
    (datos.area || []).join(', '), 'no-verificado', '', linkToken, '', false,
    datos.correo || '' // correo de registro (campo 1) — es el correo de contacto real
  ]);

  // Alerta a la empresa: nuevo proveedor
  const nombreParaAviso = nombreProveedor(datos.nombreComercial, datos.razonSocial);
  const correos = obtenerCorreosAlerta();
  if (correos.length > 0) {
    MailApp.sendEmail({
      to: correos.join(','),
      subject: 'Nuevo proveedor registrado - ' + nombreParaAviso,
      body: `Se registró un nuevo proveedor: ${nombreParaAviso}.\nPor favor verificar en el sistema.`
    });
  }
  registrarAlerta(`Nuevo proveedor registrado: ${nombreParaAviso}`, 'info');

  return { ok: true, id, linkToken };
}

function actualizarEstadoProveedor(id, estado, camposConError, observacion) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Proveedor no encontrado' };

  const colEstado = 24; // ajustar según el orden real de columnas
  const colCampos = 25;
  const colToken = 26;
  const colObservacion = 27; // nueva columna "Observación" al final del Sheet

  sheet.getRange(rowIndex + 1, colEstado + 1).setValue(estado);
  sheet.getRange(rowIndex + 1, colCampos + 1).setValue((camposConError || []).join(', '));
  sheet.getRange(rowIndex + 1, colObservacion + 1).setValue(observacion || '');

  const razonSocial = data[rowIndex][2];
  const nombreComercial = data[rowIndex][3];
  const nombreParaAviso = nombreProveedor(nombreComercial, razonSocial);
  const correoProveedor = data[rowIndex][29] || data[rowIndex][9];
  const token = data[rowIndex][colToken];
  const linkCorreccion = `${FORM_URL}?token=${token}`;

  if (estado === 'no-verificado') {
    const listaErrores = (camposConError || [])
      .map(c => '- ' + (ETIQUETAS_CAMPOS[c] || c))
      .join('\n');
    const mensajeObservacion = observacion ? `\nDetalle: ${observacion}\n` : '';
    MailApp.sendEmail({
      to: correoProveedor,
      subject: 'Corrección requerida - Registro de proveedor Acerimallas',
      body: `Estimado proveedor,\n\nRevisamos su registro y encontramos observaciones en los siguientes datos:\n\n${listaErrores}\n${mensajeObservacion}\nEl resto de su información ya está correcta, no es necesario volver a enviarla.\n\nPor favor ingrese al siguiente link para corregir únicamente lo señalado:\n${linkCorreccion}\n\nGracias.`
    });
    registrarAlerta(`Proveedor ${nombreParaAviso} marcado como no verificado.`, 'warning');
  } else if (estado === 'verificado') {
    MailApp.sendEmail({
      to: correoProveedor,
      subject: 'Registro verificado - Acerimallas',
      body: `Estimado proveedor,\n\nSu registro ha sido verificado y aprobado exitosamente.\n\nGracias.`
    });
    registrarAlerta(`Proveedor ${nombreParaAviso} verificado.`, 'success');
  }

  return { ok: true };
}

function listarProveedores() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  data.shift(); // quita encabezados
  const activos = data.filter(row => row[28] !== true);
  return { ok: true, proveedores: activos };
}

function listarProveedoresEliminados() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  data.shift();
  const eliminados = data.filter(row => row[28] === true);
  return { ok: true, proveedores: eliminados };
}

// Envía el proveedor a la carpeta de Eliminados (no borra sus datos ni su
// historial, solo lo oculta de las listas normales hasta que se recupere)
function eliminarProveedor(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Proveedor no encontrado' };

  sheet.getRange(rowIndex + 1, 29).setValue(true);
  const nombreParaAviso = nombreProveedor(data[rowIndex][3], data[rowIndex][2]);
  registrarAlerta(`Proveedor ${nombreParaAviso} enviado a la carpeta de eliminados.`, 'info');
  return { ok: true };
}

// Restaura al proveedor exactamente con el estado que tenía antes de borrarlo
function recuperarProveedor(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Proveedor no encontrado' };

  sheet.getRange(rowIndex + 1, 29).setValue(false);
  const nombreParaAviso = nombreProveedor(data[rowIndex][3], data[rowIndex][2]);
  registrarAlerta(`Proveedor ${nombreParaAviso} fue recuperado.`, 'success');
  return { ok: true };
}

// Usado por el formulario público cuando el proveedor abre el link de corrección
function obtenerProveedorPorToken(token) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  const fila = data.find(row => row[26] === token);
  if (!fila) return { ok: false, error: 'Link no válido o expirado' };
  return { ok: true, proveedor: fila };
}

// Guarda la corrección del proveedor sobre su mismo registro (no crea uno nuevo)
function actualizarProveedorPorToken(token, datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[26] === token);
  if (rowIndex === -1) return { ok: false, error: 'Link no válido o expirado' };

  const filaActual = data[rowIndex];

  const archivos = {};
  ['archivoRuc', 'archivoRepLegal', 'archivoNombramiento', 'archivoCertBancario'].forEach(campo => {
    if (datos[campo] && datos[campo].base64) {
      archivos[campo] = subirArchivoADrive(datos[campo], `${campo}_${datos.razonSocial}`);
    }
  });

  const nuevaFila = [
    filaActual[0], filaActual[1], datos.razonSocial, datos.nombreComercial, datos.ruc,
    datos.telefono1, datos.telefono2, datos.ciudad, datos.provincia, datos.correoRetenciones,
    datos.direccion, datos.representanteLegal, datos.contacto1, datos.contacto2, datos.formaPago,
    datos.entidadBancaria, datos.tipoCuenta, datos.numeroCuenta, datos.titularCuenta,
    archivos.archivoRuc || filaActual[19],
    archivos.archivoRepLegal || filaActual[20],
    archivos.archivoNombramiento || filaActual[21],
    archivos.archivoCertBancario || filaActual[22],
    (datos.area || []).join(', '),
    'corregido', // el proveedor ya reenvió su corrección, queda pendiente de re-revisión
    '', // se limpian los campos con error ya corregidos
    token, // conserva el mismo link para el proveedor
    '', // se limpia la observación anterior
    filaActual[28], // conserva el estado de eliminado tal cual estaba
    datos.correo || filaActual[29] // correo de registro (puede confirmarlo o dejarlo igual)
  ];

  sheet.getRange(rowIndex + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);

  const nombreParaAviso = nombreProveedor(datos.nombreComercial, datos.razonSocial);
  const correos = obtenerCorreosAlerta();
  if (correos.length > 0) {
    MailApp.sendEmail({
      to: correos.join(','),
      subject: 'Proveedor corrigió sus datos - ' + nombreParaAviso,
      body: `El proveedor ${nombreParaAviso} corrigió su información y la reenvió. Por favor verificar nuevamente en el sistema.`
    });
  }
  registrarAlerta(`El proveedor ${nombreParaAviso} ya corrigió sus datos, verificar.`, 'info');

  return { ok: true };
}

// ---------------------------------------------------------------------
// EQUIPOS Y MANTENIMIENTOS
// ---------------------------------------------------------------------
function agregarEquipo(datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const id = Utilities.getUuid();
  sheet.appendRow([
    id,
    datos.proveedorId || '',
    datos.nombre,
    datos.ubicacion,
    datos.serie || '',
    datos.proximoMantenimiento || '',
    datos.tipo || 'otro',
    JSON.stringify(datos.datosEspecificos || {}),
    false
  ]);
  return { ok: true, id };
}

// Reescribe los datos editables de un equipo existente (mantiene id, dueño y
// estado de eliminado tal como estaban)
function actualizarEquipo(id, datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Equipo no encontrado' };

  const filaActual = data[rowIndex];
  const nuevaFila = [
    filaActual[0], filaActual[1], datos.nombre, datos.ubicacion, filaActual[4],
    datos.proximoMantenimiento || '', datos.tipo || filaActual[6],
    JSON.stringify(datos.datosEspecificos || {}), filaActual[8]
  ];
  sheet.getRange(rowIndex + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);
  return { ok: true };
}

// Actualiza solo la columna "proximoMantenimiento" (col. 6) de un equipo.
// Se usa para reprogramar automáticamente después de registrar una visita.
function actualizarProximoMantenimiento(equipoId, nuevaFecha) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === equipoId);
  if (rowIndex === -1) return { ok: false, error: 'Equipo no encontrado' };

  sheet.getRange(rowIndex + 1, 6).setValue(nuevaFecha);

  // Como cambia la fecha, se limpian las marcas de "ya avisado" guardadas en
  // datosEspecificos para que el nuevo ciclo pueda volver a generar sus alertas
  let datosEspecificos = {};
  try { datosEspecificos = data[rowIndex][7] ? JSON.parse(data[rowIndex][7]) : {}; } catch (e) {}
  delete datosEspecificos._alertaSemanaPara;
  delete datosEspecificos._alertaHoyPara;
  sheet.getRange(rowIndex + 1, 8).setValue(JSON.stringify(datosEspecificos));

  return { ok: true };
}

function listarEquipos(proveedorId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  data.shift();
  let activos = data.filter(row => row[8] !== true);
  if (proveedorId) activos = activos.filter(row => row[1] === proveedorId);
  return { ok: true, equipos: activos };
}

function listarEquiposEliminados() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  data.shift();
  const eliminados = data.filter(row => row[8] === true);
  return { ok: true, equipos: eliminados };
}

// Envía el equipo a la carpeta de Eliminados (no borra su historial de
// mantenimientos, solo lo oculta hasta que se recupere)
function eliminarEquipo(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Equipo no encontrado' };

  sheet.getRange(rowIndex + 1, 9).setValue(true);
  registrarAlerta(`Equipo "${data[rowIndex][2]}" enviado a la carpeta de eliminados.`, 'info');
  return { ok: true };
}

function recuperarEquipo(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Equipo no encontrado' };

  sheet.getRange(rowIndex + 1, 9).setValue(false);
  registrarAlerta(`Equipo "${data[rowIndex][2]}" fue recuperado.`, 'success');
  return { ok: true };
}

function agregarMantenimiento(datos) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Mantenimientos');
  const id = Utilities.getUuid();
  let pdfUrl = '';

  if (datos.pdfBase64) {
    // Caso B: PDF externo ya elaborado
    pdfUrl = subirArchivoADrive({ base64: datos.pdfBase64, nombre: datos.pdfNombre, tipo: 'application/pdf' }, datos.pdfNombre);
  } else {
    // Caso A: se genera el PDF a partir de fecha/tipo/observación/detalle/fotos/firma
    pdfUrl = generarPdfMantenimiento(datos);
  }

  sheet.appendRow([
    id, datos.equipoId, datos.fecha, datos.tipo, datos.observacion || '', datos.detalle || '',
    pdfUrl, datos.origen, datos.modoPago || '', datos.costo || ''
  ]);
  return { ok: true, id, pdfUrl };
}

function listarMantenimientos(equipoId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Mantenimientos');
  const data = sheet.getDataRange().getValues();
  data.shift();
  const filtrados = equipoId ? data.filter(row => row[1] === equipoId) : data;
  return { ok: true, mantenimientos: filtrados };
}

// Genera el PDF del mantenimiento con el mismo estilo del reporte de
// referencia: banner oscuro de encabezado, franjas grises por sección,
// datos del equipo, servicio, diagnóstico, trabajo realizado, evidencias
// fotográficas y firma. Se arma como HTML y se convierte a PDF.
function generarPdfMantenimiento(datos) {
  const escapar = (t) => String(t || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  let fotosHtml = '';
  if (datos.fotosBase64 && datos.fotosBase64.length) {
    const imgs = datos.fotosBase64.map(f =>
      `<img src="data:image/png;base64,${f}" style="max-width:260px;max-height:260px;border:1px solid #E5E3DD;border-radius:8px;margin:0 10px 10px 0;">`
    ).join('');
    fotosHtml = `
      ${seccionHtml('EVIDENCIAS FOTOGRÁFICAS')}
      <div style="padding:4px 0 16px;">${imgs}</div>
    `;
  }

  let firmaHtml = '';
  if (datos.firmaBase64) {
    firmaHtml = `
      ${seccionHtml('FIRMA DEL CLIENTE')}
      <div style="padding:10px 0;">
        <img src="data:image/png;base64,${datos.firmaBase64}" style="max-width:200px;max-height:100px;">
      </div>
    `;
  }

  // Franja divisoria de sección: el conversor de HTML a PDF de Apps Script
  // ignora tanto el background-color de CSS como el bgcolor de las tablas,
  // pero sí respeta imágenes — por eso la franja se hace con una imagen real
  // (un PNG de 1x1 gris claro, estirado a lo ancho de la página).
  const FRANJA_GRIS_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR42mN4//Y5AAWSAsSRR9hjAAAAAElFTkSuQmCC';

  function seccionHtml(texto) {
    return `
      <div style="margin:20px 0 8px;">
        <img src="data:image/png;base64,${FRANJA_GRIS_PNG}" width="535" height="4">
        <div style="padding:6px 0 0;">
          <font size="2" color="#1F1E1B"><b>${texto}</b></font>
        </div>
      </div>
    `;
  }

  const html = `
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: Arial, sans-serif; color: #1F1E1B; margin: 0; }
        .contenido { padding: 20px 30px 30px; }
        .fila { display: flex; padding: 4px 0; font-size: 12px; }
        .fila .label { width: 160px; color: #6B6459; }
        .fila .valor { font-weight: bold; }
        .texto { font-size: 12px; padding: 2px 0 10px; }
      </style>
    </head>
    <body>
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr><td bgcolor="#1F1E1B" style="padding:24px 30px;">
          <font color="#F8F7F4" size="5"><b>Acerimallas - Mantenimiento de Equipos</b></font><br>
          <font color="#A8A39A" size="2">Reporte de visita técnica</font><br>
          <font color="#6B6459" size="1">Generado: ${new Date().toLocaleString('es-EC')}</font>
        </td></tr>
      </table>
      <div class="contenido">

        ${seccionHtml('INFORMACIÓN DEL EQUIPO')}
        <div class="fila"><span class="label">Equipo:</span><span class="valor">${escapar(datos.equipoNombre) || 'No especificado'}</span></div>
        <div class="fila"><span class="label">Ubicación:</span><span class="valor">${escapar(datos.equipoUbicacion) || 'No especificada'}</span></div>

        ${seccionHtml('SERVICIO')}
        <div class="fila"><span class="label">Tipo:</span><span class="valor">${escapar(datos.tipo)}</span></div>
        <div class="fila"><span class="label">Fecha:</span><span class="valor">${escapar(datos.fecha)}</span></div>
        <div class="fila"><span class="label">Modo de pago:</span><span class="valor">${escapar(datos.modoPago) || 'No aplica'}</span></div>
        <div class="fila"><span class="label">Costo:</span><span class="valor">$${escapar(datos.costo) || '0.00'}</span></div>

        ${seccionHtml('DIAGNÓSTICO / ESTADO INICIAL')}
        <div class="texto">${escapar(datos.observacion) || 'No especificado'}</div>

        ${seccionHtml('TRABAJO REALIZADO')}
        <div class="texto">${escapar(datos.detalle) || 'No especificado'}</div>

        ${fotosHtml}
        ${firmaHtml}
      </div>
    </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, 'text/html', `Mantenimiento_${datos.tipo}_${datos.fecha}.html`);
  const pdfBlob = blob.getAs('application/pdf');
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob);
  return pdfFile.getUrl();
}

// ---------------------------------------------------------------------
// ALERTAS AUTOMÁTICAS DE MANTENIMIENTO
// Revisa todos los equipos y, si a un mantenimiento le falta exactamente
// 7 días o es HOY, registra la alerta en el sistema (Log_Alertas) y manda
// el correo a los destinatarios configurados. Para que esto corra solo,
// hay que instalar un disparador diario UNA SOLA VEZ (ver crearTriggerDiario).
// ---------------------------------------------------------------------
function revisarAlertasMantenimiento() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const correos = obtenerCorreosAlerta();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const proximoStr = row[5];
    if (!proximoStr) continue;

    const proximo = new Date(proximoStr);
    proximo.setHours(0, 0, 0, 0);
    const diffDias = Math.round((proximo - hoy) / (1000 * 60 * 60 * 24));

    let datosEspecificos = {};
    try { datosEspecificos = row[7] ? JSON.parse(row[7]) : {}; } catch (e) {}

    // "clave" para no mandar la misma alerta dos veces mientras la fecha no cambie
    const clave = String(proximoStr);
    let cambio = false;

    if (diffDias === 7 && datosEspecificos._alertaSemanaPara !== clave) {
      enviarAlertaMantenimiento(row, correos, 'en 1 semana (' + proximoStr + ')');
      datosEspecificos._alertaSemanaPara = clave;
      cambio = true;
    } else if (diffDias === 0 && datosEspecificos._alertaHoyPara !== clave) {
      enviarAlertaMantenimiento(row, correos, 'HOY');
      datosEspecificos._alertaHoyPara = clave;
      cambio = true;
    }

    if (cambio) sheet.getRange(i + 1, 8).setValue(JSON.stringify(datosEspecificos));
  }
}

function enviarAlertaMantenimiento(row, correos, cuando) {
  const nombre = row[2], ubicacion = row[3];
  const mensaje = `Mantenimiento de "${nombre}" (${ubicacion}) programado ${cuando}.`;

  if (correos.length > 0) {
    MailApp.sendEmail({
      to: correos.join(','),
      subject: `Recordatorio de mantenimiento - ${nombre}`,
      body: mensaje
    });
  }
  registrarAlerta(mensaje, 'warning');
}

// Ejecutar ESTA función UNA SOLA VEZ manualmente desde el editor de Apps Script
// (elígela en el desplegable de funciones y dale "Ejecutar") para instalar el
// disparador que corre revisarAlertasMantenimiento todos los días a las 8am.
function crearTriggerDiario() {
  // Evita crear el disparador dos veces si ya existe
  const yaExiste = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'revisarAlertasMantenimiento');
  if (yaExiste) return;

  ScriptApp.newTrigger('revisarAlertasMantenimiento')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

// ---------------------------------------------------------------------
// UTILIDADES
// ---------------------------------------------------------------------
function subirArchivoADrive(archivo, nombreBase) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const blob = Utilities.newBlob(Utilities.base64Decode(archivo.base64), archivo.tipo || 'application/pdf', nombreBase + '.pdf');
  const file = folder.createFile(blob);
  return file.getUrl();
}

// Trae los certificados PDF adjuntos de un proveedor directo desde Drive
// (en base64), para poder unirlos en un solo PDF descargable desde la ficha.
function obtenerArchivosProveedor(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(row => row[0] === id);
  if (rowIndex === -1) return { ok: false, error: 'Proveedor no encontrado' };

  const row = data[rowIndex];
  const campos = {
    archivoRuc: row[19],
    archivoRepLegal: row[20],
    archivoNombramiento: row[21],
    certificadoBancario: row[22]
  };

  const archivos = {};
  Object.keys(campos).forEach(key => {
    const url = campos[key];
    if (!url) return;
    const match = String(url).match(/[-\w]{25,}/); // extrae el ID de Drive de la URL guardada
    if (!match) return;
    try {
      const blob = DriveApp.getFileById(match[0]).getBlob();
      archivos[key] = {
        base64: Utilities.base64Encode(blob.getBytes()),
        nombre: blob.getName(),
        tipo: blob.getContentType()
      };
    } catch (e) {
      // Si el archivo ya no existe o no es accesible, simplemente se omite
    }
  });

  return { ok: true, archivos };
}

function obtenerCorreosAlerta() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log_Alertas');
  // Se puede guardar la lista de correos en una celda fija (ej. A1) como JSON,
  // o crear una hoja "Config" dedicada. Por simplicidad, usar el default:
  return CORREOS_EMPRESA_DEFAULT;
}

function registrarAlerta(mensaje, tipo) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log_Alertas');
  sheet.appendRow([Utilities.getUuid(), new Date(), mensaje, tipo, false]);
}

// Devuelve las últimas 50 alertas, más nuevas primero
function listarAlertas() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log_Alertas');
  const data = sheet.getDataRange().getValues();
  data.shift(); // encabezados
  const alertas = data
    .map(row => ({ id: row[0], fecha: row[1], mensaje: row[2], tipo: row[3], leida: row[4] === true }))
    .reverse()
    .slice(0, 50);
  return { ok: true, alertas };
}

// Marca como leídas todas las alertas (se llama al abrir el panel de la campana)
function marcarAlertasLeidas() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log_Alertas');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] !== true) sheet.getRange(i + 1, 5).setValue(true);
  }
  return { ok: true };
}

function enviarCorreoPrueba(correos) {
  MailApp.sendEmail({
    to: correos.join(','),
    subject: 'Correo de prueba - Sistema Acerimallas',
    body: 'Este es un correo de prueba. Si lo recibiste, la configuración de alertas está funcionando correctamente.'
  });
  return { ok: true };
}
