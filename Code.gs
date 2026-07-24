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
const CORREOS_EMPRESA_DEFAULT = ['arrowrelax@gmail.com']; // se puede sobreescribir desde la hoja de config

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
    case 'agregarMantenimiento':
      resultado = agregarMantenimiento(body.datos);
      break;
    case 'enviarCorreoPrueba':
      resultado = enviarCorreoPrueba(body.correos);
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
    case 'listarMantenimientos':
      resultado = listarMantenimientos(e.parameter.equipoId);
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
    (datos.area || []).join(', '), 'no-verificado', '', linkToken
  ]);

  // Alerta a la empresa: nuevo proveedor
  const correos = obtenerCorreosAlerta();
  if (correos.length > 0) {
    MailApp.sendEmail({
      to: correos.join(','),
      subject: 'Nuevo proveedor registrado - ' + datos.razonSocial,
      body: `Se registró un nuevo proveedor: ${datos.razonSocial}.\nPor favor verificar en el sistema.`
    });
  }
  registrarAlerta(`Nuevo proveedor registrado: ${datos.razonSocial}`, 'info');

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
  const correoProveedor = data[rowIndex][9];
  const token = data[rowIndex][colToken];
  const linkCorreccion = `TU_URL_DEL_FORMULARIO_PUBLICADO?token=${token}`;

  if (estado === 'no-verificado') {
    const mensajeObservacion = observacion ? `\n${observacion}\n` : '';
    MailApp.sendEmail({
      to: correoProveedor,
      subject: 'Corrección requerida - Registro de proveedor Acerimallas',
      body: `Estimado proveedor,\n\nSe encontraron observaciones en su registro.${mensajeObservacion}\nPor favor ingrese al siguiente link para corregir:\n${linkCorreccion}\n\nGracias.`
    });
    registrarAlerta(`Proveedor ${razonSocial} marcado como no verificado.`, 'warning');
  } else if (estado === 'verificado') {
    MailApp.sendEmail({
      to: correoProveedor,
      subject: 'Registro verificado - Acerimallas',
      body: `Estimado proveedor,\n\nSu registro ha sido verificado y aprobado exitosamente.\n\nGracias.`
    });
    registrarAlerta(`Proveedor ${razonSocial} verificado.`, 'success');
  }

  return { ok: true };
}

function listarProveedores() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Proveedores');
  const data = sheet.getDataRange().getValues();
  data.shift(); // quita encabezados
  return { ok: true, proveedores: data };
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
    JSON.stringify(datos.datosEspecificos || {})
  ]);
  return { ok: true, id };
}

function listarEquipos(proveedorId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Equipos');
  const data = sheet.getDataRange().getValues();
  data.shift();
  const filtrados = proveedorId ? data.filter(row => row[1] === proveedorId) : data;
  return { ok: true, equipos: filtrados };
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

  sheet.appendRow([id, datos.equipoId, datos.fecha, datos.tipo, datos.observacion || '', datos.detalle || '', pdfUrl, datos.origen]);
  return { ok: true, id, pdfUrl };
}

function listarMantenimientos(equipoId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Mantenimientos');
  const data = sheet.getDataRange().getValues();
  data.shift();
  const filtrados = equipoId ? data.filter(row => row[1] === equipoId) : data;
  return { ok: true, mantenimientos: filtrados };
}

// Genera un PDF simple con los datos del mantenimiento (fotos y firma se
// insertan como imágenes). Para un diseño más elaborado, usar una
// plantilla de Google Docs y reemplazar marcadores con DocumentApp.
function generarPdfMantenimiento(datos) {
  const doc = DocumentApp.create(`Mantenimiento_${datos.tipo}_${datos.fecha}`);
  const body = doc.getBody();
  body.appendParagraph('Reporte de Mantenimiento').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Fecha: ${datos.fecha}`);
  body.appendParagraph(`Tipo de servicio: ${datos.tipo}`);
  body.appendParagraph(`Observación: ${datos.observacion || '-'}`);
  body.appendParagraph(`Detalle de lo resuelto: ${datos.detalle || '-'}`);

  if (datos.fotosBase64 && datos.fotosBase64.length) {
    body.appendParagraph('Fotos:');
    datos.fotosBase64.forEach(fotoB64 => {
      const blob = Utilities.newBlob(Utilities.base64Decode(fotoB64), 'image/png', 'foto.png');
      body.appendImage(blob);
    });
  }
  if (datos.firmaBase64) {
    body.appendParagraph('Firma:');
    const blob = Utilities.newBlob(Utilities.base64Decode(datos.firmaBase64), 'image/png', 'firma.png');
    body.appendImage(blob);
  }

  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs('application/pdf');
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true); // borra el Doc intermedio, solo deja el PDF
  return pdfFile.getUrl();
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

function obtenerCorreosAlerta() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log_Alertas');
  // Se puede guardar la lista de correos en una celda fija (ej. A1) como JSON,
  // o crear una hoja "Config" dedicada. Por simplicidad, usar el default:
  return CORREOS_EMPRESA_DEFAULT;
}

function registrarAlerta(mensaje, tipo) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Log_Alertas');
  sheet.appendRow([new Date(), mensaje, tipo]);
}

function enviarCorreoPrueba(correos) {
  MailApp.sendEmail({
    to: correos.join(','),
    subject: 'Correo de prueba - Sistema Acerimallas',
    body: 'Este es un correo de prueba. Si lo recibiste, la configuración de alertas está funcionando correctamente.'
  });
  return { ok: true };
}
