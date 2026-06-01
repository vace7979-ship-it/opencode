// ============================================================
// CONFIGURACIÓN GLOBAL
// ============================================================
const SHEET_NAME = 'Usuarios Gestoria';
const EMAIL_DESTINO = 'eloy.vazquez@cajadepac.org.mx';

// Índices de columna (0-based)
const COL_ZONA       = 0; // A
const COL_GESTOR     = 1; // B
const COL_NOMBRE     = 2; // C
const COL_NUM        = 3; // D
const COL_SUCURSAL   = 4; // E
const COL_VALIDACION = 5; // F
const COL_PUESTO     = 6; // G

const MAX_USUARIO_CHARS = 12;

// ============================================================
// ENTRY POINT
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gestión de Usuarios – Caja de Pac')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// HELPERS
// ============================================================
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Hoja "' + SHEET_NAME + '" no encontrada.');
  return sheet;
}

// ============================================================
// LECTURA DE DATOS
// ============================================================
function getSucursales() {
  const data = getSheet().getDataRange().getValues();
  const set = new Set();
  for (let i = 1; i < data.length; i++) {
    const s = String(data[i][COL_SUCURSAL]).trim();
    if (s && s !== '') set.add(s);
  }
  return Array.from(set).sort();
}

function getUsuariosBySucursal(sucursal) {
  const data = getSheet().getDataRange().getValues();
  const usuarios = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_SUCURSAL]).trim().toUpperCase() === sucursal.trim().toUpperCase()) {
      usuarios.push({
        rowIndex:   i + 1,
        zona:       data[i][COL_ZONA],
        gestor:     data[i][COL_GESTOR],
        nombre:     data[i][COL_NOMBRE],
        numero:     data[i][COL_NUM],
        sucursal:   data[i][COL_SUCURSAL],
        validacion: data[i][COL_VALIDACION],
        puesto:     data[i][COL_PUESTO]
      });
    }
  }
  return usuarios;
}

// ============================================================
// ACTUALIZACIÓN DE USUARIOS
// ============================================================
function updateUsuario(rowIndex, campo, valor) {
  valor = String(valor).trim();

  const colMap = {
    zona:       COL_ZONA       + 1,
    gestor:     COL_GESTOR     + 1,
    nombre:     COL_NOMBRE     + 1,
    numero:     COL_NUM        + 1,
    sucursal:   COL_SUCURSAL   + 1,
    validacion: COL_VALIDACION + 1,
    puesto:     COL_PUESTO     + 1
  };

  if (!colMap[campo]) return { success: false, error: 'Campo no válido: ' + campo };

  if (campo === 'gestor') {
    if (valor.length > MAX_USUARIO_CHARS) {
      return { success: false, error: 'El usuario no puede exceder ' + MAX_USUARIO_CHARS + ' caracteres.' };
    }
    if (!/^[A-Za-z0-9\-]+$/.test(valor)) {
      return { success: false, error: 'El usuario solo puede contener letras, números y guiones.' };
    }
  }

  try {
    getSheet().getRange(rowIndex, colMap[campo]).setValue(valor);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// SOLICITUD DE NUEVO USUARIO
// ============================================================
function solicitarUsuario(datos) {
  const { sucursal, tieneZona, zona, nombreCompleto, tipo, prefijo, quienSolicita } = datos;

  // Validaciones básicas
  if (!sucursal || !nombreCompleto || !quienSolicita || !tipo) {
    return { success: false, error: 'Faltan campos obligatorios.' };
  }

  const usuarioSugerido = (prefijo || '').toUpperCase();
  if (usuarioSugerido && usuarioSugerido.length > MAX_USUARIO_CHARS) {
    return { success: false, error: 'El usuario sugerido no puede exceder ' + MAX_USUARIO_CHARS + ' caracteres.' };
  }

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  const asunto = '🆕 Nueva Solicitud de Usuario – ' + sucursal + ' | ' + tipo;

  const cuerpoHtml = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:#1a73e8;color:#fff;padding:20px;border-radius:8px 8px 0 0">
        <h2 style="margin:0">Nueva Solicitud de Usuario</h2>
        <p style="margin:4px 0 0">Sistema de Gestión de Usuarios – Caja de Pac</p>
      </div>
      <div style="border:1px solid #ddd;border-top:none;padding:24px;border-radius:0 0 8px 8px">
        <table style="width:100%;border-collapse:collapse">
          <tr><td style="padding:8px;color:#666;width:180px">Fecha de solicitud</td><td style="padding:8px;font-weight:bold">${fecha}</td></tr>
          <tr style="background:#f8f9fa"><td style="padding:8px;color:#666">Sucursal</td><td style="padding:8px;font-weight:bold">${sucursal}</td></tr>
          <tr><td style="padding:8px;color:#666">Nombre completo</td><td style="padding:8px;font-weight:bold">${nombreCompleto}</td></tr>
          <tr style="background:#f8f9fa"><td style="padding:8px;color:#666">Tipo de puesto</td><td style="padding:8px;font-weight:bold">${tipo}</td></tr>
          <tr><td style="padding:8px;color:#666">¿Tiene zona a cargo?</td><td style="padding:8px;font-weight:bold">${tieneZona ? 'Sí – Zona: ' + (zona || 'por asignar') : 'No'}</td></tr>
          <tr style="background:#f8f9fa"><td style="padding:8px;color:#666">Usuario sugerido</td><td style="padding:8px;font-weight:bold;font-family:monospace">${usuarioSugerido || 'Por asignar'}</td></tr>
          <tr><td style="padding:8px;color:#666">Solicitado por</td><td style="padding:8px;font-weight:bold">${quienSolicita}</td></tr>
        </table>
        <div style="margin-top:20px;padding:12px;background:#fff3cd;border-radius:6px;border-left:4px solid #ffc107">
          <strong>Acción requerida:</strong> Favor de dar de alta al usuario en el sistema y registrarlo en la hoja <em>Usuarios Gestoria</em>.
        </div>
      </div>
    </div>
  `;

  const cuerpoTexto = [
    'NUEVA SOLICITUD DE USUARIO',
    '===========================',
    'Fecha: ' + fecha,
    'Sucursal: ' + sucursal,
    'Nombre completo: ' + nombreCompleto,
    'Tipo de puesto: ' + tipo,
    'Tiene zona: ' + (tieneZona ? 'Sí – ' + (zona || 'por asignar') : 'No'),
    'Usuario sugerido: ' + (usuarioSugerido || 'Por asignar'),
    'Solicitado por: ' + quienSolicita
  ].join('\n');

  try {
    GmailApp.sendEmail(EMAIL_DESTINO, asunto, cuerpoTexto, { htmlBody: cuerpoHtml });
    return { success: true, mensaje: 'Solicitud enviada correctamente a ' + EMAIL_DESTINO };
  } catch (e) {
    return { success: false, error: 'Error al enviar correo: ' + e.message };
  }
}
