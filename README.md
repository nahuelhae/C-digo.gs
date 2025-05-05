// Mostrar el formulario en un Google Site script llamado Código.gs
function doGet() {
  return HtmlService.createHtmlOutputFromFile('GarantiaForm').setTitle("Formulario Garantía");
}

// LOGIN
function loginUsuario(usuario, contrasena) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST");
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] === usuario && datos[i][3] === contrasena) {
      return { razon: datos[i][1], email: datos[i][4], rol: datos[i][5] };
    }
  }
  return { error: "Usuario o contraseña incorrectos." };
}

// VALIDAR VIN
function validarVIN(vin) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos");
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === vin) {
      return { motor: datos[i][1], modelo: datos[i][2], marca: datos[i][3] };
    }
  }
  return { error: "VIN no encontrado." };
}

// OBTENER N° DE PEDIDO
function obtenerNumeroPedido() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Garantia");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return "GR01";
  const ultimo = hoja.getRange(ultimaFila, 2).getValue(); // Col B
  const num = parseInt(ultimo.replace("GR", "")) + 1;
  return "GR" + String(num).padStart(2, "0");
}

// ENVIAR CORREO DE CONFIRMACIÓN
function enviarConfirmacionPorEmail(data, numeroPedido) {
  const hojaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST");
  const datosUsuarios = hojaUsuarios.getDataRange().getValues();
  let emailUsuario = "";

  // Buscar el correo del usuario
  for (let i = 1; i < datosUsuarios.length; i++) {
    if (datosUsuarios[i][1] === data.usuario) { // Comparar por Razón Social
      emailUsuario = datosUsuarios[i][4]; // Columna E: Email
      break;
    }
  }

  if (!emailUsuario) {
    Logger.log("No se encontró el email del usuario.");
    return;
  }

  // Preparar datos del correo
  const asunto = `Confirmación de Reclamo - N° de Pedido ${numeroPedido}`;
  const linksFotos = [
    data.urlFotoPieza,
    ...(data.urlsFotoRepuestos ? data.urlsFotoRepuestos.split(", ") : []),
    data.urlFotoTablero,
  ].filter(Boolean); // Eliminar valores vacíos

  const cuerpo = `
    Hola ${data.usuario},

    Tu reclamo ha sido registrado con éxito.

    📍 N° de Pedido: ${numeroPedido}
    📍 VIN: ${data.vin}
    📍 Tipo de Reclamo: ${data.tipoReclamo}
    📍 Kilometraje: ${data.kms}
    📍 Descripción: ${data.fallo}
    📍 Repuesto/s utilizado/s: ${data.repuestoPrincipal}, ${data.repuestos}
    📍 Archivo/s adjunto/s: ${linksFotos.length} 🔗

    Nuestro equipo revisará tu solicitud y te informaremos a la brevedad.

    Saludos,
    Equipo de Garantías.
  `;

  // Enviar el correo
  MailApp.sendEmail({
    to: emailUsuario,
    cc: "ngalvan@simpa.com.ar", // Copia al administrador
    subject: asunto,
    body: cuerpo,
  });

  Logger.log(`Correo enviado a ${emailUsuario} con copia a ngalvan@simpa.com.ar`);
}

// REGISTRO DEL RECLAMO
function registrarReclamo(data) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Garantia");

  // Validar kilometraje previo
  const datos = hoja.getDataRange().getValues();
  let kmMax = 0;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][3] === data.vin) { // VIN en col D
      kmMax = Math.max(kmMax, parseInt(datos[i][7] || 0)); // KM en col H
    }
  }
  if (parseInt(data.kms) < kmMax) {
    return {
      error: true,
      message: "El KM ingresado debe ser mayor o igual al último registrado (" + kmMax + ").",
      kmMax: kmMax
    };
  }

  const numeroPedido = obtenerNumeroPedido();

  // Formatear fecha como "dd/MM/yyyy"
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  hoja.appendRow([
    fecha,
    numeroPedido,
    data.usuario,
    data.vin,
    data.motor,
    data.modelo,
    data.marca,
    parseInt(data.kms),
    data.tipoReclamo,
    data.fallo,
    data.repuestoPrincipal,
    data.urlFotoPieza,
    data.repuestos,
    data.urlsFotoRepuestos,
    data.urlFotoTablero
  ]);

  // Enviar correo de confirmación
  enviarConfirmacionPorEmail(data, numeroPedido);

  return { error: false, message: "Reclamo registrado exitosamente con N°: " + numeroPedido };
}

// SUBIDA DE ARCHIVOS
function subirArchivo(base64, nombre, tipo) {
  let folderId = "";
  if (tipo === "pieza") folderId = "1KM2XZCeZAIP5fv8z3at_HzkQ77WkhF2v";
  if (tipo === "repuestos") folderId = "1Ylzp1Cf6tCBN8uzqMRoPAz27gvVERSoc";
  if (tipo === "tablero") folderId = "1kNlP1OMfh8CLOHkF68PwPraH9h1KCPnS";

  const carpeta = DriveApp.getFolderById(folderId);
  const contentType = base64.match(/^data:(.+);base64,/)[1];
  const bytes = Utilities.base64Decode(base64.split(",")[1]);
  const blob = Utilities.newBlob(bytes, contentType, nombre);
  const archivo = carpeta.createFile(blob);
  return archivo.getUrl();
}
