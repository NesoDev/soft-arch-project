// Obtener los departamentos disponibles para invitar inquilinos
function getDepartamentosParaInvitaciones() {
  var sheetDep = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');
  if (!sheetDep) return [];

  var data = sheetDep.getDataRange().getValues();
  var departamentos = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][6].toString().trim().toLowerCase() === "desocupado") { // Estado en columna G
      departamentos.push({
        id: data[i][4], // ID Departamento (columna E)
        nombre: data[i][1] // Nombre del Departamento (columna B)
      });
    }
  }

  return departamentos;
}

// Obtener todas las invitaciones registradas en la hoja
function obtenerInvitaciones() {
  var sheetInv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  if (!sheetInv) return [];

  var data = sheetInv.getDataRange().getValues();
  var invitaciones = [];

  for (var i = 1; i < data.length; i++) {
    invitaciones.push({
      id: data[i][0], // ID Invitación
      departamento: data[i][1], // ID Departamento
      propietario: data[i][2], // Correo del propietario
      inquilino: data[i][3], // Correo del inquilino
      estado: data[i][4], // Estado de la invitación
      fechaEnvio: data[i][5] // Fecha de envío
    });
  }

  return invitaciones;
}

function enviarInvitacion(idDepartamento, emailInquilino) {
  var sheetInv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  var sheetDep = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');
  var sheetInq = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inquilinos');

  if (!sheetInv || !sheetDep || !sheetInq) return 'Error: No se encontraron las hojas necesarias.';

  // **Paso 1: Obtener la lista de correos de los inquilinos**
  var dataInq = sheetInq.getDataRange().getValues();
  var listaCorreos = dataInq.slice(1).map(row => row[1].trim().toLowerCase()); // Columna B (Correos)

  // **Paso 2: Validar si el correo ingresado está en la lista**
  if (!listaCorreos.includes(emailInquilino.trim().toLowerCase())) {
    return `Error: El correo ${emailInquilino} no está registrado como inquilino.`;
  }

  // **Paso 3: Verificar si el departamento sigue desocupado**
  var dataDep = sheetDep.getDataRange().getValues();
  var departamentoNombre = '';
  var departamentoEstado = '';

  for (var i = 1; i < dataDep.length; i++) {
    if (dataDep[i][4] == idDepartamento) { // ID Departamento en columna E
      departamentoNombre = dataDep[i][1]; // Nombre en columna B
      departamentoEstado = dataDep[i][6]; // Estado en columna G
      break;
    }
  }

  if (departamentoEstado.toLowerCase() !== "desocupado") {
    return 'Error: Este departamento ya no está disponible.';
  }

  // **Paso 4: Verificar si ya hay una invitación pendiente para este departamento**
  var dataInv = sheetInv.getDataRange().getValues();
  for (var i = 1; i < dataInv.length; i++) {
    if (dataInv[i][1] == idDepartamento && dataInv[i][4] == 'Pendiente') {
      return 'Ya existe una invitación pendiente para este departamento.';
    }
  }

  // **Paso 5: Generar ID único basado en timestamp y número aleatorio**
  var timestamp = new Date().getTime();
  var randomNum = Math.floor(1000000 + Math.random() * 9000000);
  var idInvitacion = timestamp.toString() + "-" + randomNum.toString();

  // **Paso 6: Registrar la invitación en la hoja**
  var fechaEnvio = new Date();
  sheetInv.appendRow([idInvitacion, idDepartamento, Session.getActiveUser().getEmail(), emailInquilino, 'Pendiente', fechaEnvio, '']);

  // **Paso 7: Enviar correo al inquilino**
  var scriptUrl = ScriptApp.getService().getUrl();
  var aceptarUrl = scriptUrl + "?accion=aceptar&id=" + idInvitacion;
  var rechazarUrl = scriptUrl + "?accion=rechazar&id=" + idInvitacion;

  var mensajeHtml = `
    <h2>Invitación para Departamento</h2>
    <p>Has recibido una invitación para formar parte del departamento <b>${departamentoNombre}</b>.</p>
    <p>Puedes aceptar o rechazar la invitación utilizando los siguientes enlaces:</p>
    <a href="${aceptarUrl}" style="display:inline-block;padding:10px;background-color:green;color:white;text-decoration:none;margin-right:10px;">Aceptar</a>
    <a href="${rechazarUrl}" style="display:inline-block;padding:10px;background-color:red;color:white;text-decoration:none;">Rechazar</a>
  `;

  MailApp.sendEmail({
    to: emailInquilino,
    subject: "Invitación a Departamento",
    htmlBody: mensajeHtml
  });

  return 'Invitación enviada con éxito.';
}

// Manejar la respuesta de la invitación (aceptar o rechazar)
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('main');

  // Si no hay parámetros, renderiza la página principal
  if (!e.parameter.accion || !e.parameter.id) {
    return template.evaluate().setTitle("Gestión de Invitaciones");
  }

  var accion = e.parameter.accion;
  var idInvitacion = e.parameter.id;
  var sheetInv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  var sheetDep = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');

  if (!sheetInv || !sheetDep) {
    return ContentService.createTextOutput("No se encontraron las hojas necesarias.");
  }

  var dataInv = sheetInv.getDataRange().getValues();
  for (var i = 1; i < dataInv.length; i++) {
    if (dataInv[i][0] == idInvitacion && dataInv[i][4] == 'Pendiente') {
      var fechaRespuesta = new Date();
      sheetInv.getRange(i + 1, 5).setValue(accion.charAt(0).toUpperCase() + accion.slice(1));
      sheetInv.getRange(i + 1, 6).setValue(fechaRespuesta);

      if (accion == 'aceptar') {
        var dataDep = sheetDep.getDataRange().getValues();
        for (var j = 1; j < dataDep.length; j++) {
          if (dataDep[j][4] == dataInv[i][1]) {
            sheetDep.getRange(j + 1, 6).setValue('Ocupado');
            break;
          }
        }
        return ContentService.createTextOutput("Has aceptado la invitación.");
      } else {
        return ContentService.createTextOutput("Has rechazado la invitación.");
      }
    }
  }

  return ContentService.createTextOutput("Invitación no encontrada o ya gestionada.");
}

// Cancelar una invitación si está en estado pendiente
function cancelarInvitacion(idInvitacion) {
  var sheetInv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  if (!sheetInv) return 'Error: No se encontró la hoja de invitaciones.';

  var data = sheetInv.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == idInvitacion && data[i][4] == 'Pendiente') {
      sheetInv.getRange(i + 1, 5).setValue('Cancelada');
      return 'Invitación cancelada con éxito.';
    }
  }
  return 'No se encontró una invitación pendiente con ese ID.';
}