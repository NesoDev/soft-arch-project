function getDepartamentosParaInvitaciones() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var departamentos = [];

  for (var i = 1; i < data.length; i++) {
    var estado = data[i][6].toString().trim().toLowerCase(); // Normaliza el estado
    if (estado === "desocupado") {
      departamentos.push({
        id: data[i][4], // ID Departamento
        nombre: data[i][1], // Nombre del Departamento
      });
    }
  }
  return departamentos;
}

function obtenerInvitaciones() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var invitaciones = [];

  for (var i = 1; i < data.length; i++) {
    invitaciones.push({
      id: data[i][0],
      departamento: data[i][1],
      inquilino: data[i][3],
      estado: data[i][4],
    });
  }
  return invitaciones;
}

function enviarInvitacion(idDepartamento, emailInquilino) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Invitaciones');
    sheet.appendRow(['ID', 'ID Departamento', 'ID Arrendador', 'ID Arrendatario', 'Estado', 'Fecha Envío', 'Fecha Respuesta']);
  }

  var invitaciones = sheet.getDataRange().getValues();
  for (var i = 1; i < invitaciones.length; i++) {
    if (invitaciones[i][1] == idDepartamento && invitaciones[i][4] == 'Pendiente') {
      return 'Ya existe una invitación pendiente para este departamento';
    }
  }

  var fechaEnvio = new Date();
  var nuevaFila = sheet.getLastRow() + 1;
  sheet.appendRow([nuevaFila, idDepartamento, Session.getActiveUser().getEmail(), emailInquilino, 'Pendiente', fechaEnvio, '']);

  var scriptUrl = ScriptApp.getService().getUrl();
  var aceptarUrl = scriptUrl + "?accion=aceptar&id=" + nuevaFila;
  var rechazarUrl = scriptUrl + "?accion=rechazar&id=" + nuevaFila;

  var mensajeHtml = `
    <h2>Invitación para Departamento</h2>
    <p>Has recibido una invitación para formar parte del departamento ID: ${idDepartamento}.</p>
    <p>Puedes aceptarla o rechazarla utilizando los siguientes botones:</p>
    <a href="${aceptarUrl}" style="display:inline-block;padding:10px;background-color:green;color:white;text-decoration:none;margin-right:10px;">Aceptar</a>
    <a href="${rechazarUrl}" style="display:inline-block;padding:10px;background-color:red;color:white;text-decoration:none;">Rechazar</a>
  `;

  MailApp.sendEmail({
    to: emailInquilino,
    subject: "Invitación a Departamento",
    htmlBody: mensajeHtml
  });

  return 'Invitación enviada con éxito';
}

function doGet(e) {
  if (!e.parameter.accion) {
    return HtmlService.createTemplateFromFile('main').evaluate()
      .setTitle("Rendo - Gestión de Arrendamientos");
  }

  var accion = e.parameter.accion;
  var idInvitacion = e.parameter.id;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  if (!sheet) return ContentService.createTextOutput("No hay hoja de invitaciones creada");

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == idInvitacion && data[i][4] == 'Pendiente') {
      var fechaRespuesta = new Date();
      sheet.getRange(i + 1, 5).setValue(accion.charAt(0).toUpperCase() + accion.slice(1));
      sheet.getRange(i + 1, 7).setValue(fechaRespuesta);

      if (accion == 'aceptar') {
        var sheetDepa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');
        var departamentos = sheetDepa.getDataRange().getValues();
        for (var j = 1; j < departamentos.length; j++) {
          if (departamentos[j][0] == data[i][1]) {
            sheetDepa.getRange(j + 1, 5).setValue('Ocupado'); // Se asegura de actualizar el estado correctamente
            break;
          }
        }
        return ContentService.createTextOutput("Has aceptado la invitación. El departamento ha sido asignado.");
      } else {
        return ContentService.createTextOutput("Has rechazado la invitación.");
      }
    }
  }
  return ContentService.createTextOutput("Invitación no encontrada o ya gestionada.");
}

function cancelarInvitacion(idInvitacion) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Invitaciones');
  if (!sheet) return 'No hay hoja de invitaciones creada';

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == idInvitacion && data[i][4] == 'Pendiente') {
      sheet.getRange(i + 1, 5).setValue('Cancelada');
      return 'Invitación cancelada con éxito';
    }
  }
  return 'No se encontró una invitación pendiente con ese ID';
}