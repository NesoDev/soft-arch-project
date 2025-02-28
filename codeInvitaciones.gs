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

function obtenerInvitaciones() {
  Logger.log("🚀 FUNCIÓN obtenerInvitaciones() SE ESTÁ EJECUTANDO!");
  
  var ss = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w");
  var sheetInv = ss.getSheetByName('Invitaciones');

  if (!sheetInv) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Invitaciones'.");
    return [{ error: "No se encontró la hoja de invitaciones" }];
  }

  var data = sheetInv.getDataRange().getValues();
  if (!data || data.length <= 1) {
    Logger.log("⚠️ ADVERTENCIA: No hay invitaciones registradas.");
    return [{ mensaje: "No hay invitaciones registradas." }];
  }

  var usuarioAutenticado = Session.getActiveUser().getEmail().trim().toLowerCase();
  Logger.log("📌 Usuario autenticado: " + usuarioAutenticado);

  var invitaciones = [];

  for (var i = 1; i < data.length; i++) {
    var propietario = (data[i][2] || "").trim().toLowerCase();

    if (propietario === usuarioAutenticado) {
      invitaciones.push({
        id: data[i][0] || '',
        departamento: data[i][1] || '',
        propietario: propietario,
        inquilino: data[i][3] || '',
        estado: data[i][4] || '',
        fechaEnvio: data[i][5] || ''
      });
    }
  }

  if (invitaciones.length === 0) {
    Logger.log("⚠️ NO HAY INVITACIONES PARA ESTE USUARIO.");
    return [{ mensaje: "No tienes invitaciones registradas." }];
  }

  Logger.log("✅ INVITACIONES ENCONTRADAS: " + JSON.stringify(invitaciones));
  return JSON.stringify(invitaciones);
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