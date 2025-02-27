function doGet() {
  return HtmlService.createTemplateFromFile('main').evaluate();
}

function cargarHtml(nombreVista) {
  return include(nombreVista);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUsuarioActual() {
  var email = Session.getActiveUser().getEmail();
  Logger.log("🟢 Usuario detectado: " + email);
  return email;
}

function verificarUsuario() {
  var email = getUsuarioActual().trim().toLowerCase();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Propietarios');

  if (!sheet) {
    Logger.log("❌ ERROR: La hoja 'Propietarios' no existe.");
    return { email: email, esPropietario: false };
  }

  var data = sheet.getDataRange().getValues();
  var esPropietario = false;

  for (var i = 1; i < data.length; i++) {
    var emailPropietario = data[i][1].toString().trim().toLowerCase(); // Normalizar el email
    if (emailPropietario === email) {
      esPropietario = true;
      break;
    }
  }

  Logger.log("🔍 Validación de usuario (" + email + "): " + (esPropietario ? "✅ Propietario registrado" : "❌ No registrado"));
  return { email: email, esPropietario: esPropietario };
}