function doGet() {
  var usuario = verificarUsuario();

  if (usuario.esPropietario) {
    return HtmlService.createTemplateFromFile('gestionPagosPropietarios').evaluate();
  } else if (usuario.esInquilino) {
    return HtmlService.createTemplateFromFile('gestionPagosInquilino').evaluate();
  } else {
    return HtmlService.createTemplateFromFile('main').evaluate();
  }
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPropietarios = ss.getSheetByName('Propietarios');
  var sheetInquilinos = ss.getSheetByName('Inquilinos');

  if (!sheetPropietarios || !sheetInquilinos) {
    Logger.log("❌ ERROR: No se encontraron las hojas de usuarios.");
    return { email: email, esPropietario: false, esInquilino: false };
  }

  var dataPropietarios = sheetPropietarios.getDataRange().getValues();
  var dataInquilinos = sheetInquilinos.getDataRange().getValues();

  var esPropietario = false;
  var esInquilino = false;

  for (var i = 1; i < dataPropietarios.length; i++) {
    if (dataPropietarios[i][1].toString().trim().toLowerCase() === email) {
      esPropietario = true;
      break;
    }
  }

  if (!esPropietario) { // Solo si NO es propietario, verificamos si es inquilino
    for (var j = 1; j < dataInquilinos.length; j++) {
      if (dataInquilinos[j][1].toString().trim().toLowerCase() === email) {
        esInquilino = true;
        break;
      }
    }
  }

  Logger.log("🔍 Validación de usuario (" + email + "): " + 
              (esPropietario ? "✅ Propietario" : esInquilino ? "🏠 Inquilino" : "❌ No registrado"));

  return { email: email, esPropietario: esPropietario, esInquilino: esInquilino };
}