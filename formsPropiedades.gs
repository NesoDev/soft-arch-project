function actualizarPermisosEncuestados() {
  var hojaPropietarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios_Propietarios"); // Hoja donde están los propietarios
  var formId = "1mOBjbe2FDVPVqMBIS35gr4mCJVHrYz78IXnoMr7jYz4";
  var formFile = DriveApp.getFileById(formId);
  var datos = hojaPropietarios.getDataRange().getValues();
  var correosPropietarios = [];

  for (var i = 1; i < datos.length; i++) { 
    if (datos[i][1]) {
      correosPropietarios.push(datos[i][1]);
    }
  }

  var permisosActuales = formFile.getEditors();
  for (var i = 0; i < permisosActuales.length; i++) {
    formFile.removeEditor(permisosActuales[i]);
  }

  for (var j = 0; j < correosPropietarios.length; j++) {
    formFile.addViewer(correosPropietarios[j]);
  }

  Logger.log("Se actualizaron los permisos de encuestado para: " + correosPropietarios.join(", "));
}