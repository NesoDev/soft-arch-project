function actualizarPermisosEncuestados() {
  var hojaPropietarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios_Propietarios");

  var formIds = ["1Bqba8OCw1fSXJksvaxIBmcnEquM8VpHI64dfbB8Rats"];
  var datos = hojaPropietarios.getDataRange().getValues();
  var correosPropietarios = [];

  for (var i = 1; i < datos.length; i++) { 
    if (datos[i][2] && typeof datos[i][2] === "string") {
      correosPropietarios.push(datos[i][2].trim());
    }
  }

  formIds.forEach(function(formId) {
    try {
      var formFile = DriveApp.getFileById(formId);
      var permisosActuales = formFile.getEditors().map(editor => editor.getEmail());

      correosPropietarios.forEach(function(correo) {
        if (!permisosActuales.includes(correo) && correo.includes("@")) {
          formFile.addViewer(correo);
        }
      });

      Logger.log("Se actualizaron los permisos para el formulario " + formId + " con los correos: " + correosPropietarios.join(", "));
    } catch (error) {
      Logger.log("Error al actualizar permisos para el formulario " + formId + ": " + error.message);
    }
  });
}

/*function actualizarPermisosEncuestados() {
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
}*/