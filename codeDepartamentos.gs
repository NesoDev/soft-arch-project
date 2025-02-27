function addDepartamento(nombre, idPropiedad, email, alquiler, diaPago, garantia) {
  var ss = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w");
  var sheet = ss.getSheetByName('Departamentos');

  if (!sheet) {
    Logger.log("❌ ERROR: La hoja 'Departamentos' no existe.");
    return "❌ Error: La hoja 'Departamentos' no se encontró.";
  }

  // Generar ID único para el departamento
  var timestamp = new Date();
  var randomNum = Math.floor(Math.random() * (99999 - 10000 + 1)) + 10000;
  var idDepartamento = timestamp.getFullYear().toString() +
                       ('0' + (timestamp.getMonth() + 1)).slice(-2) +
                       ('0' + timestamp.getDate()).slice(-2) +
                       ('0' + timestamp.getHours()).slice(-2) +
                       ('0' + timestamp.getMinutes()).slice(-2) +
                       ('0' + timestamp.getSeconds()).slice(-2) +
                       randomNum;

  // Insertar nueva fila
  var lastRow = sheet.getLastRow();
  sheet.insertRowAfter(lastRow);
  
  var nuevaFila = [timestamp, nombre, idPropiedad, email, idDepartamento, "Desocupado", alquiler, diaPago, garantia];
  sheet.getRange(lastRow + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);

  Logger.log("✅ Departamento agregado correctamente con ID: " + idDepartamento);
  return "✅ Departamento agregado correctamente.";
}

function getDepartamentos() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w");
  var sheet = ss.getSheetByName('Departamentos');

  if (!sheet) {
    Logger.log("❌ ERROR: La hoja 'Departamentos' no existe.");
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var departamentos = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][3] == email) {
      departamentos.push({ 
        id: data[i][4], 
        nombre: data[i][1], 
        idPropiedad: data[i][2], 
        alquiler: data[i][6],
        diaPago: data[i][7],
        garantia: data[i][8],
        estado: data[i][5]
      });
    }
  }

  return departamentos;
}

function editDepartamento(idDepartamento, nuevoNombre, nuevoAlquiler, nuevoDiaPago, nuevaGarantia, nuevoEstado) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');
  var cell = sheet.getRange("E:E").createTextFinder(idDepartamento).findNext();

  if (!cell) {
    return "❌ Error: No se encontró el departamento.";
  }

  var rowIndex = cell.getRow();
  
  sheet.getRange(rowIndex, 2).setValue(nuevoNombre);
  sheet.getRange(rowIndex, 6).setValue(nuevoEstado);
  sheet.getRange(rowIndex, 7).setValue(nuevoAlquiler);
  sheet.getRange(rowIndex, 8).setValue(nuevoDiaPago);
  sheet.getRange(rowIndex, 9).setValue(nuevaGarantia);

  return "✏️ Departamento actualizado correctamente.";
}

function deleteDepartamento(idDepartamento) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departamentos');
  var cell = sheet.getRange("E:E").createTextFinder(idDepartamento).findNext();

  if (!cell) {
    return "❌ Error: No se encontró el departamento.";
  }

  var rowIndex = cell.getRow();
  sheet.deleteRow(rowIndex);

  return "❌ Departamento eliminado correctamente.";
}
