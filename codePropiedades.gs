function addPropiedad(nombre, direccion, referencia, cantidadDeptos, area, numPisos, serviciosStr, estacionamiento, fileData) {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w");
  var sheet = ss.getSheetByName('Propiedades');

  if (!sheet) {
    Logger.log("❌ ERROR: La hoja 'Propiedades' no existe.");
    return "❌ Error: La hoja 'Propiedades' no se encontró.";
  }

  var lastRow = sheet.getLastRow();
  
  sheet.insertRowAfter(lastRow);
  var newRow = lastRow + 1;

  var imageUrl = "Sin imagen";
    if (fileData) {
        var folder = DriveApp.getFolderById("1dsA0jIWyKE7AzMbjuiqEVtEj1GqxKwjI"); // Reemplaza con tu carpeta en Drive
        var blob = Utilities.newBlob(Utilities.base64Decode(fileData), "image/png", nombre + ".png");
        var file = folder.createFile(blob);
        imageUrl = file.getUrl();
    }

  sheet.getRange(newRow, 2).setValue(new Date());       // Fecha de registro
  sheet.getRange(newRow, 3).setValue(nombre);           // Nombre de la propiedad
  sheet.getRange(newRow, 4).setValue(direccion);        // Dirección
  sheet.getRange(newRow, 5).setValue(referencia);       // Referencia
  sheet.getRange(newRow, 6).setValue(cantidadDeptos);   // Cantidad de departamentos
  sheet.getRange(newRow, 7).setValue(area);             // Área en m²
  sheet.getRange(newRow, 8).setValue(serviciosStr);         // Número de pisos
  sheet.getRange(newRow, 9).setValue(numPisos); // Servicios (Luz, Agua)
  sheet.getRange(newRow, 10).setValue(estacionamiento); // Estacionamiento (Sí/No)
  sheet.getRange(newRow, 11).setValue(email);           // Email del usuario que registró
  sheet.getRange(newRow, 12).setValue(imageUrl);        // URL de la imagen en Google Drive

  Logger.log("📌 Datos insertados en columnas específicas en la fila: " + newRow);
  onFormSubmitPropietario();
  return "✅ Propiedad agregada correctamente en la fila correcta.";
}

// Obtener las propiedades de un usuario autenticado
function getPropiedades() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w");
  var sheet = ss.getSheetByName('Propiedades');

  if (!sheet) {
    Logger.log("❌ ERROR: La hoja 'Propiedades' no existe en esta hoja de cálculo.");
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var propiedades = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][10] == email) {
      propiedades.push({ 
        id: data[i][0],                  // ID de la propiedad
        nombre: data[i][2],              // Nombre de la propiedad
        direccion: data[i][3],           // Dirección
        referencia: data[i][4],          // Referencia
        cantidadDeptos: data[i][5],      // Cantidad de departamentos
        area: data[i][6],                // Área en m²
        servicios: data[i][7] ? data[i][7].split(",") : [],  // Servicios convertidos a array
        numPisos: data[i][8],            // Número de pisos
        estacionamiento: data[i][9],     // Estacionamiento (Sí/No)
        foto: data[i][11]
      });
    }
  }

  Logger.log("📌 Propiedades obtenidas desde la hoja correcta: " + JSON.stringify(propiedades));
  return propiedades;
}

function editPropiedad(propiedadId, nuevoNombre, nuevaDireccion) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Propiedades');
  var cell = sheet.getRange("E:E").createTextFinder(propiedadId).findNext(); // Suponiendo que la columna E contiene los IDs
  
  if (!cell) {
    return "❌ Error: No se encontró la propiedad.";
  }

  var rowIndex = cell.getRow(); // Obtener la fila real

  sheet.getRange(rowIndex, 2).setValue(nuevoNombre); // Columna B (Nombre)
  sheet.getRange(rowIndex, 3).setValue(nuevaDireccion); // Columna C (Dirección)

  return "✏️ Propiedad actualizada correctamente.";
}

function deletePropiedad(propiedadId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Propiedades');
  var cell = sheet.getRange("E:E").createTextFinder(propiedadId).findNext(); // Suponiendo que la columna E contiene los IDs

  if (!cell) {
    return "❌ Error: No se encontró la propiedad.";
  }

  var rowIndex = cell.getRow(); // Obtener la fila real
  sheet.deleteRow(rowIndex);

  return "❌ Propiedad eliminada correctamente.";
}