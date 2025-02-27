
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

  sheet.getRange(newRow, 2).setValue(new Date());       
  sheet.getRange(newRow, 3).setValue(nombre);           
  sheet.getRange(newRow, 4).setValue(direccion);       
  sheet.getRange(newRow, 5).setValue(referencia);     
  sheet.getRange(newRow, 6).setValue(cantidadDeptos);   
  sheet.getRange(newRow, 7).setValue(area);            
  sheet.getRange(newRow, 8).setValue(serviciosStr);         
  sheet.getRange(newRow, 9).setValue(numPisos); 
  sheet.getRange(newRow, 10).setValue(estacionamiento); 
  sheet.getRange(newRow, 11).setValue(email);           
  sheet.getRange(newRow, 12).setValue(imageUrl);        

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
        id: data[i][0],                  
        nombre: data[i][2],              
        direccion: data[i][3],           
        referencia: data[i][4],          
        cantidadDeptos: data[i][5],      
        area: data[i][6],               
        servicios: data[i][7] ? data[i][7].split(",") : [], 
        numPisos: data[i][8],            
        estacionamiento: data[i][9],     
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