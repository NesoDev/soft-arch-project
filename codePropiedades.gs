
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
        var rootFolder = DriveApp.getRootFolder(); 
        var folderName= "Fotos Rendo - "+email; 
        var folders = rootFolder.getFoldersByName(folderName);
        var userFolder;

        if (folders.hasNext()) {
          userFolder = folders.next(); 
        } else {
          userFolder = rootFolder.createFolder(folderName); 
        }
        
        var blob = Utilities.newBlob(Utilities.base64Decode(fileData), "image/png", nombre + ".png");
        var file = userFolder.createFile(blob);
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

function editPropiedad(propiedadId, nuevoNombre, nuevaDireccion, nuevaReferencia, nuevaCantidadDeptos, nuevaAreaProp, nuevosServicios, nuevoNumPisosProp, nuevoEstacionamiento, nuevaFotoProp) {
  var email = Session.getActiveUser().getEmail();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Propiedades');
  var cell = sheet.getRange("A:A").createTextFinder(propiedadId).findNext(); 
  
  if (!cell) {
    return "❌ Error: No se encontró la propiedad.";
  }

  var rowIndex = cell.getRow(); 

  sheet.getRange(rowIndex, 3).setValue(nuevoNombre); 
  sheet.getRange(rowIndex, 4).setValue(nuevaDireccion); 
  sheet.getRange(rowIndex, 5).setValue(nuevaReferencia); 
  sheet.getRange(rowIndex, 6).setValue(nuevaCantidadDeptos); 
  sheet.getRange(rowIndex, 7).setValue(nuevaAreaProp); 
  sheet.getRange(rowIndex, 8).setValue(nuevosServicios); 
  sheet.getRange(rowIndex, 9).setValue(nuevoNumPisosProp); 
  sheet.getRange(rowIndex, 10).setValue(nuevoEstacionamiento); 

  var imageUrl = sheet.getRange(rowIndex, 12).getValue(); 
  if (nuevaFotoProp) {
    var rootFolder = DriveApp.getRootFolder();
    var folderName = "Fotos Rendo - "+email; 
    var folders = rootFolder.getFoldersByName(folderName);
    var userFolder;

    if (folders.hasNext()) {
      userFolder = folders.next(); 
    } else {
      userFolder = rootFolder.createFolder(folderName); 
    }
    
    var blob = Utilities.newBlob(Utilities.base64Decode(nuevaFotoProp), "image/png", nuevoNombre + ".png");
    var file = userFolder.createFile(blob);
    imageUrl = file.getUrl();
  }

  sheet.getRange(rowIndex, 12).setValue(imageUrl);

  return "✏️ Propiedad actualizada correctamente.";
}


function deletePropiedad(propiedadId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Propiedades');
  var cell = sheet.getRange("A:A").createTextFinder(propiedadId).findNext(); 

  if (!cell) {
    return "❌ Error: No se encontró la propiedad.";
  }

  var rowIndex = cell.getRow(); 
  sheet.deleteRow(rowIndex);

  return "❌ Propiedad eliminada correctamente.";
}


function getPropiedadPorId(propiedadId) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propiedades");
  var datos = hoja.getDataRange().getValues();
  
  for (var i = 1; i < datos.length; i++) { 
    if (datos[i][0] == propiedadId) { 
      return {
        id: datos[i][0],
        nombre: datos[i][2],
        direccion: datos[i][3],
        referencia: datos[i][4],
        cantidadDeptos: datos[i][5],
        area: datos[i][6],
        servicios: datos[i][7].split(", "), 
        numPisos: datos[i][8],
        estacionamiento: datos[i][9],
        foto: datos[i][11] 
      };
    }
  }
  return null; 
}

