function addDepartamento(nombre, idPropiedad, alquiler, diaPago, garantia) {
  var email = Session.getActiveUser().getEmail(); // Obtener correo aquí en el backend
  var ss = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w");
  var sheet = ss.getSheetByName('Departamentos');

  if (!sheet) {
    Logger.log("❌ ERROR: La hoja 'Departamentos' no existe.");
    return "❌ Error: La hoja 'Departamentos' no se encontró.";
  }

  // Generar un ID único para el departamento
  var timestamp = new Date();
  var randomNum = Math.floor(Math.random() * 90000) + 10000;
  var idDepartamento = timestamp.getFullYear().toString() +
                       ('0' + (timestamp.getMonth() + 1)).slice(-2) +
                       ('0' + timestamp.getDate()).slice(-2) +
                       ('0' + timestamp.getHours()).slice(-2) +
                       ('0' + timestamp.getMinutes()).slice(-2) +
                       ('0' + timestamp.getSeconds()).slice(-2) +
                       randomNum;

  // Nuevo campo: Correo del inquilino (inicialmente vacío)
  var correoInquilino = "";
  
  // Estado inicial (Siempre "Desocupado" al crearse)
  var estado = "Desocupado";

  // Obtener la última fila con datos
  var lastRow = sheet.getLastRow();
  sheet.insertRowAfter(lastRow);
  
  // Datos a insertar en la hoja de cálculo
  var nuevaFila = [
    timestamp,    // Marca temporal
    nombre,       // Nombre del departamento
    idPropiedad,  // ID de la propiedad a la que pertenece
    email,        // Correo del usuario que lo creó
    idDepartamento, // ID del departamento generado
    correoInquilino, // Correo del inquilino asociado (vacío por defecto)
    estado,       // Estado (Desocupado por defecto)
    alquiler,     // Monto del alquiler
    diaPago,      // Día de pago
    garantia      // Monto de la garantía
  ];

  // Insertar datos en la hoja de cálculo
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
    if (data[i][3] == email) { // Filtrar solo los departamentos del usuario actual
      
      var correoInquilino = data[i][5]; // Columna F
      var estado = data[i][6]; // Columna G

      // Si el estado está vacío por algún error, recalcularlo
      if (!estado) {
        estado = correoInquilino ? "Ocupado" : "Desocupado";
        sheet.getRange(i + 1, 7).setValue(estado); // Guardar el estado corregido en la hoja
      }

      departamentos.push({ 
        id: data[i][4],         // ID del departamento
        nombre: data[i][1],     // Nombre del departamento
        idPropiedad: data[i][2],// ID de la propiedad a la que pertenece
        alquiler: data[i][7],   // Monto del alquiler
        diaPago: data[i][8],    // Día de pago
        garantia: data[i][9],   // Monto de la garantía
        estado: estado,         // Estado ahora viene desde la hoja
        correoInquilino: correoInquilino // Nuevo campo
      });
    }
  }

  return departamentos;
}

function editDepartamento(idDepartamento, nuevoNombre, nuevoAlquiler, nuevoDiaPago, nuevaGarantia, nuevaPropiedad) {
  var sheet = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w").getSheetByName('Departamentos');
  var cell = sheet.getRange("E:E").createTextFinder(idDepartamento).findNext();

  if (!cell) {
    return "❌ Error: No se encontró el departamento.";
  }

  var rowIndex = cell.getRow();

  // Actualizar los valores permitidos
  if (nuevoNombre) sheet.getRange(rowIndex, 2).setValue(nuevoNombre);
  if (nuevaPropiedad) sheet.getRange(rowIndex, 3).setValue(nuevaPropiedad);
  if (nuevoAlquiler) sheet.getRange(rowIndex, 8).setValue(nuevoAlquiler);
  if (nuevoDiaPago) sheet.getRange(rowIndex, 9).setValue(nuevoDiaPago);
  if (nuevaGarantia) sheet.getRange(rowIndex, 10).setValue(nuevaGarantia);

  return "✏️ Departamento actualizado correctamente.";
}

function deleteDepartamento(idDepartamento) {
  var sheet = SpreadsheetApp.openById("1Y1DybjwnF4MmCAHiydnr-fHy2DE7FTekctV0mIWAu8w").getSheetByName('Departamentos');
  var cell = sheet.getRange("E:E").createTextFinder(idDepartamento).findNext();

  if (!cell) {
    return "❌ Error: No se encontró el departamento.";
  }

  var rowIndex = cell.getRow();
  sheet.deleteRow(rowIndex);

  return "❌ Departamento eliminado correctamente.";
}
