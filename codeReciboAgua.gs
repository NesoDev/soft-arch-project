function obtenerReciboPorId(idRecibo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RecibosAgua");

  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'RecibosAgua'.");
    return JSON.stringify({ error: "No se encontró la hoja 'RecibosAgua'." });
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idRecibo) { // Verifica si el ID del recibo coincide
      var recibo = {
        idRecibo: data[i][0],
        idPropietario: data[i][1],
        idPropiedad: data[i][2],
        tarifaAplicada: data[i][3],
        mesFacturado: data[i][4],
        fechaEmision: data[i][5],
        fechaVencimiento: data[i][6],
        consumo: data[i][7],
        cargoFijo: data[i][8],
        aguaPotable: data[i][9],
        alcantarillado: data[i][10],
        montoTotal: data[i][11]
      };

      Logger.log("✅ RECIBO ENCONTRADO: " + JSON.stringify(recibo));
      return JSON.stringify(recibo); // 🔹 Devuelve el objeto en formato JSON
    }
  }

  Logger.log("❌ ERROR: No se encontró el recibo con ID " + idRecibo);
  return JSON.stringify({ error: "Recibo no encontrado" }); // 🔹 Siempre devuelve JSON válido
}

function obtenerPropiedadesPorUsuario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Propiedades");

  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Propiedades'.");
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var correoUsuario = Session.getActiveUser().getEmail();
  var propiedades = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][10] === correoUsuario) { // Filtra por propietario
      propiedades.push({
        idPropiedad: data[i][0], // ID de la propiedad
        nombre: data[i][2] // Nombre de la propiedad
      });
    }
  }

  Logger.log("✅ PROPIEDADES OBTENIDAS: " + JSON.stringify(propiedades));
  return propiedades;
}

function obtenerCorreoUsuario() {
  return Session.getActiveUser().getEmail();
}

function getTarifasAgua() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("TarifasAgua");
  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'TarifasAgua'.");
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var tarifasSet = new Set(); // Usamos un Set para evitar duplicados

  for (var i = 1; i < data.length; i++) {
    var categoria = data[i][0].trim(); // Eliminamos espacios innecesarios
    if (categoria) {
      tarifasSet.add(categoria); // Agregamos solo si no está repetido
    }
  }

  var tarifas = Array.from(tarifasSet); // Convertimos el Set en un Array

  Logger.log("✅ TARIFAS ÚNICAS OBTENIDAS: " + JSON.stringify(tarifas));
  return tarifas;
}

function crearReciboAgua(idPropietario, idPropiedad, tarifaAplicada, mesFacturado, fechaEmision, fechaVencimiento, consumo, cargoFijo, aguaPotable, alcantarillado, montoTotal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RecibosAgua");
  
  if (!sheet) {
    sheet = ss.insertSheet("RecibosAgua"); // Si no existe, la crea
    sheet.appendRow(["ID Recibo", "ID Propietario", "ID Propiedad", "Tarifa Aplicada", "Mes Facturado", "Fecha Emisión", "Fecha Vencimiento", "Consumo (m³)", "Cargo Fijo (S/.)", "Agua Potable (S/.)", "Alcantarillado (S/.)", "Monto Total (S/.)"]);
  }

  // Generar un ID único para el recibo
  var timestamp = new Date().getTime(); 
  var randomNum = Math.floor(Math.random() * (9999999 - 1000000) + 1000000);
  var idRecibo = timestamp + "-" + randomNum;

  // Validación de datos (Evitamos valores vacíos o negativos)
  if (!idPropietario || !idPropiedad || !tarifaAplicada || !mesFacturado || !fechaEmision || !fechaVencimiento || consumo < 0 || cargoFijo < 0 || aguaPotable < 0 || alcantarillado < 0 || montoTotal < 0) {
    Logger.log("❌ ERROR: Datos inválidos en el recibo.");
    return "Error: Datos inválidos. Verifique los valores ingresados.";
  }

  // Insertar los datos en la hoja
  sheet.appendRow([idRecibo, idPropietario, idPropiedad, tarifaAplicada, mesFacturado, fechaEmision, fechaVencimiento, consumo, cargoFijo, aguaPotable, alcantarillado, montoTotal]);

  Logger.log("✅ RECIBO REGISTRADO: " + idRecibo);
  return "Recibo registrado con éxito.";
}

function obtenerRecibosAgua(idPropietario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RecibosAgua");

  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'RecibosAgua'.");
    return JSON.stringify([]); // 🔹 Devuelve un array vacío en formato JSON
  }

  var data = sheet.getDataRange().getValues();
  var recibos = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === idPropietario) { // Filtra por propietario
      recibos.push({
        idRecibo: data[i][0],
        idPropietario: data[i][1],
        idPropiedad: data[i][2],
        tarifaAplicada: data[i][3],
        mesFacturado: data[i][4],
        fechaEmision: data[i][5],
        fechaVencimiento: data[i][6],
        consumo: data[i][7],
        cargoFijo: data[i][8],
        aguaPotable: data[i][9],
        alcantarillado: data[i][10],
        montoTotal: data[i][11]
      });
    }
  }

  Logger.log("✅ RECIBOS ENCONTRADOS: " + JSON.stringify(recibos));
  return JSON.stringify(recibos); // 🔹 Asegura que el retorno sea siempre un JSON válido
}

function editarReciboAgua(idRecibo, idPropietario, nuevaPropiedad, nuevaTarifa, nuevoMesFacturado, nuevaFechaEmision, nuevaFechaVencimiento, nuevoConsumo, nuevoCargoFijo, nuevoAguaPotable, nuevoAlcantarillado, nuevoMontoTotal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RecibosAgua");

  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'RecibosAgua'.");
    return "Error: No se encontró la hoja de recibos.";
  }

  var data = sheet.getDataRange().getValues();
  var reciboEncontrado = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idRecibo && data[i][1] === idPropietario) {
      sheet.getRange(i + 1, 3).setValue(nuevaPropiedad); // ✅ Actualiza la propiedad
      sheet.getRange(i + 1, 4).setValue(nuevaTarifa);
      sheet.getRange(i + 1, 5).setValue(nuevoMesFacturado);
      sheet.getRange(i + 1, 6).setValue(nuevaFechaEmision);
      sheet.getRange(i + 1, 7).setValue(nuevaFechaVencimiento);
      sheet.getRange(i + 1, 8).setValue(nuevoConsumo);
      sheet.getRange(i + 1, 9).setValue(nuevoCargoFijo);
      sheet.getRange(i + 1, 10).setValue(nuevoAguaPotable);
      sheet.getRange(i + 1, 11).setValue(nuevoAlcantarillado);
      sheet.getRange(i + 1, 12).setValue(nuevoMontoTotal);

      reciboEncontrado = true;
      break;
    }
  }

  if (reciboEncontrado) {
    Logger.log("✅ RECIBO EDITADO: " + idRecibo);
    return "Recibo actualizado con éxito.";
  } else {
    Logger.log("❌ ERROR: No se encontró el recibo con ID " + idRecibo);
    return "Error: No se encontró el recibo.";
  }
}

function eliminarReciboAgua(idRecibo, idPropietario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RecibosAgua");

  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'RecibosAgua'.");
    return "Error: No se encontró la hoja de recibos.";
  }

  var data = sheet.getDataRange().getValues();
  var reciboEncontrado = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idRecibo && data[i][1] === idPropietario) { // Verifica que el recibo pertenece al propietario
      sheet.deleteRow(i + 1); // Borra la fila completa
      reciboEncontrado = true;
      break;
    }
  }

  if (reciboEncontrado) {
    Logger.log("✅ RECIBO ELIMINADO: " + idRecibo);
    return "Recibo eliminado con éxito.";
  } else {
    Logger.log("❌ ERROR: No se encontró el recibo con ID " + idRecibo + " para este propietario.");
    return "Error: No se encontró el recibo o no pertenece al propietario.";
  }
}
