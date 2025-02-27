function obtenerHojaRecibosLuz() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RecibosLuz");

  if (!sheet) {
    Logger.log("❌ ERROR: No se encontró la hoja 'RecibosLuz'. Creando nueva hoja...");
    sheet = ss.insertSheet("RecibosLuz");
    sheet.appendRow([
      "Marca temporal", "ID Recibo", "CorreoPropietario", "ID Propiedad",
      "Precio por kWh", "Energía a facturar", "Cargo Fijo", "Mant. y Repo.",
      "Alumbrado Público", "Interés comp.", "Fecha de vencimiento", "Monto Total"
    ]);
  }
  return sheet;
}

function obtenerCorreoUsuario() {
  return Session.getActiveUser().getEmail();
}

function obtenerReciboLuzPorId(idRecibo) {
  var sheet = obtenerHojaRecibosLuz();
  var data = sheet.getDataRange().getValues();
  var correoUsuario = obtenerCorreoUsuario();

  for (var i = 1; i < data.length; i++) { // 🔹 Saltamos la primera fila (encabezados)
    if (data[i][1] === idRecibo && data[i][2] === correoUsuario) { // 🔹 Filtra por ID y usuario
      var recibo = {
        marcaTemporal: data[i][0],
        idRecibo: data[i][1],
        correoPropietario: data[i][2],
        idPropiedad: data[i][3],
        precioKWh: data[i][4],
        energiaFacturar: data[i][5],
        cargoFijo: data[i][6],
        mantRepo: data[i][7],
        alumbradoPublico: data[i][8],
        interesComp: data[i][9],
        fechaVencimiento: data[i][10],
        montoTotal: data[i][11]
      };

      Logger.log("✅ RECIBO ENCONTRADO: " + JSON.stringify(recibo));
      return JSON.stringify(recibo);
    }
  }

  Logger.log("❌ ERROR: Recibo no encontrado - ID: " + idRecibo);
  return JSON.stringify({ error: "Recibo no encontrado" });
}

function crearReciboLuz(idPropiedad, precioKWh, energiaFacturar, cargoFijo, mantRepo, alumbradoPublico, interesComp, fechaVencimiento, montoTotal) {
  var sheet = obtenerHojaRecibosLuz();
  var correoPropietario = obtenerCorreoUsuario(); // 🔹 Obtener el usuario actual
  
  var timestamp = new Date().toISOString(); // 🔹 Marca temporal
  var idRecibo = new Date().getTime() + "-" + Math.floor(Math.random() * 100000); // 🔹 ID único
  
  // Insertamos el nuevo recibo
  sheet.appendRow([
    timestamp, idRecibo, correoPropietario, idPropiedad, 
    precioKWh, energiaFacturar, cargoFijo, mantRepo,
    alumbradoPublico, interesComp, fechaVencimiento, montoTotal
  ]);

  Logger.log("✅ RECIBO DE LUZ CREADO: " + idRecibo);
  return "Recibo registrado con éxito.";
}

function obtenerRecibosLuz() {
  var sheet = obtenerHojaRecibosLuz();
  var data = sheet.getDataRange().getValues();
  var correoUsuario = obtenerCorreoUsuario();
  var recibos = [];

  for (var i = 1; i < data.length; i++) { // 🔹 Saltamos la fila de encabezado
    if (data[i][2] === correoUsuario) { // 🔹 Filtra por propietario
      recibos.push({
        marcaTemporal: data[i][0],
        idRecibo: data[i][1],
        correoPropietario: data[i][2],
        idPropiedad: data[i][3],
        precioKWh: data[i][4],
        energiaFacturar: data[i][5],
        cargoFijo: data[i][6],
        mantRepo: data[i][7],
        alumbradoPublico: data[i][8],
        interesComp: data[i][9],
        fechaVencimiento: data[i][10],
        montoTotal: data[i][11]
      });
    }
  }

  Logger.log("✅ RECIBOS DE LUZ OBTENIDOS: " + JSON.stringify(recibos));
  return JSON.stringify(recibos); // 🔹 Devuelve un JSON con los recibos
}

function editarReciboLuz(idRecibo, idPropiedad, precioKWh, energiaFacturar, cargoFijo, mantRepo, alumbradoPublico, interesComp, fechaVencimiento, montoTotal) {
  var sheet = obtenerHojaRecibosLuz();
  var data = sheet.getDataRange().getValues();
  var correoUsuario = obtenerCorreoUsuario();
  var reciboEncontrado = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === idRecibo && data[i][2] === correoUsuario) { // 🔹 Verifica ID y propietario
      sheet.getRange(i + 1, 4).setValue(idPropiedad);
      sheet.getRange(i + 1, 5).setValue(precioKWh);
      sheet.getRange(i + 1, 6).setValue(energiaFacturar);
      sheet.getRange(i + 1, 7).setValue(cargoFijo);
      sheet.getRange(i + 1, 8).setValue(mantRepo);
      sheet.getRange(i + 1, 9).setValue(alumbradoPublico);
      sheet.getRange(i + 1, 10).setValue(interesComp);
      sheet.getRange(i + 1, 11).setValue(fechaVencimiento);
      sheet.getRange(i + 1, 12).setValue(montoTotal);

      reciboEncontrado = true;
      break;
    }
  }

  if (reciboEncontrado) {
    Logger.log("✅ RECIBO DE LUZ EDITADO: " + idRecibo);
    return "Recibo actualizado con éxito.";
  } else {
    Logger.log("❌ ERROR: No se encontró el recibo o no pertenece al usuario.");
    return "Error: No se encontró el recibo o no tienes permisos.";
  }
}

function eliminarReciboLuz(idRecibo) {
  var sheet = obtenerHojaRecibosLuz();
  var data = sheet.getDataRange().getValues();
  var correoUsuario = obtenerCorreoUsuario();
  var reciboEncontrado = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === idRecibo && data[i][2] === correoUsuario) {
      sheet.deleteRow(i + 1); // 🔹 Borra la fila completa
      reciboEncontrado = true;
      break;
    }
  }

  if (reciboEncontrado) {
    Logger.log("✅ RECIBO DE LUZ ELIMINADO: " + idRecibo);
    return "Recibo eliminado con éxito.";
  } else {
    Logger.log("❌ ERROR: No se encontró el recibo o no pertenece al usuario.");
    return "Error: No se encontró el recibo o no tienes permisos.";
  }
}