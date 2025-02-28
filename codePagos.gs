function registrarPago(correoInquilino, tipoPago, mesPago, idRecibo, idPropiedad, idDepartamento, montoPagado, comprobante) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("Pagos");

  if (!sheetPagos) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Pagos'.");
    return "Error: No se encontró la hoja de pagos.";
  }

  var timestamp = new Date().toISOString();
  var idPago = generarIDPago();

  sheetPagos.appendRow([
    timestamp, idPago, correoInquilino, tipoPago, mesPago, idRecibo, idPropiedad, idDepartamento, montoPagado, "No Pagado", comprobante
  ]);

  Logger.log("✅ PAGO REGISTRADO: " + idPago);
  return "Pago registrado con éxito.";
}

function generarIDPago() {
  var timestamp = new Date().getTime();
  var randomNum = Math.floor(Math.random() * 10000);
  return timestamp + "-" + randomNum;
}

function obtenerRecibosPorPropiedad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRecibosLuz = ss.getSheetByName("RecibosLuz");
  var sheetRecibosAgua = ss.getSheetByName("RecibosAgua");
  var sheetDepartamentos = ss.getSheetByName("Departamentos");

  if (!sheetRecibosLuz || !sheetRecibosAgua || !sheetDepartamentos) {
    Logger.log("❌ ERROR: No se encontraron las hojas de recibos.");
    return JSON.stringify([]);
  }

  var correoInquilino = Session.getActiveUser().getEmail();
  var dataDepartamentos = sheetDepartamentos.getDataRange().getValues();
  var dataRecibosLuz = sheetRecibosLuz.getDataRange().getValues();
  var dataRecibosAgua = sheetRecibosAgua.getDataRange().getValues();

  var propiedadInquilino = "";
  for (var i = 1; i < dataDepartamentos.length; i++) {
    if (dataDepartamentos[i][5] === correoInquilino) { 
      propiedadInquilino = dataDepartamentos[i][2]; // ID de la propiedad
      break;
    }
  }

  if (!propiedadInquilino) {
    Logger.log("⚠️ No se encontró la propiedad del inquilino.");
    return JSON.stringify([]);
  }

  var recibosDisponibles = [];

  dataRecibosLuz.concat(dataRecibosAgua).forEach(function(row) {
    if (row[3] === propiedadInquilino) { // ID Propiedad en recibo
      recibosDisponibles.push({
        idRecibo: row[1],
        tipo: row[0].includes("RecibosLuz") ? "Luz" : "Agua",
        montoTotal: row[row.length - 1]
      });
    }
  });

  return JSON.stringify(recibosDisponibles);
}

function validarPago(idPago) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("Pagos");

  if (!sheetPagos) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Pagos'.");
    return "Error: No se encontró la hoja de pagos.";
  }

  var dataPagos = sheetPagos.getDataRange().getValues();
  for (var i = 1; i < dataPagos.length; i++) {
    if (dataPagos[i][1] === idPago) { 
      sheetPagos.getRange(i + 1, 10).setValue("Pagado"); // Cambia estado a "Pagado"
      Logger.log("✅ PAGO VALIDADO: " + idPago);
      return "Pago validado con éxito.";
    }
  }

  Logger.log("❌ ERROR: No se encontró el pago.");
  return "Error: No se encontró el pago.";
}

function listarPagos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("Pagos");

  if (!sheetPagos) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Pagos'.");
    return JSON.stringify([]);
  }

  var correoPropietario = Session.getActiveUser().getEmail();
  var dataPagos = sheetPagos.getDataRange().getValues();
  var pagosPropietario = [];

  for (var i = 1; i < dataPagos.length; i++) {
    if (dataPagos[i][6] === correoPropietario) { 
      pagosPropietario.push({
        idPago: dataPagos[i][1],
        correoInquilino: dataPagos[i][2],
        tipoPago: dataPagos[i][3],
        mesPago: dataPagos[i][4],
        montoPagado: dataPagos[i][8],
        estadoPago: dataPagos[i][9],
        comprobante: dataPagos[i][10]
      });
    }
  }

  Logger.log("✅ PAGOS DEL PROPIETARIO: " + JSON.stringify(pagosPropietario));
  return JSON.stringify(pagosPropietario);
}

function enviarRecordatorios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("Pagos");

  if (!sheetPagos) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Pagos'.");
    return "Error: No se encontró la hoja de pagos.";
  }

  var dataPagos = sheetPagos.getDataRange().getValues();
  var correosInquilinos = new Set();

  for (var i = 1; i < dataPagos.length; i++) {
    if (dataPagos[i][9] === "No Pagado") { 
      correosInquilinos.add(dataPagos[i][2]);
    }
  }

  correosInquilinos.forEach(function(correo) {
    MailApp.sendEmail({
      to: correo,
      subject: "Recordatorio de Pago Pendiente",
      body: "Estimado inquilino, tiene pagos pendientes en el sistema. Por favor, realice el pago lo antes posible."
    });
  });

  Logger.log("✅ RECORDATORIOS ENVIADOS");
  return "Recordatorios enviados con éxito.";
}


function listarPagosInquilino() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("Pagos");

  if (!sheetPagos) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Pagos'.");
    return JSON.stringify([]);
  }

  var correoInquilino = Session.getActiveUser().getEmail();
  var dataPagos = sheetPagos.getDataRange().getValues();
  var pagosInquilino = [];

  for (var i = 1; i < dataPagos.length; i++) {
    if (dataPagos[i][2] === correoInquilino) { 
      pagosInquilino.push({
        idPago: dataPagos[i][1],
        idDepartamento: dataPagos[i][7],
        montoPagado: dataPagos[i][8],
        estadoPago: dataPagos[i][9],
        mesPago: dataPagos[i][4]
      });
    }
  }

  Logger.log("✅ PAGOS DEL INQUILINO: " + JSON.stringify(pagosInquilino));
  return JSON.stringify(pagosInquilino);
}


function subirComprobante(idPago, base64String) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPagos = ss.getSheetByName("Pagos");

  if (!sheetPagos) {
    Logger.log("❌ ERROR: No se encontró la hoja 'Pagos'.");
    return "Error: No se encontró la hoja de pagos.";
  }

  var folder = DriveApp.getFolderById("19L3YqJOOyz5ypPTFGU4wNX1TcSeTLEm8");
  var fileName = "comprobante_" + idPago + ".png";
  var blob = Utilities.newBlob(Utilities.base64Decode(base64String), "image/png", fileName);
  var file = folder.createFile(blob);
  var fileUrl = file.getUrl();

  var dataPagos = sheetPagos.getDataRange().getValues();
  for (var i = 1; i < dataPagos.length; i++) {
    if (dataPagos[i][1] === idPago) { 
      sheetPagos.getRange(i + 1, 11).setValue(fileUrl); // Guarda el URL del comprobante
      Logger.log("✅ COMPROBANTE GUARDADO: " + fileUrl);
      return "Comprobante subido con éxito.";
    }
  }

  Logger.log("❌ ERROR: No se encontró el pago.");
  return "Error: No se encontró el pago.";
}