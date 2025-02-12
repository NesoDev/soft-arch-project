function triggerAgregarTarifaFormulario(e) {
    var hojaModificada = e.source.getActiveSheet().getName();
    
    if (hojaModificada === "TarifasAgua") {
        Logger.log("🔄 Se detectó un cambio en 'TarifasAgua', actualizando formulario...");
        actualizarOpcionesTarifasEnFormulario();
    }
}

function triggerCalcularMontos(e) {
    var sheet = e.source.getSheetByName("Respuestas");
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
  
    if (sheet.getName() !== "Respuestas") return;

    var colLecturaAnterior = 9;
    var colLecturaActual = 10;
    var colCargoFijo = 11;
    var colConsumo = 12;
    var colMontoTotal = 15;

    var consumoValor = sheet.getRange(row, colConsumo).getValue();
    var montoTotalValor = sheet.getRange(row, colMontoTotal).getValue();

    if (!consumoValor || !montoTotalValor) {
        Logger.log(`🆕 Nueva respuesta detectada en la fila ${row}. Calculando montos...`);
        calcularMontosParaFila(row);
        return;
    }

    if ([colLecturaAnterior, colLecturaActual, colCargoFijo].includes(col)) {
        Logger.log(`🔄 Cambio detectado en la fila ${row}, columna ${col}. Recalculando montos...`);
        calcularMontosParaFila(row);
    }
}