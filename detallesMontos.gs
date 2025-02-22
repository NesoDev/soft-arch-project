function calcularMontosParaFila(row) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRespuestas = ss.getSheetByName("RecibosAgua");
    var sheetTarifas = ss.getSheetByName("TarifasAgua");

    var dataTarifas = sheetTarifas.getDataRange().getValues();
    var filaRespuestas = sheetRespuestas.getRange(row, 1, 1, sheetRespuestas.getLastColumn()).getValues()[0];

    if (!filaRespuestas) return;

    var indicesRespuestas = {
        "Tarifa": 1, "Lectura Anterior": 8, "Lectura Actual": 9, "Cargo Fijo": 10,
        "Consumo m3": 11, "Agua Potable S/": 12, "Alcantarillado S/": 13, "Monto Total S/": 14
    };

    var tarifaSeleccionada = filaRespuestas[indicesRespuestas["Tarifa"]];
    var lecturaAnterior = filaRespuestas[indicesRespuestas["Lectura Anterior"]];
    var lecturaActual = filaRespuestas[indicesRespuestas["Lectura Actual"]];
    var cargoFijo = filaRespuestas[indicesRespuestas["Cargo Fijo"]];

    if (!tarifaSeleccionada || isNaN(lecturaAnterior) || isNaN(lecturaActual)) return;

    var consumo = lecturaActual - lecturaAnterior;
    var tarifaRangos = [];

    for (var j = 1; j < dataTarifas.length; j++) {
        if (dataTarifas[j][0] === tarifaSeleccionada) {
            tarifaRangos.push({
                rangoInferior: dataTarifas[j][2],
                rangoSuperior: dataTarifas[j][3],
                aguaPotable: dataTarifas[j][4],
                alcantarillado: dataTarifas[j][5]
            });
        }
    }

    if (tarifaRangos.length === 0) {
        Logger.log(`⚠️ No se encontró tarifa para ${tarifaSeleccionada}`);
        return;
    }

    tarifaRangos.sort((a, b) => a.rangoInferior - b.rangoInferior);

    var consumoRestante = consumo;
    var totalAgua = 0;
    var totalAlcantarillado = 0;

    for (var t = 0; t < tarifaRangos.length; t++) {
        var tramo = tarifaRangos[t];

        if (consumoRestante <= 0) break;

        var consumoEnEsteTramo = Math.min(consumoRestante, tramo.rangoSuperior - tramo.rangoInferior);
        totalAgua += consumoEnEsteTramo * tramo.aguaPotable;
        totalAlcantarillado += consumoEnEsteTramo * tramo.alcantarillado;

        consumoRestante -= consumoEnEsteTramo;
    }

    var montoTotal = totalAgua + totalAlcantarillado + cargoFijo;

    sheetRespuestas.getRange(row, indicesRespuestas["Consumo m3"] + 1).setValue(consumo);
    sheetRespuestas.getRange(row, indicesRespuestas["Agua Potable S/"] + 1).setValue(totalAgua);
    sheetRespuestas.getRange(row, indicesRespuestas["Alcantarillado S/"] + 1).setValue(totalAlcantarillado);
    sheetRespuestas.getRange(row, indicesRespuestas["Monto Total S/"] + 1).setValue(montoTotal);

    Logger.log(`✅ Montos recalculados para fila ${row}: Consumo ${consumo} m³ | Agua S/.${totalAgua} | Alcantarillado S/.${totalAlcantarillado} | Total S/.${montoTotal}`);
}