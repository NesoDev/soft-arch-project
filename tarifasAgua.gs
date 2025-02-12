function agregarNuevaTarifa(nombreTarifa) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TarifasAgua');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === nombreTarifa) {
            Logger.log(`⚠️ La tarifa "${nombreTarifa}" ya existe.`);
            return;
        }
    }

    sheet.appendRow([nombreTarifa, "", "", "", "", ""]);

    Logger.log(`✅ Tarifa "${nombreTarifa}" creada exitosamente en la hoja.`);
}

function agregarRangoATarifa(nombreTarifa, rangoInferior, rangoSuperior, tarifaAgua, tarifaAlcantarillado) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TarifasAgua');
    var data = sheet.getDataRange().getValues();
    
    var idsExistentes = [];
    var filaVacia = -1;
    var tarifaExiste = false;

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === nombreTarifa) {
            tarifaExiste = true;

            if (data[i][1] === "") {
                filaVacia = i + 1;
            } else {
                idsExistentes.push(data[i][1]);

                if (data[i][2] == rangoInferior && data[i][3] == rangoSuperior) {
                    Logger.log(`❌ Error: El rango ${rangoInferior} - ${rangoSuperior} m³ ya existe en la tarifa "${nombreTarifa}".`);
                    return;
                }
            }
        }
    }

    if (!tarifaExiste) {
        Logger.log(`❌ Error: La tarifa "${nombreTarifa}" no existe. Debes crearla antes.`);
        return;
    }

    var nuevoId = 1;
    while (idsExistentes.includes(nuevoId)) {
        nuevoId++;
    }

    if (filaVacia !== -1) {
        sheet.getRange(filaVacia, 2).setValue(nuevoId);
        sheet.getRange(filaVacia, 3).setValue(rangoInferior);
        sheet.getRange(filaVacia, 4).setValue(rangoSuperior);
        sheet.getRange(filaVacia, 5).setValue(tarifaAgua);
        sheet.getRange(filaVacia, 6).setValue(tarifaAlcantarillado);

        Logger.log(`✅ Se sobrescribió la fila vacía de "${nombreTarifa}" con ID ${nuevoId}: ${rangoInferior} - ${rangoSuperior} m³.`);
    } else {
        sheet.appendRow([nombreTarifa, nuevoId, rangoInferior, rangoSuperior, tarifaAgua, tarifaAlcantarillado]);
        Logger.log(`✅ Nuevo rango añadido a "${nombreTarifa}" con ID ${nuevoId}: ${rangoInferior} - ${rangoSuperior} m³.`);
    }
}

function eliminarRangoTarifa(nombreTarifa, idRango) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TarifasAgua');
    var data = sheet.getDataRange().getValues();

    for (var i = data.length - 1; i > 0; i--) {
        if (data[i][0] === nombreTarifa && data[i][1] == idRango) {
            sheet.deleteRow(i + 1);
            Logger.log(`✅ Rango con ID ${idRango} eliminado de "${nombreTarifa}".`);
            return;
        }
    }
    Logger.log(`⚠️ No se encontró el rango con ID ${idRango} en "${nombreTarifa}".`);
}

function eliminarTarifaCompleta(nombreTarifa) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TarifasAgua');
    var data = sheet.getDataRange().getValues();

    for (var i = data.length - 1; i > 0; i--) {
        if (data[i][0] === nombreTarifa) {
            sheet.deleteRow(i + 1);
        }
    }

    Logger.log(`✅ Tarifa completa "${nombreTarifa}" eliminada.`);
}