function actualizarOpcionesTarifasEnFormulario() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TarifasAgua');
    var data = sheet.getDataRange().getValues();
    
    var form = FormApp.openById("1yAGRVMTfgIyQZVZ4Wemu-onkaxadWlbCb5eb477tfSE");
    var items = new Set();

    for (var i = 1; i < data.length; i++) {
        if (data[i][0]) {
            items.add(data[i][0]);
        }
    }

    var opciones = Array.from(items);

    var preguntas = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
    for (var i = 0; i < preguntas.length; i++) {
        var pregunta = preguntas[i].asMultipleChoiceItem();
        if (pregunta.getTitle().includes("Seleccione la categoría de tarifa de agua:")) {
            pregunta.setChoices(opciones.map(opcion => pregunta.createChoice(opcion)));
            Logger.log("✅ Opciones del formulario actualizadas correctamente.");
            return;
        }
    }

    Logger.log("⚠️ No se encontró la pregunta en el formulario.");
}