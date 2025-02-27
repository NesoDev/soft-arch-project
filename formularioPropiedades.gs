/**
 * Nueva función pero sin el parametro e
 */

const NOMBRE = 'Propiedades';
function onFormSubmitPropietario() { 
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE);
  
  if (!hoja) {
    Logger.log("Error: No se encontró la hoja con el nombre '" + NOMBRE + "'");
    return;
  }

  var ultimaFila = hoja.getLastRow();
  
  if (ultimaFila < 2) { // Suponiendo que la primera fila es de encabezados
    Logger.log("No hay suficientes datos en la hoja.");
    return;
  }

  var correo = hoja.getRange(ultimaFila, 11).getValue(); // Columna 2 (ajusta según la estructura real)

  if (correo && typeof correo === "string" && correo.includes("@")) {
    Logger.log("Nuevo propietario detectado: " + correo);
    
    // Llamar a la función para asignar el ID
    generarIDParaFila(hoja);
    
  } else {
    Logger.log("No se detectó un correo válido en la respuesta.");
  }
}

/*** 
 * Vieja función de abajo
*/

/*function onFormSubmitPropietario(e) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propietarios");
  var datos = e.values;

  if (!datos) {
    Logger.log("No se detectaron datos en la respuesta del formulario.");
    return;
  }

  //var correo = datos[1];
  var correo = datos[3];

  if (correo && correo.includes("@")) {
    Logger.log("Nuevo propietario detectado: " + correo);
    
    // Llamar a la función para asignar el ID
    generarIDParaFila();
    
    // Actualizar permisos en el formulario
    actualizarPermisosEncuestados();
  } else {
    Logger.log("No se detectó un correo válido en la respuesta.");
  }
}*/

function actualizarPermisosEncuestados(hojaParaID) {
  var hojaPropietarios = hojaParaID;
  //var hojaPropietarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propietarios");
  var formId = "1olFjjiGeQ6QHL3JxVZ5na5MeXFtM_uYhZ8pOJNiBaEY";
  var formFile = DriveApp.getFileById(formId);
  var datos = hojaPropietarios.getDataRange().getValues();
  var correosPropietarios = [];

  // Obtener correos de propietarios
  for (var i = 1; i < datos.length; i++) { 
    if (datos[i][1] && datos[i][1].includes("@")) { 
      correosPropietarios.push(datos[i][1]);
    }
  }

  if (correosPropietarios.length > 0) {
    formFile.addViewers(correosPropietarios);
  }

  Logger.log("Se actualizaron los permisos de encuestado para: " + correosPropietarios.join(", "));
}

function generarIDParaFila(hojaParaID) {
  var hoja = hojaParaID;
  var ultimaFila = hoja.getLastRow(); // Obtiene la última fila con datos

  if (ultimaFila < 2) { 
    Logger.log("No hay datos suficientes en la hoja.");
    return;
  }

  var marcaTemporal = hoja.getRange(ultimaFila, 2).getValue(); // Columna "Marca temporal"

  // Formatear fecha en yyyyMMddHHmmss
  var anio = marcaTemporal.getFullYear();
  var mes = ('0' + (marcaTemporal.getMonth() + 1)).slice(-2);
  var dia = ('0' + marcaTemporal.getDate()).slice(-2);
  var horas = ('0' + marcaTemporal.getHours()).slice(-2);
  var minutos = ('0' + marcaTemporal.getMinutes()).slice(-2);
  var segundos = ('0' + marcaTemporal.getSeconds()).slice(-2);

  var idBase = anio + mes + dia + horas + minutos + segundos;

  // Generar número aleatorio entre 1000 y 9999
  var numeroAleatorio = Math.floor(Math.random() * (9999 - 1000 + 1)) + 1000;

  // Concatenar para formar el ID final
  var idFinal = idBase + numeroAleatorio;

  // Escribir el ID en la columna "ID" (suponiendo que es la columna 5)
  hoja.getRange(ultimaFila, 1).setValue(idFinal);

  Logger.log("Se generó el ID: " + idFinal);
}


