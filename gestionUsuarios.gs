function agregarUsusarios(nombre, correo, rol, estado = 'Activo'){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const lastRow = sheet.getLastRow();
  const id = lastRow; //Genera un id automatico

  sheet.appendRow([id, nombre, correo, rol, estado]);
  Logger.log(`Usuario ${nombre} agregado con éxito.`);
}


function obtenerUsuarios(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const data = sheet.getDataRange().getValues();
  return data.slice(1);
}

function eliminarUsuario(id){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const data = sheet.getDataRange().getValues();

  for (let i= 0; i< data.length; i++){
    if(data[i][0] == id){
      sheet.deleteRow(i + 1);
      Logger.log(`Usuario con ID ${id} eliminado con éxito`);
      return;
    }
  }
  Logger.log(`Usuario con ID ${id} no encontrado.`);
}