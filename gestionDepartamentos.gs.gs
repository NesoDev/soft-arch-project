function onOpen(e) {
  filtrarPropiedadesPorUsuario();
  var ui = SpreadsheetApp.getUi();
  var menuPrincipal = ui.createMenu('Gestión Inmobiliaria');
  var submenuPropiedades = ui.createMenu('Propiedades')
      .addItem("Agregar Propiedad", "abrirFormulario")
      .addItem("Editar Propiedad", "editarPropiedad")
      .addItem("Eliminar Propiedad", "eliminarPropiedad")
      .addItem("Ver Propiedades", "verDepartamentos")

  var submenuDepartamentos = ui.createMenu('Departamentos')
      .addItem('Agregar Departamento', 'addDepartment')
      .addItem('Editar Departamento', 'editDepartment')
      .addItem('Eliminar Departamento', 'deleteDepartment');
  
  // Agregar submenús al menú principal
  menuPrincipal.addSubMenu(submenuPropiedades);
  menuPrincipal.addSubMenu(submenuDepartamentos);
  menuPrincipal.addToUi();
}

function addDepartment() {
  const ui = SpreadsheetApp.getUi();
  // URL del formulario 
  const formularioUrl = "https://docs.google.com/forms/d/e/1FAIpQLScHwdHxnDbnOHW_qY0-C_cdVPgkxsYoWvPX-Qi74v8RQSAaXA/viewform"; // Reemplaza con la URL de tu formulario
  const htmlOutput = HtmlService.createHtmlOutput(
    `<style>
      .add-estate{
         all:none; 
         display:block;
         background-color:#145DA0;
         color: white;
         font-size: 16px;
         font-weight: bold;
         text-decoration: none;
         border: 1px solid #ccc;
         padding: 15px 5px;
         border-radius: 10px;
         text-align:center;
       }
       .add-estate:hover {
         background-color: #2E8BC0;
       }
    </style>
    <p>Haz clic en el siguiente enlace para agregar una nuevo departamento:</p>
     <a href="${formularioUrl}" target="_blank" class="add-estate">Crear departamento</a>`
  )
    .setWidth(300) 
    .setHeight(150);
  ui.showModalDialog(htmlOutput, "Agregar Departamento");
}
function generarIDDepartamento() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Departamentos");
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const idLast = sheet.getRange(lastRow - 1, 1).getValue();
    const id = (idLast === "" || isNaN(idLast)) ? 1 : idLast + 1;
    sheet.getRange(lastRow, 1).setValue(id);
  }
}
function getSheetDepartment() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Departamentos");

    if (!sheet) {
        Logger.log("❌ No se encontró la hoja de Departamentos");
        return null;
    }
    return sheet;
}
function editDepartment() {
    var sheet = getSheetDepartment();
    if (!sheet) return;

    var ui = SpreadsheetApp.getUi();
    var id = parseInt( ui.prompt("Ingrese el id del departamento a editar:").getResponseText());
    var data = sheet.getDataRange().getValues();
    for(let i =1 ; i<data.length  ;i++){
      if(data[i][0]===id){
        var dataDepartment = data[i]
        var template = HtmlService.createTemplateFromFile("FormDepartamento")
        var timeZone = Session.getScriptTimeZone();
        template.id=id;
        template.condition=dataDepartment[2];
        template.cost = dataDepartment[3];
        template.rentalDate = Utilities.formatDate(dataDepartment[4],timeZone ,"yyyy-MM-dd");
        template.serviceDate = Utilities.formatDate(dataDepartment[5],timeZone ,"yyyy-MM-dd");;
        template.warranty = dataDepartment[6];
        template.state = dataDepartment[7];
        template.photo= dataDepartment[8];
        var htmlOutput = template.evaluate().setWidth(500).setHeight(600);
        ui.showModalDialog(htmlOutput, "Editar Departamento");
        return;
      }
    }
    ui.alert("departamento no encontrada.");
}
function editionDepartment(dataForm) {
  Logger.log("Datos recibidos: " + JSON.stringify(dataForm));

  if (!dataForm || !dataForm.id) {
    Logger.log("❌ Error: formData es undefined o no tiene ID");
    return "❌ Error: formData es undefined o no tiene ID";
  }

  var sheet = getSheetDepartment();
  var data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === parseInt(dataForm.id)) {
      sheet.getRange(i + 1, 3, 1, 7).setValues([[
        dataForm.condition,
        dataForm.cost,
        dataForm.rentalDate,
        dataForm.serviceDate,
        dataForm.warranty,
        dataForm.state,
        dataForm.photo,
      ]]);

      Logger.log("✅ Datos actualizados correctamente.");
      return "✅ Datos actualizados correctamente.";
    }
  }

  Logger.log("❌ Departamento no actualizado.");
  return "❌ Departamento no actualizado.";
}

function deleteDepartment() {
    var sheet = getSheetDepartment();
    //if (!sheet) return;

    var ui = SpreadsheetApp.getUi();
    var id = ui.prompt("Ingrese el ID del departamento a eliminar:").getResponseText();
    var data = sheet.getDataRange().getValues();
    for(let i=1 ; i< data.length;i++){
      if(data[i][0]==id){
        sheet.deleteRow(i+1);
        ui.alert("Departamento eliminado");
        return ;
      }
    }
    ui.alert("Departamento no encontrado")
}
function requestAuthorization() {
  Logger.log("🚀 El script ha sido autorizado correctamente.");
}

