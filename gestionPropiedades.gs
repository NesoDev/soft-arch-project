


function mostrarMenuGestion(){
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

function obtenerHoja() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = spreadsheet.getSheetByName("Propiedades");

    if (!hoja) {
        Logger.log("❌ No se encontró la hoja de Propiedades");
        return null;
    }
    return hoja;
}

function abrirFormulario() {
    var ui = SpreadsheetApp.getUi();
    var urlFormulario = "https://docs.google.com/forms/d/e/1FAIpQLScE5DflEXqQYp9K-naRKoeXhpgPVYi-KBobl8TSqbhv23hwOQ/viewform?usp=header";  
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
          margin-left:25px;
          margin-top:50px;
        }
        .add-estate:hover {
          background-color: #2E8BC0;
        }
      </style>
      <p>Haz clic en el siguiente enlace para agregar una nueva propiedad:</p>
      <a href="${urlFormulario}" target="_blank" class="add-estate">Crear propiedad</a>`
    )
        .setWidth(400)
        .setHeight(200);
    ui.showModalDialog(htmlOutput, "Formulario para agregar propiedad");
}



function agregarIDAutoincremental(e) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propiedades");
    var ultimaFila = hoja.getLastRow();

    if (ultimaFila > 1) { 
        var idAnterior = hoja.getRange(ultimaFila - 1, 1).getValue(); 
        var nuevoID = (idAnterior === "" || isNaN(idAnterior)) ? 1 : idAnterior + 1; 
        hoja.getRange(ultimaFila, 1).setValue(nuevoID); 
    }
}

function concatenarDireccion(e) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propiedades"); 
    var fila = e.range.getRow(); 

    
    var columnaCalle = 3; 
    var columnaNumero = 4; 
    var columnaDistrito = 5; 
    var columnaCiudad = 6; 
    var columnaCodigoPostal = 7; 
    var columnaDireccion = 8; 

  
    var valores = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];

    
    if (valores[columnaCalle - 1] && valores[columnaNumero - 1] && valores[columnaDistrito - 1] && valores[columnaCiudad - 1] && valores[columnaCodigoPostal - 1]) {
        var direccion = valores[columnaCalle - 1] + " " + valores[columnaNumero - 1] + ", " + valores[columnaDistrito - 1] + ", " + valores[columnaCiudad - 1] + " - " + valores[columnaCodigoPostal - 1];
        
        hoja.getRange(fila, columnaDireccion).setValue(direccion);
    }
}



function editarPropiedad() {
    var hoja = obtenerHoja();
    if (!hoja) return;

    var ui = SpreadsheetApp.getUi();
    var id = ui.prompt("Ingrese el ID de la propiedad a editar:").getResponseText();
    var datos = hoja.getDataRange().getValues();

    for (var i = 1; i < datos.length; i++) {
        if (datos[i][0] == id) {
            var propiedad = datos[i]; 
            var template = HtmlService.createTemplateFromFile("FormularioEdicion");
            template.id = id;
            template.nombre = propiedad[1];
            template.calle = propiedad[2];
            template.numero = propiedad[3];
            template.distrito = propiedad[4];
            template.ciudad = propiedad[5];
            template.codigoPostal = propiedad[6];
            template.referencia = propiedad[8];
            template.cantidadDepartamentos = propiedad[9];
            template.area = propiedad[10];
            template.servicios = propiedad[11];
            template.pisos = propiedad[12];
            template.estacionamiento = propiedad[13];
            template.foto = propiedad[14];
            var htmlOutput = template.evaluate().setWidth(500).setHeight(600);
            ui.showModalDialog(htmlOutput, "Editar Propiedad");
            return;
        }
    }
    ui.alert("Propiedad no encontrada.");
}



function procesarEdicionPropiedad(formData) {
    var hoja = obtenerHoja();
    var datos = hoja.getDataRange().getValues();

    for (var i = 1; i < datos.length; i++) {
        if (datos[i][0] == formData.id) {
            
            var direccionCompleta = formData.calle + " " + formData.numero + ", " + formData.distrito + ", " + formData.ciudad + " - " + formData.codigoPostal;

            
            hoja.getRange(i + 1, 2, 1, 14).setValues([[ 
                formData.nombre, 
                formData.calle,
                formData.numero,
                formData.distrito,
                formData.ciudad,
                formData.codigoPostal,
                direccionCompleta,
                formData.referencia,
                formData.cantidadDepartamentos,
                formData.area,
                formData.servicios,
                formData.pisos,
                formData.estacionamiento,
                formData.foto
            ]]);

            SpreadsheetApp.getUi().alert("Propiedad actualizada correctamente.");
            return;
        }
    }
    SpreadsheetApp.getUi().alert("Error: Propiedad no encontrada.");
}














function eliminarPropiedad() {
    var hoja = obtenerHoja();
    var ui = SpreadsheetApp.getUi();
    var id = ui.prompt("Ingrese el ID de la propiedad a eliminar:").getResponseText();
    var datos = hoja.getDataRange().getValues();

    for (var i = 1; i < datos.length; i++) {
        if (datos[i][0] == id) {
            hoja.deleteRow(i + 1);
            ui.alert("Propiedad eliminada correctamente.");
            return;
        } 
    }
    ui.alert("Propiedad no encontrada.");
}

function verDepartamentos() {
    var hoja = obtenerHoja();
    if (!hoja) return;
  
    var datos = hoja.getDataRange().getValues();
    var propiedades = [];
  
    for (var i = 1; i < datos.length; i++) {
        if (!hoja.isRowHiddenByUser(i + 1)) { // Verifica si la fila está oculta
            propiedades.push({ id: datos[i][0], nombre: datos[i][1] });
        }
    }
  
    var template = HtmlService.createTemplateFromFile("listaPropiedades");
    template.propiedades = JSON.stringify(propiedades);
  
    var htmlOutput = template.evaluate().setWidth(500).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Ver Departamentos");
}












function agregarIdArrendador(e) {
    var hojaRespuestas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Propiedades"); 
    var hojaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios_Propietarios");

    if (!hojaRespuestas || !hojaUsuarios) {
        Logger.log("Error: No se encontró la hoja de respuestas o la hoja de usuarios.");
        return;
    }

    var ultimaFila = hojaRespuestas.getLastRow(); 
    var ultimaColumna = hojaRespuestas.getLastColumn();
    var primeraFila = hojaRespuestas.getRange(1, 1, 1, ultimaColumna).getValues()[0]; 

    var columnaCorreo = primeraFila.indexOf("Dirección de correo electrónico") + 1; 
    if (columnaCorreo === 0) {
        Logger.log("Error: No se encontró la columna 'Dirección de correo electrónico'.");
        return;
    }

   
    var correoUsuario = hojaRespuestas.getRange(ultimaFila, columnaCorreo).getValue();  

    if (!correoUsuario) {
        Logger.log("Error: No se pudo obtener el correo electrónico del usuario.");
        return;
    }
    var idUsuario = obtenerIdUsuarioPorCorreo(hojaUsuarios, correoUsuario);
    if (!idUsuario) {
        Logger.log("No se encontró el usuario con correo: " + correoUsuario);
        return; 
    }

   
    var columnaIdArrendador = primeraFila.indexOf("id_arrendador") + 1; 

    if (columnaIdArrendador === 0) { 
        
        columnaIdArrendador = ultimaColumna + 1;
        hojaRespuestas.getRange(1, columnaIdArrendador).setValue("id_arrendador"); 
    }

    
    hojaRespuestas.getRange(ultimaFila, columnaIdArrendador).setValue(idUsuario);
    Logger.log("Se asignó el ID de arrendador " + idUsuario + " en la fila " + ultimaFila);
}




function obtenerIdUsuarioPorCorreo(hojaUsuarios, correo) {
    var datosUsuarios = hojaUsuarios.getDataRange().getValues();

    for (var i = 1; i < datosUsuarios.length; i++) { 
        if (datosUsuarios[i][2] == correo) { 
            return datosUsuarios[i][0]; 
        }
    }
    return null; 
}




function filtrarDepartamentos(idPropiedad) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Departamentos");
    if (!hoja) {
        SpreadsheetApp.getUi().alert("No se encontró la hoja de Departamentos");
        return;
    }

    var datos = hoja.getDataRange().getValues();
    hoja.showSheet();

    hoja.showRows(2, datos.length - 1);
    var hayDepartamentos = false; 

    for (var i = 1; i < datos.length; i++) {
        var idDepartamento = datos[i][9]; 
        if (idDepartamento == idPropiedad) {
            hayDepartamentos = true;  
        } else {
            hoja.hideRows(i + 1);
        }
    }

    if (!hayDepartamentos) {
        SpreadsheetApp.getUi().alert("Propiedad sin departamentos asociados.");
    }
}



function filtrarPropiedadesPorUsuario() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaUsuarios = ss.getSheetByName("Usuarios_Propietarios");
    var hojaPropiedades = ss.getSheetByName("Propiedades");

    if (!hojaUsuarios || !hojaPropiedades) {
        SpreadsheetApp.getUi().alert("No se encontró una de las hojas necesarias.");
        return;
    }

    
    var usuario = Session.getActiveUser().getEmail();
    var datosUsuarios = hojaUsuarios.getDataRange().getValues();
    var idPropietario = null;

    Logger.log(datosUsuarios);
    Logger.log("Correo usuario activo: " + usuario);
    
    for (var i = 1; i < datosUsuarios.length; i++) {
        if (datosUsuarios[i][2] === usuario) { 
            idPropietario = datosUsuarios[i][0]; 
            break;
        }
    }

    if (!idPropietario) {
        SpreadsheetApp.getUi().alert("No se encontró el usuario en la hoja de Usuarios_Propietarios.");
        return;
    }

    
    var datosPropiedades = hojaPropiedades.getDataRange().getValues();

    for (var j = 1; j < datosPropiedades.length; j++) {
        var idArrendador = datosPropiedades[j][17]; 
        
        if (idArrendador != idPropietario) {
            hojaPropiedades.hideRows(j + 1); 
        } else {
            hojaPropiedades.showRows(j + 1); 
        }
    }
}
