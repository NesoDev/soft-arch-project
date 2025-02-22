
function asignarPropiedadAPropietario() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaPropiedades = ss.getSheetByName('Propiedades');
    var hojaPropietarios = ss.getSheetByName('Usuarios_Propietarios');
    var hojaAsignaciones = ss.getSheetByName("Asignaciones");

    if (!hojaPropiedades || !hojaPropietarios || !hojaAsignaciones) return;

    var ui = SpreadsheetApp.getUi();
    var idPropiedad = ui.prompt("Ingrese el ID de la propiedad a asignar").getResponseText().trim();
    var correoPropietario = ui.prompt("Ingrese el correo del propietario").getResponseText().trim();

    var propietarios = hojaPropietarios.getDataRange().getValues();
    var propiedad = hojaPropiedades.getDataRange().getValues();

    var propietarioEncontrado = propietarios.find(row => row[2] === correoPropietario);
    var propiedadEncontrada = propiedad.find(row => row[0] == idPropiedad);

    if (!propietarioEncontrado) {
        ui.alert("El propietario con el correo ingresado no existe.");
        return;
    }
    if (!propiedadEncontrada) {
        ui.alert("No se encontró la propiedad con el ID ingresado.");
        return;
    }

    // Obtener la última fila de la hoja de Asignaciones
    var ultimaFila = hojaAsignaciones.getLastRow() + 1;

    // Registrar la asignación
    hojaAsignaciones.getRange(ultimaFila, 1, 1, 7).setValues([[
        ultimaFila - 1, 
        propietarioEncontrado[0], 
        propietarioEncontrado[3], 
        propiedadEncontrada[0], 
        propiedadEncontrada[1], 
        new Date(), 
        "Activo"
    ]]);

    ui.alert("Propiedad asignada correctamente al propietario.");
}


function registrarAlquiler() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaDepartamentos = ss.getSheetByName("Departamentos");
    var hojaInquilinos = ss.getSheetByName("Inquilinos");
    var hojaAlquileres = ss.getSheetByName("Alquileres");

    if (!hojaDepartamentos || !hojaInquilinos || !hojaAlquileres) return;

    var ui = SpreadsheetApp.getUi();
    var idDepartamento = ui.prompt("Ingrese el ID del departamento a alquilar").getResponseText().trim();
    var correoInquilino = ui.prompt("Ingrese el correo del inquilino").getResponseText().trim();
    var fechaInicio = ui.prompt("Ingrese la fecha de inicio (YYYY-MM-DD)").getResponseText().trim();
    var fechaFin = ui.prompt("Ingrese la fecha de fin (YYYY-MM-DD)").getResponseText().trim();
    var monto = ui.prompt("Ingrese el monto del alquiler").getResponseText().trim();

    var inquilinos = hojaInquilinos.getDataRange().getValues();
    var departamentos = hojaDepartamentos.getDataRange().getValues();

    var inquilinoEncontrado = inquilinos.find(row => row[2] === correoInquilino);
    var departamentoEncontrado = departamentos.find(row => row[0] == idDepartamento);

    if (!inquilinoEncontrado) {
        ui.alert("El inquilino con el correo ingresado no existe.");
        return;
    }
    if (!departamentoEncontrado) {
        ui.alert("No se encontró el departamento con el ID ingresado.");
        return;
    }

    // Obtener la última fila de la hoja de Alquileres
    var ultimaFila = hojaAlquileres.getLastRow() + 1;

    // Registrar el alquiler
    hojaAlquileres.getRange(ultimaFila, 1, 1, 10).setValues([[
        ultimaFila - 1, 
        departamentoEncontrado[1], 
        departamentoEncontrado[2], 
        departamentoEncontrado[0], 
        inquilinoEncontrado[0], 
        inquilinoEncontrado[3], 
        fechaInicio, 
        fechaFin, 
        monto, 
        "Activo"
    ]]);

    ui.alert("Alquiler registrado correctamente.");
}



/**
 * 
 * Esto lo edito despues
 * 
 * 
//TODO: hacer un menu context para implementar esto
function crearFormulario(arrendadorEmail) {
  var formulario = FormApp.create('Formulario de Inquilinos de ' + arrendadorEmail);
  formulario.addTextItem().setTitle('Nombre del Inquilino');
  formulario.addTextItem().setTitle('Correo del Inquilino');
  formulario.addTextItem().setTitle('Teléfono del Inquilino');
  formulario.addParagraphTextItem().setTitle('Dirección del Inquilino');
  
  // Enviar notificación al arrendador
  MailApp.sendEmail(arrendadorEmail, 'Formulario de Inquilinos Creado', 
                    'El formulario para registrar inquilinos ha sido creado con éxito. ' + 
                    'Accede a él: ' + formulario.getEditUrl());
}

// onOpen para crear hojas específicas por arrendador
function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaArrendador = ss.getSheetByName('Arrendadores');
  
  if (!hojaArrendador) return;
  
  // Obtener el correo electrónico del usuario
  var emailUsuario = Session.getActiveUser().getEmail();
  
  // Comprobar si el correo es de un arrendador registrado
  var arrendador = obtenerArrendador(emailUsuario);
  
  if (arrendador) {
    // Crear una hoja para el arrendador si no existe
    if (!ss.getSheetByName(arrendador.nombre)) {
      var hoja = ss.insertSheet(arrendador.nombre);
      hoja.appendRow(['Nombre Inquilino', 'Correo Inquilino', 'Teléfono', 'Dirección']);
    }
    // Hacer visible solo la hoja correspondiente al arrendador
    ss.setActiveSheet(ss.getSheetByName(arrendador.nombre));
  } else {
    SpreadsheetApp.getUi().alert('Acceso denegado. Este usuario no es arrendador.');
  }
}

// Función para obtener el arrendador por correo
function obtenerArrendador(email) {
  var hojaArrendadores = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Arrendadores');
  var datos = hojaArrendadores.getDataRange().getValues();
  
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === email) {  // Compara con el correo del arrendador
      return { nombre: datos[i][0], propietario: datos[i][1] };
    }
  }
  return null;  // Si no lo encuentra
}

// Crear hoja resumen con Propietarios e Inquilinos
function generarResumen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaResumenPropietarios = ss.getSheetByName('Propietarios');
  var hojaResumenInquilinos = ss.getSheetByName('Inquilinos');
  
  // Si las hojas no existen, crear nuevas
  if (!hojaResumenPropietarios) {
    hojaResumenPropietarios = ss.insertSheet('Propietarios');
    hojaResumenPropietarios.appendRow(['Correo Propietario', 'Nombre']);
  }
  
  if (!hojaResumenInquilinos) {
    hojaResumenInquilinos = ss.insertSheet('Inquilinos');
    hojaResumenInquilinos.appendRow(['Nombre Inquilino', 'Correo Inquilino', 'Correo Propietario']);
  }
  
  // Limpiar hojas existentes antes de actualizarlas
  hojaResumenPropietarios.clearContents();
  hojaResumenInquilinos.clearContents();
  
  // Obtener datos de arrendadores y su hoja respectiva
  var hojaArrendadores = ss.getSheetByName('Arrendadores');
  var datosArrendadores = hojaArrendadores.getDataRange().getValues();
  
  for (var i = 1; i < datosArrendadores.length; i++) {
    var arrendadorEmail = datosArrendadores[i][0];
    hojaResumenPropietarios.appendRow([arrendadorEmail, datosArrendadores[i][1]]);
    
    // Obtener inquilinos del arrendador
    var hojaInquilinos = ss.getSheetByName(datosArrendadores[i][1]);
    if (hojaInquilinos) {
      var datosInquilinos = hojaInquilinos.getDataRange().getValues();
      for (var j = 1; j < datosInquilinos.length; j++) {
        hojaResumenInquilinos.appendRow([datosInquilinos[j][0], datosInquilinos[j][1], arrendadorEmail]);
      }
    }
  }
}

// Función para restringir la edición a solo la hoja de cada arrendador
function onEdit(e) {
  var ss = e.source;
  var hoja = e.range.getSheet();
  var emailUsuario = Session.getActiveUser().getEmail();
  
  // Verificar si la hoja es la correspondiente al arrendador
  var arrendador = obtenerArrendador(emailUsuario);
  
  if (arrendador && hoja.getName() !== arrendador.nombre) {
    // Si la hoja no corresponde al arrendador, revertir el cambio
    e.range.setValue(e.oldValue);
    SpreadsheetApp.getUi().alert('No tienes permiso para editar esta hoja.');
  }
}

// Función para notificar al arrendador cuando se registre un nuevo inquilino
function notificarInquilinoRegistrado(arrendadorEmail, inquilinoNombre, inquilinoCorreo) {
  var asunto = 'Nuevo Inquilino Registrado';
  var mensaje = 'Hola, \n\nSe ha registrado un nuevo inquilino en tu propiedad.\n\n' +
                'Nombre del Inquilino: ' + inquilinoNombre + '\n' +
                'Correo del Inquilino: ' + inquilinoCorreo + '\n\n' +
                'Gracias por usar nuestro sistema.\n\n' +
                'Saludos, \nEl equipo de Gestión de Arrendamientos';
  
  MailApp.sendEmail(arrendadorEmail, asunto, mensaje);
}

// Función para registrar un inquilino
function registrarInquilino() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var emailUsuario = Session.getActiveUser().getEmail();
  
  var arrendador = obtenerArrendador(emailUsuario);
  
  if (arrendador && hoja.getName() === arrendador.nombre) {
    var ui = SpreadsheetApp.getUi();
    var respuesta = ui.prompt('Registrar Inquilino', 'Ingrese el nombre, correo y teléfono del inquilino, separados por coma:', ui.ButtonSet.OK_CANCEL);
    
    if (respuesta.getSelectedButton() === ui.Button.OK) {
      var datosInquilino = respuesta.getResponseText().split(',');
      if (datosInquilino.length === 3) {
        hoja.appendRow([datosInquilino[0], datosInquilino[1], datosInquilino[2]]);
        
        // Notificar al arrendador sobre el nuevo inquilino
        notificarInquilinoRegistrado(arrendador.email, datosInquilino[0], datosInquilino[1]);
        
        ui.alert('Inquilino registrado y notificación enviada al arrendador.');
      } else {
        ui.alert('Por favor ingresa la información correctamente: Nombre, Correo, Teléfono');
      }
    }
  } else {
    SpreadsheetApp.getUi().alert('No tienes permiso para registrar inquilinos en esta hoja.');
  }
}

*/
