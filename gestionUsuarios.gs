const ADMINISTRADORES = new Set([ // Lista de correos de administradores
  'thulobruno@gmail.com',
  'bruno.pumapillo@unmsm.edu.pe'
]);
// **Configurar**: Cambiar nombres de hojas según tu proyecto.
const HOJA_USUARIOS = 'GestionUsuarios';
const PROPIETARIOS = 'Usuarios_Propietarios';
const INQUILINOS = 'Usuarios_Inquilinos';
const HOJA_HISTORIAL = 'Historial';

function onOpen(e) {
  
  mostrarMenuGestion();
  var usuarioActual = e && e.user ? e.user.getEmail() : Session.getActiveUser().getEmail(); 
  

  //if (usuarioActual !== ADMINISTRADOR) return; // Si no es el admin, no ejecuta el menú
  // Verificar si el usuario actual está en la lista de administradores
  if (!ADMINISTRADORES.has(usuarioActual)) return;
  var ui = SpreadsheetApp.getUi();

  // Submenú de Usuarios
  var submenuUsuariosPropietarios = ui.createMenu('Propietarios')
    .addItem('Ver Propietarios', 'mostrarUsuariosPropietarios')
    .addItem('Eliminar Propietarios', 'eliminarUsuariosPropietarios');

  var submenuUsuariosInquilinos = ui.createMenu('Inquilinos')
    .addItem('Ver Inquilinos', 'mostrarUsuariosInquilinos')
    .addItem('Eliminar Inquilinos', 'eliminarUsuariosInquilinos');

  // Menú Principal
  var menuPrincipal = ui.createMenu('Gestión de Usuarios')
    .addSubMenu(submenuUsuariosPropietarios)
    .addSubMenu(submenuUsuariosInquilinos)
    .addToUi();
}


function mostrarUsuariosPropietarios(){
  mostrarUsuarios(PROPIETARIOS);
}

function mostrarUsuariosInquilinos(){
  mostrarUsuarios(INQUILINOS);
}

function eliminarUsuariosPropietarios(){
  eliminarUsuario(PROPIETARIOS);
}

function eliminarUsuariosInquilinos(){
  eliminarUsuario(INQUILINOS);
}

function generarIDUsuariosPropietarios(){
  generarID(PROPIETARIOS);
}

function generarIDUsuariosInquilinos(){
  generarID(INQUILINOS);
}

function mostrarUsuarios(hojaSheet) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojaSheet);
  if (!hoja) return SpreadsheetApp.getUi().alert('Hoja de usuarios no encontrada.');

  var datos = hoja.getDataRange().getValues();
  var htmlOutput = `
  <html>
  <head>
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
  </head>
  <body>
    <div class="p-6">
      <h2 class="text-2xl font-semibold mb-4">Lista de Usuarios</h2>
      <table class="min-w-full table-auto border-collapse border border-gray-300">
        <thead class="bg-gray-100">
          <tr>
            <th class="px-4 py-2 text-left text-sm font-medium text-gray-600">ID</th>
            <th class="px-4 py-2 text-left text-sm font-medium text-gray-600">Marca Temporal</th>
            <th class="px-4 py-2 text-left text-sm font-medium text-gray-600">Correo</th>
            <th class="px-4 py-2 text-left text-sm font-medium text-gray-600">Nombre y Apellidos</th>
            <th class="px-4 py-2 text-left text-sm font-medium text-gray-600">Celular</th>
          </tr>
        </thead>
        <tbody>
          ${datos.slice(1).map(row => 
            row[0] && row[1] && row[2] && row[3] ? `
              <tr class="border-b hover:bg-gray-50">
                <td class="px-4 py-2 text-sm text-gray-800">${row[0]}</td>
                <td class="px-4 py-2 text-sm text-gray-800">${row[1]}</td>
                <td class="px-4 py-2 text-sm text-gray-800">${row[2]}</td>
                <td class="px-4 py-2 text-sm text-gray-800">${row[3]}</td>
                <td class="px-4 py-2 text-sm text-gray-800">${row[4]}</td>
              </tr>
            ` : ''
          ).join('')}
        </tbody>
      </table>
    </div>
  </body>
  </html>
  `;

  var html = HtmlService.createHtmlOutput(htmlOutput)
    .setWidth(800)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Usuarios Registrados');
}


// Eliminar usuario por correo y registrar acción en 'Historial'
function eliminarUsuario(hojaSheet) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojaSheet);
  if (!hoja) return mostrarAlerta('error', 'Hoja de usuarios no encontrada.');

  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('Eliminar Usuario', 'Ingrese el correo del usuario:', ui.ButtonSet.OK_CANCEL);

  // Si el usuario acepta
  if (respuesta.getSelectedButton() === ui.Button.OK) {
    var correo = respuesta.getResponseText().trim();
    if (!correo) {
      return mostrarAlerta('error', 'Por favor ingrese un correo válido.');
    }

    var datos = hoja.getDataRange().getValues();
    var filaEliminada = null;
    
    // Buscar el usuario por correo en la columna de correos (columna 1)
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][1] === correo) {  // Comparar con la columna "Correo"
        filaEliminada = i + 1;  // Guardamos la fila a eliminar
        var eliminado = `Nombre: ${datos[i][2]}, Correo: ${datos[i][1]}, Celular: ${datos[i][3]}`;
        
        // Eliminar la fila
        hoja.deleteRow(i + 1);

        // Registrar acción en el historial
        registrarHistorial(`Usuario eliminado: ${eliminado}`);
        return mostrarAlerta('success', 'Usuario eliminado con éxito.');
      }
    }

    // Si no se encontró el usuario, mostramos un mensaje
    return mostrarAlerta('error', 'Usuario no encontrado.');
  }
}

// Función para mostrar alerta personalizada con Tailwind
function mostrarAlerta(tipo, mensaje) {
  var color = tipo === 'success' ? 'bg-green-500' : (tipo === 'error' ? 'bg-red-500' : 'bg-yellow-500');
  var htmlOutput = `
    <html>
      <head>
        <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
      </head>
      <body class="flex items-center justify-center min-h-screen bg-gray-100">
        <div class="w-96 p-6 ${color} text-white rounded-lg shadow-lg">
          <h2 class="text-xl font-semibold mb-4">Notificación</h2>
          <p>${mensaje}</p>
          <div class="mt-4 text-right">
            <button onclick="google.script.host.close()" class="bg-white text-gray-800 px-4 py-2 rounded-lg shadow">Cerrar</button>
          </div>
        </div>
      </body>
    </html>
  `;

  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(htmlOutput)
    .setWidth(400)
    .setHeight(250);
  
  ui.showModalDialog(html, 'Notificación');
}


// Registrar acciones en la hoja 'Historial'
function registrarHistorial(accion) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_HISTORIAL);
  if (!hoja) return;
  hoja.appendRow([new Date(), accion, Session.getActiveUser().getEmail()]);
}

function generarID(hojaSheet){
  //Hoja
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hojaSheet);
  //Obtener la ultima fila
  var ultimaFila = hoja.getLastRow();
  var idConsecutivo = ultimaFila - 1;
  //Colocar el valor 
  hoja.getRange(ultimaFila, 1).setValue(idConsecutivo);
}

var GestionUsuarios = {
  mostrarUsuarios: mostrarUsuarios,
  eliminarUsuario: eliminarUsuario
};
