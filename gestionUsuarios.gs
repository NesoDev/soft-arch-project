const NOMBRE_DE_HOJA = "GestionUsuarios"
const HOJA_HISTORIAL = "Historial"
const CORREO_ADMIN = "bruno.pumapillo@unmsm.edu.pe"
//Para hacer tests 👇
const CORREO_USUARIO_GENERICO = "thulobruno@gmail.com"

// Función para obtener todos los usuarios
function getUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_DE_HOJA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const users = data.slice(1).map(row => {
    let user = {};
    headers.forEach((header, index) => {
      user[header] = row[index];
    });
    return user;
  });
  return users;
}

// Función para eliminar un usuario
function deleteUser(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_DE_HOJA);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIndex = headers.indexOf("Correo");

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][emailIndex] === email) {
      sheet.deleteRow(i + 1);
      logAction(`Usuario eliminado: ${email}`);
      sendNotification(email, "Su cuenta ha sido eliminada por el administrador.");
      break;
    }
  }
}

// Función para registrar acciones en el historial
function logAction(action) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_DE_HOJA);
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action]);
}
// prueba
// Función para enviar notificaciones al usuario
function sendNotification(email, message) {
  MailApp.sendEmail(email, "Notificación del sistema", message);
}

// Función principal para la interfaz de administrador
function manageUsers() {
  const userEmail = Session.getActiveUser().getEmail();
  if (userEmail !== CORREO_ADMIN) {
    throw new Error("Acceso denegado. Solo el administrador puede realizar esta acción.");
  }

  const users = getUsers();
  // Aquí puedes implementar la lógica para mostrar la lista de usuarios en una interfaz
  // Por ejemplo, usando HTML Service para crear una interfaz web.
  Logger.log(users);
}

// Función para buscar usuarios
function searchUsers(query) {
  const users = getUsers();
  return users.filter(user => 
    user.Nombre.includes(query) || user.Rol.includes(query)
  );
}

// Función para mostrar la interfaz
function showInterface() {
  const html = HtmlService.createHtmlOutputFromFile("Interface")
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Gestión de Usuarios");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Administrador")
    .addItem("Gestionar Usuarios", "showInterface")
    .addToUi();
}

function exampleUsage() {
  manageUsers(); // Mostrar usuarios
  deleteUser("usuario@example.com"); // Eliminar usuario
}