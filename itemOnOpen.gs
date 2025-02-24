function onOpen() {
  
  var ui = SpreadsheetApp.getUi();
  
  // Crear el menú de Gestión Inmobiliaria
  var menuGestionInmobiliaria = GestionPropiedades.mostrarMenuGestion();
  
  // Crear el menú de Gestión de Usuarios (solo para administradores)
  var menuGestionUsuarios = GestionUsuarios.gestionUsuariosItem();
  
  // Combinar ambos menús en uno solo
  if (menuGestionUsuarios) {
    menuGestionInmobiliaria.addSubMenu(menuGestionUsuarios);
  }
  
  // Agregar el menú combinado a la interfaz
  menuGestionInmobiliaria.addToUi();


}
