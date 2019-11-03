function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  var menu = ui.createMenu('Foire Expo');
  menu.addItem('Mettre à jour le message pour la presse', 'messagePresseConcoursBovins');
  menu.addItem("Mise à jour de l'onglet CONCOURS", 'updateOngletConcours');
  menu.addItem("Mise à jour de l'onglet LOTS", 'updateOngletLots');
  menu.addToUi();
}

