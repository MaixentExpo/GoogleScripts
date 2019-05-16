function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Foire Expo')
      //.addItem('Mettre à jour le message pour la presse', 'menuItem1')
      .addItem('Mettre à jour le message pour la presse', 'messagePresseConcoursVins')
      .addToUi();
}
