function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  var menu = ui.createMenu('Foire Expo');
  menu.addItem('Mettre à jour le message pour la presse', 'messagePresseConcoursBovins');
  menu.addItem("Mise à jour de l'onglet CONCOURS", 'updateOngletConcours');
  menu.addItem("Mise à jour de l'onglet LOTS", 'updateOngletLots');
  menu.addToUi();
}

// Class Couleur 1.0
var Couleur = function() {
  this.couleurs = [ "#e8f5e9" // green
                   ,"#e3f2fd" // blue
                   ,"#fffde7" // yellow
                   ,"#fbe9e7" // deep orange
                   ,"#e0f7fa" // cyan
                   ,"#f1f8e9" // light green
                   ,"#fce4ec" // pink
                   ,"#e1f5fe" // light blue
                   ,"#ede7f6" // deep purple
                   ,"#eceff1" // blue grey
                   ,"#e8eaf6" // indigo
                   ,"#f3e5f5" // purple
                   ,"#f9fbe7" // lime
                   ,"#fff3e0" // orange
                   ,"#fff8e1" // amber
                   ,"#efebe9" // brown
                   ,"#e0f2f1" // teal
                   ,"#ffebee" // red
                   ,"#fafafa" // grey
                  ];
  this.iCouleur = -1;
  this.couleur = "#fafafa";
  this.new_couleur = function () {
    this.iCouleur++
    if ( this.iCouleur >= this.couleurs.length ) {
      this.iCouleur = 0;
    } // endif
    this.couleur = this.couleurs[this.iCouleur];
  } // end new_couleur
} // end class Couleur
