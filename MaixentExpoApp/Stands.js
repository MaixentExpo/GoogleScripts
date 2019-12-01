/**
 * PublipostageBadge24 est en fait un publipostage de vignettes
 * Référence : https://github.com/MaixentExpo/GoogleScripts/edit/master/MaixentExpoApp/PublipostageBadge24.js
 * à raison de 6 vignettes par page mais avec 2 vignettes par enregistrement
 * La diapo Modèle contiendra des champs sous la forme {NomDeLaColonneXY} X n° de ligne Y n° de colonne
 * Le diaporama en sortie sera créé avec le même nom que le modèle avec un suffixe " - Pub"
 * Les enregistrements pourront être filtrés sur une colonne de la table
 * Les paramètres de la fonction :
 * sheetId     : Id du tableur
 * sheetName   : nom de la feuille qui contient les données du tableur
 * IndexName   : nom de la colonne qui contient le n° des stands
 */
function stands_grand_hall_fusionEmplacements(sheetId, sheetName, indexName) {
  var properties = PropertiesService.getScriptProperties();
  // Ouverture de la feuille
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getSheetByName(sheetName)
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  // chargement global de toutes les données de la feuille
  // pour optimiser les ressources du serveur de Google
  var sValues = sheet.getRange(1, 1, iLastRow, iLastCol).getValues();
  // Lecture de la ligne d'entête pour mémoriser le nom des colonnes et leur position
  var iCols = {};
  var sCell = "";
  var iRow = 0;
  var iCol;
  for(iCol=0; iCol<iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim();
    if ( sCell != "" ) {
      iCols[sCell] = iCol;
    } // endif
  } // endfor

  // On ne prendra que les lignes qui correspondent au critère filterName filterValue
  var sDatas = [];
  iLastRow = sValues.length;
  for (iRow=1; iRow < iLastRow; iRow++) {   
      sDatas.push(sValues[iRow]);
  } // endfor
  iLastRow = sDatas.length;

  // Récupération de la diapo Modèle
  var fileModele = DriveApp.getFileById(SlidesApp.getActivePresentation().getId());
  // Création du Diaporama en sortie
  var sCopyName = ""
  if ( fileModele.getName().match(" Modèle") ) {
    sCopyName = fileModele.getName().replace(" Modèle", " Pub");
  } else {
    sCopyName = fileModele.getName() + "- Pub";
  } // endif
  var fileCopy = fileModele.makeCopy(sCopyName);
  var oDiaporamaCible = SlidesApp.openById(fileCopy.getId());
  var oDiapoCibles = oDiaporamaCible.getSlides();
  
  // OK, maintenant on fusionne les données dans la diapo cible
  var sKey = "";
  var sIndexValue;
  // parcours des données
  for(iRow=0; iRow<iLastRow; iRow++) {
    // Parcours des colonnes du tableur et fusion des données dans le slide
    for( var key in iCols) {
      sIndexValue = ("" + sDatas[iRow][iCols[indexName]]).trim();
      sKey = "{" + key + "_" + sIndexValue + "}"; 
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();
      oDiapoCibles[0].replaceAllText(sKey, sCell);
      oDiapoCibles[0].replaceAllText("{$date}", fx_frenchDate(new Date()));
    } // endfor key
  } // endfor row tableur
  // Signalement des emplacements vides et coloriage du fond des emplacements
  var shapes = oDiapoCibles[0].getShapes();
  var countShapes = shapes.length; 
  for ( var iShape=0; iShape < countShapes; iShape++) {
    var tt = shapes[iShape].getText().asString();
    if ( tt.indexOf("{") != -1 ) {
      // emplacement vide
      shapes[iShape].getFill().setTransparent();
      var num = tt.match("{.*_(.*)}");
      if ( num ) {
        shapes[iShape].getText().setText(" " + num[1] + "\u00a0");
      } 
    } else {
      // emplacement occupé de type "Viticulture" ou "Gastronomie"
      var inscr = tt.match(/\((.*)\)/);
      if ( inscr ) {
        // recherche de l'emplacement dans le tableur
        for (iRow=0; iRow<iLastRow; iRow++) {
          var ss = sDatas[iRow][iCols["INSCR"]].toString()
          if ( inscr[1] == ss ) {
            var secteur = sDatas[iRow][iCols["Secteur"]]
            if ( secteur == "Viticulture" ) {
              // stand Viticulture
              shapes[iShape].getFill().setSolidFill("#ead1dc") // magenta clair 3
            } else {
              // stand Gastronomie
              shapes[iShape].getFill().setSolidFill("#fce5cd") // orange clair 3
            } // endif
          } // endif inscr
        } // end for row
      } // endif inscr
    } // endif
  }
  
  oDiaporamaCible.saveAndClose();
}
