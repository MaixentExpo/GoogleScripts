/**
 * MaixentExpoApp : Scripts réutilisables
 * - ajouter la bibliothèque "MTHmRiwDeD0HAcidrAQ9BrTiZW5tJ2woG" dans Ressources/Bibliothèques
 * - mettre le préfixe MaixentExpoApp devant la fonction que vous voulez utiliser
 * MaixentExpoApp.LaFonction
 */

function attribuerlesPlaces() {
  SpreadsheetApp.getActiveSpreadsheet().toast("attribuerlesPlaces.","Toast message"); 
} // end

/**
 * createDiaporamaFromSlide est en fait un publipostage de diapositives
 * La diapo Modèle contiendra des champs sous la forme {NomDeLaColonne}
 * à remplir avec les données des colonnes de la feuille du tableur courant
 * Le diaporama en sortie sera créé avec le même nom que le modèle avec un suffixe " - Pub"
 * à raison d'une diapo par enregistrement filtrés
 * Les enregistrements pourront être filtrés sur une colonne
 * Les paramètres de la fonction :
 * sheetId     : Id du tableur ou nom du tableur
 * sheetName   : nom de la feuille qui contient les données du tableur
 * filterName  : nom de la colonne sur laquelle le filtre sera effectué
 * filterValue : expression régulière de filtrage sur la colonne
 */
function createDiaporamaFromSlide(sheetId, sheetName, filterName, filterValue) {
  var properties = PropertiesService.getScriptProperties();
  // Ouverture de la feuille
  var spreadsheet = sheetId.length > 15 ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.openById(properties.getProperty(sheetId));
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
  for(iCol=0; iCol<=iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim();
    if ( sCell != "" ) {
      iCols[sCell] = iCol;
    } // endif
  } // endfor

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
  
  // On ne prend que les lignes qui correspondent au critère filterName filterValue
  var sDatas = [];
  iLastRow = sValues.length;
  var reFilter = new RegExp(filterValue, 'g');
  for (iRow=1; iRow < iLastRow; iRow++) {   
    if ( sValues[iRow][iCols[filterName]].match(reFilter, 'g') != null ) {
      sDatas.push(sValues[iRow]);
    } // endif
  } // endfor
  iLastRow = sDatas.length;
  // duplication de la 1ère diapo autant que d'enregistrement-1
  for (iRow=1; iRow < iLastRow; iRow++) {   
    oDiaporamaCible.appendSlide(oDiapoCibles[0]);
  } // endfor

  // OK, maintenant on fusionne les données dans les diapos
  oDiapoCibles = oDiaporamaCible.getSlides();
  var iDiapo = 0;
  for(iRow=0; iRow<iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for( var key in iCols) {
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();     
      oDiapoCibles[iDiapo].replaceAllText("{$date}", frenchDate(new Date()));
      oDiapoCibles[iDiapo].replaceAllText("{" + key + "}", sCell)
    } // endfor key
    iDiapo++;
  } // endfor tableur
  oDiaporamaCible.saveAndClose();
}

/**
 * createDiaporamaFromSlide24 est en fait un publipostage de vignettes
 * à raison de 6 vignettes par page mais avec 2 vignettes par enregistrement
 * La diapo Modèle contiendra des champs sous la forme {NomDeLaColonneXY} X n° de ligne Y n° de colonne
 * Le diaporama en sortie sera créé avec le même nom que le modèle avec un suffixe " - Pub"
 * Les enregistrements pourront être filtrés sur une colonne de la table
 * Les paramètres de la fonction :
 * sheetId     : Id du tableur ou nom du tableur
 * sheetName   : nom de la feuille qui contient les données du tableur
 * filterName  : nom de la colonne sur laquelle le filtre sera effectué
 * filterValue : expression régulière de filtrage sur la colonne
 */
function createDiaporamaFromSlide24(sheetId, sheetName, filterName, filterValue) {
  var properties = PropertiesService.getScriptProperties();
  // Ouverture de la feuille
  var spreadsheet = sheetId.length > 15 ? SpreadsheetApp.openById(sheetId) : SpreadsheetApp.openById(properties.getProperty(sheetId));
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
  
  // On ne prend que les lignes qui correspondent au critère filterName filterValue
  var sDatas = [];
  iLastRow = sValues.length;
  var reFilter = new RegExp(filterValue, 'g');
  for (iRow=1; iRow < iLastRow; iRow++) {   
    if ( sValues[iRow][iCols[filterName]].match(reFilter, 'g') != null ) {
      sDatas.push(sValues[iRow]);
    } // endif
  } // endfor
  iLastRow = sDatas.length;
  // duplication de la 1ère diapo autant que d'enregistrement / 4
  for (iRow=4; iRow < iLastRow; iRow += 4) {   
    oDiaporamaCible.appendSlide(oDiapoCibles[0]);
  } // endfor

  // OK, maintenant on fusionne les données dans les diapos
  oDiapoCibles = oDiaporamaCible.getSlides();
  var iLigne = 1;
  var iDiapo = 0;
  var sKey = "";
  for(iRow=0; iRow<iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for( var key in iCols) {
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();
      for (var iCol=1; iCol<3; iCol++ ) {
        sKey = "{" + key + iLigne + iCol + "}"; 
        oDiapoCibles[iDiapo].replaceAllText(sKey, sCell);
        oDiapoCibles[iDiapo].replaceAllText("{$date}", frenchDate(new Date()));
      } // endfor
    } // endfor key
    iLigne++;
    if ( iLigne > 4 ) {
      iDiapo++;
      iLigne=1;
    } // endif
  } // endfor tableur
  oDiaporamaCible.saveAndClose();
}