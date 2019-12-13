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
function diapo_createDiaporamaFromSlide(sheetId, sheetName, filterName, filterValue) {
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
    if ( ("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null ) {
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
      oDiapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
      oDiapoCibles[iDiapo].replaceAllText("{" + key + "}", sCell)
    } // endfor key
    iDiapo++;
  } // endfor tableur
  oDiaporamaCible.saveAndClose();
}

function fx_fx_frenchDate(date) {
  var month = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var m = month[date.getMonth()];
  var dateStringFr = date.getDate() + ' ' + m + ' ' + date.getFullYear();
  return dateStringFr
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
function diapo_createDiaporamaFromSlide24(sheetId, sheetName, filterName, filterValue) {
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
    if ( ("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null ) {
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
        oDiapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
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
 * filterName  : nom de la colonne sur laquelle le filtre sera effectué
 * filterValue : expression régulière de filtrage sur la colonne
 */
function diapo_publipostageBadge24(sheetId, sheetName, filterName, filterValue) {
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
        oDiapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
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

function diapoSommaire() {
  var presentation = SlidesApp.getActivePresentation()
  var slides = presentation.getSlides()
  var qSlides = slides.length
  
  for (var iSlide=0; iSlide<qSlides; iSlide++ ) {
    var slide = slides[iSlide];
    var elements = slide.pageElements;
    for (var iElement = 0; iElement < elements.length; iElement++) {
      if (elements[iElement].shape && elements[iElement].shape.placeholder
          && (elements[iElement].shape.placeholder.type == 'TITLE'
              || elements[iElement].shape.placeholder.type == 'CENTERED_TITLE')) {
        var texts = elements[iElement].shape.text.textElements;
        var shapeText = "";
        for (var k = 0; k < texts.length; k++) {
          if (texts[k].autoText) {
             shapeText += texts[k].autoText.content;
          }
          if (texts[k].textRun) {
             shapeText += texts[k].textRun.content;
          }
        }
        Logger.log(shapeText);
      }
    }
  } 
}