/**
 * PublipostageSheet est en fait un publipostage d'une feuille de calcul
 * !!!! Problème pour remplacer le texte riche des cellules !!!!
 * Source : https://github.com/MaixentExpo/GoogleScripts/blob/master/MaixentExpoApp/PublipostageSheet.js
 * La diapo Modèle contiendra des champs sous la forme {NomDeLaColonne}
 * à remplir avec les données des colonnes de la feuille SheetName
 * Le PDF en sortie sera créé avec le même nom que le modèle avec un suffixe " - Pub"
 * à raison d'une page par enregistrement filtrés
 * Les enregistrements pourront être filtrés sur une colonne filterName avec la valeur filterValue
 * Les paramètres de la fonction :
 * sheetId     : Id du tableur
 * sheetName   : nom de la feuille qui contient les données du tableur
 * filterName  : nom de la colonne sur laquelle le filtre sera effectué
 * filterValue : expression régulière de filtrage sur la colonne
 */
function publipostageSheet(sheetId, sheetName, filterName, filterValue) {
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
  for(iCol=0; iCol<=iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim();
    if ( sCell != "" ) {
      iCols[sCell] = iCol;
    } // endif
  } // endfor

  // On ne prend que les lignes qui correspondent au critère filterName filterValue
  // pour alimenter sData
  var sDatas = [];
  iLastRow = sValues.length;
  var reFilter = new RegExp(filterValue, 'g');
  for (iRow=1; iRow < iLastRow; iRow++) {   
    if ( (""+ sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null ) {
      sDatas.push(sValues[iRow]);
    } // endif
  } // endfor
  iLastRow = sDatas.length;
  
  // Récupération de la Feuille Modèle
  var fileModele = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  // Création du publipostage en sortie
  var sCopyName = ""
  if ( fileModele.getName().match(" Modèle") ) {
    sCopyName = fileModele.getName().replace(" Modèle", " Pub");
  } else {
    sCopyName = fileModele.getName() + "- Pub";
  } // endif
  var fileCopy = fileModele.makeCopy(sCopyName);
  var oSheetCible = SpreadsheetApp.openById(fileCopy.getId());
  var oSheetCibles = oSheetCible.getSheets();
  

  // duplication de la 1ère diapo autant que d'enregistrement-1
  for (iRow=1; iRow < iLastRow; iRow++) {   
    oSheetCibles[0].copyTo(oSheetCible).setName("n " + iRow);
  } // endfor

  // OK, maintenant on fusionne les données dans les feuilles créées
  var oSheetCibles = oSheetCible.getSheets();
  var iSheet = 0;
  var textFinder;
  for(iRow=0; iRow<iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for( var key in iCols) {
      // Creates  a text finder.
      //var textFinder = oSheetCibles[iSheet]t.createTextFinder("{$date}");
      //textFinder.replaceAllWith(frenchDate(new Date()));
      textFinder = oSheetCibles[iSheet].createTextFinder("{" + key + "}");
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();     
      textFinder.replaceAllWith(sCell);
    } // endfor key
    iSheet++;
  } // endfor tableur
  oSheetCible.saveAndClose();
  //oSheetCible.showSheet();
}

/**
 * Présente une date sous la forme "12 avril 2019"
 * var maDate = new Date();
 * var maDateFrench = frenchDate(maDate)
 * @param {*} date 
 */
function frenchDate(date) {
  var month = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var m = month[date.getMonth()];
  var dateStringFr = date.getDate() + ' ' + m + ' ' + date.getFullYear();
  return dateStringFr
}
