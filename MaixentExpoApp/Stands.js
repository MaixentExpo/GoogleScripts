/**
 * stands_grand_hall_fusionEmplacements réalise la fusion de données dans le slide
 * 
 */
function stands_fusionPlan() {
  var ui = SpreadsheetApp.getUi(); //
  var yesnoConfirm = ui.alert(
    "FUSION DES DONNEES",
    'Veuillez confirmer par Oui ou Non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  var piloteSheet = SpreadsheetApp.getActiveSheet()
  // lecture des paramètres
  var sheetUrl = piloteSheet.getRange("B2").getValue()
  var sheetName = piloteSheet.getRange("B3").getValue()
  var firstLineData = piloteSheet.getRange("B4").getValue()
  var filterName = piloteSheet.getRange("B5").getValue()
  var filterValue = piloteSheet.getRange("B6").getValue()
  var folderUrl = piloteSheet.getRange("B7").getValue()
  var slideName = piloteSheet.getRange("B8").getValue()
  var inPdf = piloteSheet.getRange("B9").getValue()
  var modele = piloteSheet.getRange("B10").getValue()

  // Ouverture de la feuille
  var spreadsheet = SpreadsheetApp.openByUrl(sheetUrl)
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

  // On ne prend que les lignes qui correspondent au critère filterName filterValue
  var sDatas = []
  iLastRow = sValues.length
  var reFilter = new RegExp(filterValue, 'g')
  for (iRow = parseInt(firstLineData) - 1; iRow < iLastRow; iRow++) {
    if (("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow])
    } // endif
  } // endfor
  iLastRow = sDatas.length
  if (iLastRow == 0) {
    ui.alert("FUSION", "Aucun enregistrement trouvé", ui.ButtonSet.OK)
    return
  }

  // Récupération de la diapo Modèle
  var diaporamaModele = SlidesApp.openByUrl(modele)
  // Duplication du diaporama modèle
  var fileModele = DriveApp.getFileById(diaporamaModele.getId())
  const regex = /.*\/folders\/(.*)/g
  var folderId = regex.exec(folderUrl)[1]
  var folderCopy = DriveApp.getFolderById(folderId)
  var fileCopy = fileModele.makeCopy(slideName, folderCopy)
  var diaporamaCible = SlidesApp.openById(fileCopy.getId())
  var diapoCibles = diaporamaCible.getSlides()
  
  // OK, maintenant on fusionne les données dans la diapo cible
  var sKey = "";
  var sIndexValue;
  // parcours des données
  for(iRow=0; iRow<iLastRow; iRow++) {
    // Parcours des colonnes du tableur et fusion des données dans le slide
    for( var key in iCols) {
      sIndexValue = ("" + sDatas[iRow][iCols[filterName]]).trim(); // filterName sert d'index
      sKey = "{" + key + "_" + sIndexValue + "}"; 
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();
      diapoCibles[0].replaceAllText(sKey, sCell);
      diapoCibles[0].replaceAllText("{$date}", fx_frenchDate(new Date()));
    } // endfor key
  } // endfor row tableur
  // Signalement des emplacements vides et coloriage du fond des emplacements
  var shapes = diapoCibles[0].getShapes();
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
  diaporamaCible.saveAndClose();

  var url = ""
  if (inPdf) {
    var blob = DriveApp.getFileById(fileCopy.getId()).getBlob()
    var pdfFile = DriveApp.createFile(blob)
    // le fichier a été crée dans la racine du répertoire de l'utilisateur
    // un fichier peur avoir plusieurs répertoires
    DriveApp.getFolderById(folderId).addFile(pdfFile); // ajout du répertoire cible
    DriveApp.getRootFolder().removeFile(pdfFile); // suppresion du répertoire racine du fichier
    var url = pdfFile.getUrl()
    fileCopy.setTrashed(true)
  } else {
    var url = fileCopy.getUrl()
  }

  // Ecriture du lien dans la feuille Pilote
  piloteSheet.getRange("B15").setValue(url)

  // affichage d'un panneau pour ouvrir le document crée
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="' + url + '" target="_blank">Voir le résultat</a>')
    .setWidth(300)
    .setHeight(100)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Script terminé")

}
