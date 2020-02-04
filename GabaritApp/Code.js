/**
 * GabaritApp : Scripts de publipostage
 * à partir d'un document maître (le gabarit) qui précisera
 * - {$fichier: id du fichier des données}
 * - {$feuille: le nom de l'onglet de la feuille}
 * - {$ligne: le n° de la 1ère ligne des données}
 * - {$doc1: l'id du 1er document à assembler}
 * - {$doc2: l'id du 2ème document à assembler}
 * - {$doc3: l'id du 3ème document à assembler}
 * - {doc n: ...}
 */

function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Foire Expo')
      .addItem('Lancer le publipostage', 'gabarit')
      .addToUi();
}


function gabarit() {
  function fx_frenchDate(date) {
    var month = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
    var m = month[date.getMonth()];
    var dateStringFr = date.getDate() + ' ' + m + ' ' + date.getFullYear();
    return dateStringFr
  }
  var gabarit = SlidesApp.getActivePresentation()
  // lecture des paramètres
  var slideGabarit = gabarit.getSlides()[0]
  var table = slideGabarit.getTables()[0]
  
  // Lecture des paramètres
  var ip = 0
  var sheetId = table.getCell(ip++, 1).getText().asString().replace(/\n/,'')
  var sheetName = table.getCell(ip++, 1).getText().asString().replace(/\n/,'')
  var firstLineData = table.getCell(ip++, 1).getText().asString().replace(/\n/,'')
  var filterName = table.getCell(ip++, 1).getText().asString().replace(/\n/,'')
  var filterValue = table.getCell(ip++, 1).getText().asString().replace(/\n/,'')
  var docs = []
  if ( table.getCell(ip++, 1).getText().getLength() > 1 )
    docs.push(table.getCell(ip-1, 1).getText().asString().replace(/\n/,''))
  if ( table.getCell(ip++, 1).getText().getLength() > 1 )
    docs.push(table.getCell(ip-1, 1).getText().asString().replace(/\n/,''))
  if ( table.getCell(ip++, 1).getText().getLength() > 1 )
    docs.push(table.getCell(ip-1, 1).getText().asString().replace(/\n/,''))
  if ( table.getCell(ip++, 1).getText().getLength() > 1 )
    docs.push(table.getCell(ip-1, 1).getText().asString().replace(/\n/,''))
  if ( table.getCell(ip++, 1).getText().getLength() > 1 )
    docs.push(table.getCell(ip-1, 1).getText().asString().replace(/\n/,''))

  // Chargement des données
  var spreadsheet = SpreadsheetApp.openById(sheetId)
  var sheet = spreadsheet.getSheetByName(sheetName)
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  // chargement global de toutes les données de la feuille
  // pour optimiser les ressources du serveur de Google
  var sValues = sheet.getRange(1, 1, iLastRow, iLastCol).getValues()
  // Lecture de la ligne d'entête pour mémoriser le nom des colonnes et leur position
  var iCols = {}
  var sCell = ""
  var iRow = 0
  var iCol
  for (iCol = 0; iCol < iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim()
    if (sCell != "") {
      iCols[sCell] = iCol
    } // endif
  } // endfor
  // On ne prend que les lignes qui correspondent au critère filterName filterValue
  var sDatas = []
  iLastRow = sValues.length
  var reFilter = new RegExp(filterValue, 'g')
  for (iRow = parseInt(firstLineData)-1; iRow < iLastRow; iRow++) {
    if (("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow])
    } // endif
  } // endfor
  iLastRow = sDatas.length
  
  // Récupération de la lettre et des annexes
  var slidesAll = []
  for( var idoc=0; idoc<docs.length; idoc++) {
    var presentation = SlidesApp.openById(docs[idoc])
    var slides = presentation.getSlides()
    for (var islide=0; islide<slides.length; islide++) {
      slidesAll.push(slides[islide])
    }
  }

  // Récupération de la diapo Gabarit pour en créer une copie - Pub
  var fileModele = DriveApp.getFileById(SlidesApp.getActivePresentation().getId())
  // Création du Diaporama en sortie
  var sCopyName = ""
  if (fileModele.getName().match(" Gabarit")) {
    sCopyName = fileModele.getName().replace(" Gabarit", " Pub")
  } else {
    sCopyName = fileModele.getName() + "- Pub"
  } // endif
  var fileCopy = fileModele.makeCopy(sCopyName)
  var oDiaporamaCible = SlidesApp.openById(fileCopy.getId())

  // duplication des diapos autant de fois que d'enregistrements
  for (iRow = 0; iRow < iLastRow; iRow++) {
    for (var islide=0; islide<slidesAll.length; islide++) {
      oDiaporamaCible.appendSlide(slidesAll[islide])
    }
  } // endfor
  // Suppression de la 1ère diapo qui correspond au gabarit
  var firstSlide = oDiaporamaCible.getSlides().shift()
  firstSlide.remove()
  
  var oDiapoCibles = oDiaporamaCible.getSlides()
  var qDiapo = oDiapoCibles.length
  
  // OK, maintenant on fusionne les données dans les diapos
  var iDiapo = 0
  for (iRow = 0; iRow < iLastRow; iRow++) {
    for (var islide=0; islide<slidesAll.length; islide++) {
      // Recherche des colonnes dans le document et fusion des données
      for (var key in iCols) {
        sCell = ("" + sDatas[iRow][iCols[key]]).trim()
        //oDiapoCibles[islide].replaceAllText("{$date}", fx_frenchDate(new Date()));
        oDiapoCibles[iDiapo].replaceAllText("{" + key + "}", sCell)
      } // endfor key
      iDiapo++
    }
  } // endfor tableur
  oDiaporamaCible.saveAndClose();

  // affichage d'un panneau pour ouvrir le document crée
  var url = oDiaporamaCible.getUrl()
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="' + url + '" target="_blank">Voir le résultat</a>')
    .setWidth(300)
    .setHeight(100)
  SlidesApp.getUi().showModalDialog(htmlOutput, "Script terminé")
  
}