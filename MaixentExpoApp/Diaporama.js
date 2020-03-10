
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
 * firstLine   : n° de la 1ère ligne des data
 */
function diapo_createDiaporamaFromSlide(sheetId, sheetName, filterName, filterValue) {
  return diapo_createDiaporamaFromSlide2(sheetId, sheetName, filterName, filterValue, 2)
}
function diapo_createDiaporamaFromSlide2(sheetId, sheetName, filterName, filterValue, firstLine) {
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
  for (iCol = 0; iCol <= iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim();
    if (sCell != "") {
      iCols[sCell] = iCol;
    } // endif
  } // endfor
  // On ne prend que les lignes qui correspondent au critère filterName filterValue
  var sDatas = [];
  iLastRow = sValues.length;
  var reFilter = new RegExp(filterValue, 'g');
  for (iRow = firstLine - 1; iRow < iLastRow; iRow++) {
    if (("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow]);
    } // endif
  } // endfor
  iLastRow = sDatas.length;

  // Récupération de la diapo Modèle
  var fileModele = DriveApp.getFileById(SlidesApp.getActivePresentation().getId());
  // Création du Diaporama en sortie
  var sCopyName = ""
  if (fileModele.getName().match(" Modèle")) {
    sCopyName = fileModele.getName().replace(" Modèle", " Pub");
  } else {
    sCopyName = fileModele.getName() + "- Pub";
  } // endif
  var fileCopy = fileModele.makeCopy(sCopyName);
  var oDiaporamaCible = SlidesApp.openById(fileCopy.getId());
  var oDiapoCibles = oDiaporamaCible.getSlides();

  // duplication de la 1ère diapo autant que d'enregistrement-1
  for (iRow = 1; iRow < iLastRow; iRow++) {
    oDiaporamaCible.appendSlide(oDiapoCibles[0]);
  } // endfor

  // OK, maintenant on fusionne les données dans les diapos
  oDiapoCibles = oDiaporamaCible.getSlides();
  var iDiapo = 0;
  for (iRow = 0; iRow < iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for (var key in iCols) {
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();
      oDiapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
      oDiapoCibles[iDiapo].replaceAllText("{" + key + "}", sCell)
    } // endfor key
    iDiapo++;
  } // endfor tableur
  oDiaporamaCible.saveAndClose();

  // affichage d'un panneau pour ouvrir le document crée
  var url = oDiaporamaCible.getUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="' + url + '" target="_blank">Voir le résultat</a>')
    .setWidth(300)
    .setHeight(100);
  SlidesApp.getUi().showModalDialog(htmlOutput, "Script terminé");

  return oDiaporamaCible.getId();
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
  for (iCol = 0; iCol < iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim();
    if (sCell != "") {
      iCols[sCell] = iCol;
    } // endif
  } // endfor

  // Récupération de la diapo Modèle
  var fileModele = DriveApp.getFileById(SlidesApp.getActivePresentation().getId());
  // Création du Diaporama en sortie
  var sCopyName = ""
  if (fileModele.getName().match(" Modèle")) {
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
  for (iRow = 1; iRow < iLastRow; iRow++) {
    if (("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow]);
    } // endif
  } // endfor
  iLastRow = sDatas.length;
  // duplication de la 1ère diapo autant que d'enregistrement / 4
  for (iRow = 4; iRow < iLastRow; iRow += 4) {
    oDiaporamaCible.appendSlide(oDiapoCibles[0]);
  } // endfor

  // OK, maintenant on fusionne les données dans les diapos
  oDiapoCibles = oDiaporamaCible.getSlides();
  var iLigne = 1;
  var iDiapo = 0;
  var sKey = "";
  for (iRow = 0; iRow < iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for (var key in iCols) {
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();
      for (var iCol = 1; iCol < 3; iCol++) {
        sKey = "{" + key + iLigne + iCol + "}";
        oDiapoCibles[iDiapo].replaceAllText(sKey, sCell);
        oDiapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
      } // endfor
    } // endfor key
    iLigne++;
    if (iLigne > 4) {
      iDiapo++;
      iLigne = 1;
    } // endif
  } // endfor tableur
  oDiaporamaCible.saveAndClose();

  // affichage d'un panneau pour ouvrir le document crée
  var url = oDiaporamaCible.getUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="' + url + '" target="_blank">Voir le résultat</a>')
    .setWidth(300)
    .setHeight(100);
  SlidesApp.getUi().showModalDialog(htmlOutput, "Script terminé");

  return oDiaporamaCible.getId();
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
  for (iCol = 0; iCol < iLastCol; iCol++) {
    sCell = ("" + sValues[iRow][iCol]).trim();
    if (sCell != "") {
      iCols[sCell] = iCol;
    } // endif
  } // endfor

  // Récupération de la diapo Modèle
  var fileModele = DriveApp.getFileById(SlidesApp.getActivePresentation().getId());
  // Création du Diaporama en sortie
  var sCopyName = ""
  if (fileModele.getName().match(" Modèle")) {
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
  for (iRow = 1; iRow < iLastRow; iRow++) {
    if (sValues[iRow][iCols[filterName]].match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow]);
    } // endif
  } // endfor
  iLastRow = sDatas.length;
  // duplication de la 1ère diapo autant que d'enregistrement / 4
  for (iRow = 4; iRow < iLastRow; iRow += 4) {
    oDiaporamaCible.appendSlide(oDiapoCibles[0]);
  } // endfor

  // OK, maintenant on fusionne les données dans les diapos
  oDiapoCibles = oDiaporamaCible.getSlides();
  var iLigne = 1;
  var iDiapo = 0;
  var sKey = "";
  for (iRow = 0; iRow < iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for (var key in iCols) {
      sCell = ("" + sDatas[iRow][iCols[key]]).trim();
      for (var iCol = 1; iCol < 3; iCol++) {
        sKey = "{" + key + iLigne + iCol + "}";
        oDiapoCibles[iDiapo].replaceAllText(sKey, sCell);
        oDiapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
      } // endfor
    } // endfor key
    iLigne++;
    if (iLigne > 4) {
      iDiapo++;
      iLigne = 1;
    } // endif
  } // endfor tableur
  oDiaporamaCible.saveAndClose();

  // affichage d'un panneau pour ouvrir le document crée
  var url = oDiaporamaCible.getUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<a href="' + url + '" target="_blank">Voir le résultat</a>')
    .setWidth(300)
    .setHeight(100);
  SlidesApp.getUi().showModalDialog(htmlOutput, "Script terminé");

  return oDiaporamaCible.getId();
}

/**
 * diapo_publipostage : Scripts de publipostage
 * Les parmètres du publipostage sont dans une feuille à partir de la cellule B2
 * Fichier des données     : l'url du tableur des données
 * Feuille des données     : la feuille qui contient les données
 * Ligne des données       : à partir de cette ligne
 * Colonne à filtrer       : colonne qui sert au filtrage éventuel
 * Valeur du filtre        : expression régulière du filtre
 * Répertoire du résultat  : répertoire du fichier résultat
 * Nom du fichier résultat : nom du fichier résultat
 * Convertir en Pdf        : option pour convertir en pdf le résultat
 * Lettre                  : url du 1er fichier à inclure dans le résultat
 * Annexe 1                : url de l'annexe 1
 * Annexe 2                : etc
 * Annexe 3                : etc
 * Annexe 4                : etc
 */
function diapo_publipostage() {
  var ui = SpreadsheetApp.getUi(); //
  var yesnoConfirm = ui.alert(
    "PUBLIPOSTAGE",
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
  var docs = []
  docs.push(piloteSheet.getRange("B10").getValue())
  docs.push(piloteSheet.getRange("B11").getValue())
  docs.push(piloteSheet.getRange("B12").getValue())
  docs.push(piloteSheet.getRange("B13").getValue())
  docs.push(piloteSheet.getRange("B14").getValue())

  // Chargement des données
  var spreadsheet = SpreadsheetApp.openByUrl(sheetUrl)
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
  for (iRow = parseInt(firstLineData) - 1; iRow < iLastRow; iRow++) {
    if (("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow])
    } // endif
  } // endfor
  iLastRow = sDatas.length
  if (iLastRow == 0) {
    ui.alert("PUBLIPOSTAGE", "Aucun enregistrement trouvé", ui.ButtonSet.OK)
    return
  }

  // Récupération de la lettre et des annexes
  var slidesAll = []
  var diaporamaModele = null // le 1er document servira de modèle au diaporama cible
  for (var idoc in docs) {
    if (docs[idoc] != "") {
      var presentation = SlidesApp.openByUrl(docs[idoc])
      if (diaporamaModele == null) {
        diaporamaModele = presentation
      }
      var slides = presentation.getSlides()
      for (var islide = 0; islide < slides.length; islide++) {
        slidesAll.push(slides[islide])
      }
    }
  }

  // Duplication du diaporama modèle
  var fileModele = DriveApp.getFileById(diaporamaModele.getId())
  const regex = /.*\/folders\/(.*)/g
  var folderId = regex.exec(folderUrl)[1]
  var folderCopy = DriveApp.getFolderById(folderId)
  var fileCopy = fileModele.makeCopy(slideName, folderCopy)
  var diaporamaCible = SlidesApp.openById(fileCopy.getId())
  // suppression des slides
  var slides = diaporamaCible.getSlides()
  for (var i in slides) {
    slides[i].remove()
  }

  // duplication des diapos autant de fois que d'enregistrements
  for (iRow = 0; iRow < iLastRow; iRow++) {
    for (var islide = 0; islide < slidesAll.length; islide++) {
      diaporamaCible.appendSlide(slidesAll[islide])
    }
  } // endfor

  var diapoCibles = diaporamaCible.getSlides()

  // OK, maintenant on fusionne les données dans les diapos
  var iDiapo = 0
  for (iRow = 0; iRow < iLastRow; iRow++) {
    for (var islide = 0; islide < slidesAll.length; islide++) {
      // Recherche des colonnes dans le document et fusion des données
      for (var key in iCols) {
        sCell = ("" + sDatas[iRow][iCols[key]]).trim()
        diapoCibles[iDiapo].replaceAllText("{$date}", fx_frenchDate(new Date()));
        diapoCibles[iDiapo].replaceAllText("{" + key + "}", sCell)
        diapoCibles[iDiapo].replaceAllText("{" + key + " €}", Utilities.formatString("%.2f €", parseInt(sCell)) )
      } // endfor key
      iDiapo++
    }
  } // endfor tableur
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

/**
 * diapo_publipostage : Scripts d'édition d'étiquettes
 * à raison de 14 étiquettes par page
 * Les parmètres du publipostage sont dans une feuille à partir de la cellule B2
 * Fichier des données     : l'url du tableur des données
 * Feuille des données     : la feuille qui contient les données
 * Ligne des données       : à partir de cette ligne
 * Colonne à filtrer       : colonne qui sert au filtrage éventuel
 * Valeur du filtre        : expression régulière du filtre
 * Répertoire du résultat  : répertoire du fichier résultat
 * Nom du fichier résultat : nom du fichier résultat
 * Convertir en Pdf        : option pour convertir en pdf le résultat
 * Slide des étiquettes    : url du fichier slide planche des étiquettes 
 */
function diapo_etiquettes() {
  var ui = SpreadsheetApp.getUi(); //
  var yesnoConfirm = ui.alert(
    "PUBLIPOSTAGE EN PDF DES ETIQUETTES",
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

  // Chargement des données
  var spreadsheet = SpreadsheetApp.openByUrl(sheetUrl)
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
  for (iRow = parseInt(firstLineData) - 1; iRow < iLastRow; iRow++) {
    if (("" + sValues[iRow][iCols[filterName]]).match(reFilter, 'g') != null) {
      sDatas.push(sValues[iRow])
    } // endif
  } // endfor
  iLastRow = sDatas.length
  if (iLastRow == 0) {
    ui.alert("PUBLIPOSTAGE", "Aucun enregistrement trouvé", ui.ButtonSet.OK)
    return
  }

  // Récupération de la lettre et des annexes
  var diaporamaModele = SlidesApp.openByUrl(modele)

  // Duplication du diaporama modèle
  var fileModele = DriveApp.getFileById(diaporamaModele.getId())
  const regex = /.*\/folders\/(.*)/g
  var folderId = regex.exec(folderUrl)[1]
  var folderCopy = DriveApp.getFolderById(folderId)
  var fileCopy = fileModele.makeCopy(slideName, folderCopy)
  var diaporamaCible = SlidesApp.openById(fileCopy.getId())
  var diapoCibles = diaporamaCible.getSlides()

  // duplication des diapos à raison de 14 enregistrements par slide
  var qPage = Math.floor(iLastRow / 14)
  for (var islide = 0; islide < qPage; islide++) {
    diaporamaCible.appendSlide(diapoCibles[0])
  } // endfor
  diapoCibles = diaporamaCible.getSlides()

  // OK, maintenant on fusionne les données dans les étiquettes
  var iPage = 0
  var iEtiquette = 1
  for (iRow = 0; iRow < iLastRow; iRow++) {
    // Recherche des colonnes dans le document et fusion des données
    for (var key in iCols) {
      sCell = ("" + sDatas[iRow][iCols[key]]).trim()
      diapoCibles[iPage].replaceAllText("{" + key + iEtiquette + "}", sCell)
    } // endfor
    iEtiquette++
    if (iEtiquette > 14) {
      iPage++;
      iEtiquette = 1;
    } // endif  
  } // endfor tableur
  // sur la dernière page on efface les {...} qui restent
  for (; iEtiquette < 15; iEtiquette++) {
    for (var key in iCols) {
      diapoCibles[iPage].replaceAllText("{" + key + iEtiquette + "}", "")
    } // endfor
  } // endfor
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