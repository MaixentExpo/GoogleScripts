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
  for (iCol = 0; iCol <= iLastCol; iCol++) {
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
}

function fx_fx_frenchDate(date) {
  var month = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
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
}

/**
 * diapoSommaire
 */
function diapoSommaire() {
  var presentation = SlidesApp.getActivePresentation();
  // Retrieve slides as images
  var id = presentation.getId();
  //var accessToken = ScriptApp.getOAuthToken();
  // Recup de l'id de chaque dispo
  var pageObjectIds = presentation.getSlides().map(function (e) { return e.getObjectId() });
  // construction des Urls
  // url: "https://slides.googleapis.com/v1/presentations/" + id + "/pages/" + pageObjectId + "/thumbnail?access_token=" + accessToken,
  var reqUrls = pageObjectIds.map(function (pageObjectId) {
    return {
      method: "get",
      url: "https://slides.googleapis.com/v1/presentations/" + id + "/pages/" + pageObjectId + "/thumbnail",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
  });
  // Soumission des Urls
  var reqBlobs = UrlFetchApp.fetchAll(reqUrls).map(function (e) {
    var r = JSON.parse(e);
    return {
      method: "get",
      url: r.contentUrl
    };
  });
  var reqClean = [];
  for (var i=0; i<reqBlobs.length; i++) {
    if ( typeof reqBlobs[i].url === 'undefined' ) {
      ;
    } else {
      reqClean.push(reqBlobs[i]);
    }
  } // endfor
  // Recup des Images générées dans blobs
  var blobs = UrlFetchApp.fetchAll(reqClean).map(function (e) {
    return e.getBlob()
  });

  // Ajout de slides Sommaire
  var col = 5; // Number of columns
  var row = 4; // Number of rows
  var wsize = 130; // Size of width of each image (pixels)
  var sep = 5; // Space of each image (pixels)

  var ph = presentation.getPageHeight(); // 540 px
  var pw = presentation.getPageWidth();  // 720 px
  var leftOffset = (pw - ((wsize * col) + (sep * (col - 1)))) / 2;
  if (leftOffset < 0) throw new Error("Images are sticking out from a slide.");
  var len = col * row;
  var loops = Math.ceil(blobs.length / (col * row));
  for (var loop = 0; loop < loops; loop++) {
    var ns = presentation.insertSlide(loop);
    var topOffset, top;
    var left = leftOffset;
    for (var i = len * loop; i < len + (len * loop); i++) {
      if (i === blobs.length) break;
      var image = ns.insertImage(blobs[i]);
      var w = image.getWidth();
      var h = image.getHeight();
      var hsize = h * wsize / w;
      if (i === 0 || i % len === 0) {
        topOffset = (ph - ((hsize * row) + sep)) / 2;
        if (topOffset < 0) throw new Error("Images are sticking out from a slide.");
        top = topOffset;
      }
      image.setWidth(wsize).setHeight(hsize).setTop(top).setLeft(left).getObjectId();
      //if (i === col - 1 + (loop * len)) {
      if ( loop % col === 0 ) {
        top = topOffset + hsize + sep;
        left = leftOffset;
      } else {
        left += wsize + sep;
      }
    }
  }
  presentation.saveAndClose();
} // end function diapoSommaire