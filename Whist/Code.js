function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Whist')
  .addItem('Afficher les colonnes cachées', 'demasquerLesColonnes')
  .addItem('Initialiser le jeu', 'initJeu')
  .addItem('Attribuer les médailles', 'calculMedailles')
    .addToUi();
}

/**
 * Démasquer les colonnes
 */
function demasquerLesColonnes() {
  // Ouverture de la feuille courante
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getActiveSheet()
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  // parcours de toutes les colonnes
  for (var ic = 1; ic <= iLastCol; ic++) {
    if (sheet.isColumnHiddenByUser(ic)) {
      sheet.unhideColumn(sheet.getRange(1,ic, 1, iLastCol-ic+1));
      break
    } // endif
  }
}

/**
 * Initilaisation du jeu
 * en fonction des joueurs renseignés :
 * - va cacher les colonnes des joueurs non renseignés
 * - va remettre à blanc les colonnes Dem. et Plis
 */
function initJeu() {
  // Ouverture de la feuille courante
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getActiveSheet()
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var range = sheet.getRange(1, 1, 1, iLastCol) // 1 ligne
  var values = range.getValues()
  // rang des colonnes
  var iColJoueur = 0

  var iRow = 0
  var iCol = 0
  var sCell = ""
  var qJoueurs = 0
  for (iCol = 0; iCol < iLastCol; iCol++) {
    sCell = values[0][iCol];
    if (sCell == "c1") {
      iColJoueur = iCol;
      qJoueurs = 1
    }
    if (sCell == "c2") qJoueurs = 2
    if (sCell == "c3") qJoueurs = 3
    if (sCell == "c4") qJoueurs = 4
    if (sCell == "c5") qJoueurs = 5
    if (sCell == "c6") qJoueurs = 6
    if (sCell == "c7") qJoueurs = 7
    if (sCell == "c8") qJoueurs = 8
    if (sCell == "c9") qJoueurs = 9
    if (sCell == "c10") qJoueurs = 10
    if (sCell == "c11") qJoueurs = 11
  } // endfor
  var qColonnes = qJoueurs // colonnes qu'il faudra initialiser à blanc au début

  // Chargement du tableau des joueurs
  range = sheet.getRange(2, 1, 1, iLastCol) // 2ème ligne
  values = range.getValues()
  var ij = 0
  for (iRow = 0; iRow < values.length; iRow++) {
    for (ij = 0; ij < qJoueurs; ij++) {
      // Si le nom du joueur est à blanc on ajuste le nombre de joueur
      if (values[iRow][iColJoueur + ij * 4].length == 0) {
        qJoueurs = ij
        break
      } // endif
    } // end for
  } // endfor
  var ic = iColJoueur + ij * 4 + 1 // colonne à masquer plus loin
  for (; ij < qColonnes; ij++) {
    range.getCell(1, iColJoueur + ij * 4 + 1).setBackground("white");
  } // end for
  // on masque les colonnes non utilisées
  sheet.hideColumns(ic, iLastCol-ic+1);
  // Mise à blanc des colonnes Dem. et Plis
  for (ic=iColJoueur; ic < iLastCol; ic+=4) {
    sheet.getRange(4, ic+1, iLastRow-3, 2).clear({contentsOnly: true}) // Dem. et Plis
  } // end for

}

/**
 * Calcul des médailles Or Argent Bronze Chocolat
 */
function calculMedailles() {
  // Ouverture de la feuille courante
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getActiveSheet()
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var range = sheet.getRange(1, 1, 1, iLastCol) // 1 ligne
  var values = range.getValues()
  // rang des colonnes
  var iColJoueur = 0
  var iColPlis = 0

  var iRow = 0
  var iCol = 0
  var sCell = ""
  var qJoueurs = 0
  for (iCol = 0; iCol < iLastCol; iCol++) {
    sCell = values[0][iCol];
    if (sCell == "plis") iColPlis = iCol;
    if (sCell == "c1") {
      iColJoueur = iCol;
      qJoueurs = 1
    }
    if (sCell == "c2") qJoueurs = 2
    if (sCell == "c3") qJoueurs = 3
    if (sCell == "c4") qJoueurs = 4
    if (sCell == "c5") qJoueurs = 5
    if (sCell == "c6") qJoueurs = 6
    if (sCell == "c7") qJoueurs = 7
    if (sCell == "c8") qJoueurs = 8
    if (sCell == "c9") qJoueurs = 9
    if (sCell == "c10") qJoueurs = 10
    if (sCell == "c11") qJoueurs = 11
  } // endfor
  var qColonnes = qJoueurs // colonnes qu'il faudra initialiser à blanc au début

  // Chargement du tableau des joueurs
  range = sheet.getRange(2, 1, 1, iLastCol) // 2ème ligne
  values = range.getValues()
  var aJoueurs = []
  var ij = 0
  for (iRow = 0; iRow < values.length; iRow++) {
    for (ij = 0; ij < qJoueurs; ij++) {
      // Si le nom du joueur est à blanc on ajuste le nombre de joueur
      if (values[iRow][iColJoueur + ij * 4].length > 0) {
        var joueur = {}
        joueur["id"] = ij
        joueur["name"] = values[iRow][iColJoueur + ij * 4]
        aJoueurs[ij] = joueur
      } else {
        qJoueurs = ij
        break
      } // endif
    } // end for
  } // endfor

  // lecture des totaux
  range = sheet.getRange(4, 1, iLastRow - 3, iLastCol) // 3 lignes d'entête à sauter
  values = range.getValues()
  var iPlis = 0
  var qLignes = 0 // nombre de parties remplies
  for (iRow = 0; iRow < values.length; iRow++) {
    iPlis = values[iRow][iColPlis]
    if (iPlis == "")
      break
    qLignes++
    // boucle sur les joueurs
    for (ij = 0; ij < qJoueurs; ij++) {
      aJoueurs[ij]["total"] = values[iRow][iColJoueur + ij * 4 + 3]
    } // end for
  } // endfor
  // Tri du tableau
  aJoueurs.sort(function (a, b) { // 
    return b.total - a.total;
  });

  // Attribution des médailles Or Argent Bronze
  var iPointsOr = -1
  var iPointsArgent = -1
  var iPointsBronze = -1
  for (ij = 0; ij < qJoueurs; ij++) {
    if (iPointsOr == -1) {
      iPointsOr = aJoueurs[ij]["total"]
      aJoueurs[ij]["medaille"] = "or"
      continue
    } // endif
    if (aJoueurs[ij]["total"] == iPointsOr) {
      aJoueurs[ij]["medaille"] = "or"
      continue
    } // endif
    if (iPointsArgent == -1) {
      iPointsArgent = aJoueurs[ij]["total"]
      aJoueurs[ij]["medaille"] = "argent"
      continue
    } // endif
    if (aJoueurs[ij]["total"] == iPointsArgent) {
      aJoueurs[ij]["medaille"] = "argent"
      continue
    } // endif
    if (iPointsBronze == -1) {
      iPointsBronze = aJoueurs[ij]["total"]
      aJoueurs[ij]["medaille"] = "bronze"
      continue
    } // endif
    if (aJoueurs[ij]["total"] == iPointsBronze) {
      aJoueurs[ij]["medaille"] = "bronze"
      continue
    } // endif
  } // endfor

  // Attribution de la médaille de Chocolat
  var iPointsChocolat = -1
  for (ij = qJoueurs - 1; ij > 0; ij--) {
    if (iPointsChocolat == -1) {
      iPointsChocolat = aJoueurs[ij]["total"]
      if (!aJoueurs[ij]["medaille"]) {
        aJoueurs[ij]["medaille"] = "chocolat"
      } // endif
      continue
    } // endif
    if (aJoueurs[ij]["total"] == iPointsChocolat) {
      if (!aJoueurs[ij]["medaille"]) {
        aJoueurs[ij]["medaille"] = "chocolat"
      } // endif
      continue
    } // endif
  } // endfor

  // Mise à jour des couleurs
  range = sheet.getRange(2, 1, 1, iLastCol) // 2ème ligne
  values = range.getValues()
  for (iRow = 0; iRow < values.length; iRow++) {
    for (ij = 0; ij < qJoueurs; ij++) {
      if (aJoueurs[ij].medaille == "or") {
        range.getCell(1, iColJoueur + aJoueurs[ij].id * 4 + 1).setBackground("yellow");
      } else if (aJoueurs[ij].medaille == "argent") {
        range.getCell(1, iColJoueur + aJoueurs[ij].id * 4 + 1).setBackground("lightGrey");
      } else if (aJoueurs[ij].medaille == "bronze") {
        range.getCell(1, iColJoueur + aJoueurs[ij].id * 4 + 1).setBackground("darkKhaki");
      } else if (aJoueurs[ij].medaille == "chocolat") {
        range.getCell(1, iColJoueur + aJoueurs[ij].id * 4 + 1).setBackground("chocolate");
      } else {
        range.getCell(1, iColJoueur + aJoueurs[ij].id * 4 + 1).setBackground("white");
      } // endif
    } // end for
    // fond blanc pour le reste des colonnes
    var ic = iColJoueur + ij * 4 + 1 // colonne à masquer plus loin
    for (; ij < qColonnes; ij++) {
      range.getCell(1, iColJoueur + ij * 4 + 1).setBackground("white");
    } // end for
    // on masque les colonnes non utilisées - fonction abandonée car certaines cellules sont protégées
    //sheet.hideColumns(ic, iLastCol-ic+1);
  } // endfor
} // end function
