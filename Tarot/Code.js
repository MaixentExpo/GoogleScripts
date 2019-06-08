function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Tarots')
      .addItem('Initialiser le jeux', 'initJeux')
      .addItem('Calculer les parties', 'calculParties')
      .addToUi();
}

function initJeux() {

}

function calculParties() {
  // Ouverture de la feuille courante
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getActiveSheet()
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var range = sheet.getRange(1,1,2,iLastCol) // 2ère lignes
  var values = range.getValues()
  // rang des colonnes
  var iColPartie = 0
  var iColJoueur = 0 // colonne j1,j2,j3,j4,j5
  var iColResultat = 0
  var iColPreneur = 0
  var iColPartenaire = 0
  var iColContrat = 0
  var iColBouts = 0
  var iColPoints = 0
  var iColPetit = 0
  var iColPoignee = 0
  var iColDoublePoignee = 0
  var iColTriplePoignee = 0
  var iColChelemAnnonce = 0
  var iColChelemNonAnnonce = 0
  
  var iRow = 0
  var iCol = 0
  var sCell = ""
  for(iCol=0; iCol<iLastCol; iCol++){
    sCell = values[0][iCol];
    if ( sCell == "partie" ) iColPartie = iCol;
    if ( sCell == "j1" ) iColJoueur = iCol;
    if ( sCell == "donneur" ) iColPartie = iCol;
    if ( sCell == "resultat" ) iColResultat = iCol;
    if ( sCell == "preneur" ) iColPreneur = iCol;
    if ( sCell == "partenaire" ) iColPartenaire = iCol;
    if ( sCell == "contrat" ) iColContrat = iCol;
    if ( sCell == "bouts" ) iColBouts = iCol;
    if ( sCell == "points" ) iColPoints = iCol;
    if ( sCell == "petit" ) iColPetit = iCol;
    if ( sCell == "poignee" ) iColPoignee = iCol;
    if ( sCell == "doublePoignee" ) iColDoublePoignee = iCol;
    if ( sCell == "triplePoignee" ) iColTriplePoignee = iCol;
    if ( sCell == "chelemAnnonce" ) iColChelemAnnonce = iCol;
    if ( sCell == "chelemNonAnnonce" ) iColChelemNonAnnonce = iCol
  } // endfor

  // tableau des joueurs
  var aJoueurs = []
  for(iCol=iColJoueur; iCol<iColJoueur+5; iCol++){
    sCell = values[1][iCol]; // 2ème ligne d'entête
    if ( sCell.length > 0 ) {
      aJoueurs.push(sCell)
    }
  } // end for
  var qJoueurs = aJoueurs.length // nombre de joueurs

  // lecture des ligne parties
  range = sheet.getRange(3,1,iLastRow,iLastCol) // 2 lignes d'entête à sauter
  values = range.getValues()
  var sPreneur = ""
  var sPartenaire = ""
  var sContrat = ""
  var iNombreBouts = 0
  var iPoints = 0
  var sPetit = ""
  var sPoignee = ""
  var sDoublePoignee = ""
  var sTriplePoignee = ""
  var bChelemAnnonce = false
  var bChelemNonAnnonce = false

  var iScoreCible = 0
  var iScore = 0
  for(iRow=0; iRow<iLastRow; iRow++){
    sPreneur = values[iRow][iColPreneur]
    if ( sPreneur == "" ) 
      continue
    sPartenaire = values[iRow][iColPartenaire]
    sContrat = values[iRow][iColContrat]
    iNombreBouts = values[iRow][iColBouts]
    iPoints = values[iRow][iColPoints]
    sPetit = values[iRow][iColPetit]
    sPoignee = values[iRow][iColPoignee]
    sDoublePoignee = values[iRow][iColDoublePoignee]
    sTriplePoignee = values[iRow][iColTriplePoignee]
    bChelemAnnonce = values[iRow][iColChelemAnnonce]
    bChelemNonAnnonce = values[iRow][iColChelemNonAnnonce]
    // Calcul du Score à atteindre
    switch (iNombreBouts) {
      case 0:
        iScoreCible = 56
        break
      case 1:
        iScoreCible = 51
        break
      case 2:
        iScoreCible = 41
        break
      case 3:
        iScoreCible = 36
        break              
    }
    // Calcul du coefficient
    var coeff = 1
    switch (sContrat) {
      case "Petite":
        coeff = 1
        break
      case "Garde":
        coeff = 2
        break
      case "Garde Sans":
        coeff = 4
        break
      case "Garde Contre":
        coeff = 6
        break
    }
    // calcul du score
    iPoints = iPoints - iScoreCible
    iScore = iPoints < 0 ? (iPoints-25)*coeff : (iPoints+25)*coeff
    // Calcul du score / chelem
    if ( bChelemAnnonce ) iScore = iScore > 0 ? iScore + 400 : iScore - 200
    if ( bChelemNonAnnonce ) iScore += 200
    // Calcul petit au bout
    if ( sPetit == sPreneur || sPetit == sPartenaire ) 
      iScore = iScore > 0 ? iScore + 10*coeff : iScore + 10*coeff
    else
      if ( sPetit.length > 0)
        iScore = iScore > 0 ? iScore - 10*coeff : iScore - 10*coeff
    // Calcul poignée
    if ( sPoignee == sPreneur || sPoignee == sPartenaire ) 
      iScore = iScore > 0 ? iScore + 20 : iScore - 20
    else
      if ( sPoignee.length > 0)
        iScore = iScore > 0 ? iScore - 20 : iScore - 20
    // Calcul double poignée
    if ( sDoublePoignee == sPreneur || sDoublePoignee == sPartenaire ) 
      iScore = iScore > 0 ? iScore + 30 : iScore - 30
    else
      if ( sDoublePoignee.length > 0)
        iScore = iScore > 0 ? iScore - 30 : iScore - 30
    // Calcul triple poignée
    if ( sTriplePoignee == sPreneur || sTriplePoignee == sPartenaire ) 
      iScore = iScore > 0 ? iScore + 40 : iScore - 40
    else
      if ( sTriplePoignee.length > 0)
        iScore = iScore > 0 ? iScore - 40 : iScore - 40

    values[iRow][iColResultat] = iScore

  } // endfor
  
  // Mise à jour des colonne résultat j1 j2 j3 j4
  var aCumuls = []
  for(iRow=0; iRow<iLastRow; iRow++) { 
    if ( values[iRow][iColPreneur].length < 1 ) 
      continue
    sheet.getRange(iRow+3,iColResultat+1).setValue(values[iRow][iColResultat])
    // Distribution des points aux joueurs
    for (var ij = 0; ij<qJoueurs; ij++) {
      if ( aJoueurs[ij] == values[iRow][iColPreneur]) {
        switch (qJoueurs) {
          case 4:
            iPoints = values[iRow][iColResultat] * 3
            break
          case 5:
            iPoints = values[iRow][iColResultat] * 2
            break;
        } // end switch
      } else if ( aJoueurs[ij] == values[iRow][iColPartenaire]) {
        iPoints = values[iRow][iColResultat] * 1
      } else {
        iPoints = values[iRow][iColResultat] * -1
      }// endif
      // Cumul des points
      if (aCumuls[ij]) 
        aCumuls[ij] += iPoints 
      else 
        aCumuls[ij] = iPoints
      sheet.getRange(iRow+3,iColJoueur+1+ij).setValue(aCumuls[ij])
    } // endfor
  } // endfor
} // end function
