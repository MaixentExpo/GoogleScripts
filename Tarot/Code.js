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
  var range = sheet.getRange(1,1,1,iLastCol) // 1ère ligne
  var values = range.getValues()
  // rang des colonnes
  var iColPartie = 0
  var iColResultat = 0
  var iColDonneur = 0
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
    if ( sCell == "Partie" ) iColPartie = iCol;
    if ( sCell == "Donneur" ) iColPartie = iCol;
    if ( sCell == "Résultat" ) iColResultat = iCol;
    if ( sCell == "Preneur" ) iColPreneur = iCol;
    if ( sCell == "Partenaire" ) iColPartenaire = iCol;
    if ( sCell == "Contrat" ) iColContrat = iCol;
    if ( sCell == "Nombre de bouts" ) iColBouts = iCol;
    if ( sCell == "Points" ) iColPoints = iCol;
    if ( sCell == "Petit au bout" ) iColPetit = iCol;
    if ( sCell == "Poignée" ) iColPoignee = iCol;
    if ( sCell == "Double Poignée" ) iColDoublePoignee = iCol;
    if ( sCell == "Triple Poignée" ) iColTriplePoignee = iCol;
    if ( sCell == "Chelem annoncé" ) iColChelemAnnonce = iCol;
    if ( sCell == "Chelem non annoncé" ) iColChelemNonAnnonce = iCol
  } // endfor

  // lecture des ligne parties
  range = sheet.getRange(2,1,iLastRow,iLastCol)
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
  
  // Mise à jour de la colonne résultat
  for(iRow=0; iRow<iLastRow; iRow++) { 
    sheet.getRange(iRow+2,iColResultat+1).setValue(values[iRow][iColResultat])
  }
}
