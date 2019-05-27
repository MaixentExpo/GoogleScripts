/*
  ComiteCommissions
  Retourne la liste des commisions affectée au Menmbres
*/
function ComiteCommissions() {
  // Ouverture du tableur conteneur du script
  var spreadsheet = SpreadsheetApp.getActive();
  // Ouverture de la feuille MEMBRES
  var sheet = spreadsheet.getSheetByName("MEMBRES")
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var valuesMembres = sheet.getRange(1,1,iLastRow,iLastCol).getValues();
  // Recherche de la position des colonnes dans valuesComite
  var iColNom = 0;
  var iColCommissions = 0; 
  var sCell = ""
  var iRow = 0
  var iCol = 0
  for(iRow=0,iCol=0; iCol<iLastCol; iCol++){
    sCell = valuesMembres[iRow][iCol];
    if ( sCell == "Nom" ) iColNom = iCol;
    if ( sCell == "Commissions" ) iColCommissions = iCol;
  } // endfor
  
  // Ouverture de la feuiile COMMISSIONS
  var sheetCommissions = spreadsheet.getSheetByName("COMMISSIONS")
  var iLastColCommissions = sheetCommissions.getLastColumn()
  var iLastRowCommissions = sheetCommissions.getLastRow()
  var valuesCommissions = sheetCommissions.getRange(2,1,iLastRowCommissions,iLastColCommissions).getValues();

  // Mise à jour de la colonne Commissions de la feuille COMITE
  var sNom = ""
  var sCommissions = ""
  var oCell
  var bAffecte = false
  var aBolds = []
  var jBold = {}
  for(iRow=1; iRow<iLastRow; iRow++) {
    sNom = valuesMembres[iRow][iColNom]
    if (sNom == "") 
      break
    sCommissions = ""
    bAffecte = false
    aBolds = Array()
    // Recherche du nom dans les commissions
    for(var ir=0; ir<iLastRowCommissions-2; ir++) {
      // boucle des affectations
      for(var ic=1; ic<iLastColCommissions; ic++) {
        if (sNom == valuesCommissions[ir][ic]) {
          if (bAffecte) {
            sCommissions += " & "
          } // endif
          bAffecte = true
          if ( ic == 1 ) {
            // Colonne responsable commission
            jBold["start"] = sCommissions.length
          }
          if ( valuesCommissions[ir][0] == "Présidence" ) {
            if ( ic > 1 ) {
              sCommissions += "Vice-présidence"
            } else {
              sCommissions += "Présidence"
            } // endif
          } else {
            sCommissions += valuesCommissions[ir][0]  
          } // endif
          if ( ic == 1 ) {
            jBold["end"] = sCommissions.length
            aBolds.push(jBold)
          }
        } // endif
      } // end for
    } // end for
    sheet.getRange(iRow+1, iColCommissions+1).setRichTextValue(getRichTextBold(sCommissions, aBolds))
    valuesMembres[iRow][iColCommissions] = sCommissions
  } // endfor
  
  
} // end ComiteCommissions

function getRichTextBold(textValue, aBolds) {
  var bold = SpreadsheetApp.newTextStyle().setBold(true).build()
  var red = SpreadsheetApp.newTextStyle().setForegroundColor("red").build()
  var textRich = SpreadsheetApp.newRichTextValue()
  textRich = textRich.setText(textValue)
  var jBold = {}
  for (var i=0; i<aBolds.length; i++) {
    jBold = aBolds[i]
    textRich = textRich.setTextStyle(jBold.start, jBold.end, bold)
    textRich = textRich.setTextStyle(jBold.start, jBold.end, red)
  }
  return textRich.build()
}
