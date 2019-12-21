/*
  ComiteCommissions
  Retourne la liste des commisions affectée au Membres
*/
function comite_Commissions() {
  var oVar = {
    style_normal : SpreadsheetApp.newTextStyle().setBold(false).build(),
    style_red : SpreadsheetApp.newTextStyle().setForegroundColor("red").build(),
  }
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
  // Nettoyage des commissions
  sheet.getRange(2, iColCommissions+1).clear({contentsOnly: true})
  
  // Ouverture de la feuiile COMMISSIONS
  var sheetCommissions = spreadsheet.getSheetByName("COMMISSIONS")
  var iLastColCommissions = sheetCommissions.getLastColumn()
  var iLastRowCommissions = sheetCommissions.getLastRow()
  var valuesCommissions = sheetCommissions.getRange(2,1,iLastRowCommissions,iLastColCommissions).getValues();

  // Mise à jour de la colonne Commissions de la feuille COMITE
  var sNom = ""
  var sCommissions = ""
  var bAffecte = false
  var aStyles = []
  for(iRow=1; iRow<iLastRow; iRow++) {
    sNom = valuesMembres[iRow][iColNom]
    if (sNom == "") 
      break
    sCommissions = ""
    bAffecte = false
    aStyles = Array()
    // Recherche du nom dans les commissions
    for(var ir=0; ir<iLastRowCommissions-1; ir++) {
      // boucle des affectations
      for(var ic=1; ic<iLastColCommissions; ic++) {
        var oStyle = {}
        if (sNom == valuesCommissions[ir][ic]) {
          if (bAffecte) {
            sCommissions += " & "
          } // endif
          bAffecte = true
          if ( ic == 1 ) {
            // Colonne responsable commission
            oStyle["start"] = sCommissions.length
          }
          sCommissions += valuesCommissions[ir][0]
          if ( ic == 1 ) {
            oStyle["end"] = sCommissions.length
            aStyles.push(oStyle)
          }
        } // endif
      } // end for
    } // end for
    sheet.getRange(iRow+1, iColCommissions+1).setRichTextValue(getRichTextBold(sCommissions, aStyles, oVar))
    valuesMembres[iRow][iColCommissions] = sCommissions
  } // endfor
  
} // end ComiteCommissions

function getRichTextBold(textValue, aStyles, oVar) {
  var textRich = SpreadsheetApp.newRichTextValue()
  textRich = textRich.setText(textValue)
  var oStyle = {}
  for (var i=0; i<aStyles.length; i++) {
    oStyle = aStyles[i]
    textRich = textRich.setTextStyle(oStyle.start, oStyle.end, oVar.style_bold)
    textRich = textRich.setTextStyle(oStyle.start, oStyle.end, oVar.style_red)
  }
  return textRich.build()
}
