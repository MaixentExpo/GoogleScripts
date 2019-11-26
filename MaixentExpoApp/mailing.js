/*
  Script de Mailing
*/
function prepareMessage() {
  prepareMessageResultat("15gPNGjf_Sga1Ips11NccEeSUq7X8eJjEiOHPMooT5jw", "VINS", "RESULTAT");
}

/*
  messagePresseConcoursVins
  Retourne le message à envoyer à la presse
  Les colonnes devront êtres préalablement triées sur Medaille, Couleur, Vin
*/
function prepareMessageResultat(source_file_id, sheet_name, cible_range_name) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var yesnoConfirm = ui.alert(
     "Préparer le message",
     'Veuillez confirmer par oui ou non',
      ui.ButtonSet.YES_NO);
  if ( yesnoConfirm != ui.Button.YES ) return;

  // Ouverture de la feuille VINS
  var spreadsheet_source = SpreadsheetApp.openById(source_file_id);
  var sheet = spreadsheet_source.getSheetByName(sheet_name)
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var range = sheet.getRange(1,1,1,iLastCol); // 1ère ligne
  var values = range.getValues();
  var iVignoble = 0;
  var iProducteur = 0;
  var iRegion = 0;
  var iVin = 0;
  var iCouleur = 0;
  var iMedaille = 0;
  var iTriMedaille = 0;
  var sCell = "";
  var iRow = 0;
  var iCol = 0
  for(; iCol<iLastCol; iCol++){
    sCell = values[iRow][iCol];
    if ( sCell == "Vignoble" ) iVignoble = iCol+1;
    if ( sCell == "Producteur" ) iProducteur = iCol+1;
    if ( sCell == "Region" ) iRegion = iCol+1;
    if ( sCell == "Vin" ) iVin = iCol+1;
    if ( sCell == "Couleur" ) iCouleur = iCol+1;
    if ( sCell == "Medaille" ) iMedaille = iCol+1;
    if ( sCell == "TriMedaille" ) iTriMedaille = iCol+1;
  } // endfor
  
  // TRI des colonnes Medaille, Couleur, Vin
  range = sheet.getRange(2, 1, iLastRow, iLastCol);
  range.sort([{column: iTriMedaille, ascending: true}, {column: iCouleur, ascending: true}, {column: iVin, ascending: true}]);

  // Boucle avec rupture sur les médailles
  var sVignoble = "";
  var sProducteur = "";
  var sRegion = "";
  var sVin = "";
  var sCouleur = "";
  var sMedaille = "";
  var sRuptureMedaille = "";
  var sResultat = "";
  // var rich = SpreadsheetApp.newRichTextValue();
  // var iPos = 0;
  var bFirst = true;
  iRow = 2; // on commence sur la 2ème ligne
  for(; iRow<iLastRow; iRow++) {
    sMedaille = sheet.getRange(iRow, iMedaille).getValue().trim();
    if ( sMedaille == "" ) 
      break;
    sVignoble = sheet.getRange(iRow, iVignoble).getValue();
    sProducteur = sheet.getRange(iRow, iProducteur).getValue();
    sRegion = sheet.getRange(iRow, iRegion).getValue();
    sVin = sheet.getRange(iRow, iVin).getValue();
    sCouleur = sheet.getRange(iRow, iCouleur).getValue();
    
    if ( sMedaille != sRuptureMedaille ) {
      if (bFirst == false) sResultat += "\n\n"
      sResultat += sMedaille == "or" ? "Médaille d'or" : sMedaille == "argent" ? "Médaille d'argent" : "Médaille de bronze";
      sResultat += " :\n";
      bFirst = true;
      sRuptureMedaille = sMedaille;
    } // endif
    if (bFirst == false) sResultat += "\n"
    bFirst = false
    sResultat += "- " + sProducteur + " (" + sVin + ") " + sVignoble
  } // endfor

  // Ouverture du fichier mailing courant
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Copie du resulat dans la cible
  var cible = spreadsheet.getRange(cible_range_name);
  cible.setValue(sResultat);
  
} // end messagePresseConcoursBovins

function recupEmail() {
  fx_recupEmail("15gPNGjf_Sga1Ips11NccEeSUq7X8eJjEiOHPMooT5jw", "EMAILS_JURYS", "CC");
}

