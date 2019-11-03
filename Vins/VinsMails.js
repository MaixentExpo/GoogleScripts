/*
  messagePresseConcoursVins
  Retourne le message à envoyer à la presse
  Les colonnes devront êtres préalablement triées sur Medaille, Couleur, Vin
*/
function messagePresseConcoursVins(e) {
  // Ouverture de la feuille VINS
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName("VINS")
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
  var s = "";
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
      s += '\n';
      s += sMedaille == "or" ? "Médaille d'or" : sMedaille == "argent" ? "Médaille d'argent" : "Médaille de bronze";
      s += " : ";
      bFirst = true;
      sRuptureMedaille = sMedaille;
    } // endif
    if (bFirst == false) s += ", "
    bFirst = false
    s += sProducteur + " (" + sVin + ") " + sVignoble
      
  } // endfor
  //return(s);
  
  // Mise à jour de la cellule MAIL_PRESSE
  var sheetMail = spreadsheet.getSheetByName("MAILS")
  sCell = sheetMail.getRange("MAIL_PRESSE");
  sCell.setValue(s);
  
} // end messagePresseConcoursBovins
