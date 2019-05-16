/*
  messagePresseConcoursBovins
  Retourne le message à envoyer à la presse
  Les colonnes devront êtres préalablement triées sur Categorie, Section, Classification
  On ne prend que les lignes avec Travail non vide
*/
function messagePresseConcoursBovins() {
  // Ouverture de la feuille BOVINS
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName("BOVINS")
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  // Recherche de la position des colonnes sur la 1ère ligne
  
  var iCategorie = 0;
  var iSection = 0;
  var iClassification = 0;
  var iEleveur = 0;
  var iTravail = 0;
  
  var sCategorie = "";
  var sSection = "";
  var sClassification = "";
  var sEleveur = "";
  var sTravail = "";

  var sCell = "";
  var iRow = 1;
  var iCol = 1
  for(; iCol<=iLastCol; iCol++){
    sCell = sheet.getRange(iRow, iCol).getValue().trim();
    if ( sCell == "Categorie" ) iCategorie = iCol;
    if ( sCell == "Section" ) iSection = iCol;
    if ( sCell == "Classification" ) iClassification = iCol;
    if ( sCell == "Eleveur" ) iEleveur = iCol;
    if ( sCell == "NumTravail" ) iTravail = iCol;
  } // endfor

  // TRI des colonnes Categorie, Section, Classification
  var range = sheet.getRange(2, 1, iLastRow, iLastCol);
  range.sort([{column: iCategorie, ascending: true}, {column: iSection, ascending: true}, {column: iClassification, ascending: true}]);

  var s = "";
  // Grand prix d'excellence
  iRow = 2;
  for(; iRow<=iLastRow; iRow++) {
    sTravail = sheet.getRange(iRow, iTravail).getValue().trim();
    if ( sTravail == "" ) break;
    sClassification = sheet.getRange(iRow, iClassification).getValue().trim();
    if ( sClassification.indexOf("Grand Prix") == -1 ) 
      continue;  
    sEleveur = sheet.getRange(iRow, iEleveur).getValue().trim();
    sCategorie = sheet.getRange(iRow, iCategorie).getValue().trim();
    s += sClassification + " " + sCategorie + " : " + sEleveur;
  } // endfor
  // Prix d'excellence
  iRow = 2;
  for(; iRow<=iLastRow; iRow++) {
    sTravail = sheet.getRange(iRow, iTravail).getValue().trim();
    if ( sTravail == "" ) break;
    sClassification = sheet.getRange(iRow, iClassification).getValue().trim();
    if ( sClassification.indexOf("Excellence") == -1 ) 
      continue;  
    if ( sClassification.indexOf("Excellence") > 7 )
      continue;
    sEleveur = sheet.getRange(iRow, iEleveur).getValue().trim();
    sCategorie = sheet.getRange(iRow, iCategorie).getValue().trim();
    s += "\n";
    s += sClassification + " " + sCategorie + " : " + sEleveur;
  } // endfor

  // Boucle avec rupture sur Catégorie, Section
  var sRuptureCategorie = "";
  var sRuptureSection = "";  
  var bSection = true;
  iRow = 2; // on commence sur la 2ème ligne
  for(; iRow<iLastRow; iRow++) {
    sTravail = sheet.getRange(iRow, iTravail).getValue().trim();
    if ( sTravail == "" ) break;
    sCategorie = sheet.getRange(iRow, iCategorie).getValue().trim();
    if ( sCategorie.indexOf("peser") != -1 || sCategorie.indexOf("absent") != -1 ) 
      continue;
    sSection = "" + sheet.getRange(iRow, iSection).getValue();
    if ( sCategorie != sRuptureCategorie ) {
      s += '\n';
      s += "Race " + sCategorie + " : ";
      bSection = true;
      sRuptureCategorie = sCategorie;
      sRuptureSection = "";
    } // endif
    if ( sSection != sRuptureSection ) {
      if ( bSection == false ) s += '. ';
      s += sSection == "1" ? "1re section : " : sSection + "e section : ";       
      bSection = true;
      sRuptureSection = sSection;
    } // endif
    sClassification = "" + sheet.getRange(iRow, iClassification).getValue().trim();
    sEleveur = "" + sheet.getRange(iRow, iEleveur).getValue().trim();
    if (bSection == false) s += ", "
    bSection = false
    s += sClassification.indexOf("Honneur") != -1 || sClassification.indexOf("Excellence") != -1 ? sClassification : sClassification + " prix";
    s += " " + sEleveur
  } // endfor

  // Mise à jour de la cellule MAIL_PRESSE
  spreadsheet.getSheetByName("MAILS").getRange("MAIL_PRESSE").setValue(s);  
  
} // end messagePresseConcoursBovins
