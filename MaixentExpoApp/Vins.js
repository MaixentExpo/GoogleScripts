/*
  messagePresseConcoursVins
  Retourne le message à envoyer à la presse
  Les colonnes devront êtres préalablement triées sur Medaille, Couleur, Vin
*/
function vins_prepareMessageResultat(source_file_id, sheet_name, cible_range_name) {
    var ui = SpreadsheetApp.getUi(); // Same variations.
    var yesnoConfirm = ui.alert(
       "Préparer le message",
       'Veuillez confirmer par Oui ou Non',
        ui.ButtonSet.YES_NO);
    if ( yesnoConfirm != ui.Button.YES ) return;
  
    // Ouverture de la feuille VINS
    var spreadsheet_source = SpreadsheetApp.openById(fx_getIdFromUrl(source_file_id));
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
  
  /**
 * repartirBouteillesSurTables
 * source : https://github.com/MaixentExpo/GoogleScripts/edit/master/MaixentExpoApp/repartirBouteillesSurTables.js
 * Fonction qui organise la répartition des BOUTEILLES à déguster
 * sur les différentes TABLES du concours de VINS
 * La répartition se basera sur la classification du vin référencée par la colonne "Couleur" dans la feuille
 * A blanc, B rosé, C rouge, D moelleux, E pétillant, F champagne
 * La lette qui précède la couleur servira à référencer la bouteille (A01 par exemple pour nommée la 1ère bouteille de blanc)
 * La lettre servira aussi à trier les couleurs de façon à goûter les bouteilles dans l'ordre alphabétique
 * Le nombre de Tables est calculé en fonction du nombre de bouteilles et du NB_BOUTEILLES_PAR_TABLE
 * Les colonnes lues dans la feuille VINS sont :
 * - Vignoble
 * - Couleur
 * La fonction va mettre à jour les colonnes suivantes :
 * - Bouteille : la référence de la bouteille pour le concours
 * - Table     : le n° de la table pour le concours
 */
function vins_repartirBouteillesSurTables() {
  var ui = SpreadsheetApp.getUi(); //
  var yesnoConfirm = ui.alert(
    "REPARTIR LES BOUTEILLES",
    'Veuillez confirmer par Oui ou Non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  // Ouverture du Spreadsheet courant
  var spreadsheet = SpreadsheetApp.getActive()

  // Récupéartion des paramètres
  var NB_BOUTEILLES_PAR_TABLE = spreadsheet.getRangeByName("PL_BOUTEILLES_PAR_TABLE").getValue()

  // Ouverture de la feuille VINS
  var sheet = spreadsheet.getSheetByName("VINS")
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var range = sheet.getRange(1, 1, 1, iLastCol);
  var values = range.getValues()
  var iColVignoble = 0
  var iColCouleur = 0
  var iColBouteille = 0
  var iColTable = 0
  var sCell = ""
  var iRow = 0
  var iCol = 0
  for(; iCol<=iLastCol; iCol++){
    sCell = values[iRow][iCol];
    if ( sCell == "Vignoble" ) iColVignoble = iCol+1
    if ( sCell == "Couleur" ) iColCouleur = iCol+1
    if ( sCell == "Bouteille" ) iColBouteille = iCol+1
    if ( sCell == "Table" ) iColTable = iCol+1
  } // endfor
  
  // TRI des colonnes sur Couleur et Vignoble
  range = sheet.getRange(2, 1, iLastRow, iLastCol);
  range.sort([{column: iColCouleur, ascending: true},{column: iColVignoble, ascending: true}])
  
  // Calcul du nombre de tables à raison de NB_BOUTEILLES_PAR_TABLE bouteilles par table
  var qTable = Math.floor((iLastRow - 1) / NB_BOUTEILLES_PAR_TABLE) // on ne prend que la valeur entière

  // Boucle avec rupture sur les couleurs
  var sCouleur = ""
  var sruptureCouleur = ""
  var iBouteille = 0
  var lettreCouleur = ""
  var sBouteille = ""
  var iTable = 0;
  // Les données calculées Bouteille et Table sont d'abord enregistrées en mémoire
  // car les performances sont désastreuses si des accès en lecture et écriture sur la feuille sont réalisés en alternance
  var sUpdates = [] 
  for(iRow=2; iRow<=iLastRow; iRow++) { // on commence à la ligne 2
    sCouleur = sheet.getRange(iRow, iColCouleur).getValue().trim()  
    if ( sCouleur != sruptureCouleur ) {
      sruptureCouleur = sCouleur
      iBouteille = 0
      lettreCouleur = sCouleur.substring(0,1);
    } // endif
    iBouteille++;
    sBouteille = lettreCouleur + Utilities.formatString("%02d", iBouteille)
    iTable++
    // Mise à jour des colonnes Bouteilles et Tables
    sUpdates.push([sBouteille, iTable]);
    if ( iTable >= qTable )
      iTable = 0      
  } // endfor
  // Mise à jour des colonnes de la feuille
  for(iRow=2; iRow<=iLastRow; iRow++) { 
    sheet.getRange(iRow, iColBouteille).setValue(sUpdates[iRow-2][0])
    sheet.getRange(iRow, iColTable).setValue(sUpdates[iRow-2][1])
  } // endfor
    
} // end repartirBouteillesSurTables
