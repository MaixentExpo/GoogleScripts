/**
 * repartirBouteillesSurTables
 * source : https://github.com/MaixentExpo/GoogleScripts/edit/master/MaixentExpoApp/repartirBouteillesSurTables.js
 * Fonction qui organise la répartition des BOUTEILLES à déguster
 * sur les différentes TABLES du concours de VINS
 * La répartition se basera sur la classification du vin référencée par la colonne "Couleur" dans la feuille
 * A blanc, B rosé, C rouge, D moelleux, E pétillant, F champagne
 * La lette qui précède la couleur servira à référencer la bouteille A01 pour par exemple nommée la 1ère bouteille de blanc
 * La lettre servi aussi à trier les couleurs de façon à goûter les bouteilles dans l'ordre alphabétique
 * Le nombre de Tables est calculé en fonction du nombre de bouteilles et du NB_BOUTEILLES_PAR_TABLE
 * Les colonnes lues dans la feuille VINS sont :
 * - Vignoble
 * - Couleur
 * La fonction va mettre à jour les colonnes suivantes :
 * - Bouteille : la référence de la bouteille pour le concours
 * - Table     : le n° de la table pour le concours
 */
function repartirBouteillesSurTables() {
  // Constantes de la fonction
  var NB_BOUTEILLES_PAR_TABLE = 6

  // Ouverture de la feuille VINS
  var spreadsheet = SpreadsheetApp.getActive()
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
  // car les performances sont désastreuses si des accès en lecture et écriture sur la feuille sont réalisés dans la même boucle
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