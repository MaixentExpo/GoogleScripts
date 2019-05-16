/**
 * Création d'un onglet Lots à partir de l'onglet Bovins
 * Colonnes : Eleveurs, Travail, Race, Sexe, N°, Catégorie, Classification
 * La feuille devra être triée sur le N°
 * Rupture sur Catégorie Section
 */
function updateOngletLots() {
  // Ouverture du tableur courant
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Ouverture de la feuille BOVINS
  var sheetBovins = spreadsheet.getSheetByName("BOVINS")
  var iLastColBovins = sheetBovins.getLastColumn()
  var iLastRowBovins = sheetBovins.getLastRow()
  var sEleveur = "";
  var sCategorie = "";
  var sTravail = "";
  var sRace = "";
  var sSexe = "";
  var sNumero = "";
  var sSection = "";
  var sClassification = "";
  var iColEleveur = 0;
  var iColCategorie = 0;
  var iColTravail = 0;
  var iColRace = 0;
  var iColSexe = 0;
  var iColNumero = 0;
  var iColSection = 0;
  var iColClassification = 0;
  // Chargement de BOVINS en mémoire dans sBovins
  var sRange = sheetBovins.getRange(1, 1, iLastRowBovins, iLastColBovins).getValues();
  var iLengthBovins = sRange.length;
  // Récupération de la position des colonnes
  var sCell = "";
  var iRow = 0;
  var iCol = 0
  for(; iCol<iLastColBovins; iCol++){
    sCell = sRange[iRow][iCol];
    if ( sCell == "Eleveur" ) iColEleveur = iCol;
    if ( sCell == "Categorie" ) iColCategorie = iCol;
    if ( sCell == "Section" ) iColSection = iCol;
    if ( sCell == "Classification" ) iColClassification = iCol;
    if ( sCell == "NumTravail" ) iColTravail = iCol;
    if ( sCell == "Race" ) iColRace = iCol;
    if ( sCell == "Sexe" ) iColSexe = iCol;
    if ( sCell == "N°" ) iColNumero = iCol;
  } // endfor
  // Filtre - on ne prend que les ligne avec une classification
  var sBovins = [];
  for (iRow=1; iRow<iLengthBovins; iRow++) {
    if ( sRange[iRow][iColClassification] != "" ) {
      sBovins.push(sRange[iRow]);
    } // endif
  } // endfor
  var iLengthBovins = sBovins.length;
  
  // Ouverture de la feuille LOTS
  var sheetLots = spreadsheet.getSheetByName("LOTS")
  var iLastColLots = sheetLots.getLastColumn()
  var iLastRowLots = sheetLots.getLastRow()
  // Indice des colonnes
  var iCategorie = 1;
  var iSection = 1;
  var iEleveur = 1;
  var iTravail = 2;
  var iRace = 3;
  var iSexe = 4;
  var iClassification = 5;
  var iLot = 6;
  var rowHeight = 28;
  
  // Nettoyage de la feuille de LOTS
  if ( iLastRowLots > 0 ) sheetLots.getRange(1, 1, iLastRowLots, iLastColLots).clear();
  
  // largeur des colonnes
  sheetLots.setColumnWidth(iEleveur, 200);
  sheetLots.setColumnWidth(iTravail, 64);
  sheetLots.setColumnWidth(iRace, 64);
  sheetLots.setColumnWidth(iSexe, 64);
  sheetLots.setColumnWidth(iClassification, 200);
  sheetLots.setColumnWidth(iLot, 400);
  // Lecture de BOVINS et écriture dans LOTS
  // Titre de la feuille
  sheetLots.getRange(1,1,1,1).setValue("CONCOURS BOVINS " + new Date().getFullYear() + " - LISTE POUR LA REMISE DES PRIX");
  sheetLots.setRowHeight(1, rowHeight);
  sheetLots.getRange(1,1,1,6).mergeAcross().setHorizontalAlignment("center")
  .setFontColor("red").setFontSize(13).setBackground("white").setBorder(false, false, false, false, false, false);
  
  var ruptureCategorie = "";
  var ruptureSection = "";
  var ir = 2;
  var oCouleur = new Couleur();
  for (iRow=0; iRow<iLengthBovins; iRow++) {
    sCategorie = sBovins[iRow][iColCategorie];
    sSection = sBovins[iRow][iColSection];
    if ( sCategorie != ruptureCategorie || sSection != ruptureSection ) {
      // Rupture Catégorie ou Section
      if ( sCategorie != ruptureCategorie ) {
        // changement de couleur
        oCouleur.new_couleur();
      } // endif
      ruptureCategorie = sCategorie;
      ruptureSection = sSection;
      // ligne à blanc
      if ( ir > 2 )  
        sheetLots.getRange(ir,1,1,6).mergeAcross().setHorizontalAlignment("center").setValue("")
        .setBackground("white").setBorder(null, null, null, null, null, null);
        sheetLots.setRowHeight(ir, rowHeight);
      // Titre Catégorie - Section
      ir++;
      sheetLots.getRange(ir,1,1,1).setValue(sCategorie + " - section n° " + sSection);
      sheetLots.getRange(ir,1,1,6).mergeAcross();
      sheetLots.setRowHeight(ir, rowHeight+8);
      sheetLots.getRange(ir,1,1,6).setFontColor("black").setFontSize(15).setHorizontalAlignment("center").setBackground(oCouleur.couleur)
      .setBorder(true, true, true, true, true, true);
      // Sous-titre
      ir++;
      sheetLots.getRange(ir,1,1,4).breakApart();
      sheetLots.getRange(ir,iEleveur).setValue("Eleveur");
      sheetLots.getRange(ir,iTravail).setValue("N° Travail");
      sheetLots.getRange(ir,iRace).setValue("Race");
      sheetLots.getRange(ir,iSexe).setValue("Sexe");
      sheetLots.getRange(ir,iClassification).setValue("Classification");
      sheetLots.getRange(ir,iLot).setValue("Commentaires");
      sheetLots.setRowHeight(ir, rowHeight*2/3);
      sheetLots.getRange(ir,1,1,6).setFontColor("black").setFontSize(8).setHorizontalAlignment("center").setBackground(oCouleur.couleur)
      .setBorder(true, true, true, true, true, true);
      ir++;
    } // endif
    sEleveur = sBovins[iRow][iColEleveur];
    sTravail = sBovins[iRow][iColTravail];
    sRace = sBovins[iRow][iColRace];
    sSexe = sBovins[iRow][iColSexe];
    sClassification = sBovins[iRow][iColClassification];
    sheetLots.getRange(ir,iEleveur).setValue(sEleveur).setHorizontalAlignment("left");
    sheetLots.getRange(ir,iTravail).setValue(sTravail).setNumberFormat("0000").setHorizontalAlignment("center");
    sheetLots.getRange(ir,iRace).setValue(sRace).setHorizontalAlignment("center");
    sheetLots.getRange(ir,iSexe).setValue(sSexe).setHorizontalAlignment("center");
    sheetLots.getRange(ir,iClassification).setValue(sClassification).setHorizontalAlignment("center");
    sheetLots.getRange(ir,iLot).setValue("");
    sheetLots.setRowHeight(ir, rowHeight);
    sheetLots.getRange(ir,1,1,6).setFontColor("black").setFontSize(13).setBackground("white")
    .setBorder(true, true, true, true, true, true);
    ir++;
  } // end for
  
  sheetLots.activate();
}
