/**
 * Création d'un onglet Concours à partir de l'onglet Bovins
 * Colonnes : Eleveurs, Travail, Race, Sexe, N°, Catégorie, Classification
 * La feuille devra être  triée sur le N°
 * Rupture sur catégorie
 */
function updateOngletConcours() {
  // Ouverture du tableur courant
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Ouverture de la feuille BOVINS
  var sheetBovins = spreadsheet.getSheetByName("BOVINS");
  var iLastColBovins = sheetBovins.getLastColumn();
  var iLastRowBovins = sheetBovins.getLastRow();
  var sCategorie = "";
  var sTravail = "";
  var sRace = "";
  var sSexe = "";
  var sNumero = "";
  var sSection = "";
  var sClassification = "";
  var iColCategorie = 0;
  var iColTravail = 0;
  var iColSexe = 0;
  var iColRace = 0;
  var iColNumero = 0;
  var iColSection = 0;
  var iColClassification = 0;
  // Chargement de BOVINS en mémoire dans sBovins
  var sRange = sheetBovins.getRange(1, 1, iLastRowBovins, iLastColBovins).getValues();
  var iLengthBovins = sRange.length;
  // Récupération du nom des colonnes
  var sCell = "";
  var iRow = 0;
  var iCol = 0
  for(; iCol<iLastColBovins; iCol++){
    sCell = sRange[iRow][iCol];
    if ( sCell == "Categorie" ) iColCategorie = iCol;
    if ( sCell == "Section" ) iColSection = iCol;
    if ( sCell == "Classification" ) iColClassification = iCol;
    if ( sCell == "NumTravail" ) iColTravail = iCol;
    if ( sCell == "Race" ) iColRace = iCol;
    if ( sCell == "Sexe" ) iColSexe = iCol;
    if ( sCell == "N°" ) iColNumero = iCol;
  } // endfor
  // Filtre - on ne prend que les ligne avec un N°
  var sBovins = [];
  for (iRow=1; iRow<iLengthBovins; iRow++) {
    if ( sRange[iRow][iColNumero] != "" ) {
      sBovins.push(sRange[iRow]);
    } // endif
  } // endfor
  var iLengthBovins = sBovins.length;
  
  // Ouverture de la feuille CONCOURS
  var sheetConcours = spreadsheet.getSheetByName("CONCOURS")
  var iLastColConcours = sheetConcours.getLastColumn();
  var iLastRowConcours = sheetConcours.getLastRow();
  // Indice des colonnes
  var iCategorie = 1;
  var iTravail = 1;
  var iSexe = 2;
  var iRace = 3;
  var iNumero = 4;
  var iSection = 1;
  var iClassification = 5;
  var rowHeight = 28;
  
  // Nettoyage de la feuille de CONCOURS
  if ( iLastRowConcours > 0 ) sheetConcours.getRange(1, 1, iLastRowConcours, iLastColConcours).clear();
  
  // largeur des colonnes
  sheetConcours.setColumnWidth(iTravail, 64);
  sheetConcours.setColumnWidth(iRace, 64);
  sheetConcours.setColumnWidth(iSexe, 64);
  sheetConcours.setColumnWidth(iNumero, 64);
  sheetConcours.setColumnWidth(iClassification, 296);
  // Lecture de BOVINS et écriture dans CONCOURS
  // Titre de la feuille
  sheetConcours.getRange(1,1,1,1).setValue("CONCOURS BOVINS " + new Date().getFullYear() + " - LISTE POUR LE JURY");
  sheetConcours.setRowHeight(1, rowHeight);
  sheetConcours.getRange(1,1,1,5).mergeAcross().setHorizontalAlignment("center")
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
        sheetConcours.getRange(ir,1,1,5).mergeAcross().setHorizontalAlignment("center").setValue("")
        .setBackground("white").setBorder(null, null, null, null, null, null);
        sheetConcours.setRowHeight(ir, rowHeight);
      // Titre Catégorie - Section
      ir++;
      sheetConcours.getRange(ir,1,1,1).setValue(sCategorie + " - section n° " + sSection);
      sheetConcours.getRange(ir,1,1,5).mergeAcross();
      sheetConcours.setRowHeight(ir, rowHeight+8);
      sheetConcours.getRange(ir,1, 1, 5).setFontColor("black").setFontSize(15).setHorizontalAlignment("center").setBackground(oCouleur.couleur)
      .setBorder(true, true, true, true, true, true);
      // Sous-titre
      ir++;
      sheetConcours.getRange(ir,1,1,5).breakApart();
      sheetConcours.getRange(ir,iTravail).setValue("N° Travail");
      sheetConcours.getRange(ir,iRace).setValue("Race");
      sheetConcours.getRange(ir,iSexe).setValue("Sexe");
      sheetConcours.getRange(ir,iNumero).setValue("N°");
      sheetConcours.getRange(ir,iClassification).setValue("Classification");
      sheetConcours.setRowHeight(ir, rowHeight*2/3);
      sheetConcours.getRange(ir,1, 1, 5).setFontColor("black").setFontSize(8).setHorizontalAlignment("center").setBackground(oCouleur.couleur)
      .setBorder(true, true, true, true, true, true);
      ir++;
    } // endif
    sTravail = sBovins[iRow][iColTravail];
    sRace = sBovins[iRow][iColRace];
    sSexe = sBovins[iRow][iColSexe];
    sNumero = sBovins[iRow][iColNumero];
    sClassification = sBovins[iRow][iColClassification];
    sheetConcours.getRange(ir,iTravail).setValue(sTravail).setNumberFormat("0000");
    sheetConcours.getRange(ir,iRace).setValue(sRace);
    sheetConcours.getRange(ir,iSexe).setValue(sSexe);
    sheetConcours.getRange(ir,iNumero).setValue(sNumero);
    sheetConcours.getRange(ir,iClassification).setValue("");
    sheetConcours.setRowHeight(ir, rowHeight);
    sheetConcours.getRange(ir,1, 1, 5).setFontColor("black").setFontSize(13).setHorizontalAlignment("center").setBackground("white")
    .setBorder(true, true, true, true, true, true);
    ir++;
  } // end for
  
  // Ajout manuel des prix spéciaux définis dans la plage nommée PRIX_SPECIAUX
  var sPrixSpeciaux = spreadsheet.getRangeByName("PRIX_SPECIAUX").getValues();
  var iLengthPrixSpeciaux = sPrixSpeciaux.length;
  oCouleur.new_couleur();
  for ( iRow=0; iRow<iLengthPrixSpeciaux; iRow++) {
    // ligne à blanc
    sheetConcours.getRange(ir,1,1,5).mergeAcross().setHorizontalAlignment("center").setValue("")
    .setBackground("white").setBorder(null, null, null, null, null, null);
    sheetConcours.setRowHeight(ir, rowHeight);
    ir++;
    // Titre
    sheetConcours.getRange(ir,1,1,1).setValue(sPrixSpeciaux[iRow][0]);
    sheetConcours.getRange(ir,1,1,5).mergeAcross()
    sheetConcours.getRange(ir,1,1,5).setFontColor("black").setHorizontalAlignment("center").setFontSize(15).setBackground(oCouleur.couleur)
    .setBorder(true, true, true, true, true, true);
    sheetConcours.setRowHeight(ir, rowHeight+8);
    ir++;
    // Sous-titre
    sheetConcours.getRange(ir,1,1,5).breakApart();
    sheetConcours.getRange(ir,iTravail).setValue("N° Travail");
    sheetConcours.getRange(ir,iRace).setValue("Race");
    sheetConcours.getRange(ir,iSexe).setValue("Sexe");
    sheetConcours.getRange(ir,iNumero).setValue("N°");
    sheetConcours.getRange(ir,iClassification).setValue("Classification");
    sheetConcours.setRowHeight(ir, rowHeight*2/3);
    sheetConcours.getRange(ir,1, 1, 5).setFontColor("black").setFontSize(8).setHorizontalAlignment("center").setBackground(oCouleur.couleur)
    .setBorder(true, true, true, true, true, true);
    ir++;
    // Données
    sheetConcours.setRowHeight(ir, rowHeight);
    sheetConcours.getRange(ir,1, 1, 5).setFontColor("black").setFontSize(13).setHorizontalAlignment("center").setBackground("white")
    .setBorder(true, true, true, true, true, true);
    ir++;
  } // endfor

  sheetConcours.activate();
}
