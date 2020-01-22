/**
 * bovins_prepareMessageResultatConcours
 * Retourne le message à envoyer à la presse dans "cible_range_name"
 * Les colonnes devront êtres préalablement triées sur Categorie, Section, Classification
 * On ne prend que les lignes avec Travail non vide
 * @param {String} source_file_id : id du Spreadsheet
 * @param {String} sheet_name : nom de l'onglet des données à analyser
 * @param {String} cible_range_name : plage nommée de réception du résultat du Spreadsheet courant
 */
function bovins_prepareMessageResultatConcours(source_file_id, sheet_name, cible_range_name) {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
     "Préparer le message",
     'Veuillez confirmer par oui ou non',
      ui.ButtonSet.YES_NO);
  if ( yesnoConfirm != ui.Button.YES ) return;

  // Ouverture de la feuille BOVINS
  var spreadsheet_source = SpreadsheetApp.openById(source_file_id);
  var sheet = spreadsheet_source.getSheetByName(sheet_name)
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
  s += "\n";
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

  // Ouverture du fichier mailing courant
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Copie du resulat dans la cible
  var cible = spreadsheet.getRange(cible_range_name);
  cible.setValue(s);

} // end bovins_prepareMessageResultatConcours

/**
 * Création d'un onglet Concours à partir de l'onglet Bovins
 * Colonnes : Eleveurs, Travail, Race, Sexe, N°, Catégorie, Classification
 * La feuille devra être  triée sur le N°
 * Rupture sur catégorie
 */
function bovins_updateOngletConcours() {
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

/**
 * Création d'un onglet Lots à partir de l'onglet Bovins
 * Colonnes : Eleveurs, Travail, Race, Sexe, N°, Catégorie, Classification
 * La feuille devra être triée sur le N°
 * Rupture sur Catégorie Section
 */
function bovins_updateOngletLots() {
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

