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

/**
 * Fonction : attribuerLesPlacesAuRepas
 * Source : https://github.com/MaixentExpo/GoogleScripts/edit/master/MaixentExpoApp/AttribuerLesPlacesAuRepas.js
 * Script qui va attribuer les places des convives sur les tables de la salle
 * La feuille "Inscriptions" comportera au moins les colonnes suivantes :
 * en lecture
 * Groupe : Nom éventuel d'un groupe pour regrouper des sous-groupes de personnes
 * Nom    : Nom de la personne qui a inscrit un sous-groupe de personnes
 * Nombre : Nombre de personne pour le sous-groupe
 * Zone souhaitée : zone souhaitée par le groupe (nom d'une REF_ZONE)
 * en écriture
 * Table       : N° de la table attribué par le système
 * Place début : à partir de ce n° de place
 * Place fin   : jusqu'à ce n° de place
 * La feuille "Technique" indiquera :
 * REF_ZONES : la plage nommée des différentes zones sur la tables
 * La feuille "Tables" représentera graphiquement les tables avec leur n° de place et les REF_ZONES
 * Chaque table sera une plage nommée TABLE_1, TABLE_2, .. TABLE_N
 * nom du groupe qui va occuper la place impair, n°place impair, ref zone, n°place pair, nom du groupe qui va occuper la place pair
 */
function bovins_attribuerLesPlacesAuRepas() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast("Calcul en cours...");
  // Ouverture des feuilles et chargement des données en mémoire

  // Onglet Inscriptions
  var sheetInscription = spreadsheet.getSheetByName("Inscriptions");
  var iLastRowInscription = sheetInscription.getLastRow();
  var iLastColInscription = sheetInscription.getLastColumn();
  var oRangeInscriptions = sheetInscription.getRange(1, 1, iLastRowInscription, iLastColInscription);
  var sRangeInscriptions = oRangeInscriptions.getValues();
  var iRowDataInscriptions = 4;
  // Lecture de la ligne d'entête pour mémoriser le nom des colonnes et leur position
  var iCols = {};
  var sCell = "";
  var iRow = 0;
  var iCol = 0;
  for(iCol=0; iCol<iLastColInscription; iCol++) {
    sCell = ("" + sRangeInscriptions[iRow][iCol]).trim();
    if ( sCell != "" ) {
      iCols[sCell] = iCol;
    } // endif
  } // endfor
  // on supprimme les 3 lignes d'entête
  var sInscriptions = [];
  for ( iRow=3; iRow < iLastRowInscription; iRow++ ) {
    if ( sRangeInscriptions[iRow][iCols["Nom"]] != "" ) 
      sInscriptions.push(sRangeInscriptions[iRow]);
  } // endfor
  var iLengthInscriptions = sInscriptions.length;

  // Lecture des plages nommées de l'onglet Tables
  // Plage nommée TABLES_N : nom impair, n°place impair, ref zone, n°place pair, nom pair
  // pour construire le tableau sPlaces[] n°table, RefZone, n°place
  var sPlaces = [];
  var iColNumPlace=0, iColTablePlace=1, iColZonePlace=2, iColNomPlace=3;
  var numPlace=0, tablePlace=0, zonePlace=0, nomPlace="";
  var iRowPlace = 0, iLengthPlaces = 0;

  var sheetTables = spreadsheet.getSheetByName("Tables");
  var plTables = sheetTables.getNamedRanges();
  var plLengthPlaces = plTables.length;
  var iColNomImpair = 0, iColNumImpair = 1, iColRefZone = 2, iColNumPair = 3, iColNomPair = 4;
  var iNomImpair = 0, iNumImpair = 1, iRefZone = "", iNumPair = 3, iNomPair = 4;
  plTables.sort(function (a,b) { // tri TABLE_1, TABLE_2, ...
    return a.getName() > b.getName() ? 1 : a.getName() < b.getName() ? -1 : 0;
  });
  iRowPlace = 0;
  var plName = "", plRanges, iTable = "", plRefZone = ""; 
  for ( var ipl=0; ipl < plLengthPlaces; ipl++) {
    plName = plTables[ipl].getName();
    iTable = plName.match(/TABLE_(\d)/, 'g')[1];
    plRanges = plTables[ipl].getRange().getValues();
    for (var ir=0; ir<plRanges.length; ir++) {
      if (plRanges[ir][iColRefZone]!="") 
        plRefZone = plRanges[ir][iColRefZone];
      sPlaces[iRowPlace++] = new Array(plRanges[ir][iColNumImpair], iTable, plRefZone, "");
      sPlaces[iRowPlace++] = new Array(plRanges[ir][iColNumPair], iTable, plRefZone, "");
    } // endfor
  } // endfor
  iLengthPlaces = sPlaces.length;

  // Affectation de la même zone aux inscrits du même groupe
  // + maj du groupe avec le nom si le groupe est vide
  var groupeZones = {};
  var groupeIns = "";
  var nomIns = "";
  var tableZoneIns = "";
  for ( iRow=0; iRow < iLengthInscriptions; iRow++ ) {
    groupeIns = sInscriptions[iRow][iCols["Groupe"]];
    nomIns = sInscriptions[iRow][iCols["Nom"]];
    tableZoneIns = sInscriptions[iRow][iCols["TableZone"]];
    if ( groupeIns != "" ) {
      if ( groupeZones[groupeIns] == null ) {
        groupeZones[groupeIns] = tableZoneIns;
      } else {
        sInscriptions[iRow][iCols["TableZone"]] = groupeZones[groupeIns];
      } // endif
    } else {
      sInscriptions[iRow][iCols["Groupe"]] = nomIns;
    } // endif
  } // endfor


  // Affectation des places des inscrits avec zone prédéfinie
  var nombreIns = 0;
  var tableIns = 0;
  var zoneIns = 0
  for ( iRow=0; iRow < iLengthInscriptions; iRow++ ) {
    groupeIns = sInscriptions[iRow][iCols["Groupe"]];
    nomIns = sInscriptions[iRow][iCols["Nom"]];
    tableZoneIns = sInscriptions[iRow][iCols["TableZone"]];
    if ( tableZoneIns != "" ) {
      nombreIns = sInscriptions[iRow][iCols["Nombre"]];
      // recup de la table et zone souhaitées
      var rf = tableZoneIns.match(/T(\d)(\d)/, 'g');
      tableIns = rf[1];
      zoneIns = rf[2];
      // remplissage des places
      for (iRowPlace=0; iRowPlace<iLengthPlaces; iRowPlace++) {
        place = sPlaces[iRowPlace][iColNumPlace];
        tablePlace = sPlaces[iRowPlace][iColTablePlace];
        zonePlace = sPlaces[iRowPlace][iColZonePlace];
        nomPlace = sPlaces[iRowPlace][iColNomPlace];
        if ( nomPlace != "" )
          continue;
        // place disponible
        // est-ce la bonne table et bonne zone ?
        if ( tablePlace == tableIns && zonePlace >= zoneIns ) { // on peut déborder sur la zone suivante de la même table
          sPlaces[iRowPlace][iColNomPlace] = nomIns;
          nombreIns--;
        }
        if ( nombreIns < 1 ) 
          break; // toutes les places demandées ont été attribuées
      } // endfor
    } // endif
  } // endfor

  // Calcul du nombre de places pour le groupe
  var tmpGroupes = {};
  for ( iRow=0; iRow < iLengthInscriptions; iRow++ ) {
    groupeIns = sInscriptions[iRow][iCols["Groupe"]];
    nombreIns = sInscriptions[iRow][iCols["Nombre"]];
    if ( tmpGroupes[groupeIns] == null ) tmpGroupes[groupeIns] = 0;
    tmpGroupes[groupeIns] += nombreIns;    
  } // endfor
  // transformation du tableau associatif en tableau simple
  // {"groupe1": 12, "groupe2": 2} en [ ["groupe1"][12],["groupe2"][2] ]
  var iGroupes = [];
  for (var g in tmpGroupes) {
    iGroupes.push([g, tmpGroupes[g]]);
  }
  // Tri de iGroupes
  iGroupes.sort(function (a, b) {
    return b[1] - a[1];
  })

  // Affectation des places des inscrits en zone libre
  // on traite groupe par groupe à partir des plus gros
  var groupeName = "", groupeCount = 0;
  var iTablePlace = 0, iStartPlace = -1, qplace = 0;
  for ( var ig=0, igmax=iGroupes.length ; ig<igmax; ig++) {
    groupeName = iGroupes[ig][0];
    groupeCount = iGroupes[ig][1];
    for ( iRow=0; iRow < iLengthInscriptions; iRow++ ) {
      if ( sInscriptions[iRow][iCols["TableZone"]] == "" 
        && sInscriptions[iRow][iCols["Groupe"]] == groupeName ) { 
        // affectation des places du groupe groupeName
        // recherche de la table qui dispose du nombre de places disponibles pour le groupe
        iStartPlace = -1;
        iTablePlace = -1
        qplace = 0;
        for (iRowPlace=0; iRowPlace<iLengthPlaces; iRowPlace++) {
          nomPlace = sPlaces[iRowPlace][iColNomPlace];
          tablePlace = sPlaces[iRowPlace][iColTablePlace];
          if ( nomPlace == "" ) { // place disponible
            if ( iTablePlace == -1 ) iTablePlace = tablePlace;
            if ( iTablePlace == tablePlace ) {
              if ( iStartPlace == -1 ) iStartPlace = iRowPlace;
              qplace += 1;
            } else {
              // recherche sur table suivante
              iTablePlace = tablePlace;
              iStartPlace = -1;
              qplace = 0;
            } // endif
          } else {
            // dent creuse avec espace insuffisant 
            iTablePlace = -1
            iStartPlace = -1;
            qplace = 0;
          } // endif
          if ( qplace == groupeCount ) {
            // la table avec suffisament de place a été trouvée, youpii !!!
            // attribution des places à partir de iStartPlace
            for ( var ir=0; ir < iLengthInscriptions; ir++ ) {
              nombreIns = sInscriptions[ir][iCols["Nombre"]];
              if ( sInscriptions[ir][iCols["Groupe"]] == groupeName ) { 
                while ( nombreIns > 0) {
                  sPlaces[iStartPlace++][iColNomPlace] = sInscriptions[ir][iCols["Nom"]];
                  nombreIns--;
                } // end while
              } // endif
            } // endfor
            // sortie des 2 for, le groupeName est traité
            iRowPlace = iLengthPlaces;
            iRow = iLengthInscriptions;
          } // endif
        } // endfor Place
        if ( qplace == 0 ) {
          // le groupe n'a pas pu être logé
          SpreadsheetApp.getUi().alert(Utilities.formatString("Le groupe [%s] ne loge pas sur une table", groupeName));
        }
      } // endif groupeName 
    } // endfor Inscriptions
  } // endfor iGroupes

  // Mise à jour de l'onglet Tables et Inscriptions
  var tablePlaceRupture = "";
  var nomPlaceRupture = "";
  var iRowRange = 0;
  var cell;
  var color = "";
  for (iRowPlace=0; iRowPlace<iLengthPlaces; iRowPlace++) {
    numPlace = sPlaces[iRowPlace][iColNumPlace];
    nomPlace = sPlaces[iRowPlace][iColNomPlace];
    tablePlace = sPlaces[iRowPlace][iColTablePlace];
    if ( nomPlaceRupture != nomPlace ) {
      nomPlaceRupture = nomPlace;
      color = color == "aliceblue" ? "lavender" : "aliceblue";
    } // endif
    if ( tablePlaceRupture != tablePlace ) {
      // changement de table donc de plage nommée
      tablePlaceRupture = tablePlace;
      plRanges = plTables[tablePlace-1].getRange();
      iRowRange = 0;
    } // endif
    if ( numPlace % 2 ) {
      cell = plRanges.getCell(iRowRange+1, iColNomImpair+1);
    } else {
      cell = plRanges.getCell(iRowRange+1, iColNomPair+1);
      iRowRange++;
    } // endif
    cell.setValue(nomPlace);
    cell.setFontSize(8);
    cell.setWrap(true);
    if ( nomPlace == "" ) {
      cell.setBackground("darkgray");
    } else {
      cell.setBackground(color);
    } // endif
  } // endfor

  // Mise à jour de l'onglet Inscriptions
  for ( iRow=0; iRow<iLengthInscriptions; iRow++) {
    for (iRowPlace=0; iRowPlace<iLengthPlaces; iRowPlace++) {
      if ( sInscriptions[iRow][iCols["Nom"]] == sPlaces[iRowPlace][iColNomPlace]) {
        oRangeInscriptions.getCell(iRow+iRowDataInscriptions, iCols["Table"]+1).setValue(sPlaces[iRowPlace][iColTablePlace]);
        oRangeInscriptions.getCell(iRow+iRowDataInscriptions, iCols["Debut"]+1).setValue(sPlaces[iRowPlace][iColNumPlace]);
        oRangeInscriptions.getCell(iRow+iRowDataInscriptions, iCols["Fin"]+1).setValue(
          sPlaces[iRowPlace][iColNumPlace] + sInscriptions[iRow][iCols["Nom"]+1] -1);
        break;
      } // endif
    } // endfor
  } // endfor

  spreadsheet.toast("Calcul terminé");
}
