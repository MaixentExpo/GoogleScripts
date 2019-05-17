/**
 * attribuerLesPlacesAuRepas
 * 
 */
function attribuerLesPlacesAuRepas() {
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