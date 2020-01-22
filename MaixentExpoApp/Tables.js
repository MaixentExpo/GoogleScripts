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
  for (iCol = 0; iCol < iLastColInscription; iCol++) {
    sCell = ("" + sRangeInscriptions[iRow][iCol]).trim();
    if (sCell != "") {
      iCols[sCell] = iCol;
    } // endif
  } // endfor
  // on supprimme les 3 lignes d'entête
  var sInscriptions = [];
  for (iRow = 3; iRow < iLastRowInscription; iRow++) {
    if (sRangeInscriptions[iRow][iCols["Nom"]] != "")
      sInscriptions.push(sRangeInscriptions[iRow]);
  } // endfor
  var iLengthInscriptions = sInscriptions.length;

  // Lecture des plages nommées de l'onglet Tables
  // Plage nommée TABLES_N : nom impair, n°place impair, ref zone, n°place pair, nom pair
  // pour construire le tableau sPlaces[] n°table, RefZone, n°place
  var sPlaces = [];
  // [[1, "A", "A1", ""], [2, "A", "A1", ""], [3, "A", "A1", ""]
  // [[numPlace, tablePlace, zonePlace, nomPlace]]
  var iColNumPlace = 0, iColTablePlace = 1, iColZonePlace = 2, iColNomPlace = 3;
  var numPlace = 0, tablePlace = 0, zonePlace = 0, nomPlace = "";
  var iRowPlace = 0, iLengthPlaces = 0;

  var sheetTables = spreadsheet.getSheetByName("Tables");
  var plTables = sheetTables.getNamedRanges();
  var plLengthPlaces = plTables.length;
  var iColNomImpair = 0, iColNumImpair = 1, iColRefZone = 2, iColNumPair = 3, iColNomPair = 4;
  plTables.sort(function (a, b) { // tri TABLE_1, TABLE_2, ...
    return a.getName() > b.getName() ? 1 : a.getName() < b.getName() ? -1 : 0;
  });
  iRowPlace = 0;
  var plName = "", plRanges, sTable = "", plRefZone = "";
  for (var ipl = 0; ipl < plLengthPlaces; ipl++) {
    plName = plTables[ipl].getName();
    sTable = plName.match(/TABLE_(.)/, 'g')[1]; // A B C .. G
    plRanges = plTables[ipl].getRange().getValues();
    for (var ir = 0; ir < plRanges.length; ir++) {
      if (plRanges[ir][iColRefZone] != "")
        plRefZone = plRanges[ir][iColRefZone];
      if (plRanges[ir][iColNumImpair] != "") {
        sPlaces[iRowPlace++] = new Array(plRanges[ir][iColNumImpair], sTable, plRefZone, "");
        sPlaces[iRowPlace++] = new Array(plRanges[ir][iColNumPair], sTable, plRefZone, "");
      }
    } // endfor
  } // endfor
  iLengthPlaces = sPlaces.length;

  // Affectation de la même zone aux inscrits du même groupe
  // + maj du groupe avec le nom si le groupe est vide
  var groupeZones = {};
  var groupeIns = "";
  var nomIns = "";
  var tableZoneIns = "";
  for (iRow = 0; iRow < iLengthInscriptions; iRow++) {
    groupeIns = sInscriptions[iRow][iCols["Groupe"]];
    nomIns = sInscriptions[iRow][iCols["Nom"]];
    tableZoneIns = sInscriptions[iRow][iCols["TableZone"]];
    if (groupeIns != "") {
      if (groupeZones[groupeIns] == null) {
        groupeZones[groupeIns] = tableZoneIns;
      } else {
        sInscriptions[iRow][iCols["TableZone"]] = groupeZones[groupeIns];
      } // endif
    } else {
      sInscriptions[iRow][iCols["Groupe"]] = nomIns;
    } // endif
  } // endfor

  // Tri des inscriptions sur TableZone
  sInscriptions.sort(function (a, b) { // tri A1 A2 B1 B2 blanc...
    if (a[iCols["TableZone"]] > b[iCols["TableZone"]] && a[iCols["TableZone"]] != "")
      return 1
    else if (a[iCols["TableZone"]] < b[iCols["TableZone"]] && b[iCols["TableZone"]] != "")
      return -1
    else return 0;
    //return a[iCols["TableZone"]] > b[iCols["TableZone"]] ? 1 : a[iCols["TableZone"]] < b[iCols["TableZone"]] ? -1 : 0;
  });


  // Affectation des places des inscrits avec zone prédéfinie
  var nombreIns = 0;
  var tableIns = 0;
  var zoneIns = 0
  for (iRow = 0; iRow < iLengthInscriptions; iRow++) {
    groupeIns = sInscriptions[iRow][iCols["Groupe"]];
    nomIns = sInscriptions[iRow][iCols["Nom"]];
    tableZoneIns = sInscriptions[iRow][iCols["TableZone"]];
    if (tableZoneIns != "") {
      nombreIns = sInscriptions[iRow][iCols["Nombre"]];
      // recup de la table et zone souhaitées A1 A2 B1 B2...G2
      var rf = tableZoneIns.match(/([ABCDEFG])(\d)/, 'g');
      tableIns = rf[1];
      zoneIns = rf[2];
      // remplissage des places
      for (iRowPlace = 0; iRowPlace < iLengthPlaces; iRowPlace++) {
        tablePlace = sPlaces[iRowPlace][iColTablePlace];
        zonePlace = sPlaces[iRowPlace][iColZonePlace];
        nomPlace = sPlaces[iRowPlace][iColNomPlace];
        if (nomPlace != "")
          continue;
        // place disponible
        // est-ce la bonne table et bonne zone ?
        if (tablePlace == tableIns && zonePlace >= zoneIns) { // on peut déborder sur la zone suivante de la même table
          // Option affichage du nom du groupe ou du nom de l'inscrit
          //sPlaces[iRowPlace][iColNomPlace] = nomIns;
          sPlaces[iRowPlace][iColNomPlace] = groupeIns;
          nombreIns--;
        }
        if (nombreIns < 1)
          break; // toutes les places demandées ont été attribuées
      } // endfor
      if ( nombreIns > 0 ) {
        // il manque des places pour le groupe
        var mess = Utilities.formatString("Le groupe [%s] ne loge pas sur la table [%s]", groupeIns, tableZoneIns)
        SpreadsheetApp.getUi().alert(mess);
      }
    } // endif
  } // endfor

  // Calcul du nombre de places pour le groupe
  var tmpGroupes = {};
  for (iRow = 0; iRow < iLengthInscriptions; iRow++) {
    groupeIns = sInscriptions[iRow][iCols["Groupe"]];
    nombreIns = sInscriptions[iRow][iCols["Nombre"]];
    if (tmpGroupes[groupeIns] == null) tmpGroupes[groupeIns] = 0;
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
  var sTablePlace = 0, iStartPlace = -1, qplace = 0;
  for (var ig = 0, igmax = iGroupes.length; ig < igmax; ig++) {
    groupeName = iGroupes[ig][0];
    groupeCount = iGroupes[ig][1];
    for (iRow = 0; iRow < iLengthInscriptions; iRow++) {
      if (sInscriptions[iRow][iCols["TableZone"]] == ""
        && sInscriptions[iRow][iCols["Groupe"]] == groupeName) {
        // affectation des places du groupe groupeName
        // recherche de la table qui dispose du nombre de places disponibles pour le groupe
        iStartPlace = -1; // varie de 0 à le nombre total de places dans la salle
        sTablePlace = ""
        qplace = 0;
        for (iRowPlace = 0; iRowPlace < iLengthPlaces; iRowPlace++) {
          // [[1, "A", "A1", ""], [2, "A", "A1", ""], [3, "A", "A1", ""]
          // [[numPlace, tablePlace, zonePlace, nomPlace]]
          nomPlace = sPlaces[iRowPlace][iColNomPlace];
          tablePlace = sPlaces[iRowPlace][iColTablePlace];
          if (nomPlace == "") { // place disponible
            if (sTablePlace == "") {
              sTablePlace = tablePlace;
              // iStartPlace = iRowPlace;
            }
            if (sTablePlace == tablePlace) {
              if (iStartPlace == -1) iStartPlace = iRowPlace;
              qplace += 1;
            } else {
              // recherche sur table suivante
              sTablePlace = tablePlace;
              iStartPlace = iRowPlace;
              qplace = 0;
            } // endif
          } else {
            // dent creuse avec espace insuffisant 
            sTablePlace = ""
            iStartPlace = -1;
            qplace = 0;
          } // endif
          if (qplace == groupeCount) {
            // la table avec suffisament de place a été trouvée, youpii !!!
            // attribution des places à partir de iStartPlace
            for (var ir = 0; ir < iLengthInscriptions; ir++) {
              nombreIns = sInscriptions[ir][iCols["Nombre"]];
              if (sInscriptions[ir][iCols["Groupe"]] == groupeName) {
                while (nombreIns > 0) {
                  // Option affichage du nom du groupe ou du nom de l'inscrit
                  //sPlaces[iStartPlace++][iColNomPlace] = sInscriptions[ir][iCols["Nom"]];
                  sPlaces[iStartPlace++][iColNomPlace] = sInscriptions[ir][iCols["Groupe"]];
                  nombreIns--;
                } // end while
              } // endif
            } // endfor
            // sortie des 2 for, le groupeName est traité
            iRowPlace = iLengthPlaces;
            iRow = iLengthInscriptions;
          } // endif
        } // endfor Place
        if (qplace == 0) {
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
  for (iRowPlace = 0; iRowPlace < iLengthPlaces; iRowPlace++) {
    numPlace = sPlaces[iRowPlace][iColNumPlace];
    nomPlace = sPlaces[iRowPlace][iColNomPlace];
    tablePlace = sPlaces[iRowPlace][iColTablePlace];
    if (nomPlaceRupture != nomPlace) {
      nomPlaceRupture = nomPlace;
      color = color == "aliceblue" ? "lavender" : "aliceblue";
    } // endif
    if (tablePlaceRupture != tablePlace) {
      // changement de table donc de plage nommée
      tablePlaceRupture = tablePlace;
      plRanges = plTables[tablePlace.charCodeAt(0)-65].getRange(); // A code 65 en ascii
      iRowRange = 0;
    } // endif
    if (numPlace % 2) {
      cell = plRanges.getCell(iRowRange + 1, iColNomImpair + 1);
    } else {
      cell = plRanges.getCell(iRowRange + 1, iColNomPair + 1);
      iRowRange++;
    } // endif
    cell.setValue(nomPlace);
    cell.setFontSize(8);
    cell.setWrap(true);
    if (nomPlace == "") {
      cell.setBackground("darkgray");
    } else {
      cell.setBackground(color);
    } // endif
  } // endfor

  // Mise à jour de l'onglet Inscriptions
  for (iRow = 0; iRow < iLengthInscriptions; iRow++) {
    for (iRowPlace = 0; iRowPlace < iLengthPlaces; iRowPlace++) {
      if (sInscriptions[iRow][iCols["Nom"]] == sPlaces[iRowPlace][iColNomPlace]) {
        oRangeInscriptions.getCell(iRow + iRowDataInscriptions, iCols["Table"] + 1).setValue(sPlaces[iRowPlace][iColTablePlace]);
        oRangeInscriptions.getCell(iRow + iRowDataInscriptions, iCols["Debut"] + 1).setValue(sPlaces[iRowPlace][iColNumPlace]);
        oRangeInscriptions.getCell(iRow + iRowDataInscriptions, iCols["Fin"] + 1).setValue(
          sPlaces[iRowPlace][iColNumPlace] + sInscriptions[iRow][iCols["Nom"] + 1] - 1);
        break;
      } // endif
    } // endfor
  } // endfor

  spreadsheet.toast("Calcul terminé");
}
