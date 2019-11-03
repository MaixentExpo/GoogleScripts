/*
Initialisation du calendrier
Préalable :
- copier coller la feuille N-1 dans N

*/

function InitialiserLeCalendrier() {
  // ATTENTION : fonction qui sera exécutée sur la feuille courante
  // Ouverture du tableur conteneur du script
  var spreadsheet = SpreadsheetApp.getActive();
  // Ouverture de la feuille courante
  var sheet = spreadsheet.getActiveSheet();
  
  // Mise à jour autorisée ?
  var ok = sheet.getCell(0,0)
  if ( sheet.getRange(0,0).getValue() != "*") return;
  
  var annee = parseInt(sheet.getName())
  var ijs = new Date(annee, 0, 1).getDay() // jour semaine du 1er janvier
  var sjs = { 1:"L", 2:"M", 3:"M", 4:"J", 5:"V", 6:"S", 0:"D"} // jour semaine
  var idj = { 0:31, 1:28, 2:31, 3:30, 4:31, 5:30, 6:31, 7:31, 8:30, 9:31, 10:30, 11:31 } // dernier jour des mois
  idj[1] = new Date(annee, 2, 0).getDate()
  var jf = JoursFeries(annee)
  
  // nettoyage des cellules
  sheet.getRange(0,1,31,48).setValue("").setBackground("white")

  var iq, imois, irow, icol, cell0, cell1, cell2, cell3;
  for ( iq=1, imois=0; imois < 12; imois++ ) {
    icol = imois * 4
    for ( irow = 1; irow < 32; irow++ ) {
      cell0 = sheet.getCell(irow, icol)
      cell1 = sheet.getCell(irow, icol+1)
      cell2 = sheet.getCell(irow, icol+2)
      cell3 = sheet.getCell(irow, icol+3)
      cell0.setBackground("white")
      if ( ijs == 0 || ijs == 6 ) { // week-end
        cell0.setBackground("white")
        cell1.setBackground("#fff2cc")
        cell2.setBackground("#fff2cc")
        cell3.setBackground("white")
      } else {
        cell0.setBackground("white")
        cell1.setBackground("#efefef")
        cell2.setBackground("#efefef")
        cell3.setBackground("white")
      } // endif
    } // end for row
      
  } // end for mois
  
} // end InitialiserLeCalendrier

/**
 * Retourne un dictionnaire du quantième des jours fériés de l'année
 * @param {*} an 
 */
function JoursFeries (an) {
	var JourAn = new Date(an, "00", "01")
	var FeteTravail = new Date(an, "04", "01")
	var Victoire1945 = new Date(an, "04", "08")
	var FeteNationale = new Date(an,"06", "14")
	var Assomption = new Date(an, "07", "15")
	var Toussaint = new Date(an, "10", "01")
	var Armistice = new Date(an, "10", "11")
	var Noel = new Date(an, "11", "25")
	
	var G = an%19
	var C = Math.floor(an/100)
	var H = (C - Math.floor(C/4) - Math.floor((8*C+13)/25) + 19*G + 15)%30
	var I = H - Math.floor(H/28)*(1 - Math.floor(H/28)*Math.floor(29/(H + 1))*Math.floor((21 - G)/11))
	var J = (an*1 + Math.floor(an/4) + I + 2 - C + Math.floor(C/4))%7
	var L = I - J
	var MoisPaques = 3 + Math.floor((L + 40)/44)
	var JourPaques = L + 28 - 31*Math.floor(MoisPaques/4)
	var Paques = new Date(an, MoisPaques-1, JourPaques)
	var LundiPaques = new Date(an, MoisPaques-1, JourPaques+1)
	var Ascension = new Date(an, MoisPaques-1, JourPaques+39)
	var Pentecote = new Date(an, MoisPaques-1, JourPaques+49)
  var LundiPentecote = new Date(an, MoisPaques-1, JourPaques+50)
  
  var jf = {}
  jf[1] = JourAn
  jf[Math.floor(FeteTravail.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = FeteTravail
  jf[Math.floor(Victoire1945.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Victoire1945
  jf[Math.floor(FeteNationale.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = FeteNationale
  jf[Math.floor(Assomption.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Assomption
  jf[Math.floor(Toussaint.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Toussaint
  jf[Math.floor(Armistice.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Armistice
  jf[Math.floor(Noel.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Noel
  jf[Math.floor(Paques.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Paques
  jf[Math.floor(LundiPaques.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = LundiPaques
  jf[Math.floor(Ascension.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Ascension
  jf[Math.floor(Pentecote.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = Pentecote
  jf[Math.floor(LundiPentecote.getTime() - JourAn.getTime()) / (24 * 3600 * 1000)] = LundiPentecote
	
    return jf
}
