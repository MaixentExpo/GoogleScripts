/*
Initialisation du calendrier
Préalable :
- copier coller la feuille N-1 dans N
- mettre une * dans la 1ère cellule

*/

function InitialiserLeCalendrier() {
  // ATTENTION : fonction qui sera exécutée sur la feuille courante
  // Ouverture du tableur conteneur du script
  var spreadsheet = SpreadsheetApp.getActive()
  // Ouverture de la feuille courante
  var sheet = spreadsheet.getActiveSheet()
  
  // Mise à jour autorisée ?
  if ( sheet.getRange(1,1).getValue() != "*") return;
  
  var annee = parseInt(sheet.getName())
  var jourSemaine = { 1:"L", 2:"M", 3:"M", 4:"J", 5:"V", 6:"S", 0:"D"} // jour semaine
  var jourFeries = JoursFeries(annee)
  
  // Déclaration de la plage à mettre à jour
  var range = sheet.getRange(2,1,31,48)
  // nettoyage de la plage
  range.setValue("").setBackground("white")
  var icol, irow, cell1, cell2, cell3;
  var dateCourante = new Date(annee, "00", "01")
  var iq = 1 // quantième du jour de l'année
  var imonth = 0
  while ( dateCourante.getFullYear() == annee ) {
    imonth = dateCourante.getMonth()
    icol = imonth * 4 + 1
    irow = 1
    while ( dateCourante.getMonth() == imonth ) {
      cell1 = range.getCell(irow, icol+1)
      cell2 = range.getCell(irow, icol+2)
      cell3 = range.getCell(irow, icol+3)
      if ( dateCourante.getDay() == 0 || jourFeries[iq] != null) { // dimanche ou jour férié
        cell1.setBackground("#fff2cc").setValue(jourSemaine[dateCourante.getDay()])
        cell2.setBackground("#fff2cc").setValue(dateCourante.getDate())
        cell3.setBackground("#fff2cc")
      } else {
        cell1.setBackground("#efefef").setValue(jourSemaine[dateCourante.getDay()])
        cell2.setBackground("#efefef").setValue(dateCourante.getDate())
        cell3.setBackground("white")
      } // endif
      dateCourante.setDate(dateCourante.getDate() + 1)
      iq++
      irow++
    } // end while month
  } // end while annee
  
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
  jf[Math.floor((FeteTravail.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = FeteTravail
  jf[Math.floor((Victoire1945.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Victoire1945
  jf[Math.floor((FeteNationale.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = FeteNationale
  jf[Math.floor((Assomption.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Assomption
  jf[Math.floor((Toussaint.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Toussaint
  jf[Math.floor((Armistice.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Armistice
  jf[Math.floor((Noel.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Noel
  jf[Math.floor((Paques.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Paques
  jf[Math.floor((LundiPaques.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = LundiPaques
  jf[Math.floor((Ascension.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Ascension
  jf[Math.floor((Pentecote.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = Pentecote
  jf[Math.floor((LundiPentecote.getTime() - JourAn.getTime()) / (24 * 3600 * 1000))+2] = LundiPentecote
	
  return jf
}
