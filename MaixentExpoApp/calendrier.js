/*
Initialisation du calendrier
Préalable :
- copier coller la feuille N-1 dans N
- mettre une * dans la 1ère cellule
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Macros')
      .addItem('Initialiser le calendrier (si * en A1)', 'InitialiserLeCalendrier')
      .addToUi();
}

function InitialiserLeCalendrier() {
  // ATTENTION : fonction qui sera exécutée sur la feuille courante
  // Ouverture du tableur conteneur du script
  var spreadsheet = SpreadsheetApp.getActive()
  // Ouverture de la feuille courante
  var sheet = spreadsheet.getActiveSheet()
  
  // Mise à jour autorisée ?
  if ( sheet.getRange("A1").getValue() != "*") return;
  
  var annee = parseInt(sheet.getName())
  var jourSemaine = { 1:"L", 2:"M", 3:"M", 4:"J", 5:"V", 6:"S", 0:"D"} // jour semaine
  var jourFeries = JoursFeries(annee)
  
  var dateCourante = new Date(annee, "00", "01")
  var conges = getConges(spreadsheet, dateCourante) 
  
  // Déclaration de la plage à mettre à jour
  var range = sheet.getRange("A2:AV32")
  // nettoyage de la plage
  range.setValue("").setBackground("white")
  var icol, irow, cell0, cell1, cell2, cell3;
  var iq = 1 // quantième du jour de l'année
  var imonth = 0
  while ( dateCourante.getFullYear() == annee ) {
    imonth = dateCourante.getMonth()
    icol = imonth * 4 + 1
    irow = 1
    while ( dateCourante.getMonth() == imonth ) {
      cell0 = range.getCell(irow, icol)
      cell1 = range.getCell(irow, icol+1)
      cell2 = range.getCell(irow, icol+2)
      cell3 = range.getCell(irow, icol+3)
      if ( conges[formatDateMMdd(dateCourante)] != null ) {
        cell0.setBackground("#e6b8af")
      } else {
        cell0.setBackground("white")
      } // endif
      if ( dateCourante.getDay() == 0 
        || jourFeries[formatDateMMdd(dateCourante)] != null) { // dimanche ou jour férié
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
  sheet.getRange("A1").setValue("")
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
  jf[formatDateMMdd(JourAn)] = true
  jf[formatDateMMdd(FeteTravail)] = true
  jf[formatDateMMdd(Victoire1945)] = true
  jf[formatDateMMdd(FeteNationale)] = true
  jf[formatDateMMdd(Assomption)] = true
  jf[formatDateMMdd(Toussaint)] = true
  jf[formatDateMMdd(Armistice)] = true
  jf[formatDateMMdd(Noel)] = true
  jf[formatDateMMdd(Paques)] = true
  jf[formatDateMMdd(LundiPaques)] = true
  jf[formatDateMMdd(Ascension)] = true
  jf[formatDateMMdd(Pentecote)] = true
  jf[formatDateMMdd(LundiPentecote)] = true
	
  return jf
}

function formatDateMMdd(date) {
  return date.getMonth().toString() + "_" + date.getDate().toString()
}

/**
 * Reourne un dictionnaire des congés scolaires key:MM_dd value:true
 * @param {SpreadSheet} spreadsheet 
 * @param {Date} date 
 */
function getConges(spreadsheet, date) {
  var values = spreadsheet.getRangeByName("CONGES").getValues()
  var iLastRow = values.length
  var irow, dateStart, dateEnd
  var oConges = {}
  for ( irow=0; irow < iLastRow; irow++ ) {
    dateStart = values[irow][1]
    dateEnd = values[irow][2]
    if ( dateStart.getFullYear() == date.getFullYear() ) {
      while ( formatDateMMdd(dateStart) <= formatDateMMdd(dateEnd) ) {
        oConges[formatDateMMdd(dateStart)] = true
        dateStart.setDate(dateStart.getDate() + 1)
        if ( dateStart.getFullYear() > date.getFullYear() ) {
          break;
        } // endif
      } // end while dateStart
    } // end if dateStart
  } // end for row
  return oConges  
}