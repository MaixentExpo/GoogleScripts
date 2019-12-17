/**
 * Echange des champs Répondre à, Envoyer à, En Copie, En copie cachée
 */
function mailing_echangeAdresses() {
  var spreadsheet = SpreadsheetApp.getActive();
  var replyTo = spreadsheet.getRangeByName("REPLYTO").getCell(1,1).getValue();
  var to = spreadsheet.getRangeByName("TO").getCell(1,1).getValue();
  var copy = spreadsheet.getRangeByName("COPY").getCell(1,1).getValue();
  var cc = spreadsheet.getRangeByName("CC").getCell(1,1).getValue();

  var sheet = spreadsheet.getSheetByName("MAIL")
  var replyTo2 = sheet.getRange("D2").getCell(1,1).getValue();
  var to2 = sheet.getRange("D3").getCell(1,1).getValue();
  var copy2 = sheet.getRange("D4").getCell(1,1).getValue();
  var cc2 = sheet.getRange("D5").getCell(1,1).getValue();

  spreadsheet.getRangeByName("REPLYTO").setValue(replyTo2);
  spreadsheet.getRangeByName("TO").setValue(to2);
  spreadsheet.getRangeByName("COPY").setValue(copy2);
  spreadsheet.getRangeByName("CC").setValue(cc2);

  sheet.getRange("D2").setValue(replyTo);
  sheet.getRange("D3").setValue(to);
  sheet.getRange("D4").setValue(copy);
  sheet.getRange("D5").setValue(cc);

}

function mailing_echangeAdresses2() {
  var spreadsheet = SpreadsheetApp.getActive();
  var replyTo = spreadsheet.getRangeByName("REPLYTO").getCell(1,1).getValue();
  var to = spreadsheet.getRangeByName("TO").getCell(1,1).getValue();
  var copy = spreadsheet.getRangeByName("COPY").getCell(1,1).getValue();
  var cc = spreadsheet.getRangeByName("CC").getCell(1,1).getValue();

  var sheet = spreadsheet.getSheetByName("MAIL")
  var replyTo2 = sheet.getRange("C2").getCell(1,1).getValue();
  var to2 = sheet.getRange("C3").getCell(1,1).getValue();
  var copy2 = sheet.getRange("C4").getCell(1,1).getValue();
  var cc2 = sheet.getRange("C5").getCell(1,1).getValue();

  spreadsheet.getRangeByName("REPLYTO").setValue(replyTo2);
  spreadsheet.getRangeByName("TO").setValue(to2);
  spreadsheet.getRangeByName("COPY").setValue(copy2);
  spreadsheet.getRangeByName("CC").setValue(cc2);

  sheet.getRange("C2").setValue(replyTo);
  sheet.getRange("C3").setValue(to);
  sheet.getRange("C4").setValue(copy);
  sheet.getRange("C5").setValue(cc);

}
