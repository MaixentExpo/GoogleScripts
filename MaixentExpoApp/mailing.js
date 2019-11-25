/*
  Script de Mailing
*/
function prepareMessage() {
  prepareMessageResultat("15gPNGjf_Sga1Ips11NccEeSUq7X8eJjEiOHPMooT5jw", "VINS", "RESULTAT");
}


/*
  messagePresseConcoursVins
  Retourne le message à envoyer à la presse
  Les colonnes devront êtres préalablement triées sur Medaille, Couleur, Vin
*/
function prepareMessageResultat(source_file_id, sheet_name, cible_range_name) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var yesnoConfirm = ui.alert(
     "Préparer le message",
     'Veuillez confirmer par oui ou non',
      ui.ButtonSet.YES_NO);
  if ( yesnoConfirm != ui.Button.YES ) return;

  // Ouverture de la feuille VINS
  var spreadsheet_source = SpreadsheetApp.openById(source_file_id);
  var sheet = spreadsheet_source.getSheetByName(sheet_name)
  // Recherche de la position des colonnes sur la 1ère ligne
  var iLastCol = sheet.getLastColumn()
  var iLastRow = sheet.getLastRow()
  var range = sheet.getRange(1,1,1,iLastCol); // 1ère ligne
  var values = range.getValues();
  var iVignoble = 0;
  var iProducteur = 0;
  var iRegion = 0;
  var iVin = 0;
  var iCouleur = 0;
  var iMedaille = 0;
  var iTriMedaille = 0;
  var sCell = "";
  var iRow = 0;
  var iCol = 0
  for(; iCol<iLastCol; iCol++){
    sCell = values[iRow][iCol];
    if ( sCell == "Vignoble" ) iVignoble = iCol+1;
    if ( sCell == "Producteur" ) iProducteur = iCol+1;
    if ( sCell == "Region" ) iRegion = iCol+1;
    if ( sCell == "Vin" ) iVin = iCol+1;
    if ( sCell == "Couleur" ) iCouleur = iCol+1;
    if ( sCell == "Medaille" ) iMedaille = iCol+1;
    if ( sCell == "TriMedaille" ) iTriMedaille = iCol+1;
  } // endfor
  
  // TRI des colonnes Medaille, Couleur, Vin
  range = sheet.getRange(2, 1, iLastRow, iLastCol);
  range.sort([{column: iTriMedaille, ascending: true}, {column: iCouleur, ascending: true}, {column: iVin, ascending: true}]);

  // Boucle avec rupture sur les médailles
  var sVignoble = "";
  var sProducteur = "";
  var sRegion = "";
  var sVin = "";
  var sCouleur = "";
  var sMedaille = "";
  var sRuptureMedaille = "";
  var sResultat = "";
  // var rich = SpreadsheetApp.newRichTextValue();
  // var iPos = 0;
  var bFirst = true;
  iRow = 2; // on commence sur la 2ème ligne
  for(; iRow<iLastRow; iRow++) {
    sMedaille = sheet.getRange(iRow, iMedaille).getValue().trim();
    if ( sMedaille == "" ) 
      break;
    sVignoble = sheet.getRange(iRow, iVignoble).getValue();
    sProducteur = sheet.getRange(iRow, iProducteur).getValue();
    sRegion = sheet.getRange(iRow, iRegion).getValue();
    sVin = sheet.getRange(iRow, iVin).getValue();
    sCouleur = sheet.getRange(iRow, iCouleur).getValue();
    
    if ( sMedaille != sRuptureMedaille ) {
      // rich.set
      sResultat += '\n';
      if (bFirst == false) sResultat += "\n"
      sResultat += sMedaille == "or" ? "Médaille d'or" : sMedaille == "argent" ? "Médaille d'argent" : "Médaille de bronze";
      sResultat += " :\n";
      bFirst = true;
      sRuptureMedaille = sMedaille;
    } // endif
    if (bFirst == false) sResultat += "\n"
    bFirst = false
    sResultat += "- " + sProducteur + " (" + sVin + ") " + sVignoble
  } // endfor

  // Ouverture du fichier mailing courant
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // Copie du resulat dans la cible
  var cible = spreadsheet.getRange(cible_range_name);
  cible.setValue(sResultat);
  
} // end messagePresseConcoursBovins

function recupEmail() {
  fx_recupEmail("15gPNGjf_Sga1Ips11NccEeSUq7X8eJjEiOHPMooT5jw", "EMAILS_JURYS", "CC");
}

function fx_recupEmail(source_file_id, source_range_name, cible_range_name) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var yesnoConfirm = ui.alert(
     "Récupérer les emails",
     'Veuillez confirmer par oui ou non',
      ui.ButtonSet.YES_NO);
  if ( yesnoConfirm != ui.Button.YES ) return;
  
  // Recup emails 
  var spreadsheet_source = SpreadsheetApp.openById(source_file_id);
  var values = spreadsheet_source.getRangeByName(source_range_name).getValues();
  var iLastRow = values.length;
  var emails = "";
  var isStart = true;
  for (var i=0; i<iLastRow; i++) {
    if ( values[i] == "" ) 
      continue;
    if ( ! isStart ) {
      emails += ", ";
    } // end if
    isStart = false;
    emails += values[i];
  } // end for
  
  // Copie emails dans plage EMAILS
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cell = spreadsheet.getRange(cible_range_name);
  cell.setValue(emails);
  
}

function fx_sendMail() {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
     "Envoiyer les mails",
     'Veuillez confirmer par oui ou non',
      ui.ButtonSet.YES_NO);
  if ( yesnoConfirm != ui.Button.YES ) return;
  
  // Recup des champs dans la feuille courante
  var spreadsheet = SpreadsheetApp.getActive();
  var to = spreadsheet.getRangeByName("TO").getCell(1,1).getValue();
  var copy = spreadsheet.getRangeByName("COPY").getCell(1,1).getValue();
  var cc = spreadsheet.getRangeByName("CC").getCell(1,1).getValue();
  var replyTo = spreadsheet.getRangeByName("REPLYTO").getCell(1,1).getValue();
  var subject = spreadsheet.getRangeByName("SUBJECT").getCell(1,1).getValue();
  var richText = spreadsheet.getRangeByName("BODY").getCell(1,1).getRichTextValue();
  // Message en Html enrichi
  var html = fx_htmlEncodeRichText(richText);
  // Envoi du message
  MailApp.sendEmail({
    replyTo: replyTo,
    to: to,
    copy: copy,
    cc: cc,
    subject: subject,
    htmlBody: html
  });
  
}

/**
 * Remplace la chaine {tag} du RichTextValue par une valeur
 * @param {RichTextValue} richTextValue
 * @param {String} tag 
 * @param {String} stringForReplace 
 * @return {RichTextValue}  
 */
function fx_ReplaceRichText(richTextValue, tag, stringForReplace) {
  // get an array of Runs for the given Rich Text
  var runs = richTextValue.getRuns();
  // loop the array
  for (var i = 0; i < runs.length; i++) {
    var richText = runs[i].getText();
    var style = runs[i].getTextStyle();
    var re = new RegExp(tag, "g");
    richText = richText.replace(re, stringForReplace);
    runs[i] = richText;
  }
  return richTextValue;
}

/**
 * Given a RichTextValue Object, iterate over the individual runs
 *    and call out to htmlStyleRtRun() to return the text wrapped
 *    in <span> tags with specific styling.
 * @param {RichTextValue} richTextValue a RichTextValue object
 *    from a given Cell.
 * @return {string} HTML encoded text 
 */
function fx_htmlEncodeRichText(richTextValue) {
  // create an empty string which will hold the html content
  var htmlString = "";
  // get an array of Runs for the given Rich Text
  var rtRuns = richTextValue.getRuns();
  // loop the array
  for (var i = 0; i < rtRuns.length; i++) {
    // return html version of a given run, append to existing string
    htmlString += fx_htmlStyleRtRun(rtRuns[i]);
//    Logger.log("Run # " + i + " plain text: ");
//    Logger.log(rtRuns[i].getText());
//    Logger.log("Run # " + i + " Output:");
//    Logger.log(htmlString);
  }
  return htmlString;
}

/**
 * Given a RichTextValue Run, evaluates for style attributes and 
 *    builds a <span> tag with in-line styles. 
 *    For instance:
 *    <span style="color: cyan">text</span>
 *
 * @param {RichTextValue} richTextRun an instance of a
 *    RichTextValue run
 * @return {string} inputted text wrapped in <span> tag with 
 *    applicable styling. 
 */
function fx_htmlStyleRtRun(richTextRun) {
  // string to hold the inline style key value pairs
  var styleString = "";
  // evaluate the attributes of a given Run and construct style attributes
  if (richTextRun.getTextStyle().isBold()) {
    styleString += "font-weight:bold;"
  }
  if (richTextRun.getTextStyle().isItalic()) {
    styleString += "font-style:italic;"
  }
  // fetch values for font family, size, and color attributes
  styleString += "font-family:" + richTextRun.getTextStyle().getFontFamily() +
    ";";
  styleString += "font-size:" + richTextRun.getTextStyle().getFontSize() +
    "px;";
  styleString += "color:" + richTextRun.getTextStyle().getForegroundColor() +
    ";";

  // underline and strikethrough use the same style key, text-decoration, must evaluate together, otherwise, the styling breaks. 
  // both false 
  if (!richTextRun.getTextStyle().isUnderline() && !richTextRun.getTextStyle().isStrikethrough()) {
    // do nothing
  }
  // underline true, strikethrough false
  else if (richTextRun.getTextStyle().isUnderline() && !richTextRun.getTextStyle()
    .isStrikethrough()) {
    styleString += "text-decoration: underline;";
  }
  // underline false, strikethrough true
  else if (!richTextRun.getTextStyle().isUnderline() && richTextRun.getTextStyle()
    .isStrikethrough()) {
    styleString += "text-decoration: line-through;";
  }
  // both true
  else {
    styleString += "text-decoration: line-through underline;";
  }

  // line breaks don't get converted, run regex and insert <br> to replace \n
  var richText = richTextRun.getText();
  var re = new RegExp("\n", "g");
  var richText = richText.replace(re, "<br>");

  // bring it all together
  var formattedText = '<span style="' + styleString + '">' + richText +
    '</span>';
  return formattedText;
}

/**
 * Présente une date sous la forme "12 avril 2019"
 * var maDate = new Date();
 * var maDateFrench = frenchDate(maDate)
 * @param {*} date 
 */
function fx_frenchDate(date) {
  var month = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var m = month[date.getMonth()];
  var dateStringFr = date.getDate() + ' ' + m + ' ' + date.getFullYear();
  return dateStringFr
}
