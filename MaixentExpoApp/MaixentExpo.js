/**
 * Fonctions communes javascript avec le préfixs fx_
 * Maixent.Expo@gmail.com
 */

/**
 * fx_getIdFromUrl : Extraire l'id d'une url Google Drive
 * Sheets : https://docs.google.com/spreadsheets/d/idididid/...
 * Slides : https://docs.google.com/presentation/d/idididid/...
 * Docs   : https://docs.google.com/document/d/idididid/...
 * File   : https://drive.google.com/open?id=1Cw5UE0TUC8ARRbQBmUW_U8CiM8_o_YuO
 * Drive  : https://drive.google.com/drive/folders/1MjNf3-poTAo6cF6OJ83Y3Oi9J7enPhYy
 * @param {String} urlOrId 
 */
function fx_getIdFromUrl(urlOrId) {
  const regexDoc = /.*\/d\/(.*)\/.*/g
  const regexFile = /.*\/open.id=(.*)/g
  const regexFolder = /.*\/folders\/(.*)/g
  var ids = regexDoc.exec(urlOrId)
  ids = ids == null ? regexFile.exec(urlOrId) : ids
  ids = ids == null ? regexFolder.exec(urlOrId) : ids
  var id = ids == null ? urlOrId : ids[1]
  return id
}


/**
* Présente une date sous la forme "12 avril 2019"
* var maDate = new Date();
* var maDateFrench = frenchDate(maDate)
* @param {Date} date 
*/
function fx_frenchDate(date) {
  var month = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  var m = month[date.getMonth()];
  var dateStringFr = date.getDate() + ' ' + m + ' ' + date.getFullYear();
  return dateStringFr
}

/**
 * Class Couleur
 * qui fournit un code couleur à chaque appel new_couleur()
 */
var Couleur = function () {
  this.couleurs = ["#e8f5e9" // green
    , "#e3f2fd" // blue
    , "#fffde7" // yellow
    , "#fbe9e7" // deep orange
    , "#e0f7fa" // cyan
    , "#f1f8e9" // light green
    , "#fce4ec" // pink
    , "#e1f5fe" // light blue
    , "#ede7f6" // deep purple
    , "#eceff1" // blue grey
    , "#e8eaf6" // indigo
    , "#f3e5f5" // purple
    , "#f9fbe7" // lime
    , "#fff3e0" // orange
    , "#fff8e1" // amber
    , "#efebe9" // brown
    , "#e0f2f1" // teal
    , "#ffebee" // red
    , "#fafafa" // grey
  ];
  this.iCouleur = -1;
  this.couleur = "#fafafa";
  this.new_couleur = function () {
    this.iCouleur++
    if (this.iCouleur >= this.couleurs.length) {
      this.iCouleur = 0;
    } // endif
    this.couleur = this.couleurs[this.iCouleur];
  } // end new_couleur
} // end class Couleur

function fx_recupEmail(source_IdorUrl, source_range_name, cible_range_name) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var yesnoConfirm = ui.alert(
    "Récupérer les emails",
    'Veuillez confirmer par oui ou non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  // Recup emails 
  var spreadsheet_source = SpreadsheetApp.openById(fx_getIdFromUrl(source_IdorUrl));
  var values = spreadsheet_source.getRangeByName(source_range_name).getValues();
  var iLastRow = values.length;
  var emails = "";
  var isStart = true;
  for (var i = 0; i < iLastRow; i++) {
    if (values[i] == "")
      continue;
    if (!isStart) {
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
/**
 * fx_selectEmail
 * Sélection des emails d'une colonne qui répondent au critère
 * @param {String} source_file_id id du tableur 
 * @param {String} source_range_name plage des emails du genre "MEMBRES!D2:D70"
 * @param {String} cible_range_name cellule cible de la feuille courante
 * @param {String} rangeFilter plage de valeur à tester
 * @param {String} filterValue expression régulière "Bovins" "Vins" "Tombola|Publicité" "^((?!Adhérent).)*$"
 */
function fx_selectEmail(source_file_id, source_range_name, cible_range_name, rangeFilter, filterValue) {
  var ui = SpreadsheetApp.getUi();
  var message = "Récupérer les emails (" + filterValue + ")";
  var yesnoConfirm = ui.alert(
    message,
    'Veuillez confirmer par oui ou non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  // Recup emails 
  var spreadsheet_source = SpreadsheetApp.openById(source_file_id);
  var values = spreadsheet_source.getRangeByName(source_range_name).getValues();
  var iLastRow = values.length;
  var emails = "";
  var isStart = true;
  var reFilter = new RegExp(filterValue, 'gm');
  var filterValues = spreadsheet_source.getRangeByName(rangeFilter).getValues();
  for (var i = 0; i < iLastRow; i++) {
    if (("" + filterValues[i]).match(reFilter, 'g') == null) {
      continue;
    } // endif
    if (values[i] == "")
      continue;
    if (!isStart) {
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

function fx_envoyerMessageB() {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
    "Envoyer le message",
    'Veuillez confirmer par Oui ou Non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  // Recup des champs dans la feuille courante
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  var replyTo = sheet.getRange("B2").getValue();
  var to = sheet.getRange("B3").getValue();
  var cc = sheet.getRange("B4").getValue();
  var bcc = sheet.getRange("B5").getValue();
  var pjSheet = sheet.getRange("B6").getValue();
  var pjFile1 = sheet.getRange("B7").getValue();
  var pjFile2 = sheet.getRange("B8").getValue();
  var pjFile3 = sheet.getRange("B9").getValue();
  var subject = sheet.getRange("B10").getValue();
  // Message en Html enrichi
  var richText = sheet.getRange("B11").getRichTextValue();
  var html = fx_htmlEncodeRichText(richText);
  // Pièces jointes
  var blobs = [];
  if (pjSheet != "") {
    blobs.push(fx_SpreadsheetToExcel(pjSheet));
  }
  if (pjFile1 != "") {
    blobs.push(fx_FileToPdf(pjFile1, "&portrait=false"));
  }
  if (pjFile2 != "") {
    blobs.push(fx_FileToPdf(pjFile2, "&portrait=true"));
  }
  if (pjFile3 != "") {
    blobs.push(fx_FileToPdf(pjFile3, "&portrait=true"));
  }
  // Envoi du message
  if (blobs.length > 0) {
    MailApp.sendEmail({
      replyTo: replyTo,
      to: to,
      cc: cc,
      bcc: bcc,
      subject: subject,
      htmlBody: html,
      attachments: blobs
    });
  } else {
    MailApp.sendEmail({
      replyTo: replyTo,
      to: to,
      cc: cc,
      bcc: bcc,
      subject: subject,
      htmlBody: html
    });
  } // endif

  // Historisation de l'action dans la plage LOG
  if (sheet.getRange("C11") != "") {
    var slog = sheet.getRange("C11").getValue();
    var strace = Utilities.formatString("%s par %s",
      Utilities.formatDate(new Date(),
        spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      Session.getActiveUser().getEmail());
    slog = strace + "\n" + slog;
    sheet.getRange("C11").setValue(slog);
  }

}

function fx_envoyerMessage() {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
    "Envoyer le message",
    'Veuillez confirmer par oui ou non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  // Recup des champs dans la feuille courante
  var spreadsheet = SpreadsheetApp.getActive();
  var to = spreadsheet.getRangeByName("TO").getCell(1, 1).getValue();
  var cc = spreadsheet.getRangeByName("COPY").getCell(1, 1).getValue();
  var bcc = spreadsheet.getRangeByName("CC").getCell(1, 1).getValue();
  var replyTo = spreadsheet.getRangeByName("REPLYTO").getCell(1, 1).getValue();
  var subject = spreadsheet.getRangeByName("SUBJECT").getCell(1, 1).getValue();
  var richText = spreadsheet.getRangeByName("BODY").getCell(1, 1).getRichTextValue();
  var pjSheet = spreadsheet.getRangeByName("PJ_SHEET") != null ? spreadsheet.getRangeByName("PJ_SHEET").getCell(1, 1).getValue() : "";
  var pjFile1 = spreadsheet.getRangeByName("PJ_FILE1") != null ? spreadsheet.getRangeByName("PJ_FILE1").getCell(1, 1).getValue() : "";
  var pjFile2 = spreadsheet.getRangeByName("PJ_FILE2") != null ? spreadsheet.getRangeByName("PJ_FILE2").getCell(1, 1).getValue() : "";
  var pjFile3 = spreadsheet.getRangeByName("PJ_FILE3") != null ? spreadsheet.getRangeByName("PJ_FILE3").getCell(1, 1).getValue() : "";
  // Message en Html enrichi
  var html = fx_htmlEncodeRichText(richText);
  // Envoi du message
  var blobs = [];
  if (pjSheet != "") {
    blobs.push(fx_SpreadsheetToExcel(pjSheet));
  }
  if (pjFile1 != "") {
    blobs.push(fx_FileToPdf(pjFile1, "&portrait=false"));
  }
  if (pjFile2 != "") {
    blobs.push(fx_FileToPdf(pjFile2, "&portrait=true"));
  }
  if (pjFile3 != "") {
    blobs.push(fx_FileToPdf(pjFile3, "&portrait=true"));
  }
  if (blobs.length > 0) {
    MailApp.sendEmail({
      replyTo: replyTo,
      to: to,
      cc: cc,
      bcc: bcc,
      subject: subject,
      htmlBody: html,
      attachments: blobs
    });
  } else {
    MailApp.sendEmail({
      replyTo: replyTo,
      to: to,
      cc: cc,
      bcc: bcc,
      subject: subject,
      htmlBody: html
    });
  } // endif

  // Historisation de l'action dans la plage LOG
  if (spreadsheet.getRangeByName("LOG") != null) {
    var slog = spreadsheet.getRangeByName("LOG").getCell(1, 1).getValue();
    var strace = Utilities.formatString("%s par %s",
      Utilities.formatDate(new Date(),
        spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      Session.getActiveUser().getEmail());
    slog = strace + "\n" + slog;
    spreadsheet.getRangeByName("LOG").setValue(slog);
  }

}

function fx_SpreadsheetToExcel(sheet_id) {
  // https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  var blob = null;

  try {
    var file_id = sheet_id;
    if (sheet_id.indexOf("https") > -1) {
      const regex = /.*\/d\/(.*)\/.*/g;
      file_id = regex.exec(sheet_id)[1];
    } // endif

    spreadsheet = SpreadsheetApp.openById(file_id);
    var url = "https://docs.google.com/spreadsheets/d/" + file_id + "/export?format=xlsx";
    var params = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(spreadsheet.getName() + ".xlsx");

  } catch (f) {
    Logger.log(f.toString());
  }
  return blob;
}

/**
 * Conversion fichier Google Drive en PDF
 * @param {String} FileId url du fichier ou seulement l'id
 * @param {String} parametres "&portrait=false" par exemple
 * @return {Blob} Objet du fichier converti
**/
function fx_FileToPdf(FileId, parametres) {
  // https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  var blob = null;
  try {
    var file_id = FileId;
    if (FileId.indexOf("https") > -1) {
      const regex = /.*\/d\/(.*)\/.*/g;
      file_id = regex.exec(FileId)[1];
    } // endif
    var file = DriveApp.getFileById(file_id);
    var params = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    var url_google = "https://docs.google.com/spreadsheets/d/";
    if (file.getMimeType() == MimeType.GOOGLE_DOCS) url_google = "https://docs.google.com/document/d/"
    if (file.getMimeType() == MimeType.GOOGLE_SLIDES) url_google = "https://docs.google.com/presentation/d/"

    var url = url_google + file_id + "/export?format=pdf&size=7&fzr=true" + parametres;
    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(file.getName() + ".pdf");

  } catch (f) {
    Logger.log(f.toString());
  }
  return blob;
}

/**
 * fx_exportPdf
 * Export d'un fichier en pdf dans le même répertoire
 * @param {String} UrlorId 
 */
function fx_exportPdf(UrlorId) {
  var sourceFile = DriveApp.getFileById(fx_getIdFromUrl(UrlorId))
  // un fichier peut avoir plusieurs répertoires parents
  // on ne pendra que le 1er parent
  var folders = sourceFile.getParents()
  var folder
  while (folders.hasNext()) {
    folder = folders.next();
    break;
  }
  // Quand on créé un fichier à partir d'un blob, on récupère automatiquement uns pdf, Incroyable Non ?
  var pdfFile = DriveApp.createFile(sourceFile.getBlob())
  // le fichier a été crée dans la racine du répertoire de l'utilisateur
  // on va le déplacer dans le répertoire du fichier source
  folder.addFile(pdfFile); // ajout du répertoire cible
  DriveApp.getRootFolder().removeFile(pdfFile); // suppresion du répertoire racine du fichier
  return  pdfFile.getUrl()
}

function fx_dialogCreatePdf(ui, folderId, fileId, pdfParameters) {
  var blob = fx_FileToPdf(fileId, pdfParameters);
  var folder = DriveApp.getFolderById(folderId);
  var pdfFile = folder.createFile(blob);

  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + blob.getName() + '</a></p>')
    .setWidth(300)
    .setHeight(80)
  ui.showModalDialog(htmlOutput, 'Export Successful')
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
 * Archivage des mails labelisés dans un répertoire de Drive
 * @param {String} gmailLabel Label des mails à archiver
 * @param {String} driveFolderId  id du répertoire drive
 */
function fx_saveGmailAsPDF(gmailLabel, driveFolderId) {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
    "Archiver les mails ?",
    'Veuillez confirmer par oui ou non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  var threads = GmailApp.search("in:" + gmailLabel, 0, 5)

  var mailsArchived = []

  if (threads.length > 0) {

    /* Google Drive folder where the Files would be saved */
    var folder = DriveApp.getFolderById(driveFolderId);
    // mémorisation des msgid déjà enregistrés dans le folder
    var files = folder.getFiles()
    var msgIds = []
    while (files.hasNext()) {
      var file = files.next();
      msgIds.push(file.getDescription())
    }

    /* Gmail Label that contains the queue */
    var label = GmailApp.getUserLabelByName(gmailLabel)

    for (var t = 0; t < threads.length; t++) {

      var msgs = threads[t].getMessages()
      var html = ""
      var attachments = []

      var subject = threads[t].getFirstMessageSubject()
      var isMsgFounded = false

      /* Append all the threads in a message in an HTML document */
      for (var m = 0; m < msgs.length; m++) {
        var msg = msgs[m]
        if (msgIds.indexOf(msg.getId()) + 1) {
          continue
        }
        isMsgFounded = true
        html += "De: " + msg.getFrom() + "<br />"
        html += "a&#768;: " + msg.getTo() + "<br />"
        if (msg.getCc())
          html += "En copie: " + msg.getCc() + "<br />"
        if (msg.getBcc())
          html += "En copie caché: " + msg.getBcc() + "<br />"
        html += "Date: " + Utilities.formatDate(msg.getDate(), "GMT", "dd/MM/yyyy HH:mm:ss") + "<br />"
        html += "Objet: " + msg.getSubject() + "<br />"
        html += "<hr />"
        html += msg.getBody().replace(/<img[^>]*>/g, "")
        html += "<hr />"

        var atts = msg.getAttachments()
        for (var a = 0; a < atts.length; a++) {
          attachments.push(atts[a])
        }
      }

      /* Save the attachment files and create links in the document's footer */
      if (attachments.length > 0) {
        var folderIterator = folder.getFoldersByName("pj")
        var folderAttachment
        if (folderIterator.hasNext()) {
          folderAttachment = folderIterator.next()
        } else {
          folderAttachment = folder.createFolder("pj")
        }
        var footer = "<strong>Pie&#768;ces jointes:</strong><ul>"
        for (var z = 0; z < attachments.length; z++) {
          var file = folderAttachment.createFile(attachments[z])
          footer += "<li><a href='" + file.getUrl() + "'>" + file.getName() + "</a></li>"
        }
        html += footer + "</ul>"
      }

      /* Conver the Email Thread into a PDF File */
      if (isMsgFounded) {
        var tempFile = DriveApp.createFile("temp.html", html, "text/html")
        var pdf = folder.createFile(tempFile.getAs("application/pdf")).setName("Mail - " + subject + ".pdf")
        pdf.setDescription(threads[t].getId())
        tempFile.setTrashed(true)
        mailsArchived.push(subject)
      }

      threads[t].removeLabel(label)

    }
  }
  if (mailsArchived.length > 0) {
    ui.alert("Archivage des mails", mailsArchived.toString(), ui.ButtonSet.OK)
  } else {
    ui.alert("Archivage des mails", "Aucun mail à archiver", ui.ButtonSet.OK)
  }
}


 /**
  * Archivage dans un répertoire de Drive
  * des mails avec un subject particulier
  * @param {String} rangeSubject  B10
  * @param {String} urlIforRangeOfFolder https://drive.google.com/drive/folders/14ZQMmdVPBDVU09Htw-td4rdKgR2Y6TW0 ou cellule
  * @param {String} rangeResult   B12
  */
function fx_archiveGmail(rangeSubject, urlIforRangeOfFolder, rangeResult) {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
    "Récupérer la conversation",
    'Veuillez confirmer par Oui ou Non',
    ui.ButtonSet.YES_NO);
  if (yesnoConfirm != ui.Button.YES) return;

  var spreadsheet = SpreadsheetApp.getActive();
  var subject = spreadsheet.getRange(rangeSubject).getValue()
  var threads = GmailApp.search("subject:" + subject, 0, 5)

  var mailsArchived = []

  if (threads.length > 0) {

    /* Google Drive folder where the Files would be saved */
    if ( urlIforRangeOfFolder.indexOf("https") > -1 ) {
      var urlOrId = urlIforRangeOfFolder
    } else {
      var urlOrId = spreadsheet.getRange(urlIforRangeOfFolder).getValue()
    }
    var folder = DriveApp.getFolderById(fx_getIdFromUrl(urlOrId));
    // mémorisation des msgid déjà enregistrés dans le folder
    var files = folder.getFiles()
    var msgIds = []
    while (files.hasNext()) {
      var file = files.next();
      msgIds.push(file.getDescription())
    }

    for (var t = 0; t < threads.length; t++) {

      var msgs = threads[t].getMessages()
      var html = ""
      var attachments = []

      var subject = threads[t].getFirstMessageSubject()
      var isMsgFounded = false

      /* Append all the threads in a message in an HTML document */
      for (var m = 0; m < msgs.length; m++) {
        var msg = msgs[m]
        // if (msgIds.indexOf(msg.getId()) + 1) {
        //   continue
        // }
        isMsgFounded = true
        html += "De: " + msg.getFrom() + "<br />"
        html += "a&#768;: " + msg.getTo() + "<br />"
        if (msg.getCc())
          html += "En copie: " + msg.getCc() + "<br />"
        if (msg.getBcc())
          html += "En copie caché: " + msg.getBcc() + "<br />"
        html += "Date: " + Utilities.formatDate(msg.getDate(), "GMT", "dd/MM/yyyy HH:mm:ss") + "<br />"
        html += "Objet: " + msg.getSubject() + "<br />"
        html += "<hr />"
        html += msg.getBody().replace(/<img[^>]*>/g, "")
        html += "<hr />"

        var atts = msg.getAttachments()
        for (var a = 0; a < atts.length; a++) {
          attachments.push(atts[a])
        }
      }

      /* Save the attachment files and create links in the document's footer */
      if (attachments.length > 0) {
        var folderIterator = folder.getFoldersByName("pj")
        var folderAttachment
        if (folderIterator.hasNext()) {
          folderAttachment = folderIterator.next()
        } else {
          folderAttachment = folder.createFolder("pj")
        }
        var footer = "<strong>Pie&#768;ces jointes:</strong><ul>"
        for (var z = 0; z < attachments.length; z++) {
          var file = folderAttachment.createFile(attachments[z])
          footer += "<li><a href='" + file.getUrl() + "'>" + file.getName() + "</a></li>"
        }
        html += footer + "</ul>"
      }

      /* Conver the Email Thread into a PDF File */
      if (isMsgFounded) {
        var tempFile = DriveApp.createFile("temp.html", html, "text/html")
        var pdf = folder.createFile(tempFile.getAs("application/pdf")).setName("Mail - " + subject + ".pdf")
        pdf.setDescription(threads[t].getId())

        tempFile.setTrashed(true)
        mailsArchived.push(subject)

        // copie url du pdf dans la cellule résultat
        spreadsheet.getRange(rangeResult).setValue(pdf.getUrl());
      }

    }
  }
  if (mailsArchived.length > 0) {
    ui.alert("Archivage des mails", mailsArchived.toString(), ui.ButtonSet.OK)
  } else {
    ui.alert("Archivage des mails", "Aucun mail à archiver", ui.ButtonSet.OK)
  }
}  

function recupererEmail() {

}