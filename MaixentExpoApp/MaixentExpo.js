/**
 * Fonctions communes javascript
 * MaixentExpo@gmail.com
 */

function sendMail() {
  var ui = SpreadsheetApp.getUi();
  var yesnoConfirm = ui.alert(
     "Envoiyer les mails",
     'Veuillez confirmer par oui ou non',
      ui.ButtonSet.YES_NO);
  if ( yesnoConfirm != ui.Button.YES ) return;
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  var to = spreadsheet.getRangeByName("TO").getCell(1,1).getValue();
  var copy = spreadsheet.getRangeByName("COPY").getCell(1,1).getValue();
  var cc = spreadsheet.getRangeByName("CC").getCell(1,1).getValue();
  var replyTo = spreadsheet.getRangeByName("REPLYTO").getCell(1,1).getValue();
  var subject = spreadsheet.getRangeByName("SUBJECT").getCell(1,1).getValue();
  var richText = spreadsheet.getRangeByName("BODY").getCell(1,1).getRichTextValue();
  
  var html = htmlEncodeRichText(richText);
  // https://developers.google.com/apps-script/reference/mail/mail-app
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
 * Given a RichTextValue Object, iterate over the individual runs
 *    and call out to htmlStyleRtRun() to return the text wrapped
 *    in <span> tags with specific styling.
 * @param {RichTextValue} richTextValue a RichTextValue object
 *    from a given Cell.
 * @return {string} HTML encoded text 
 */
function htmlEncodeRichText(richTextValue) {
  // create an empty string which will hold the html content
  var htmlString = "";
  // get an array of Runs for the given Rich Text
  var rtRuns = richTextValue.getRuns();
  // loop the array
  for (var i = 0; i < rtRuns.length; i++) {
    // return html version of a given run, append to existing string
    htmlString += htmlStyleRtRun(rtRuns[i]);
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
function htmlStyleRtRun(richTextRun) {
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
 * Class Couleur
 * qui fournit un code couleur à chaque appel new_couleur()
 */
var Couleur = function() {
  this.couleurs = [ "#e8f5e9" // green
                   ,"#e3f2fd" // blue
                   ,"#fffde7" // yellow
                   ,"#fbe9e7" // deep orange
                   ,"#e0f7fa" // cyan
                   ,"#f1f8e9" // light green
                   ,"#fce4ec" // pink
                   ,"#e1f5fe" // light blue
                   ,"#ede7f6" // deep purple
                   ,"#eceff1" // blue grey
                   ,"#e8eaf6" // indigo
                   ,"#f3e5f5" // purple
                   ,"#f9fbe7" // lime
                   ,"#fff3e0" // orange
                   ,"#fff8e1" // amber
                   ,"#efebe9" // brown
                   ,"#e0f2f1" // teal
                   ,"#ffebee" // red
                   ,"#fafafa" // grey
                  ];
  this.iCouleur = -1;
  this.couleur = "#fafafa";
  this.new_couleur = function () {
    this.iCouleur++
    if ( this.iCouleur >= this.couleurs.length ) {
      this.iCouleur = 0;
    } // endif
    this.couleur = this.couleurs[this.iCouleur];
  } // end new_couleur
} // end class Couleur
