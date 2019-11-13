/**
 * MaixentExpoApp : Scripts réutilisables
 * - ajouter la bibliothèque "MTHmRiwDeD0HAcidrAQ9BrTiZW5tJ2woG" dans Ressources/Bibliothèques
 * - mettre le préfixe MaixentExpoApp devant la fonction que vous voulez utiliser
 * MaixentExpoApp.LaFonction
 */

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Email', functionName: 'sendEmail'},
    {name: 'Alert...', functionName: 'alert'},
    {name: 'showPrompt...', functionName: 'showPrompt'},
    {name: 'showAlert...', functionName: 'showAlert'}
  ];
  spreadsheet.addMenu('Foire Expo', menuItems);
}


function alert() {
  Browser.msgBox('Info', 'Row does not contain two addresses.',
        Browser.Buttons.OK);
}
function sendEmail() {
  var recipient = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(recipient, 'Email from your site', 'You clicked a link!');
}
function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Let\'s get to know each other!',
      'Please enter your name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert(Utilities.formatString('Your name is %s.', text));
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}