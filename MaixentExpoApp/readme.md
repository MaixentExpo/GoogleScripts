
## Installation de clasp

    sudo apt install npm
    sudo npm install @google/clasp -g

    clasp login
    clasp status
    clasp versions
    clasp pull
    clasp push

## Activation de API Google Apps Script
https://script.google.com/home/usersettings

## Scopes nécessaire pour le projet MaixentExpoApp
```json
{
  "timeZone": "Europe/Paris",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```


## Déclarations des applications utilisatrices

### Code.js
```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Foire Expo')
      .addItem('Action 1', 'fonction1')
      .addToUi();
}
```

### appscript.json
```json
{
  "timeZone": "Europe/Paris",
  "dependencies": {
    "libraries": [{
      "userSymbol": "MaixentExpoApp",
      "libraryId": "1jhMQC2ecQ90hLSQx5hrgo6rVq_DqHfO_3i2lYH1JmMteRXb4S5GJ5DxN",
      "version": "11",
      "developmentMode": true
    }]
  },
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": []
}
```

### Exemples d'appel
```javascript
function envoyerMessage() {
    MaixentExpoApp.fx_envoyerMessage();
}
function echangerAdresses() {
    MaixentExpoApp.mailing_echangeAdresses();
}
```

