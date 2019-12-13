
## Installation de clasp

    sudo apt install npm
    sudo npm install @google/clasp -g

    clasp login
    clasp status
    clasp versions
    clasp pull
    clasp push

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

