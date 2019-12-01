/**
 * MaixentExpoApp : Scripts réutilisables
 * - ajouter la bibliothèque "1jhMQC2ecQ90hLSQx5hrgo6rVq_DqHfO_3i2lYH1JmMteRXb4S5GJ5DxN" 
 *   dans Ressources/Bibliothèques
 * ou recopier le code ci-dessous dans appsscript.json

{
  "timeZone": "Europe/Paris",
  "dependencies": {
    "libraries": [{
      "userSymbol": "MaixentExpoApp",
      "libraryId": "1jhMQC2ecQ90hLSQx5hrgo6rVq_DqHfO_3i2lYH1JmMteRXb4S5GJ5DxN",
      "version": "10",
      "developmentMode": true
    }]
  },
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": []
}

 * - mettre le préfixe MaixentExpoApp devant la fonction que vous voulez utiliser
 * Exemples :

function envoyerMessage() {
   MaixentExpoApp.fx_envoyerMessage();
}
function echangerAdresses() {
   MaixentExpoApp.mailing_echangeAdresses();
}

 */

