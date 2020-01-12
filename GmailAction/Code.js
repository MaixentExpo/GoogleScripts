
function getContextualAddOn(e) {
    // Activate temporary Gmail add-on scopes, in this case to allow
    // message metadata to be read.
    var accessToken = e.messageMetadata.accessToken;
    GmailApp.setCurrentMessageAccessToken(accessToken);

    var messageId = e.messageMetadata.messageId;
    var message = GmailApp.getMessageById(messageId);
    var subject = message.getSubject();
    var sender = message.getFrom();

    // Create a card with a single card section and two widgets.
    // Be sure to execute build() to finalize the card construction.
    var exampleCard = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
            .setTitle('Example card'))
        .addSection(CardService.newCardSection()
            .addWidget(CardService.newKeyValue()
                .setTopLabel('Subject')
                .setContent(subject))
            .addWidget(CardService.newKeyValue()
                .setTopLabel('From')
                .setContent(sender)))
        .build();   // Don't forget to build the Card!
    return [exampleCard];
  }