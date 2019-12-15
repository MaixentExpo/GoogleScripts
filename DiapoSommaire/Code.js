/**
 * Scripts pour récupérer les copies d'écrans des diapositives
 * - problème de lenteur
 * - problème de présentation des diapos
 */


function onOpen() {
  var ui = SlidesApp.getUi();
  ui.createMenu('Foire Expo')
      .addItem('Création du sommaire', 'sommaire')
      .addToUi();
}

function sommaire() {
  diapoSommaire();  
}

/**
 * diapoSommaire
 */
function diapoSommaire() {
  var presentation = SlidesApp.getActivePresentation();
  // Retrieve slides as images
  var id = presentation.getId();
  //var accessToken = ScriptApp.getOAuthToken();
  // Recup de l'id de chaque dispo
  var pageObjectIds = presentation.getSlides().map(function (e) { return e.getObjectId() });
  // construction des Urls
  // url: "https://slides.googleapis.com/v1/presentations/" + id + "/pages/" + pageObjectId + "/thumbnail?access_token=" + accessToken,
  var reqUrls = pageObjectIds.map(function (pageObjectId) {
    return {
      method: "get",
      url: "https://slides.googleapis.com/v1/presentations/" + id + "/pages/" + pageObjectId + "/thumbnail",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
  });
  // Soumission des Urls
  var reqBlobs = UrlFetchApp.fetchAll(reqUrls).map(function (e) {
    var r = JSON.parse(e);
    return {
      method: "get",
      url: r.contentUrl
    };
  });
  var reqClean = [];
  for (var i=0; i<reqBlobs.length; i++) {
    if ( typeof reqBlobs[i].url === 'undefined' ) {
      ;
    } else {
      reqClean.push(reqBlobs[i]);
    }
  } // endfor
  // Recup des Images générées dans blobs
  var blobs = UrlFetchApp.fetchAll(reqClean).map(function (e) {
    return e.getBlob()
  });

  // Ajout de slides Sommaire
  var col = 5; // Number of columns
  var row = 4; // Number of rows
  var wsize = 130; // Size of width of each image (pixels)
  var sep = 5; // Space of each image (pixels)

  var ph = presentation.getPageHeight(); // 540 px
  var pw = presentation.getPageWidth();  // 720 px
  var leftOffset = (pw - ((wsize * col) + (sep * (col - 1)))) / 2;
  if (leftOffset < 0) throw new Error("Images are sticking out from a slide.");
  var len = col * row;
  var loops = Math.ceil(blobs.length / (col * row));
  for (var loop = 0; loop < loops; loop++) {
    var ns = presentation.insertSlide(loop);
    var topOffset, top;
    var left = leftOffset;
    for (var i = len * loop; i < len + (len * loop); i++) {
      if (i === blobs.length) break;
      var image = ns.insertImage(blobs[i]);
      var w = image.getWidth();
      var h = image.getHeight();
      var hsize = h * wsize / w;
      if (i === 0 || i % len === 0) {
        topOffset = (ph - ((hsize * row) + sep)) / 2;
        if (topOffset < 0) throw new Error("Images are sticking out from a slide.");
        top = topOffset;
      }
      image.setWidth(wsize).setHeight(hsize).setTop(top).setLeft(left).getObjectId();
      //if (i === col - 1 + (loop * len)) {
      if ( loop % col === 0 ) {
        top = topOffset + hsize + sep;
        left = leftOffset;
      } else {
        left += wsize + sep;
      }
    }
  }
  presentation.saveAndClose();
} // end function diapoSommaire
