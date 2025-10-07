var clickEvent;
// The initialize function must be defined each time a new page is loaded
(function () {
    Office.initialize = function (reason) {
       // If you need to initialize something you can do so here.
    };
})();

// Wrap the writeToDoc in showNotification because showNotification is called
// in DialogHelper.js but must be defined differently when the dialog is called
// from a task pane instead of a custom menu command.
function showNotification(text) {
    writeToDoc(text);
    //Required, call event.completed to let the platform know you are done processing.
    clickEvent.completed();
}

//Notice function needs to be in global namespace
function doSomethingAndShowDialog(event) {
    clickEvent = event;
    writeToDoc("Ribbon button clicked.");
    openDialogAsIframe();
}

function writeToDoc(text) {
  Office.context.document.setSelectedDataAsync(text, function (asyncResult) {
      var error = asyncResult.error;
      if (asyncResult.status === "failed") {
          console.log("Unable to write to the document: " + asyncResult.error.message);
      }
  });
}
