"use strict";

var messageBanner;

Office.initialize = () => {
    setAutoOpenOn();
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem(TARGET.sheet);
      sheet.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
    }).catch(window.ErrorHandler.handleError);
    
    $(document).ready(function () {
        // Initialize the FabricUI notification mechanism and hide it
        var element = document.querySelector('.ms-MessageBanner');
        messageBanner = new app.notification.MessageBanner(element);
        //messageBanner.hideBanner();
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            $('#subtitle').text("Opps!");
            $("#template-description").text("Sorry, this sample requires Word 2016 or later. The button will not open a dialog.");
            $('#button-text').text("Button");
            $('#button-desc').text("Button that opens dialog only on Word 2016 or later.");
            return;
        }
        $('#action-button').click(openDialog);
        $('#action-button2').click(openDialog);
    });
};

async function handleSelectionChanged(event) {
  await Excel.run(async (context) => {
    if (event.address !== TARGET.address) return;
    await context.sync();
  }).catch(window.ErrorHandler.handleError);
};

function errorHandler(error) {
    showNotification(error);
}

function showNotification(content) {
    $("#notificationBody").text(content);
    $("#subtitle").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
}

function setAutoOpenOn() {
    Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
    Office.context.document.settings.saveAsync();
}
