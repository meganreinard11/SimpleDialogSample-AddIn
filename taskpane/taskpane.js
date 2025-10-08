"use strict";

var messageBanner;

// The initialize function must be defined each time a new page is loaded.
Office.initialize = function (reason) {
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

function errorHandler(error) {
    showNotification(error);
}

function showNotification(content) {
    $("#notificationBody").text(content);
    $("#subtitle").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
}
