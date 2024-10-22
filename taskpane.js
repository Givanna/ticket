document.getElementById('automateButton').onclick = function() {
    // Functionality to automate ticket filling goes here
    Office.context.mailbox.item.subject.setAsync("Automated Subject", function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Subject set successfully.");
        } else {
            console.error("Error setting subject: " + result.error.message);
        }
    });
};
