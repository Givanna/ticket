// taskpane.js

Office.initialize = function (reason) {
    $(document).ready(function () {
        // Check if the Office add-in is ready
        console.log("Office Add-in is ready");

        // Add an event handler for when the recipient changes
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
    });
};

function itemChanged() {
    // Get the current email item
    var item = Office.context.mailbox.item;

    // Check if the item is a message
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
        // Get the email addresses of the recipients
        item.to.getAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var recipients = result.value;
                recipients.forEach(function (recipient) {
                    if (recipient.emailAddress === "givannawright@gmail.com") {
                        // Populate the subject and body
                        item.subject.setAsync("Support Ticket", function (result) {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                console.log("Subject set successfully.");
                            } else {
                                console.error("Error setting subject: " + result.error.message);
                            }
                        });

                        item.body.setAsync("Who is this for?\nWhat is the issue?", { coercionType: Office.CoercionType.Text }, function (result) {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                console.log("Body set successfully.");
                            } else {
                                console.error("Error setting body: " + result.error.message);
                            }
                        });
                    }
                });
            } else {
                console.error("Error getting recipients: " + result.error.message);
            }
        });
    }
}
