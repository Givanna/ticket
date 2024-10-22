Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.item.to.getAsync(function(result) {
            let recipients = result.value;
            recipients.forEach(function(recipient) {
                if (recipient.emailAddress === "givannawright@gmail.com") {
                    // Set the subject to "Support Ticket"
                    Office.context.mailbox.item.subject.setAsync("Support Ticket");
                    
                    // Set the body of the email
                    let bodyContent = "Who is it for?\nWhat is the issue?";
                    Office.context.mailbox.item.body.setAsync(bodyContent, {coercionType: Office.CoercionType.Text});
                }
            });
        });
    }
});
