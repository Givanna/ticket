Office.onReady(function (info) {
  // Office is ready
  if (info.host === Office.HostType.Outlook) {
      // Ensure the user is in compose mode
      Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
  }
});

// Handler for when the item changes (when the recipient is changed)
function itemChanged() {
  // Get the current item
  var item = Office.context.mailbox.item;

  // Get the To recipients
  item.to.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          const recipients = result.value;

          // Check if the specific recipient is included
          const recipientEmail = 'givannawright@gmail.com';
          if (recipients.some(r => r.emailAddress.toLowerCase() === recipientEmail.toLowerCase())) {
              // Set the subject
              item.subject.setAsync('Support Ticket', function (result) {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                      console.log('Subject set to Support Ticket.');
                  } else {
                      console.error('Failed to set subject: ' + result.error.message);
                  }
              });

              // Set the body
              const bodyContent = `Who is this for?\nWhat is the issue?`;
              item.body.setAsync(bodyContent, { coercionType: Office.CoercionType.Text }, function (result) {
                  if (result.status === Office.AsyncResultStatus.Succeeded) {
                      console.log('Body populated with questions.');
                  } else {
                      console.error('Failed to set body: ' + result.error.message);
                  }
              });
          }
      } else {
          console.error('Failed to get recipients: ' + result.error.message);
      }
  });
}
