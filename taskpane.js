/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const item = Office.context.mailbox.item;

  // Get the recipient's email address
  const recipientEmail = "givannawright@gmail.com"; // Adjust this to your target recipient
  
  // Check if the current recipient matches
  const currentRecipients = item.to; // Get the current recipients

  if (currentRecipients.some(recipient => recipient.emailAddress === recipientEmail)) {
    // Set the subject
    item.subject = "Support Ticket";

    // Create the body content
    const bodyContent = `
      <p>What is the issue?</p>
      <p>What is the best time to reach you?</p>
    `;
    
    // Set the body content
    await item.body.setAsync(bodyContent, { coercionType: Office.CoercionType.Html });
  } else {
    // Handle case where the recipient does not match
    console.log("Recipient does not match the target email.");
  }
}
