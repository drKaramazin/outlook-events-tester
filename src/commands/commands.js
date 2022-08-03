/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

let isReady = false;
let interval$;

Office.onReady(() => {
  // If needed, Office.js is ready to be called
  isReady = true;
});

function displayDialogAsync(dialogUrl, dialogOptions) {
  return new Promise((resolve) => {
    const dialogClosed = async () => {
      resolve();
    };

    Office.context.ui.displayDialogAsync(dialogUrl, dialogOptions, (result) => {
      let dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
    });
  });
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // const message = {
  //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  //   message: "Performed action.",
  //   icon: "Icon.80x80",
  //   persistent: true,
  // };

  // Show a notification message
  // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  displayDialogAsync(window.location.origin + "/dialog.html", {
    width: 30,
    height: 20,
  }).then(() => {
    event.completed();
  });

  // Be sure to indicate when the add-in command function is complete
  // setTimeout(() => event.completed(), 30000);
}

function itemSentHolder(event) {
  displayDialogAsync(window.location.origin + "/dialog.html", {
    width: 30,
    height: 20,
  }).then(() => {
    event.completed({ allowEvent: true });
  });
}

function itemSent(event) {
  interval$ = setInterval(() => {
    if (isReady) {
      clearInterval(interval$);
      itemSentHolder(event);
    }
  }, 200);
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
g.itemSent = itemSent;
