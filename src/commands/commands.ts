/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

void Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event - The event object from the add-in command
 */
function action(event: Office.AddinCommands.Event): void {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item?.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal(): typeof globalThis | undefined {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as typeof globalThis & { action?: typeof action };

// The add-in command functions need to be available in global scope
g.action = action;
