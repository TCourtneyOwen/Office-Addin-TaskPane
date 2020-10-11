/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  var msgFrom = Office.context.mailbox.item.from;
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: msgFrom.displayName,
    icon: "Icon.80x80",
    persistent: true
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
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

function validateSendable(event) {
	Office.context.mailbox.item.loadCustomPropertiesAsync(
		function (asyncResult)
		{
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded)
			{
				var customProps = asyncResult.value;
				var isSendable = customProps.get('isSendable');
				isSendable = isSendable === undefined ? true : isSendable;
				event.completed({ allowEvent: isSendable });
			} else {
				Office.context.mailbox.item.notificationMessages.replaceAsync(
				"isSendable",
				{
					type: "informationalMessage", icon: "icon1",
					message: "Loading toggle state failed.",
					persistent: false
				}, function (result) {
					// if (result.status == "failed") {
					// 	var statusString = 'Failed ' + result.error.code + ': ' + result.error.name + ': ' + result.error.message;
					// 	console.log(statusString);
					// }
					event.completed({ allowEvent: true });
				});
			}
		}
	);
}

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.action = action;
g.validateSendable = validateSendable;
