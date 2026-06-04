Office.onReady(() => {});

function onMessageSendHandler(event) {
  const settings = Office.context.roamingSettings;

  if (settings.get('sendImmediately')) {
    settings.set('sendImmediately', false);
    settings.saveAsync(() => {
      event.completed({ allowEvent: true });
    });
    return;
  }

  const delaySeconds = settings.get('delaySeconds') || 120;
  const sendAt = new Date(Date.now() + delaySeconds * 1000);

  Office.context.mailbox.item.delayDeliveryTime.setAsync(sendAt, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("delayDeliveryTime.setAsync failed:", asyncResult.error.message);
      event.completed({ allowEvent: true });
      return;
    }
    event.completed({ allowEvent: true });
  });
}

function sendNowHandler(event) {
  Office.context.roamingSettings.set('sendImmediately', true);
  Office.context.roamingSettings.saveAsync(() => {
    Office.context.mailbox.item.notificationMessages.replaceAsync("delay-send-notice", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Send Now active — click Send to bypass the delay.",
      icon: "none",
      persistent: true
    }, () => { event.completed(); });
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("sendNowHandler", sendNowHandler);
