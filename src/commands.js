Office.onReady(() => {});

function onMessageSendHandler(event) {
  const DELAY_SECONDS = 120;
  const sendAt = new Date(Date.now() + DELAY_SECONDS * 1000);

  Office.context.mailbox.item.delayDeliveryTime.setAsync(
    sendAt,
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("delayDeliveryTime.setAsync failed:", asyncResult.error.message);
        event.completed({ allowEvent: true });
        return;
      }
      event.completed({ allowEvent: true });
    }
  );
}

globalThis.onMessageSendHandler = onMessageSendHandler;
