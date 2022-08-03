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

function holder(event, message) {
  displayDialogAsync("/dialog.html", {
    width: 30,
    height: 20,
  });
}

function onMessageComposeHandler(event) {
  holder(event, "onMessageComposeHandler");
}

function onAppointmentComposeHandler(event) {
  holder(event, "onAppointmentComposeHandler");
}

function onMessageAttachmentsChangedHandler(event) {
  holder(event, "onMessageAttachmentsChangedHandler");
}

function onAppointmentAttachmentsChangedHandler(event) {
  holder(event, "onAppointmentAttachmentsChangedHandler");
}

function onMessageRecipientsChangedHandler(event) {
  holder(event, "onMessageRecipientsChangedHandler");
}

function onAppointmentAttendeesChangedHandler(event) {
  holder(event, "onAppointmentAttendeesChangedHandler");
}

function onAppointmentTimeChangedHandler(event) {
  holder(event, "onAppointmentTimeChangedHandler");
}

function onAppointmentRecurrenceChangedHandler(event) {
  holder(event, "onAppointmentRecurrenceChangedHandler");
}

function onInfobarDismissClickedHandler(event) {
  holder(event, "onInfobarDismissClickedHandler");
}

Office.onReady(() => {
  // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
  Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
  Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);

  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
  Office.actions.associate("onAppointmentAttachmentsChangedHandler", onAppointmentAttachmentsChangedHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  Office.actions.associate("onAppointmentAttendeesChangedHandler", onAppointmentAttendeesChangedHandler);
  Office.actions.associate("onAppointmentTimeChangedHandler", onAppointmentTimeChangedHandler);
  Office.actions.associate("onAppointmentRecurrenceChangedHandler", onAppointmentRecurrenceChangedHandler);
  Office.actions.associate("onInfobarDismissClickedHandler", onInfobarDismissClickedHandler);
});
