// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console, fetch */

// sadly, no imports for event-based activation so everything has to be here

// Office is ready. Init
Office.onReady(function () {
    if (Office.context.requirements.isSetSupported('Mailbox', '1.12')) {
        Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
    }
});

function onAppointmentSendHandler(event) {
    console.log("onAppointmentSendHandler executed");
  }