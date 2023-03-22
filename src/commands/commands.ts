// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console */

import { createGroupConversation, createGroupLink } from "../api/api";
import { appendToBody, getSubject } from "../utils/mailbox";

// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

const defaultSubjectValue = "New Appointment";
let mailboxItem;

export function test() {
  const tokenExpired: boolean = true;
  let isLoggedIn: boolean = false;

  let dialog;

  console.log('open dialog');

  Office.context.ui.displayDialogAsync('https://outlook.integrations.zinfra.io/login', { height: 50, width: 50 },
    (result: Office.AsyncResult<Office.Dialog>) => {
      const dialog = result.value;
  
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args: Office.DialogParentMessageReceivedEventArgs) => {
        const messageFromDialog = args.message;
  
        console.log("messageParent:", messageFromDialog);
        Office.context.ui.messageParent(messageFromDialog);
        dialog.close();
      });
    }
  );

  if(isLoggedIn) {
    getSubject(mailboxItem, (subject) => {
      createGroupConversation(subject ?? defaultSubjectValue).then((r) => {
        createGroupLink(r).then((r) => {
          const groupLink = `<a href="${r}">${r}</a>`;
          appendToBody(mailboxItem, groupLink);
        });
      });
    });
  }
  
  // maybe can be done better ?
  // createMeetingLinkElement().then((meetingLink) => {
  //   appendToBody(mailboxItem, meetingLink);
  // });
}

function addMeetingLink() {
  test();
}

function appendDisclaimerOnSend(event) {
  // Calls the getTypeAsync method and passes its returned value to the options.coercionType parameter of the appendOnSendAsync call.
  mailboxItem.body.getTypeAsync(
    {
      asyncContext: event,
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      // Sets the disclaimer to be appended to the body of the message on send.
      const bodyFormat = asyncResult.value;
      let meetingLink = "<p>Testing wire addin</p>";

      mailboxItem.body.appendOnSendAsync(
        meetingLink,
        {
          asyncContext: asyncResult.asyncContext,
          coercionType: bodyFormat,
        },

        async (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }

          asyncResult.asyncContext.completed();
        }
      );
    }
  );
}

// Register the functions.
Office.actions.associate("test", test);
Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
Office.actions.associate("addMeetingLink", addMeetingLink);
