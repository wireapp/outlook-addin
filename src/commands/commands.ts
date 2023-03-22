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

let pendingCreateConversation = false;

export function test() {
  const tokenExpired: boolean = true;
  let isLoggedIn: boolean = false;

  let dialog;

  pendingCreateConversation = true;

  const token = localStorage.getItem('token');
  const refreshToken = localStorage.getItem('refresh_token');
  console.log('token from local storage: ', token);
  console.log('refresh token from local storage: ', refreshToken);

  if(!token && !refreshToken) {
    console.log('open dialog');

    Office.context.ui.displayDialogAsync('https://outlook.integrations.zinfra.io/login', { height: 60, width: 40 }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("dialog result failed: " + asyncResult.error.message);
      } else {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (messageEvent: Office.DialogParentMessageReceivedEventArgs) => {
          console.log('DialogMessageReceived');
          const authResult = JSON.parse(messageEvent.message) as any;
          console.log('Auth result:', authResult);
          localStorage.setItem('token', authResult.token);
          localStorage.setItem('refresh_token', authResult.refresh_token);

          if(pendingCreateConversation) {
            createGroupConversationForCurrentMeeting();
            pendingCreateConversation = false;
          }

          dialog.close();
        });
      }
    });
  } else {
    if(pendingCreateConversation) {
      createGroupConversationForCurrentMeeting();
      pendingCreateConversation = false;
    }
  }
  
  // maybe can be done better ?
  // createMeetingLinkElement().then((meetingLink) => {
  //   appendToBody(mailboxItem, meetingLink);
  // });
}

function createGroupConversationForCurrentMeeting() {
  getSubject(mailboxItem, (subject) => {
    createGroupConversation(subject ? subject : defaultSubjectValue).then((r) => {
      createGroupLink(r).then((r) => {
        const groupLink = `<a href="${r}">${r}</a>`;
        appendToBody(mailboxItem, groupLink);
      });
    });
  });
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
