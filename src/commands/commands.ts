// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console */

import { createGroupConversation, createGroupLink } from "../api/api";
import { appendToBody, getSubject, createMeetingLinkElement, getOrganizer, setLocation } from "../utils/mailbox";

// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

const defaultSubjectValue = "New Appointment";
let mailboxItem;

let pendingCreateConversation = false;

export function addMeetingLink() {
  let isLoggedIn: boolean = false;

  let dialog;

  pendingCreateConversation = true;

  const token = localStorage.getItem("token");
  const refreshToken = localStorage.getItem("refresh_token");
  console.log("token from local storage: ", token);
  console.log("refresh token from local storage: ", refreshToken);

  if (!token && !refreshToken) {
    console.log("open dialog");

    Office.context.ui.displayDialogAsync(
      "https://outlook.integrations.zinfra.io/login",
      { height: 60, width: 40 },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("dialog result failed: " + asyncResult.error.message);
        } else {
          const dialog = asyncResult.value;
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (messageEvent: Office.DialogParentMessageReceivedEventArgs) => {
              console.log("DialogMessageReceived");
              const authResult = JSON.parse(messageEvent.message) as any;
              console.log("Auth result:", authResult);
              localStorage.setItem("token", authResult.token);
              localStorage.setItem("refresh_token", authResult.refresh_token);

              if (pendingCreateConversation) {
                createGroupConversationForCurrentMeeting();
                pendingCreateConversation = false;
              }

              dialog.close();
            }
          );
        }
      }
    );
  } else {
    if (pendingCreateConversation) {
      createGroupConversationForCurrentMeeting();
      pendingCreateConversation = false;
    }
  }
}

function createGroupConversationForCurrentMeeting() {
  getSubject(mailboxItem, (subject) => {
    createGroupConversation(subject ? subject : defaultSubjectValue).then((r) => {
      if (r) {
        createGroupLink(r).then((r) => {
          getOrganizer(mailboxItem, function (organizer) {
            setLocation(mailboxItem, r, () => {});
            const groupLink = createMeetingLinkElement(r, organizer);
            appendToBody(mailboxItem, groupLink);
          });
        });
      }
    });
  });
}

function test() {
  // test some code here with the test button in UI
}

// Register the functions.
Office.actions.associate("addMeetingLink", addMeetingLink);

// TODO: remove after DEV is complete
Office.actions.associate("test", test);
