// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console */

//import { createGroupConversation, createGroupLink } from "../api/api";
import { createEvent } from "../api/api";
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

  isLoggedIn = JSON.parse(localStorage.getItem("isLoggedIn"));

  if (!isLoggedIn) {
    authorizeDialog();
  } else {
    if (pendingCreateConversation) {
      createGroupConversationForCurrentMeeting();
      pendingCreateConversation = false;
    }
  }
}

function authorizeDialog() {
  let isLoggedIn : boolean = false;

  console.log("open dialog");

  Office.context.ui.displayDialogAsync(
    "https://outlook.integrations.zinfra.io/authorize",
    { height: 70, width: 40 },
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

            if(authResult.success) {
              isLoggedIn = true;
              localStorage.setItem("isLoggedIn", String(isLoggedIn));
            }

            if (isLoggedIn && pendingCreateConversation) {
              createGroupConversationForCurrentMeeting();
              pendingCreateConversation = false;
            }

            dialog.close();
          }
        );
      }
    }
  );

  return isLoggedIn;
}

function createGroupConversationForCurrentMeeting() {
  getSubject(mailboxItem, (subject) => {
    createEvent(subject || defaultSubjectValue).then((r) => {
      if (r) {
        getOrganizer(mailboxItem, function (organizer) {
          setLocation(mailboxItem, r.link, () => {});
          const groupLink = createMeetingLinkElement(r.link, organizer);
          appendToBody(mailboxItem, groupLink);
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
