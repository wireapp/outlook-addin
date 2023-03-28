// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console */

//import { createGroupConversation, createGroupLink } from "../api/api";
import { createEvent } from "../api/api";
import { AuthResult } from "../api/types";
import { appendToBody, getSubject, createMeetingLinkElement, getOrganizer, setLocation } from "../utils/mailbox";



// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

const defaultSubjectValue = "New Appointment";
let mailboxItem;

let pendingCreateConversation = false;

async function addMeetingLink() {
  let isLoggedIn = JSON.parse(localStorage.getItem("isLoggedIn"));

  if (!isLoggedIn) {
    isLoggedIn = await authorizeDialog();
  }

  if (isLoggedIn) {
    try {
      await fetchWithAuthorizeDialog("/createGroupConversation", {
        method: "POST",
        body: JSON.stringify(mailboxItem),
        headers: {
          "Content-Type": "application/json"
        }
      });
    } catch (error) {
      if (error.status === 401) {
        localStorage.removeItem("isLoggedIn");
        isLoggedIn = await authorizeDialog();
        if (isLoggedIn) {
          await fetchWithAuthorizeDialog("/createGroupConversation", {
            method: "POST",
            body: JSON.stringify(mailboxItem),
            headers: {
              "Content-Type": "application/json"
            }
          });
        }
      } else {
        console.error(error);
      }
    }
  }
}

async function fetchWithAuthorizeDialog(url: string, options: RequestInit): Promise<Response> {
  try {
    const response = await fetch(url, options);
    if (!response.ok) {
      if (response.status === 401) {
        const isLoggedIn = await authorizeDialog();
        if (isLoggedIn) {
          return await fetch(url, options);
        } else {
          throw new Error("Authorization failed");
        }
      } else {
        throw new Error(`Request failed with status ${response.status}`);
      }
    } else {
      return response;
    }
  } catch (error) {
    console.error(error);
    throw error;
  }
}

function authorizeDialog(): Promise<boolean> {
  console.log("open dialog");

  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      "https://outlook.integrations.zinfra.io/authorize",
      { height: 70, width: 40 },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("dialog result failed: " + asyncResult.error.message);
          resolve(false);
        } else {
          const dialog = asyncResult.value;
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            (messageEvent: Office.DialogParentMessageReceivedEventArgs) => {
              console.log("DialogMessageReceived");
              const authResult = JSON.parse(messageEvent.message) as AuthResult;
              console.log("Auth result:", authResult);

              if (authResult.success) {
                localStorage.setItem("isLoggedIn", "true");
                resolve(true);
              } else {
                resolve(false);
              }

              dialog.close();
            }
          );
        }
      }
    );
  });
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
