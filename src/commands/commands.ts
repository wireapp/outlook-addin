// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console */

import { AuthResult, EventResult } from "../api/types";
import { appendToBody, getSubject, createMeetingSummary as createMeetingSummary, getOrganizer, setLocation } from "../utils/mailbox";
import { setCustomPropertyAsync, getCustomPropertyAsync } from "../utils/customproperties";
import { showNotification } from "../utils/notifications";

// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

const defaultSubjectValue = "New Appointment";
let mailboxItem;

async function addMeetingLink() {
  //try {
    const wireId = await getCustomPropertyAsync(mailboxItem, 'wireId');
    if(!wireId) {
      console.log('There is no Wire meeting for this Outlook meeting, starting process of creating it...');
      //showNotification('adding-wire-meeting', 'Adding Wire meeting...', Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, 'icon1');
      const subject = await getMailboxItemSubject(mailboxItem);
      const eventResult = await createEvent(subject || defaultSubjectValue);
      if (eventResult) {
        getOrganizer(mailboxItem, function (organizer) {
          setLocation(mailboxItem, eventResult.link, () => {});
          const meetingSummary = createMeetingSummary(eventResult.link, organizer);
          appendToBody(mailboxItem, meetingSummary);
        });
        await setCustomPropertyAsync(mailboxItem, 'wireId', eventResult.id);
        await setCustomPropertyAsync(mailboxItem, 'wireLink', eventResult.link);
      }
    } else {
      console.log('Wire meeting is already created for this Outlook meeting');
      //showNotification('wire-meeting-exists', 'Wire meeting is already created for this Outlook meeting', Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, 'icon1');
    }
  //} catch (error) {
  //  console.error(error);
  //}
}

async function createEvent(name: string): Promise<EventResult> {
  try {
    const response = await fetchWithAuthorizeDialog("/event", {
      method: "POST",
      credentials: "include",
      body: JSON.stringify({ name }),
      headers: {
        "Content-Type": "application/json",
      },
    });

    if (response.ok) {
      const result = (await response.json()) as EventResult;
      return result;
    } else {
      throw new Error(`Request failed with status ${response.status}`);
    }
  } catch (error) {
    console.error(error);
    throw error;
  }
}

async function fetchWithAuthorizeDialog(url: string, options: RequestInit): Promise<Response> {
  try {
    let isLoggedIn = JSON.parse(localStorage.getItem("isLoggedIn"));

    if (!isLoggedIn) {
      isLoggedIn = await authorizeDialog();
    }

    if (isLoggedIn) {
      const response = await fetch(url, options);

      if (response.status === 401) {
        localStorage.removeItem("isLoggedIn");
        isLoggedIn = await authorizeDialog();
        if (isLoggedIn) {
          return await fetch(url, options);
        } else {
          throw new Error("Authorization failed");
        }
      } else if (!response.ok) {
        throw new Error(`Request failed with status ${response.status}`);
      }

      return response;
    } else {
      throw new Error("Authorization failed");
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

async function getMailboxItemSubject(mailboxItem: any): Promise<string> {
  return new Promise((resolve) => {
    getSubject(mailboxItem, (result) => {
      resolve(result);
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
