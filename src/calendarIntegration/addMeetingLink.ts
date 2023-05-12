/* global Office, console */

import { appendToBody, getMailboxItemSubject, getOrganizer, setLocation } from "../utils/mailbox";
import { createMeetingSummary as createMeetingSummary } from "./createMeetingSummary";
import { setCustomPropertyAsync, getCustomPropertyAsync } from "../utils/customProperties";
import { showNotification, removeNotification } from "../utils/notifications";
import { isOutlookCalIntegrationEnabled } from "./isOutlookCalIntegrationEnabled";
import { createEvent } from "./createEvent";
import { mailboxItem } from "../commands/commands";

const defaultSubjectValue = "New Appointment";

export async function addMeetingLink(event: Office.AddinCommands.Event) {
  try {
    const isFeatureEnabled = await isOutlookCalIntegrationEnabled();

    if (!isFeatureEnabled) {
      console.log(
        "There is no Outlook calendar integration enabled for this team. Please contact your Wire system administrator."
      );

      removeNotification("wire-for-outlook-disabled");
      showNotification(
        "wire-for-outlook-disabled",
        "Wire for Outlook is disabled for your team. Please contact your Wire system administrator.",
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      );
    } else {
      console.log("Outlook calendar integration feature is enabled for this team.");

      const wireId = await getCustomPropertyAsync(mailboxItem, "wireId");
      if (!wireId) {
        console.log("There is no Wire meeting for this Outlook meeting, starting process of creating it...");
        removeNotification("adding-wire-meeting");
        showNotification(
          "adding-wire-meeting",
          "Adding Wire meeting...",
          Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator
        );
        const subject = await getMailboxItemSubject(mailboxItem);
        const eventResult = await createEvent(subject || defaultSubjectValue);
        if (eventResult) {
          getOrganizer(mailboxItem, function (organizer) {
            setLocation(mailboxItem, eventResult.link, () => {});
            const meetingSummary = createMeetingSummary(eventResult.link, organizer);
            appendToBody(mailboxItem, meetingSummary);
          });
          await setCustomPropertyAsync(mailboxItem, "wireId", eventResult.id);
          await setCustomPropertyAsync(mailboxItem, "wireLink", eventResult.link);
        }
        removeNotification("adding-wire-meeting");
      } else {
        console.log("Wire meeting is already created for this Outlook meeting");
        removeNotification("wire-meeting-exists");
        showNotification(
          "wire-meeting-exists",
          "Wire meeting is already created for this Outlook meeting",
          Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        );
      }
    }
  } catch (error) {
    console.error("Error during adding Wire meeting link", error);

    removeNotification("adding-wire-meeting");
    removeNotification("adding-wire-meeting-error");

    showNotification(
      "adding-wire-meeting-error",
      "There was error while adding wire meeting",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    );
  }

  event.completed();
}
