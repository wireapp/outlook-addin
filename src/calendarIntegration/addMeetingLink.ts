/* global Office, console */

import { appendToBody, getBody, getLocation, getMailboxItemSubject, getOrganizer, setLocation } from "../utils/mailbox";
import { createMeetingSummary } from "./createMeetingSummary";
import { setCustomPropertyAsync, getCustomPropertyAsync } from "../utils/customProperties";
import { showNotification, removeNotification } from "../utils/notifications";
import { isOutlookCalIntegrationEnabled } from "./isOutlookCalIntegrationEnabled";
import { createEvent } from "./createEvent";
import { mailboxItem } from "../commands/commands";
import { EventResult } from "../types/EventResult";
import { getUserDetails } from "../utils/userDetailsStore";

const defaultSubjectValue = "New Appointment";
let createdMeeting: EventResult;

/**
 * Checks if the feature is enabled by calling isOutlookCalIntegrationEnabled.
 *
 * @return {Promise<boolean>} Whether the feature is enabled or not.
 */
async function isFeatureEnabled(): Promise<boolean> {
  const isEnabled = await isOutlookCalIntegrationEnabled();

  if (!isEnabled) {
    console.log("Outlook calendar integration is disabled for this team. Contact your Wire system administrator.");
    removeNotification("wire-for-outlook-disabled");
    showNotification(
      "wire-for-outlook-disabled",
      "Wire for Outlook is disabled for your team. Please contact your Wire system administrator.",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    );
    return false;
  }
  return true;
}

/**
 * Fetches custom properties from the mailbox item and sets the createdMeeting object if both wireId and wireLink are present.
 *
 * @return {Promise<void>} A promise that resolves when the custom properties are fetched and the createdMeeting object is set.
 */
async function fetchCustomProperties(): Promise<void> {
  const wireId = await getCustomPropertyAsync(mailboxItem, "wireId");
  const wireLink = await getCustomPropertyAsync(mailboxItem, "wireLink");

  if (wireId && wireLink) {
    createdMeeting = { id: wireId, link: wireLink } as EventResult;
  }
}

/**
 * Creates a new meeting by calling the createEvent function with a subject obtained from the mailboxItem.
 * If the eventResult is not null, it sets the createdMeeting object to eventResult, updates the meeting details,
 * and sets the custom properties wireId and wireLink on the mailboxItem.
 *
 * @return {Promise<void>} A promise that resolves when the new meeting is created and the custom properties are set.
 */
async function createNewMeeting(): Promise<void> {
  removeNotification("adding-wire-meeting");
  showNotification(
    "adding-wire-meeting",
    "Adding Wire meeting...",
    Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator
  );

  const subject = await getMailboxItemSubject(mailboxItem);
  const eventResult = await createEvent(subject || defaultSubjectValue);

  if (eventResult) {
    createdMeeting = eventResult;
    await updateMeetingDetails(eventResult);
    await setCustomPropertyAsync(mailboxItem, "wireId", eventResult.id);
    await setCustomPropertyAsync(mailboxItem, "wireLink", eventResult.link);
  }

  removeNotification("adding-wire-meeting");
}

/**
 * Updates the meeting details by setting the location and appending the meeting summary to the body of the mailbox item.
 *
 * @param {EventResult} eventResult - The event result containing the link for the meeting.
 * @return {Promise<void>} A promise that resolves when the meeting details are updated.
 */
async function updateMeetingDetails(eventResult: EventResult): Promise<void> {

  await setLocation(mailboxItem, eventResult.link, () => {});
  // const organizer = await getOrganizer(mailboxItem);
  const user = getUserDetails();
  const meetingSummary = createMeetingSummary(eventResult.link, user.name);
  await appendToBody(mailboxItem, user.name);
  // getOrganizer(mailboxItem, async (organizer) => {
  //   const meetingSummary = createMeetingSummary(eventResult.link, organizer);
  //   await appendToBody(mailboxItem, meetingSummary);
  // });


}

/**
 * Handles an existing meeting by updating meeting details and setting custom properties.
 *
 * @return {Promise<void>} A promise that resolves when the existing meeting is handled.
 */
async function handleExistingMeeting(): Promise<void> {
  if (!createdMeeting) {
    throw new Error("createdMeeting is undefined");
  }

  const currentBody = await getBody(mailboxItem);
  const currentLocation = await getLocation(mailboxItem);
  const normalizedCurrentBody = currentBody.replace(/&amp;/g, "&");
  const normalizedMeetingLink = createdMeeting.link?.replace(/&amp;/g, "&");

  // getOrganizer(mailboxItem, async (organizer) => {
  //   if (!currentLocation) {
  //     await setLocation(mailboxItem, createdMeeting.link, () => {});
  //   }
  //   const meetingSummary = createMeetingSummary(createdMeeting.link, organizer);
  //   if (!normalizedCurrentBody.includes(normalizedMeetingLink)) {
  //     await appendToBody(mailboxItem, meetingSummary);
  //   }
  // });

  await setCustomPropertyAsync(mailboxItem, "wireId", createdMeeting.id);
  await setCustomPropertyAsync(mailboxItem, "wireLink", createdMeeting.link);
}

/**
 * Adds a meeting link to the Outlook calendar.
 *
 * @param {Office.AddinCommands.Event} event - The event object.
 * @return {Promise<void>} A promise that resolves when the meeting link is added.
 */
async function addMeetingLink(event: Office.AddinCommands.Event): Promise<void> {
  try {

    const isEnabled = await isFeatureEnabled();
    if (!isEnabled) return;


    await fetchCustomProperties();
    if (!createdMeeting) {
      await createNewMeeting();
    } else {
      await handleExistingMeeting();
    }
  } catch (error) {
    console.error("Error during adding Wire meeting link", error);
    handleAddMeetingLinkError(error);
  } finally {
    event.completed();
  }
}

/**
 * Handles errors that occur when adding a meeting link.
 *
 * @param {Error} error - The error that occurred.
 * @return {void} This function does not return anything.
 */
function handleAddMeetingLinkError(error: Error): void {
  removeNotification("adding-wire-meeting");
  removeNotification("adding-wire-meeting-error");

  if (error.message.includes("authorization failed")) {
    showNotification(
      "adding-wire-meeting-error",
      "Authorization failed.",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    );
  } else {
    showNotification(
      "adding-wire-meeting-error",
      "There was an error while adding the Wire meeting.",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    );
  }
}

export { addMeetingLink };
