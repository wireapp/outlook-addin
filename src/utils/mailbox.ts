/* global , Office, console */

export function createMeetingLinkElement(groupInviteLink, organizer) {
  const wireDownloadLink = "https://wire.com/en/download/";
  const addinDownloadLink = undefined;
  const fullInvite = `<div>
    <p>${organizer} is inviting you to join this meeting in Wire.</p>
    <p>Join meeting in Wire <a href="${groupInviteLink}">${groupInviteLink}</a></p>
    <p><a href="${wireDownloadLink}">Download Wire</a></p>
    <p><a href="${addinDownloadLink}">Get Wire add-in for Outlook</a></p>
  </div>`;

  return fullInvite;
}

/** Return */
export async function getOrganizer(item, callback) {
  const { organizer } = item;

  await organizer.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to get organizer.");
    } else {
      callback(asyncResult.value.displayName);
    }
  });
}

/** Returns value of current mailbox item subject, must pass a callback function to receive value */
export async function getSubject(item, callback) {
  const { subject } = item;

  await subject.getAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.error("Failed to get item subject");
    } else {
      callback(asyncResult.value);
    }
  });
}

/** Returns value of current mailbox item body, must pass a callback function to receive value*/
export async function getBody(item, callback) {
  const { body } = item;

  await body.getAsync(Office.CoercionType.Html, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to get HTML body.");
    } else {
      callback(asyncResult.value);
    }
  });
}

export function setBody(item, newBody) {
  const { body } = item;
  const type = { coercionType: Office.CoercionType.Html };

  body.setAsync(newBody, type, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set HTML body.", asyncResult.error.message);
    } else {
      // do something else perhaps?
    }
  });
}

export function appendToBody(item, contentToAppend) {
  getBody(item, (currentBody) => {
    setBody(item, currentBody + contentToAppend);
  });
}

export async function setLocation(item, meetlingLink, callback) {
  const { location } = item;

  location.setAsync(meetlingLink, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${asyncResult.error.message}`);
      return;
    }
    console.log(`Successfully set location to ${location}`);
    callback();
  });
}
