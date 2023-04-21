/* global Office, console */

export async function getMailboxItemSubject(item): Promise<string> {
  return new Promise((resolve) => {
    getSubject(item, (result) => {
      resolve(result);
    });
  });
}

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
