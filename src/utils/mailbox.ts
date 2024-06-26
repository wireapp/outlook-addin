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

export function getBody(item): Promise<string> {
  return new Promise((resolve, reject) => {
    const { body } = item;

    body.getAsync(Office.CoercionType.Html, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to get HTML body.");
        reject(new Error("Failed to get HTML body."));
      } else {
        resolve(asyncResult.value);
      }
    });
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

export async function appendToBody(item, contentToAppend) {
  try {
    const currentBody = await getBody(item);
    setBody(item, currentBody + contentToAppend);
  } catch (error) {
    console.error("Failed to append to body:", error);
  }
}

export async function getLocation(item): Promise<string> {
  return new Promise((resolve, reject) => {
    const { location } = item;

    location.getAsync(function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${asyncResult.error.message}`);
        reject(new Error(`Failed to get location: ${asyncResult.error.message}`));
        return;
      }
      resolve(asyncResult.value);
    });
  });
}


export async function setLocation(item, meetlingLink, callback) {
  const { location } = item;

  location.setAsync(meetlingLink, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${asyncResult.error.message}`);
      return;
    }
    callback();
  });
}
