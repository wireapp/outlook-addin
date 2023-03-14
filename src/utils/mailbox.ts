/* global , Office, console */

// async function createMeetingLinkElement() {
//   return await createGroupConversation("Success-Outlook").then((r) => {
//     createGroupLink(r).then((r) => {
//       return `<a href="${r}">${r}</a>`;
//     });
//   });
// }

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
