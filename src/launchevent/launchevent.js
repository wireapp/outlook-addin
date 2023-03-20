// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console, fetch */

// sadly, no imports for event-based activation so everything has to be here

// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

const defaultSubjectValue = "New Appointment";
let mailboxItem;

// API

// TODO: should not be hardcoded
const teamIdStaging = "34e0d2d9-db1b-4029-8e01-471a11374dd5";

const apiUrl = "https://staging-nginz-https.zinfra.io/v2";
const token = localStorage.getItem('token');

async function createGroupConversation(name) {
  const payload = {
    access: ["invite", "code"],
    access_role_v2: ["guest", "non_team_member", "team_member", "service"],
    conversation_role: "wire_member",
    name: name,
    protocol: "proteus",
    qualified_users: [],
    receipt_mode: 1,
    team: {
      managed: false,
      teamid: teamIdStaging,
    },
    users: [],
  };

  // TODO: any/model
  const response = await fetch(apiUrl + "/conversations", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify(payload),
  }).then((r) => r.json());

  return response.id;
}

function onAppointmentSendHandler(event) {
  Office.context.mailbox.item.body.getAsync("text", { asyncContext: event }, doSomething);
}

function doSomething(asyncResult) {
  const event = asyncResult.asyncContext;
  getSubject(mailboxItem, (subject) => {
    createGroupConversation(subject ?? defaultSubjectValue).then((r) => {
      createGroupLink(r).then((r) => {
        const groupLink = `<a href="${r}">${r}</a>`;
        appendToBody(mailboxItem, groupLink, event);
      });
    });
  });

  // maybe can be done better ?
  // createMeetingLinkElement().then((meetingLink) => {
  //   appendToBody(mailboxItem, meetingLink);
  // });
}

// Register the functions.

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
}

// UTILS

async function createGroupLink(conversationId) {
  // TODO: any/model
  const response = await fetch(apiUrl + `/conversations/${conversationId}/code`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    },
  }).then((r) => r.json());

  return response.data.uri;
}

/** Returns value of current mailbox item subject, must pass a callback function to receive value */
async function getSubject(item, callback) {
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
async function getBody(item, callback) {
  const { body } = item;

  await body.getAsync(Office.CoercionType.Html, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to get HTML body.");
    } else {
      callback(asyncResult.value);
    }
  });
}

function setBody(item, newBody, event) {
  const { body } = item;
  const type = { coercionType: Office.CoercionType.Html };

  body.setAsync(newBody, type, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set HTML body.", asyncResult.error.message);
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

function appendToBody(item, contentToAppend, event) {
  getBody(item, (currentBody) => {
    setBody(item, currentBody + contentToAppend, event);
  });
}
