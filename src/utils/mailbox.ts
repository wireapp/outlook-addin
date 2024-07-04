export async function getMailboxItemSubject(item) : Promise<string> {
  return new Promise((resolve,reject) => {
  const { subject } = item;

  subject.getAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      console.error("Failed to get item subject");
      reject(new Error("Failed to get item subject"));
    } else {
      resolve(asyncResult.value);
    }
  });
});
}

export async function getOrganizer(item): Promise<string> {
  return new Promise((resolve, reject) => {
    const { organizer } = item;

    organizer.getAsync(function(asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value.displayName);
      } else {
        console.error(asyncResult.error);
        reject(new Error("Failed to get organizer"));
      }
    });
  });
}

export async function getOrganizerOnMobile(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(
      "html",
      { asyncContext: Office.context.mailbox.userProfile.displayName },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.asyncContext);
        } else {
          reject(new Error("Failed to get body."));
        }
      }
    );
  });
}

export async function setSubject(item, newSubject:string){
    
  const { subject } = item;

  subject.setAsync(newSubject, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${asyncResult.error.message}`);
      return;
    }
  });
}

export async function getMeetingTime(item): Promise<string> {
  return new Promise((resolve, reject) => {
    const { start } = item;

    start.getAsync(function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${asyncResult.error.message}`);
        reject(new Error(`Failed to get start time: ${asyncResult.error.message}`));
        return;
      }
        const locale = Office.context.displayLanguage;
        const formattedDate = new Date(asyncResult.value).toLocaleString(locale, {
          year: 'numeric',
          month: 'numeric',
          day: 'numeric',
        });
        resolve(formattedDate);
    });
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

export async function getLocation(item) {
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


export async function setLocation(item, meetlingLink) {
  const { location } = item;

  location.setAsync(meetlingLink, function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${asyncResult.error.message}`);
      return;
    }
  });
}
