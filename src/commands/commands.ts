/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, Office, console */

import { AuthResult, EventResult } from "../types/types";
import {
  appendToBody,
  getSubject,
  createMeetingSummary as createMeetingSummary,
  getOrganizer,
  setLocation,
} from "../utils/mailbox";
import { setCustomPropertyAsync, getCustomPropertyAsync } from "../utils/customproperties";
import { showNotification, removeNotification } from "../utils/notifications";
import jwt_decode from "jwt-decode";

const config = window.config;

// Office is ready. Init
Office.onReady(function () {
  mailboxItem = Office.context.mailbox.item;
});

const defaultSubjectValue = "New Appointment";
let mailboxItem;

async function addMeetingLink(event: Office.AddinCommands.Event) {
  try {
    const isFeatureEnabled = isOutlookCalIntegrationEnabled();
    console.log("isOutlookCalIntegrationEnabled: ", isFeatureEnabled);
    if (!isFeatureEnabled) {
      console.log(
        "There is no Outlook calendar integration enabled for this team. Please contact your Wire system administrator."
      );
      showNotification(
        "wire-meeting-exists",
        "Wire meeting is already created for this Outlook meeting",
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      );
    } else {
      const wireId = await getCustomPropertyAsync(mailboxItem, "wireId");
      if (!wireId) {
        console.log("There is no Wire meeting for this Outlook meeting, starting process of creating it...");
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
        showNotification(
          "wire-meeting-exists",
          "Wire meeting is already created for this Outlook meeting",
          Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        );
      }
    }
  } catch (error) {
    console.error(error);
    removeNotification("adding-wire-meeting");
    showNotification(
      "adding-wire-meeting-error",
      "There was error while adding wire meeting",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
    );
  }
  event.completed();
}

async function createEvent(name: string): Promise<EventResult> {
  try {
    const teamId = await getTeamId();

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
        teamid: teamId,
      },
      users: [],
    };

    // TODO: any/model
    const response: any = await fetchWithAuthorizeDialog(new URL("/v2/conversations", config.apiBaseUrl), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (response.ok) {
      const conversationId = (await response.json()).id;

      const responseLink: any = await fetchWithAuthorizeDialog(
        new URL(`/conversations/${conversationId}/code`, config.apiBaseUrl),
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
        }
      ).then((r) => r.json());

      return { id: conversationId, link: responseLink.data.uri };
    } else {
      throw new Error(`Request failed with status ${response.status}`);
    }
  } catch (error) {
    console.error(error);
    throw error;
  }
}

async function fetchWithAuthorizeDialog(url: string | URL, options: RequestInit): Promise<Response> {
  try {
    let isLoggedIn = !!localStorage.getItem("refresh_token");

    if (!isLoggedIn) {
      isLoggedIn = await authorizeDialog();
    }

    if (isLoggedIn) {
      const token = localStorage.getItem("access_token");
      options.headers = {
        ...options.headers,
        Authorization: `Bearer ${token}`,
      };

      const response = await fetch(url, options);

      if (response.status === 401) {
        isLoggedIn = await refreshTokenExchange();

        if (!isLoggedIn) {
          localStorage.removeItem("access_token");
          localStorage.removeItem("refresh_token");

          isLoggedIn = await authorizeDialog();
        }

        if (isLoggedIn) {
          const token = localStorage.getItem("access_token");
          options.headers = {
            ...options.headers,
            Authorization: `Bearer ${token}`,
          };
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

const refreshTokenExchange = async (): Promise<boolean> => {
  const refreshToken = localStorage.getItem("refresh_token");
  const tokenEndpoint = new URL("/oauth/token", config.apiBaseUrl);
  const clientId = config.clientId;

  if (!refreshToken) {
    return false;
  }

  const body = new URLSearchParams();
  body.append("grant_type", "refresh_token");
  body.append("client_id", clientId);
  body.append("refresh_token", refreshToken);

  try {
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    });

    if (response.ok) {
      const json = await response.json();
      const { access_token, refresh_token } = json;

      localStorage.setItem("access_token", access_token);
      localStorage.setItem("refresh_token", refresh_token);

      return true;
    } else {
      return false;
    }
  } catch (error) {
    console.error("Error during refresh token exchange:", error);
    return false;
  }
};

function authorizeDialog(): Promise<boolean> {
  console.log("open dialog");

  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      new URL("/authorize.html", config.addInBaseUrl).toString(),
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
                localStorage.setItem("access_token", authResult.access_token);
                localStorage.setItem("refresh_token", authResult.refresh_token);
                resolve(true);
              } else {
                localStorage.removeItem("isLoggedIn");
                localStorage.removeItem("access_token");
                localStorage.removeItem("refresh_token");
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

export async function isTokenStillValid(token: string) {
  if (token) {
    const decodedToken = jwt_decode(token) as any;
    const currentDate = new Date();
    const currentTime = currentDate.getTime();
    return decodedToken.exp * 1000 > currentTime;
  }

  return false;
}

async function getTeamId() {
  const response: any = await fetchWithAuthorizeDialog(new URL("/self", config.apiBaseUrl), {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  }).then((r) => r.json());

  return response.team;
}

async function isOutlookCalIntegrationEnabled() {
  const response: any = await fetchWithAuthorizeDialog(new URL("/feature-configs", config.apiBaseUrl), {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  }).then((r) => r.json());

  console.log("/feature-configs", response);
  console.log("outlookCalIntegration", response.outlookCalIntegration);

  return true;
}

// Register the functions.
Office.actions.associate("addMeetingLink", addMeetingLink);
