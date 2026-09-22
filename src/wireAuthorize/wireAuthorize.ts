/* global Office, console, fetch */

import type { AuthResult } from "../types/AuthResult";
import { getAccessToken, getRefreshToken, setTokens, removeTokens } from "../utils/tokenStore";
import jwt_decode from "jwt-decode";
import type { DecodedToken } from "../types/DecodedToken";
import { config } from "../utils/config";
import { showNotification, removeNotification } from "../utils/notifications";
import { setUserDetails, removeUserDetails } from "../utils/userDetailsStore";
import { getSelf } from "../calendarIntegration/getSelf";

export async function fetchWithAuthorizeDialog(
  url: string | URL,
  options: RequestInit
): Promise<Response> {
  try {
    let isAuthenticated = isLoggedIn();

    const refreshToken = getRefreshToken();

    if (!isAuthenticated && refreshToken) {
      isAuthenticated = await refreshTokenExchange();
    }

    if (!isAuthenticated) {
      isAuthenticated = await authorizeDialog();
    }

    if (isAuthenticated) {
      const token = getAccessToken();
      options.headers = {
        ...options.headers,
        Authorization: `Bearer ${token}`,
      };

      const response = await fetch(url, options);

      if (response.status === 401) {
        isAuthenticated = await refreshTokenExchange();

        if (!isAuthenticated) {
          removeTokens();

          isAuthenticated = await authorizeDialog();
        }

        if (isAuthenticated) {
          const token = getAccessToken();
          options.headers = {
            ...options.headers,
            Authorization: `Bearer ${token}`,
          };
          return await fetch(url, options);
        } else {
          removeNotification("auth-failed");
          showNotification(
            "auth-failed",
            "Authorization failed.",
            Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
          );

          throw new Error("Authorization failed");
        }
      } else if (!response.ok) {
        removeNotification("auth-failed");
        showNotification(
          "auth-failed",
          "Authorization failed.",
          Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        );

        throw new Error(`Request failed with status ${response.status}`);
      }

      return response;
    } else {
      removeNotification("auth-failed");
      showNotification(
        "auth-failed",
        "Authorization failed.",
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      );

      throw new Error("Authorization failed");
    }
  } catch (error) {
    console.error(error);
    throw error;
  }
}

export function authorizeDialog(): Promise<boolean> {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      new URL("/authorize.html", config.addInBaseUrl).toString(),
      { height: 70, width: 40 },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("dialog result failed: ", asyncResult.error.message);
          resolve(false);
        } else {
          const dialog = asyncResult.value;
          let completed = false;
          let processingMessage = false;

          const finish = (success: boolean, closeDialog = true) => {
            if (completed) {
              return;
            }
            completed = true;

            try {
              if (!success) {
                removeTokens();
                removeUserDetails();
              }
              if (closeDialog) {
                dialog.close();
              }
            } catch (error) {
              console.error("dialog cleanup failed: ", error);
            } finally {
              resolve(success);
            }
          };

          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            finish(false, false);
          });
          dialog.addEventHandler(
            Office.EventType.DialogMessageReceived,
            async (messageEvent: Office.DialogParentMessageReceivedEventArgs) => {
              if (completed || processingMessage) {
                return;
              }
              processingMessage = true;

              try {
                const authResult = JSON.parse(messageEvent.message) as AuthResult;
                if (
                  authResult?.success !== true ||
                  typeof authResult.access_token !== "string" ||
                  !authResult.access_token ||
                  (authResult.refresh_token != null && typeof authResult.refresh_token !== "string")
                ) {
                  finish(false);
                  return;
                }

                setTokens(authResult.access_token, authResult.refresh_token);
                const user = await getSelf();
                // The dialog may have been closed while the request was pending.
                if (completed) {
                  return;
                }
                setUserDetails(user);
                finish(true);
              } catch (error) {
                console.error("dialog authorization failed: ", error);
                finish(false);
              }
            }
          );
        }
      }
    );
  });
}

async function refreshTokenExchange(): Promise<boolean> {
  const refreshToken = getRefreshToken();
  if (!refreshToken) {
    return false;
  }

  let data = new URLSearchParams();
  data.append("grant_type", "refresh_token");
  data.append("refresh_token", refreshToken);
  data.append("client_id", config.clientId);

  const response = await fetch(new URL(`${config.apiVersion}/oauth/token`, config.apiBaseUrl), {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: data,
  });

  if (response.ok) {
    const authResult: AuthResult = await response.json();
    setTokens(authResult.access_token, authResult.refresh_token);
    return true;
  } else {
    removeTokens();
    return false;
  }
}

export async function revokeOauthToken(): Promise<boolean> {
  const refreshToken = getRefreshToken();
  if (!refreshToken) {
    return false;
  }

  const payload = {
    refresh_token: refreshToken,
    client_id: config.clientId,
  };

  const response = await fetch(new URL(`${config.apiVersion}/oauth/revoke`, config.apiBaseUrl), {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (response.ok) {
    return true;
  } else {
    return false;
  }
}

export function isTokenValid(token: string): boolean {
  // null-check
  if (!token) {
    console.error("isTokenValid: token was null", token);
    return false;
  }

  // decode token
  let decodedToken: DecodedToken;
  try {
    decodedToken = jwt_decode<DecodedToken>(token);
  } catch (err) {
    console.error("isTokenValid: error decoding token", err);
    return false;
  }

  // check token
  let result: boolean;
  try {
    result = decodedToken.exp * 1000 > new Date().getTime();
  } catch {
    console.error("isTokenValid: error checking token validity");
    return false;
  }

  // return if token is valid
  return result;
}

function isLoggedIn(): boolean {
  return isTokenValid(getAccessToken());
}
