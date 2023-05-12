/* global Office, console, fetch */

import { AuthResult } from "../types/AuthResult";
import { getAccessToken, getRefreshToken, setTokens, removeTokens } from "../utils/tokenStore";
import jwt_decode from "jwt-decode";
import { DecodedToken } from "../types/DecodedToken";
import { config } from "../utils/config";
import { showNotification, removeNotification } from "../utils/notifications";

export async function fetchWithAuthorizeDialog(url: string | URL, options: RequestInit): Promise<Response> {
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

function authorizeDialog(): Promise<boolean> {
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
              const authResult = JSON.parse(messageEvent.message) as AuthResult;

              if (authResult.success) {
                setTokens(authResult.access_token, authResult.refresh_token);
                resolve(true);
              } else {
                removeTokens();
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

async function refreshTokenExchange(): Promise<boolean> {
  const refreshToken = getRefreshToken();
  if (!refreshToken) {
    return false;
  }

  const response = await fetch(new URL("/auth/refresh", config.apiBaseUrl).toString(), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ refresh_token: refreshToken }),
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

export function isTokenValid(token: string): boolean {
  if (token) {
    const decodedToken = jwt_decode<DecodedToken>(token);
    const currentDate = new Date();
    const currentTime = currentDate.getTime();
    return decodedToken.exp * 1000 > currentTime;
  }

  return false;
}

function isLoggedIn(): boolean {
  return isTokenValid(getAccessToken());
}
