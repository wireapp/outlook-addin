/* global Office, window, document, console, fetch, sessionStorage */

import { AuthResult } from "../types/AuthResult";
import { UrlParameters } from "./UrlParameters";
import { config } from "../utils/config";

document.addEventListener(
  "DOMContentLoaded",
  function () {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Outlook) {
        handleCallback();
      }
    });
  },
  false
);

async function handleCallback(): Promise<void> {
  const urlParams = getUrlParameters();
  const { code, receivedState, error } = urlParams;

  if (error) {
    console.error("Error in auth flow: ", error);
    const authResult = { success: false, error };
    sendMessageToParent(authResult);
    return;
  }

  const storedCodeVerifier = sessionStorage.getItem("code_verifier");
  const storedState = sessionStorage.getItem("state");

  if (code && receivedState && storedCodeVerifier) {
    if (receivedState !== storedState) {
      console.error("State validation failed");
      return;
    }

    try {
      const authResult = await exchangeCodeForTokens(code, storedCodeVerifier);
      sendMessageToParent(authResult);
    } catch (error) {
      console.error("Error during token exchange:", error);
    }
  }
}

function getUrlParameters(): UrlParameters {
  const urlParams = new URLSearchParams(window.location.search);
  return {
    code: urlParams.get("code"),
    receivedState: urlParams.get("state"),
    error: urlParams.get("error"),
  };
}

function sendMessageToParent(authResult: AuthResult): void {
  Office.onReady(() => {
    Office.context.ui.messageParent(JSON.stringify(authResult));
  });
}

async function exchangeCodeForTokens(code: string, codeVerifier: string): Promise<AuthResult> {
  const clientId = config.clientId;
  const redirectUri = new URL("/callback.html", config.addInBaseUrl);
  const tokenEndpoint = new URL(`${config.apiVersion}/oauth/token`, config.apiBaseUrl);

  const body = getRequestBody(code, clientId, redirectUri, codeVerifier);

  const response = await fetch(tokenEndpoint, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: body.toString(),
  });

  if (response.ok) {
    const json = await response.json();
    const { access_token, refresh_token } = json;
    return { success: true, access_token, refresh_token };
  } else {
    throw new Error("Failed to exchange authorization code for tokens");
  }
}

function getRequestBody(code: string, clientId: string, redirectUri: URL, codeVerifier: string): URLSearchParams {
  const body = new URLSearchParams();
  body.append("grant_type", "authorization_code");
  body.append("client_id", clientId);
  body.append("code", code);
  body.append("redirect_uri", redirectUri.toString());
  body.append("code_verifier", codeVerifier);

  return body;
}
