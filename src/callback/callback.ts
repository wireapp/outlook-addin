/* global global, Office, self, window */
import config from "../config";
import { AuthResult } from "../types/types";

document.addEventListener(
  "DOMContentLoaded",
  async function () {
    await handleCallback();
  },
  false
);

const handleCallback = async (): Promise<void> => {
  const urlParams = new URLSearchParams(window.location.search);
  const code = urlParams.get("code");
  const receivedState = urlParams.get("state");
  const storedCodeVerifier = sessionStorage.getItem("code_verifier");
  const storedState = sessionStorage.getItem("state");

  if (code && receivedState && storedCodeVerifier) {
    if (receivedState !== storedState) {
      console.error("State validation failed");
      return;
    }

    try {
      const authResult = await exchangeCodeForTokens(code, storedCodeVerifier);

      Office.onReady(() => {
        Office.context.ui.messageParent(JSON.stringify(authResult));
      });
    } catch (error) {
      console.error("Error during token exchange:", error);
    }
  }
};

const exchangeCodeForTokens = async (code: string, codeVerifier: string): Promise<AuthResult> => {
  const clientId = config.clientId;
  const redirectUri = new URL("/callback.html", config.addInBaseUrl);
  const tokenEndpoint = new URL("/oauth/token", config.apiBaseUrl);

  const body = new URLSearchParams();
  body.append("grant_type", "authorization_code");
  body.append("client_id", clientId);
  body.append("code", code);
  body.append("redirect_uri", redirectUri.toString());
  body.append("code_verifier", codeVerifier);

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
};
