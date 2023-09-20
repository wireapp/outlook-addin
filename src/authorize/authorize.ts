/* global document, window, sessionStorage */

import * as CryptoJS from "crypto-js";
import { config } from "../utils/config";

document.addEventListener("DOMContentLoaded", redirectToAuthorize, false);

async function redirectToAuthorize(): Promise<void> {
  const { clientId, addInBaseUrl, authorizeUrl } = config;
  const redirectUri = new URL("/callback.html", addInBaseUrl);
  const responseType = "code";
  const state = generateRandomState();
  const scope = "write:conversations write:conversations_code read:self read:feature_configs";
  const codeChallengeMethod = "S256";
  const codeVerifier = generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);

  sessionStorage.setItem("state", state);
  sessionStorage.setItem("code_verifier", codeVerifier);

  const url = new URL(authorizeUrl);
  url.searchParams.append("client_id", clientId);
  url.searchParams.append("redirect_uri", redirectUri.toString());
  url.searchParams.append("response_type", responseType);
  url.searchParams.append("state", state);
  url.searchParams.append("scope", scope);
  url.searchParams.append("code_challenge_method", codeChallengeMethod);
  url.searchParams.append("code_challenge", codeChallenge);

  window.location.href = url.href.replace("/auth?", "/auth/#/login?");
}

function generateRandomState(): string {
  return generateRandomHexString(16);
}

function generateCodeVerifier(): string {
  return generateRandomHexString(64);
}

async function generateCodeChallenge(codeVerifier: string): Promise<string> {
  const hash = CryptoJS.SHA256(codeVerifier);
  const base64Url = hash.toString(CryptoJS.enc.Base64url);

  return base64Url;
}

function generateRandomHexString(length: number): string {
  function dec2hex(dec: number): string {
    return dec.toString(16).padStart(2, "0");
  }

  const arr = new Uint8Array(Math.ceil(length / 2));
  window.crypto.getRandomValues(arr);
  return Array.from(arr, dec2hex).join("").slice(0, length);
}
