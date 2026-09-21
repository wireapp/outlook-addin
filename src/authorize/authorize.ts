/* global document, window, sessionStorage */

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

async function digestMessage(message: string): Promise<ArrayBuffer> {
  const encoder = new TextEncoder();
  const data = encoder.encode(message);
  const hash = await window.crypto.subtle.digest("SHA-256", data);
  return hash;
}

function arrayBufferToBase64URL(buffer: ArrayBuffer): string {
  const uint8Array = new Uint8Array(buffer);
  const binaryString = Array.from(uint8Array)
    .map((byte) => String.fromCodePoint(byte))
    .join("");
  const base64EncodedString = window.btoa(binaryString);
  const base64URLEncodedString = base64EncodedString.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  return base64URLEncodedString;
}

async function generateCodeChallenge(codeVerifier: string): Promise<string> {
  return arrayBufferToBase64URL(await digestMessage(codeVerifier));
}

function generateRandomHexString(length: number): string {
  function dec2hex(dec: number): string {
    return dec.toString(16).padStart(2, "0");
  }

  const arr = new Uint8Array(Math.ceil(length / 2));
  window.crypto.getRandomValues(arr);
  return Array.from(arr, dec2hex).join("").slice(0, length);
}
