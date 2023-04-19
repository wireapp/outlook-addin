import * as CryptoJS from "crypto-js";

const config = window.config;

document.addEventListener(
  "DOMContentLoaded",
  async function () {
    await redirectToAuthorize();
  },
  false
);

const redirectToAuthorize = async () => {
  const clientId = config.clientId;
  const redirectUri = new URL("/callback.html", config.addInBaseUrl);
  const responseType = "code";
  const state = await generateRandomState();
  const scope = "write:conversations write:conversations_code read:self read:feature_configs";

  const codeChallengeMethod = "S256";
  const codeVerifier = await generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);
  
  sessionStorage.setItem("state", state);
  sessionStorage.setItem("code_verifier", codeVerifier);

  const url = new URL(config.authorizeUrl);
  url.searchParams.append("client_id", clientId);
  url.searchParams.append("redirect_uri", redirectUri.toString());
  url.searchParams.append("response_type", responseType);
  url.searchParams.append("state", state);
  url.searchParams.append("scope", scope);
  url.searchParams.append("code_challenge_method", codeChallengeMethod);
  url.searchParams.append("code_challenge", codeChallenge);
  url.hash = "authorize";

  window.location.href = url.href;
};

const generateRandomState = async (): Promise<string> => {
  return generateRandomHexString(16);
};

const generateCodeVerifier = async (): Promise<string> => {
  return generateRandomHexString(64);
};

const generateCodeChallenge = async (codeVerifier: string): Promise<string> => {
  const hash = CryptoJS.SHA256(codeVerifier);
  const base64Url = hash.toString(CryptoJS.enc.Base64url);

  return base64Url;
};

const generateRandomHexString = (length) => {
  const dec2hex = (dec) => {
    return dec.toString(16).padStart(2, "0");
  };

  const arr = new Uint8Array(Math.ceil(length / 2));
  window.crypto.getRandomValues(arr);
  return Array.from(arr, dec2hex).join("").slice(0, length);
};
