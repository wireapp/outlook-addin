import * as CryptoJS from 'crypto-js';

document.addEventListener('DOMContentLoaded', async function () {
  await redirectToAuthorize();
}, false);

const redirectToAuthorize = async () => {
  const clientId = '3af4a9c5-4ae3-42f9-a168-981bbca4c56f';
  const redirectUri = 'https://outlook.integrations.zinfra.io/client/callback.html';
  const responseType = 'code';
  const state = await generateRandomState();
  const scope = 'write:conversations write:conversations_code read:self read:feature_configs';

  const codeChallengeMethod = 'S256';
  const codeVerifier = await generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);

  const url = new URL('https://wire-webapp-edge.zinfra.io/auth');
  url.searchParams.append('client_id', clientId);
  url.searchParams.append('redirect_uri', redirectUri);
  url.searchParams.append('response_type', responseType);
  url.searchParams.append('state', state);
  url.searchParams.append('scope', scope);
  url.searchParams.append('code_challenge_method', codeChallengeMethod);
  url.searchParams.append('code_challenge', codeChallenge);

  sessionStorage.setItem('state', state);
  sessionStorage.setItem('code_verifier', codeVerifier);

  window.location.href = url.href;
};

const generateRandomState = async (): Promise<string> => {
  const array = new Uint32Array(8);
  window.crypto.getRandomValues(array);

  const state = Array.from(array, (value) => value.toString(36)).join('');
  return state;
};

const generateCodeVerifier = async (): Promise<string> => {
  const array = new Uint32Array(32);
  window.crypto.getRandomValues(array);

  const verifier = Array.from(array, (value) => value.toString(36)).join('');
  return verifier;
};

const generateCodeChallenge = async (codeVerifier: string): Promise<string> => {
  const hash = CryptoJS.SHA256(codeVerifier);
  const base64Url = hash
    .toString(CryptoJS.enc.Base64)
    .replace('+', '-')
    .replace('/', '_')
    .replace(/=+$/, '');

  return base64Url;
};
