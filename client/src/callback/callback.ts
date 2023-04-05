/* global global, Office, self, window */
import { AuthResult } from '../types/types'

document.addEventListener('DOMContentLoaded', async function () {
  console.log('BEFORE handling callback');
  await handleCallback();
}, false);

const handleCallback = async (): Promise<void> => {
  const urlParams = new URLSearchParams(window.location.search);
  const code = urlParams.get('code');
  const receivedState = urlParams.get('state');
  const storedCodeVerifier = sessionStorage.getItem('code_verifier');
  const storedState = sessionStorage.getItem('state');

  console.log('handleCallback');
  console.log('code: ', code);
  console.log('receivedState: ', receivedState);
  console.log('storedCodeVerifier: ', storedCodeVerifier);
  console.log('storedState: ', storedState);
  if (code && receivedState && storedCodeVerifier) {
    if (receivedState !== storedState) {
      console.error('State validation failed');
      return;
    }

    try {
      const authResult = await exchangeCodeForTokens(code, storedCodeVerifier);

      console.log('authResult: ', authResult);
        
      Office.onReady(() => {
        Office.context.ui.messageParent(JSON.stringify(authResult));
      });
    } catch (error) {
      console.error('Error during token exchange:', error);
    }
  }
};

const exchangeCodeForTokens = async (code: string, codeVerifier: string): Promise<AuthResult> => {
  const clientId = '3af4a9c5-4ae3-42f9-a168-981bbca4c56f';
  const redirectUri = 'https://outlook.integrations.zinfra.io/client/callback.html';
  const tokenEndpoint = 'https://staging-nginz-https.zinfra.io/oauth/token';

  const body = new URLSearchParams();
  body.append('grant_type', 'authorization_code');
  body.append('client_id', clientId);
  body.append('code', code);
  body.append('redirect_uri', redirectUri);
  body.append('code_verifier', codeVerifier);

  const response = await fetch(tokenEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: body.toString(),
  });

  console.log('body.toString(): ', body.toString());

  console.log('response: ', response);

  if (response.ok) {
    const json = await response.json();
    const { access_token, refresh_token } = json;

    console.log('access_token: ', access_token);
    console.log('refresh_token: ', refresh_token);

    return { success: true, access_token, refresh_token};
  } else {
    throw new Error('Failed to exchange authorization code for tokens');
  }
};


