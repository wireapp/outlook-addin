# outlook-addin

Wire add-in for Microsoft Outlook

## App Config
```
window.config = {
  addInBaseUrl: "https://${ADDIN_HOST}",
  apiBaseUrl: "https://${API_HOST}",
  authorizeUrl: "https://${AUTHORIZE_HOST}/auth",
  clientId: "${CLIENT_ID}",
};
```

## Nginx config template
```
server {
    listen 80;
    server_name ${ADDIN_HOST};
    root /usr/share/nginx/html;
    index taskpane.html taskpane.htm;
}
```

## Entrypoint.sh
```
envsubst < /usr/share/nginx/html/manifest.xml.template > /usr/share/nginx/html/manifest.xml
envsubst < /usr/share/nginx/html/config.js.template > /usr/share/nginx/html/config.js
envsubst < /etc/nginx/conf.d/default.conf.template > /etc/nginx/conf.d/default.conf

nginx -g "daemon off;"
```

## Local Storage
- isLoggedIn
- refresh_token
- access_token

## Feature flag
 - `outlookCalIntegration` - Must be enabled in order to be able to create a group and the link.

## Authorize
- URL: [config.authorizeUrl]
- Callback: [config.addInBaseUrl]/callback.html
- Scope: write:conversations write:conversations_code read:self read:feature_configs
- State: random 16 hex chars
- Verifier: random 64 hex chars

`State` and `Verifier` saved to Session Storage under: `state` and `code_verifier` respectively

## OAuth Callback
- When called verifies the `state` parameter and exchanges `code` for the tokens
- `access_token` and `refresh_token` then stored to Local Storage

## Refresh token
- Upon 401 Add-in will go to: POST [config.apiBaseUrl]/auth/refresh and body = LocalStorage.refresh_token

## Business Logic
- 