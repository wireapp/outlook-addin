# outlook-addin

Wire add-in for Microsoft Outlook

## App Config
```
window.config = {
  addInBaseUrl: "${BASE_URL}",
  apiBaseUrl: "${WIRE_API_BASE_URL}",
  authorizeUrl: "${WIRE_AUTHORIZATION_ENDPOINT}",
  clientId: "${CLIENT_ID}",
};
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

## How to create new Service with the BE
```agsl
curl -s -X POST localhost:8080/i/oauth/clients \
    -H "Content-Type: application/json" \
    -d '{
      "application_name":"Wire Microsoft Outlook Calendar Add-in",
      "redirect_url":"https://outlook.wire.com/callback.html" 
    }'
```

## How to install the Add-in in MS Outlook
- Open an email and got to 3 dots and select Get Add-ins
![Step 1](images/step_1.png)
- Go to My Add-ins, Custom Add-ins, Add a Custom Add-in
![Step 2](images/step_2.png)
- Pick up a URL and add: https://outlook.integrations.wire.com/manifest.xml
![Step 3](images/step_3.png)
Wire button will appear in the toolbar when new event is being created