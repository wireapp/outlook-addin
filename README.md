# Wire's Microsoft Outlook Calendar Add-in

Wire add-in for Microsoft Outlook. It creates a Wire conversation and adds its invitation link to an Outlook calendar event.

Outlook loads the add-in's HTML and JavaScript from a web server using the URLs in its manifest.
For local development, Webpack serves these files over HTTPS. In production, a deployed Nginx container serves the built files; Outlook does not run that container.

## Configuration
The program is configured through environment variables listed in the [.env.template](.env.template) file.  
Depending on the deployment mode, the values are substituted differently:
- development: from a local `.env` file at build time via a Webpack plugin;
- production – at container startup via a Docker entrypoint script using `envsubst` command.

The [manifest.xml](manifest.xml.template) file describes the Office Add-in (its name, permissions, and endpoints), 
while [config.js](./src/config.js.template) provides the app with runtime configuration such as API URLs and client IDs.

The actual values for the staging environment are provided in the [.env.staging](.env.staging) file.

## Configuration
The project needs a running web-server to serve the add-in's HTML and JavaScript files.
The server must be reachable from Outlook, so it must have a public URL.

Currently it is running in AWS, this project just deploys to Quay.io, while [cailleach]([https://ops.zinfra.io/](https://github.com/zinfra/cailleach/blob/ad459f23622d7ad003bc84c03e407dab27905ced/tf-modules/k8s-outlook-addin/README.md))
actually deploys to AWS the image.

### Feature flag
`outlookCalIntegration` – Must be enabled to be able to create a group and the link.

## Local development

These steps run the add-in on your machine while connecting to Wire's staging backend.
You need Node.js and npm, access to Outlook, and a Wire staging account whose team has `outlookCalIntegration` enabled.

1. Install the dependencies:

   ```shell
   npm ci
   ```

2. Create your local configuration by copying the staging file. If you already have a `.env`, update it instead of overwriting it:

   ```shell
   cp .env.staging .env
   ```

   Webpack reads `.env` automatically; it does not load `.env.staging` directly. The `.env` file is ignored by Git.
   Change its `BASE_URL` to:

   ```dotenv
   BASE_URL=https://localhost:8080
   ```

   Keep the staging API URL, API version, and authorization endpoint from the copied file.

3. Register a development OAuth client in [staging Backoffice](https://staging-backoffice.ops.zinfra.io/swagger-ui/index.html#/default/register-oauth-client) with:

   ```json
   {
     "application_name": "Wire Calendar Outlook Add-in Local Development",
     "redirect_url": "https://localhost:8080/callback.html"
   }
   ```

   The client in `.env.staging` is registered for `https://outlook.integrations.zinfra.io/callback.html`, so it cannot be used with the localhost callback.
   Leave that registration unchanged so the deployed staging add-in keeps working.
   If Backoffice is unavailable, see [How to create a new OAuth client](#how-to-create-a-new-oauth-client) and use the localhost redirect URL.

   You can reuse another developer's client if it is registered on the same backend with exactly the same callback URL, including the scheme, hostname, port, and path.
   A client identifies the application, not an individual developer; each developer signs in with their own Wire account.

4. Replace `CLIENT_ID` in your `.env` with the returned `client_id`.
   This add-in uses OAuth with PKCE and does not use the returned client secret. Do not put the secret in `.env` or frontend code.

5. Start the HTTPS development server:

   ```shell
   npm run dev-server-local
   ```

   Trust the local development certificate if prompted. You can check that the server is reachable at `https://localhost:8080/commands.html`.
   Restart the server after changing `.env`.

6. Add the generated `dist/manifest.xml` to Outlook using the [installation instructions below](#how-to-install-the-add-in-in-ms-outlook).
   Keep the development server running while using the add-in, then sign in with your Wire staging account.

If you choose a different port, update both `BASE_URL` and the OAuth client's registered callback URL to match, restart the server, and reinstall the generated manifest in Outlook.

## Local Storage
- isLoggedIn
- refresh_token
- access_token

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

## How to create a new OAuth client

When available, use Backoffice. For example, the corresponding endpoint for Staging private API is located at https://staging-backoffice.ops.zinfra.io/swagger-ui/index.html#/default/register-oauth-client

Otherwise, connect to the backend pod of the Brig service and run:

```shell
curl -s -X POST localhost:8080/i/oauth/clients \
    -H "Content-Type: application/json" \
    -d '{
      "application_name":"Wire Microsoft Outlook Calendar Add-in",
      "redirect_url":"https://outlook.wire.com/callback.html" 
    }'
```

## How to install the Add-in in MS Outlook

- Get the manifest.xml file:
  - Production:
    ```shell
    curl https://outlook.integrations.wire.com/manifest.xml > manifest.xml
    ``` 
  - Development: follow [Local development](#local-development), then use the generated `dist/manifest.xml`.
- Open an email and go to three dots and select Get Add-ins
![Step 1](images/step_1.png)
- Go to My Add-ins, Custom Add-ins, **Add a Custom Add-in**
![Step 2](images/step_2.png)
- Choose **Add from File**.
![Step 3](images/step_3.png)

Wire button will appear in the toolbar when a new event is being created

## Troubleshooting
- If you are getting `401` error, please make sure that you have enabled the feature flag `outlookCalIntegration` for your account.
- If your browser is blocking third-party cookies, please make sure to allow them for the add-in to work properly. Or you can add `https://outlook.office.com` to the list of allowed websites.
- For local development the add-in requires HTTPS and uses a self-signed certificate. If the add-in does not load, open `https://localhost:8080/commands.html` in your browser. If your browser displays a certificate warning, accept or trust the certificate, then reload Outlook and try again.
- If authorization reports a redirect URL mismatch, check that the OAuth client is registered for exactly `${BASE_URL}/callback.html` and that `.env` contains its client ID. The deployed staging client's callback does not match localhost.
