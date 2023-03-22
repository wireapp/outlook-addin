const { default: axios } = require("axios");
const express = require("express");
const cors = require("cors");
const app = express();
const port = process.env.PORT || 4000;

app.use(cors());

app.listen(port, () => console.log(`Listening on port ${port}`));

app.get("/oauth2callback", (req, res) => {
  if (!(req.xhr || req.headers.accept.indexOf('json') > -1)) {
    axios
      .post(
        "https://staging-nginz-https.zinfra.io/oauth/token",
        {
          grant_type: "authorization_code",
          client_id: "91ab148a-4b6c-4eac-aab6-60f316912f4d",
          client_secret: "92f7b4d4613f0e70ebf391a8f108e90c536bbccecb2b3eba59de32102a9831bf",
          code: req.query.code,
          redirect_uri: "https://outlook.integrations.zinfra.io/oauth2callback",
        },
        {
          headers: {
            Accept: "application/json",
            "Content-Type": "application/x-www-form-urlencoded",
          },
        }
      )
      .then((result) => {
        console.log(result.data.access_token);
        res.send("<html><head><script src='https://appsforoffice.microsoft.com/lib/1/hosted/office.js' type='text/javascript'></script></head><body>you are authorized " + result.data.access_token +
        "<script>document.addEventListener('DOMContentLoaded', function () { localStorage.setItem('token', '" + result.data.access_token + "'); }, false);</script>" +
        "<script>Office.onReady(function() { document.addEventListener('DOMContentLoaded', function () { console.log('dialog Office messageParent'); Office.context.ui.messageParent(JSON.stringify('" + result.data.access_token + "')); console.log('AFTER dialog Office messageParent'); }, false); });</script>" +
        "</body></html>");
      })
      .catch((err) => {
        console.log('error: ' + err.response.data.message);
        res.send("<html><head></head><body>error " + err.response.data.message +
        "<script>document.addEventListener('DOMContentLoaded', function () { localStorage.setItem('token', ''); }, false);</script>" +
        "</body></html>");
      });
  } else {
    res.send("");
  }
});

app.get("/login", (req, res) => {
    res.send("<html><head></head><body>" +
    "<script>document.addEventListener('DOMContentLoaded', function () { window.location.replace('https://wire-webapp-edge.zinfra.io/auth?client_id=91ab148a-4b6c-4eac-aab6-60f316912f4d&state=boop&response_type=code&redirect_uri=https://outlook.integrations.zinfra.io/oauth2callback&scope=write%3Aconversations+write%3Aconversations_code+read%3Aself+read%3Afeature_configs#/authorize'); }, false);</script>" +
    "</body></html>");
});
