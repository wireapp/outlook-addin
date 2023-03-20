const { default: axios } = require("axios");
const express = require("express");
const cors = require("cors");
const app = express();
const port = process.env.PORT || 4000;

app.use(cors({
  origin: ['https://wire-webapp-edge.zinfra.io', 'https://staging-nginz-https.zinfra.io']
}));

app.listen(port, () => console.log(`Listening on port ${port}`));

app.get("/oauth2callback", (req, res) => {
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
      res.send("you are authorized " + result.data.access_token);
    })
    .catch((err) => {
      console.log(err.response.data.message);
      res.send("error " + err.response.data.message);
    });
});
