// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, fetch, console */

import jwt_decode from "jwt-decode";

const apiUrl = "https://staging-nginz-https.zinfra.io/v2";

export async function createGroupConversation(name: string) {
  console.log('BEGIN createGroupConversation');
  let token = localStorage.getItem('token');
  let refreshToken = localStorage.getItem('refresh_token');
  if(token) {
    console.log('token exists');
  }
  if(await isTokenStillValid(token)) {
    console.log('token is valid');
  } else {
    console.log('token is NOT valid');
  }
  if(!await isTokenStillValid(token)) {
    console.log('removing token as it is not valid');
    localStorage.removeItem('token');
    token = null;
  }
  if(!token && (await isTokenStillValid(refreshToken))) {
    console.log('refreshing token');
    const newTokens = await refreshAccessToken(refreshToken);
    console.log("new tokens: ", newTokens);
    token = newTokens.access_token;
    refreshToken = newTokens.refresh_token;
    localStorage.setItem('token', token);
    localStorage.setItem('refresh_token', refreshToken);
  }
  if(token && (await isTokenStillValid(token))) {
    console.log('token is still valid: ', token);
    const teamId = await getTeamId();
    const payload = {
      access: ["invite", "code"],
      access_role_v2: ["guest", "non_team_member", "team_member", "service"],
      conversation_role: "wire_member",
      name: name,
      protocol: "proteus",
      qualified_users: [],
      receipt_mode: 1,
      team: {
        managed: false,
        teamid: teamId,
      },
      users: [],
    };

    // TODO: any/model
    const response: any = await fetch(apiUrl + "/conversations", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify(payload),
    }).then((r) => r.json());

    return response.id;
  } else {
    return null;
  }
}

export async function createGroupLink(conversationId: string) {
  const token = localStorage.getItem('token');
  // TODO: any/model
  const response: any = await fetch(apiUrl + `/conversations/${conversationId}/code`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    },
  }).then((r) => r.json());

  return response.data.uri;
}

export async function getTeamId() {
  const token = localStorage.getItem('token');
  const response: any = await fetch(apiUrl + `/self`, {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    },
  }).then((r) => r.json());

  return response.team;
}

export async function isTokenStillValid(token: string) {
  if(token) {
    const decodedToken = jwt_decode(token) as any;
    console.log('isTokenStillValid for:');
    console.log(token);
    const currentDate = new Date();
    const currentTime = currentDate.getTime();
    console.log(decodedToken.exp * 1000 > currentTime);
    return decodedToken.exp * 1000 > currentTime;
  } 
  
  return false;
}

export async function refreshAccessToken(refresh_token: string) {
  const formBody = encodeURIComponent("refresh_token") + "=" + encodeURIComponent(refresh_token);
  const response: any = await fetch('/refreshToken/', {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: formBody,
  }).then((r) => r.json());

  return response;
}
