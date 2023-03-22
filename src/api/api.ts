// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, fetch, console */

import jwt_decode from "jwt-decode";

const apiUrl = "https://staging-nginz-https.zinfra.io/v2";

export async function createGroupConversation(name: string) {
  let token = localStorage.getItem('token');
  let refreshToken = localStorage.getItem('refresh_token');
  if((!token || !isTokenStillValid(token)) && isTokenStillValid(refreshToken)) {
    const newTokens = await refreshAccessToken(refreshToken);
    console.log("new tokens: ", newTokens);
    token = newTokens.access_token;
    refreshToken = newTokens.refresh_token;
    localStorage.setItem('token', token);
    localStorage.setItem('refresh_token', refreshToken);
  }
  const teamId = await getTeamId();
  const payload = {
    access: ["invite", "code"],
    access_role_v2: ["guest", "non_team_member", "team_member", "service"],
    conversation_role: "wire_member",
    name: name ?? "New appointment",
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
  const decodedToken = jwt_decode(token) as any;
  return decodedToken.exp * 1000 > (new Date()).getTime();
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
