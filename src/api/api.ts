// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global global, fetch, console */

// TODO: should not be hardcoded
const teamIdStaging = "34e0d2d9-db1b-4029-8e01-471a11374dd5";

const apiUrl = "https://staging-nginz-https.zinfra.io/v2";
const token = localStorage.getItem('token');

export async function createGroupConversation(name: string) {
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
      teamid: teamIdStaging,
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
