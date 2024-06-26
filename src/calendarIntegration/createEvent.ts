/* global console */

import { EventResult } from "../types/EventResult";
import { config } from "../utils/config";
import { getTeamId } from "./getSelf";
import { fetchWithAuthorizeDialog } from "../wireAuthorize/wireAuthorize";

export async function createEvent(name: string): Promise<EventResult> {
  try {
    const teamId = await getTeamId();

    const payload = {
      access: ["invite", "code"],
      access_role: ["guest", "non_team_member", "team_member", "service"],
      conversation_role: "wire_member",
      name: name,
      protocol: "proteus",
      qualified_users: [],
      receipt_mode: 1,
      team: {
        managed: false,
        teamid: teamId,
      },
    };

    const response = await fetchWithAuthorizeDialog(new URL(`${config.apiVersion}/conversations`, config.apiBaseUrl), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (response.ok) {
      const conversationId = (await response.json()).id;

      const responseLink: any = await fetchWithAuthorizeDialog(
        new URL(`${config.apiVersion}/conversations/${conversationId}/code`, config.apiBaseUrl),
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ password: null}),
        }
      ).then((r) => r.json());

      return { id: conversationId, link: responseLink.data.uri };
    } else {
      throw new Error(`Request failed with status ${response.status}`);
    }
  } catch (error) {
    console.error("Error during creating new event in Wire", error);
    throw error;
  }
}
