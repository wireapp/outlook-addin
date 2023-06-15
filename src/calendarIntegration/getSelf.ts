import { config } from "../utils/config";
import { fetchWithAuthorizeDialog } from "../wireAuthorize/wireAuthorize";
import { SelfUser } from "../types/SelfUser";

export async function getSelf(): Promise<SelfUser> {
  const response = await fetchWithAuthorizeDialog(new URL("/self", config.apiBaseUrl), {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  }).then((r) => r.json());

  const user: SelfUser = {
    email: response.email,
    handle: response.handle,
    id: response.id,
    name: response.name,
    team: response.team,
  };

  return user;
}

export async function getTeamId(): Promise<string> {
  const user = await getSelf();
  return user.team;
}
