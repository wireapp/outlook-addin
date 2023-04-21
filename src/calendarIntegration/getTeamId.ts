import { config } from "../utils/config";
import { fetchWithAuthorizeDialog } from "../wireAuthorize/wireAuthorize";

export async function getTeamId() {
  const response: any = await fetchWithAuthorizeDialog(new URL("/self", config.apiBaseUrl), {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  }).then((r) => r.json());

  return response.team;
}
