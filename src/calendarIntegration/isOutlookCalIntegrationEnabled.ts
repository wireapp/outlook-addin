/* global console */

import { FeatureConfigsResponse, Feature } from "../types/FeatureConfigsResponse";
import { config } from "../utils/config";
import { fetchWithAuthorizeDialog } from "../wireAuthorize/wireAuthorize";

export async function isOutlookCalIntegrationEnabled() {
  try {
    const response = await fetchWithAuthorizeDialog(new URL("/feature-configs", config.apiBaseUrl), {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
      },
    });

    if (response.ok) {
      const data: FeatureConfigsResponse = await response.json();
      const outlookCalIntegration: Feature | undefined = data.outlookCalIntegration;
      if (outlookCalIntegration && outlookCalIntegration.status === "enabled") {
        return true;
      }
    } else {
      console.error("Error while fetching outlookCalIntegration feature config. Status code: ", response.status);
    }
  } catch (error) {
    console.error("Error while checking outlookCalIntegration feature config: ", error);
  }

  return false;
}
