/* global console */

import { FeatureConfigsResponse, Feature } from "../types/FeatureConfigsResponse";
import { config } from "../utils/config";
import { fetchWithAuthorizeDialog } from "../wireAuthorize/wireAuthorize";

export async function isOutlookCalIntegrationEnabled() {
  try {
    const response = await fetchWithAuthorizeDialog(new URL(`${config.apiVersion}/feature-configs`, config.apiBaseUrl), {
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
      const errorMsg = `Error while fetching outlookCalIntegration feature config. Status code: ${response.status}`;
      console.error(errorMsg);
      throw new Error(errorMsg);
    }
  } catch (error) {
    const errorMsg = `Error while checking outlookCalIntegration feature config: ${error}`;
    console.error(errorMsg);
    throw new Error(errorMsg);
  }

  return false;
}
