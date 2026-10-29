/* global console */

import type { EventResult } from "../types/EventResult";
import { config } from "../utils/config";
import { fetchWithAuthorizeDialog } from "../wireAuthorize/wireAuthorize";
import { mapRecurrence } from "./mapRecurrence";

export async function createEvent(
  name: string,
  startTime: string,
  endTime: string,
  timeZone: string,
  recurrence: Office.Recurrence | null
): Promise<EventResult> {
  try {
    // Throws for unsupported patterns, so no request is sent
    const meetingRecurrence = mapRecurrence(recurrence);

    const payload = {
      title: name,
      start_time: startTime,
      end_time: endTime,
      tzid: timeZone,
      type: "scheduled",
      ...(meetingRecurrence && { recurrence: meetingRecurrence }),
    };

    const response = await fetchWithAuthorizeDialog(
      new URL(`${config.apiVersion}/meetings`, config.apiBaseUrl),
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      }
    );

    if (response.ok) {
      const meeting = await response.json();

      return { id: meeting.qualified_id.id, link: meeting.link };
    } else {
      throw new Error(`Request failed with status ${response.status}`);
    }
  } catch (error) {
    console.error("Error during creating new event in Wire", error);
    throw error;
  }
}
