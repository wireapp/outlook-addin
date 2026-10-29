import type { MeetingRecurrence } from "../types/MeetingRecurrence";

const SUPPORTED_WEEKLY_INTERVALS = [1, 2, 4];

export class UnsupportedRecurrenceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "UnsupportedRecurrenceError";
  }
}

/**
 * Maps an Outlook recurrence pattern to the recurrence format accepted by the Wire meetings API.
 * Supported patterns: daily, weekly, every two weeks and every four weeks (on a single weekday).
 *
 * @param {Office.Recurrence | null} recurrence - The recurrence of the calendar event, null for non-recurring events.
 * @return {MeetingRecurrence | null} The Wire recurrence, or null for non-recurring events.
 * @throws {UnsupportedRecurrenceError} If the recurrence pattern is not supported by Wire.
 */
export function mapRecurrence(recurrence: Office.Recurrence | null): MeetingRecurrence | null {
  if (!recurrence) {
    return null;
  }

  const { recurrenceType, recurrenceProperties } = recurrence;
  const interval = recurrenceProperties?.interval ?? 1;
  const endDate = recurrence.seriesTime.getEndDate();
  const until = endDate ? `${endDate}T23:59:59Z` : undefined;

  if (recurrenceType === "daily" && interval === 1) {
    return { frequency: "daily", ...(until && { until }) };
  }

  if (recurrenceType === "weekly" && SUPPORTED_WEEKLY_INTERVALS.includes(interval)) {
    const days = recurrenceProperties?.days ?? [];

    if (days.length <= 1) {
      return {
        frequency: "weekly",
        ...(interval !== 1 && { interval }),
        ...(until && { until }),
      };
    }
  }

  throw new UnsupportedRecurrenceError(
    `Unsupported recurrence: ${JSON.stringify({ recurrenceType, recurrenceProperties })}`
  );
}
