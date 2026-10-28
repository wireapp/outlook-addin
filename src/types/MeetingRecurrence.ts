export interface MeetingRecurrence {
  frequency: "daily" | "weekly";
  interval?: number;
  until?: string;
}
