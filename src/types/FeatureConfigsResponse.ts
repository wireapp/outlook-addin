type LockStatus = "locked" | "unlocked";
type Status = "enabled" | "disabled";

export interface Feature {
  lockStatus: LockStatus;
  status: Status;
  ttl: string;
}

export interface FeatureConfigsResponse {
  outlookCalIntegration: Feature;
  [key: string]: any;
}
