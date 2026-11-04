export type ReportingAccessMode = "sharepoint-acl" | "prototype";
export function reportingAccessMode(value: unknown): ReportingAccessMode {
  if (value === undefined) return "sharepoint-acl";
  if (value !== "sharepoint-acl" && value !== "prototype") throw new Error("Invalid reporting access mode.");
  return value;
}
export interface ReportingSettings {
  enabled: boolean;
  accessMode: ReportingAccessMode;
  siteId: string;
  policyListId: string;
  balanceListId: string;
  isAdmin: boolean;
  policyIds: string[];
  securityReady: boolean;
}
