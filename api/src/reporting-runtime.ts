import type { Metadata } from "./contracts.js";
import { CertificateObo } from "./reporting-auth.js";
import { AzureReportingStore } from "./reporting-storage.js";
import { SharePointPermissions } from "./reporting-sharepoint.js";
import { ReportingService, reportingConfiguration } from "./reporting.js";
import { AzureSyncStore } from "./sync-storage.js";
import { syncSettings } from "./sync.js";
import { credential, required } from "./storage.js";

export function reportingService(authorization: string | null) {
  const sync = syncSettings();
  const config = reportingConfiguration(sync);
  const token = credential();
  return new ReportingService(new AzureReportingStore(config.tenantId, required("BUSINESS_BLOB_URL"), token), config, sync.enabled,
    (assertLock) => new SharePointPermissions(config.destination!, config.tenantId, config.admins,
      new CertificateObo(config.tenantId, config.clientId, config.certificateId, new URL(config.destination!.siteUrl).origin, authorization, token),
      assertLock),
    async () => (await new AzureSyncStore(config.tenantId, required("BUSINESS_TABLE_URL"), required("BUSINESS_BLOB_URL"), token).current())?.job ?? null);
}
export function guardMetadataWrite(metadata: Metadata, etag: string | null, write: (leaseId?: string) => Promise<void>, authorization: string | null) {
  return reportingService(authorization).mutateMetadata(metadata, etag, write);
}
