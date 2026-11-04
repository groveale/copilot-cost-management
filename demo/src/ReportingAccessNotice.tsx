import type { ReportingAccessMode } from "../../api/src/reporting-contracts";

export function ReportingAccessNotice({ accessMode }: { accessMode: ReportingAccessMode | null }) {
  return accessMode === "prototype" ? <p className="notice" role="note">
    Prototype mode - app filtering only. SharePoint row permissions are not managed or verified.
    Anyone with read access to the underlying lists may read other policies and user balances directly.
    Use only trusted testers and non-sensitive test data. Not production isolation.
  </p> : null;
}
