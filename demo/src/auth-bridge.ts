import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

broadcastResponseToMainFrame().catch((error: unknown) => {
  console.error("Authentication redirect failed.", error);
  const status = document.getElementById("auth-status");
  if (status) {
    status.textContent = `Authentication could not complete: ${error instanceof Error ? error.message : String(error)}`;
  }
});
