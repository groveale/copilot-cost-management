import { createRoot } from "react-dom/client";
import { DemoAuth } from "./auth";
import { loadConfig } from "./config";
import { HostedPortal } from "./HostedPortal";
import "./styles.css";

async function start() {
  const root = document.getElementById("root");
  if (!root) throw new Error("Application root is missing.");
  const app = createRoot(root);
  let disposed = false;
  import.meta.hot?.dispose(() => {
    disposed = true;
    app.unmount();
  });
  try {
    const config = await loadConfig();
    if (!config.apiBaseUrl || !config.apiScope || config.apiBaseUrl.includes("configure-your-api")) {
      throw new Error("Hosted API configuration is incomplete. Ask the administrator to deploy the public bootstrap settings.");
    }
    const auth = new DemoAuth(config);
    await auth.initialize();
    const requestPage = new URLSearchParams(window.location.search).get("view") === "request";
    if (!disposed) {
      document.title = requestPage ? "Copilot Budget Portal | Request an increase" : "Copilot Budget Portal | Live budget details";
      app.render(<HostedPortal config={config} auth={auth} requestPage={requestPage} />);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (!disposed) app.render(<main className="login-shell"><section className="card login-card"><h1>Sign-in unavailable</h1><p className="error" role="alert">{message}</p><p>Contact the demo administrator to check the sign-in configuration, then reload.</p></section></main>);
  }
}

void start();
