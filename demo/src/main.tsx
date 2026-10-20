import { createRoot } from "react-dom/client";
import { DemoAuth } from "./auth";
import { BudgetPortal } from "./BudgetPortal";
import { loadConfig } from "./config";
import { RequestPage } from "./RequestPage";
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
    const auth = new DemoAuth(config);
    await auth.initialize();
    const requestPage = new URLSearchParams(window.location.search).get("view") === "request";
    if (!disposed) {
      document.title = requestPage ? "Copilot Budget Portal | Request an increase" : "Copilot Budget Portal | Live budget details";
      app.render(requestPage ? <RequestPage config={config} auth={auth} /> : <BudgetPortal config={config} auth={auth} />);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (!disposed) app.render(<main className="login-shell"><section className="card login-card"><h1>Sign-in unavailable</h1><p className="error" role="alert">{message}</p><p>Contact the demo administrator to check the sign-in configuration, then reload.</p></section></main>);
  }
}

void start();
