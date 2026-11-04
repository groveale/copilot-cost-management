import { usePreferences } from "./preferences";

export function ThemeToggle() {
  const preferences = usePreferences();
  const dark = preferences.value.appearance === "dark" || (preferences.value.appearance === "system" &&
    typeof document !== "undefined" && document.documentElement.getAttribute("data-theme") === "dark");
  return <><button type="button" className="secondary" onClick={() => preferences.save("appearance", dark ? "light" : "dark")}
    aria-label={`Switch to ${dark ? "light" : "dark"} mode`}>{dark ? "Light mode" : "Dark mode"}</button>
    {preferences.error && <span className="error" role="alert">{preferences.error}</span>}</>;
}
