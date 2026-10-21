import { useState } from "react";

export function ThemeToggle() {
  const [dark, setDark] = useState(() =>
    typeof document !== "undefined" && document.documentElement.getAttribute("data-theme") === "dark");
  function toggle() {
    const nextDark = !dark;
    document.documentElement.setAttribute("data-theme", nextDark ? "dark" : "light");
    setDark(nextDark);
  }
  return <button type="button" className="secondary" onClick={toggle}
    aria-label={`Switch to ${dark ? "light" : "dark"} mode`}>{dark ? "Light mode" : "Dark mode"}</button>;
}
