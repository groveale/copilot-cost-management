import { useEffect, useState } from "react";

export type ReportingSource = "realtime" | "sync";
export type Appearance = "light" | "dark" | "system";
export interface Preferences { source: ReportingSource; appearance: Appearance }
const keys = { source: "copilotReportingSource", appearance: "copilotAppearance" };
const changed = "copilot-preferences-changed";
const defaults: Preferences = { source: "realtime", appearance: "system" };
const message = (error: unknown) => error instanceof Error ? error.message : String(error);

export function readPreferences(storage: Pick<Storage, "getItem">): { value: Preferences; error: string } {
  const value = { ...defaults };
  try {
    const source = storage.getItem(keys.source), appearance = storage.getItem(keys.appearance);
    if (source !== null) {
      if (source !== "realtime" && source !== "sync") throw new Error("Saved reporting source is invalid. Choose it again in Settings.");
      value.source = source;
    }
    if (appearance !== null) {
      if (appearance !== "light" && appearance !== "dark" && appearance !== "system") throw new Error("Saved appearance is invalid. Choose it again in Settings.");
      value.appearance = appearance;
    }
    return { value, error: "" };
  } catch (error) { return { value, error: `Browser preferences could not be loaded. ${message(error)}` }; }
}

function browserPreferences() {
  if (typeof window === "undefined") return { value: defaults, error: "" };
  try { return readPreferences(window.localStorage); }
  catch (error) { return { value: defaults, error: `Browser preferences are unavailable. ${message(error)}` }; }
}

export function applyAppearance(appearance: Appearance) {
  const override = new URLSearchParams(window.location.search).get("scoutTheme");
  const theme = override === "light" || override === "dark" ? override :
    appearance === "system" ? (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light") : appearance;
  document.documentElement.setAttribute("data-theme", theme);
}

export function usePreferences() {
  const [state, setState] = useState(browserPreferences);
  useEffect(() => {
    const update = () => setState(browserPreferences());
    window.addEventListener(changed, update);
    window.addEventListener("storage", update);
    return () => { window.removeEventListener(changed, update); window.removeEventListener("storage", update); };
  }, []);
  useEffect(() => {
    applyAppearance(state.value.appearance);
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    const update = () => applyAppearance(state.value.appearance);
    media.addEventListener("change", update);
    return () => media.removeEventListener("change", update);
  }, [state.value.appearance]);
  function save<K extends keyof Preferences>(key: K, value: Preferences[K]) {
    try {
      window.localStorage.setItem(keys[key], value);
      setState((previous) => ({ value: { ...previous.value, [key]: value }, error: "" }));
      window.dispatchEvent(new Event(changed));
    } catch (error) {
      setState((previous) => ({ ...previous, error: `Preference was not saved or changed. ${message(error)}` }));
    }
  }
  return { ...state, save };
}
