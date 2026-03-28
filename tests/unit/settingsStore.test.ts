import { beforeEach, expect, test } from "vitest";
import { createSettingsStore } from "../../src/state/settingsStore";

beforeEach(() => {
  window.localStorage.clear();
});

test("persists api key and model choice", () => {
  const firstStore = createSettingsStore();
  firstStore.save({ apiKey: "sk-demo", model: "gpt-4.1-mini" });

  const reloadedStore = createSettingsStore();
  expect(reloadedStore.load()).toEqual({ apiKey: "sk-demo", model: "gpt-4.1-mini" });
});
