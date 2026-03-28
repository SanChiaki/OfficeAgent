import { beforeEach, expect, test } from "vitest";
import { getJson } from "../../src/state/localStorageAdapter";

beforeEach(() => {
  window.localStorage.clear();
});

test("falls back to the provided value when stored JSON is malformed", () => {
  window.localStorage.setItem("bad-key", "{not-json");

  expect(getJson("bad-key", ["fallback"])).toEqual(["fallback"]);
});
