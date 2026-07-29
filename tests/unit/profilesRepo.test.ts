import { describe, expect, it } from "vitest";
import { RESERVED_ALIASES, isValidAlias } from "../../src/storage/profilesRepo.js";

describe("isValidAlias", () => {
  it("accepts simple lowercase aliases", () => {
    expect(isValidAlias("personal")).toBe(true);
    expect(isValidAlias("work")).toBe(true);
    expect(isValidAlias("my-second-work-account")).toBe(true);
  });

  it("rejects reserved route names", () => {
    for (const reserved of RESERVED_ALIASES) {
      expect(isValidAlias(reserved)).toBe(false);
    }
  });

  it("rejects uppercase or invalid characters", () => {
    expect(isValidAlias("Personal")).toBe(false);
    expect(isValidAlias("my_account")).toBe(false);
    expect(isValidAlias("my account")).toBe(false);
    expect(isValidAlias("")).toBe(false);
  });

  it("rejects aliases starting or ending with a hyphen", () => {
    expect(isValidAlias("-personal")).toBe(false);
    expect(isValidAlias("personal-")).toBe(false);
  });

  it("rejects overly long aliases", () => {
    expect(isValidAlias("a".repeat(60))).toBe(false);
  });
});
