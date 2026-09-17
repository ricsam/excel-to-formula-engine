import { describe, expect, test } from "bun:test";
import { excelSerialToIso } from "../../src/translate/dates";

describe("excelSerialToIso", () => {
  test("converts whole serials to a date", () => {
    // 1900-01-01 is serial 1 in Excel's 1900 system.
    expect(excelSerialToIso(1)).toBe("1900-01-01");
    expect(excelSerialToIso(45306)).toBe("2024-01-15");
    expect(excelSerialToIso(45473)).toBe("2024-06-30");
  });

  test("accounts for Excel's 1900 leap year bug", () => {
    // Serial 59 is 1900-02-28; serial 60 is Excel's non-existent 1900-02-29,
    // and 61 is 1900-03-01. Serials on either side of the bug must still map to
    // the correct real date.
    expect(excelSerialToIso(59)).toBe("1900-02-28");
    expect(excelSerialToIso(61)).toBe("1900-03-01");
  });

  test("converts a fraction-only serial to a time", () => {
    expect(excelSerialToIso(0.5)).toBe("12:00:00");
    expect(excelSerialToIso(0.25)).toBe("06:00:00");
  });

  test("converts a serial with a fraction to a date and time", () => {
    expect(excelSerialToIso(45306.5)).toBe("2024-01-15 12:00:00");
  });

  test("supports the 1904 date system", () => {
    expect(excelSerialToIso(0, true)).toBe("1904-01-01");
    // The same calendar day is 1462 lower in the 1904 system.
    expect(excelSerialToIso(45306 - 1462, true)).toBe("2024-01-15");
  });

  test("passes through values that are not finite numbers", () => {
    expect(excelSerialToIso(Number.NaN)).toBe("NaN");
    expect(excelSerialToIso(Number.POSITIVE_INFINITY)).toBe("Infinity");
  });
});
