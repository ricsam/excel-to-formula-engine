/**
 * Excel date handling.
 *
 * Excel stores dates as a serial number of days since an epoch, and the engine
 * has no date type. A date-formatted number would therefore display as a bare
 * serial (45000), so it is converted to an ISO 8601 string instead, which is
 * both readable and sortable.
 */

/** Milliseconds in one day. */
const DAY_MS = 86_400_000;

/**
 * Excel's 1900 date system has a deliberate bug: it treats 1900 as a leap year,
 * so serial 60 is the non-existent 1900-02-29 and every serial after it is one
 * day ahead of the true date. Serials above 60 are therefore offset by an epoch
 * of 1899-12-30 rather than 1899-12-31.
 */
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30);
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1);

export function excelSerialToIso(serial: number, date1904 = false): string {
  if (!Number.isFinite(serial)) {
    return String(serial);
  }

  const epoch = date1904 ? EPOCH_1904_UTC : EPOCH_1900_UTC;

  // In the 1900 system serials 1-59 map to Jan 1 - Feb 28 1900 and need the
  // extra day that the leap-year bug adds back for later serials.
  const adjusted =
    !date1904 && serial < 60 && serial >= 1 ? serial + 1 : serial;

  const wholeDays = Math.floor(adjusted);
  const fraction = adjusted - wholeDays;
  const ms = epoch + wholeDays * DAY_MS + Math.round(fraction * DAY_MS);
  const date = new Date(ms);

  if (Number.isNaN(date.getTime())) {
    return String(serial);
  }

  const iso = date.toISOString();

  // A whole serial is a date with no time component; Excel represents a
  // time-only value as a fraction below 1.
  if (fraction === 0) {
    return iso.slice(0, 10);
  }
  if (wholeDays === 0) {
    return iso.slice(11, 19);
  }
  return iso.slice(0, 19).replace("T", " ");
}
