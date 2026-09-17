/**
 * Delimited-text parsing: CSV, TSV, and the semicolon-separated files a
 * comma-decimal locale produces.
 *
 * Follows RFC 4180 for quoting — a field wrapped in double quotes may contain
 * the delimiter, line breaks, and doubled quotes standing for one — because that
 * is what every spreadsheet writes and the naive `split(",")` gets wrong on the
 * first address field it meets.
 */

/** Delimiters worth guessing between. */
export const CANDIDATE_DELIMITERS = [",", "\t", ";", "|"] as const;

export type Delimiter = string;

/**
 * Split delimited text into rows of raw string fields.
 *
 * Values keep the exact text they had: trimming and type conversion are separate
 * decisions made later, where the caller's options are in scope.
 */
export function parseDelimitedText(text: string, delimiter: Delimiter): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let fieldWasQuoted = false;
  // Distinguishes "no fields yet" from "one empty field", so a trailing newline
  // does not produce a spurious final row.
  let rowHasContent = false;

  const endField = () => {
    row.push(field);
    field = "";
    fieldWasQuoted = false;
    rowHasContent = true;
  };

  const endRow = () => {
    endField();
    rows.push(row);
    row = [];
    rowHasContent = false;
  };

  for (let i = 0; i < text.length; i++) {
    const char = text[i]!;

    if (inQuotes) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          // A doubled quote inside a quoted field is one literal quote.
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
      continue;
    }

    if (char === '"' && field.length === 0) {
      inQuotes = true;
      fieldWasQuoted = true;
      continue;
    }

    if (char === delimiter) {
      endField();
      continue;
    }

    if (char === "\r") {
      // Treat CRLF and a bare CR as one line ending.
      if (text[i + 1] === "\n") i++;
      endRow();
      continue;
    }

    if (char === "\n") {
      endRow();
      continue;
    }

    field += char;
  }

  // A file not ending in a newline still has a last row; one that does should
  // not gain an empty one. A lone quoted empty field counts as content.
  if (field.length > 0 || fieldWasQuoted || rowHasContent) {
    endRow();
  }

  return rows;
}

/**
 * Guess the delimiter from the text itself.
 *
 * Scores each candidate on how *consistently* it splits the first several lines
 * rather than on how often it appears: prose full of commas scores badly because
 * its lines disagree about the field count, while a two-column TSV scores
 * perfectly on one tab per line. Quoted sections are skipped, so a comma inside
 * `"Smith, J"` never votes for comma.
 *
 * Falls back to comma, which is both the commonest and the least surprising
 * thing to be wrong about.
 */
export function detectDelimiter(
  text: string,
  candidates: readonly string[] = CANDIDATE_DELIMITERS
): Delimiter {
  const lines = takeSampleLines(text, 20);
  if (lines.length === 0) return ",";

  let best: { delimiter: string; score: number } | undefined;

  for (const delimiter of candidates) {
    const counts = lines.map((line) => countOutsideQuotes(line, delimiter));
    const present = counts.filter((count) => count > 0);
    // A delimiter that never appears is not a candidate.
    if (present.length === 0) continue;

    const first = counts[0]!;
    // Consistency across lines is the signal; ties break toward more fields,
    // which is what distinguishes the real delimiter from one that happens to
    // appear once per line inside the data.
    const consistent = counts.filter((count) => count === first).length;
    const score = consistent / counts.length + Math.min(first, 50) / 1000;

    if (!best || score > best.score) {
      best = { delimiter, score };
    }
  }

  return best?.delimiter ?? ",";
}

/**
 * The first `limit` non-empty lines, ignoring line breaks inside quoted fields.
 */
function takeSampleLines(text: string, limit: number): string[] {
  const lines: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < text.length && lines.length < limit; i++) {
    const char = text[i]!;

    if (char === '"') {
      if (inQuotes && text[i + 1] === '"') {
        current += '""';
        i++;
        continue;
      }
      inQuotes = !inQuotes;
      current += char;
      continue;
    }

    if (!inQuotes && (char === "\n" || char === "\r")) {
      if (char === "\r" && text[i + 1] === "\n") i++;
      if (current.trim().length > 0) lines.push(current);
      current = "";
      continue;
    }

    current += char;
  }

  if (current.trim().length > 0 && lines.length < limit) lines.push(current);
  return lines;
}

function countOutsideQuotes(line: string, delimiter: string): number {
  let count = 0;
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i]!;
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        i++;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && char === delimiter) count++;
  }

  return count;
}

/**
 * Strip a UTF-8 byte-order mark.
 *
 * Excel writes one on every CSV it exports, and left in place it becomes part of
 * the first header cell — where it is invisible, so the column looks right and
 * every lookup against its name silently misses.
 */
export function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
