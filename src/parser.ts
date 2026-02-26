// parser.ts — pure TypeScript, zero external imports.
// Parses fenced block content and inline xl: references into BlockConfig objects.

export interface BlockConfig {
  file?: string;               // vault-relative .xlsx path
  sheet?: string;              // sheet tab name
  cell?: string;               // A1 reference string e.g. "C5"
  range?: string;              // Range string e.g. "A1:E11"
  headers?: boolean | "auto";  // true = always; false = never; "auto" = detect from data
  format?: string;             // printf-style format string e.g. "%.2f"
}

export type ParseResult =
  | { ok: true; config: BlockConfig }
  | { ok: false; error: string };

// ── A1 notation helpers ──────────────────────────────────────────────────────

const A1_RE = /^[A-Z]+[0-9]+$/;

/** Convert A1-style address to 1-based { r, c }. Assumes already uppercase. */
export function parseA1(ref: string): { r: number; c: number } {
  const m = ref.match(/^([A-Z]+)([0-9]+)$/);
  if (!m) throw new Error(`Invalid A1 reference: "${ref}"`);
  const colStr = m[1];
  const rowStr = m[2];

  let c = 0;
  for (const ch of colStr) {
    c = c * 26 + (ch.charCodeAt(0) - 64);
  }
  const r = parseInt(rowStr, 10);

  return { r, c };
}

/** Parse "A1:D5" into normalised { r1, c1, r2, c2 } (1-based). */
export function parseRange(
  rangeStr: string
): { r1: number; c1: number; r2: number; c2: number } {
  const [start, end] = rangeStr.split(":");
  const a = parseA1(start);
  const b = parseA1(end);
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c),
  };
}

// ── Fenced block parser ──────────────────────────────────────────────────────

/**
 * Parse the body of an ```excel fenced code block.
 *
 * Format (YAML-like, one key: value per line):
 *   file:    vault-relative .xlsx path      (optional)
 *   sheet:   sheet tab name                 (optional)
 *   cell:    A1 reference, e.g. C5          (mutually exclusive with range)
 *   range:   A1:A1 range, e.g. A1:E11       (mutually exclusive with cell)
 *   headers: true | false | auto            (optional, default: auto)
 *   format:  printf-style string, e.g. %.2f (optional)
 *
 * Lines starting with # are treated as comments and ignored.
 * Unknown keys are silently ignored (forward-compatible).
 */
export function parseBlockContent(source: string): ParseResult {
  const config: BlockConfig = {};

  for (const rawLine of source.split("\n")) {
    const line = rawLine.trim();

    // Skip blank lines and comment lines
    if (!line || line.startsWith("#")) continue;

    // Split on the FIRST colon only — values may contain colons (e.g. range "A1:D5")
    const colonIdx = line.indexOf(":");
    if (colonIdx === -1) continue;

    const key   = line.slice(0, colonIdx).trim().toLowerCase();
    const value = line.slice(colonIdx + 1).trim();

    if (!value) continue; // key with no value — skip

    switch (key) {
      case "file":
        config.file = value;
        break;
      case "sheet":
        config.sheet = value;
        break;
      case "cell":
        config.cell = value.toUpperCase();
        break;
      case "range":
        config.range = value.toUpperCase();
        break;
      case "headers":
        if (/^(true|yes)$/i.test(value))       config.headers = true;
        else if (/^(false|no)$/i.test(value))  config.headers = false;
        else if (/^auto$/i.test(value))         config.headers = "auto";
        // Unrecognised values are ignored — default stays undefined (treated as auto)
        break;
      case "format":
        config.format = value;
        break;
      // Unknown keys silently ignored
    }
  }

  return validateConfig(config);
}

// ── Inline reference parser ──────────────────────────────────────────────────

/**
 * Parse the content after "xl:" in an inline backtick span.
 *
 * Supported forms (all parts separated by colons):
 *   xl:C5                              — cell, all defaults
 *   xl:C5:%.2f                         — cell with format
 *   xl:Sheet1:C5                       — cell on explicit sheet
 *   xl:Budget.xlsx:Sheet1:C5           — fully specified cell
 *   xl:A1:D5                           — range, all defaults
 *   xl:Sheet1:A1:D5                    — range on explicit sheet
 *   xl:Budget.xlsx:Sheet1:A1:D5:%.2f   — fully specified range with format
 *
 * Parsing algorithm:
 *  1. If last part starts with "%" → extract as format
 *  2. If last 2 parts are both A1 refs → range mode
 *  3. Else if last part is A1 ref → cell mode
 *  4. Remaining left parts: [] / [sheet] / [file, sheet]
 */
export function parseInline(content: string): ParseResult {
  const trimmed = content.trim();
  if (!trimmed) {
    return { ok: false, error: "Empty xl: reference." };
  }

  const parts = trimmed.split(":");
  const config: BlockConfig = {};

  // Step 1: Extract optional format suffix
  if (parts.length > 0 && parts[parts.length - 1].startsWith("%")) {
    config.format = parts.pop()!;
  }

  if (parts.length === 0) {
    return { ok: false, error: "Empty xl: reference after format specifier." };
  }

  // Step 2–3: Detect cell vs range from the rightmost parts
  const last       = parts[parts.length - 1].toUpperCase();
  const secondLast = parts.length >= 2 ? parts[parts.length - 2].toUpperCase() : null;

  if (secondLast !== null && A1_RE.test(secondLast) && A1_RE.test(last)) {
    // Range: last two parts are both A1 references
    config.range = `${secondLast}:${last}`;
    parts.splice(parts.length - 2, 2);
  } else if (A1_RE.test(last)) {
    // Cell: last part is an A1 reference
    config.cell = last;
    parts.pop();
  } else {
    return {
      ok: false,
      error: `Cannot parse xl: reference — expected a cell (e.g. C5) or range (e.g. A1:D5) at the end, got "${last}".`,
    };
  }

  // Step 4: Remaining parts are file/sheet context
  if (parts.length === 1) {
    config.sheet = parts[0].trim() || undefined;
  } else if (parts.length >= 2) {
    config.file  = parts[0].trim() || undefined;
    config.sheet = parts[1].trim() || undefined;
    // Parts beyond index 1 are ignored
  }

  return validateConfig(config);
}

// ── Shared validation ────────────────────────────────────────────────────────

function validateConfig(config: BlockConfig): ParseResult {
  if (config.cell && config.range) {
    return { ok: false, error: "Specify either 'cell' or 'range', not both." };
  }

  if (!config.cell && !config.range) {
    return { ok: false, error: "Missing required key: 'cell' or 'range'." };
  }

  if (config.cell && !A1_RE.test(config.cell)) {
    return {
      ok: false,
      error: `Invalid cell reference: "${config.cell}". Expected A1 format, e.g. "C5".`,
    };
  }

  if (config.range) {
    const rangeParts = config.range.split(":");
    if (
      rangeParts.length !== 2 ||
      !A1_RE.test(rangeParts[0]) ||
      !A1_RE.test(rangeParts[1])
    ) {
      return {
        ok: false,
        error: `Invalid range: "${config.range}". Expected format like "A1:D5".`,
      };
    }
  }

  return { ok: true, config };
}
