import * as XLSX from "xlsx";
import { BlockConfig, parseA1, parseRange } from "./parser";

// ── Ported utilities (from original main.ts) ────────────────────────────────

/** Convert 1-based row / column integers to an A1-style address string. */
function a1FromRC(r: number, c: number): string {
  let col = "";
  let n   = c;
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n   = Math.floor((n - 1) / 26);
  }
  return `${col}${r}`;
}

/** Add comma thousands separators to an integer string (no decimals). */
function addThousandsSep(intStr: string): string {
  return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/**
 * Apply a printf-style format string to a raw value.
 *
 * Format: %[flags][.precision]specifier
 *   Flags:      "," = thousands separator,  "+" = always show sign
 *   Specifiers: f (fixed), e (exponential), g (shortest), d/i (integer), s (string)
 *
 * Non-numeric values pass through unchanged for numeric specifiers.
 */
function applyFormat(raw: string | number, fmt: string): string {
  const str = typeof raw === "string" ? raw : String(raw);
  if (!fmt || !fmt.startsWith("%")) return str;

  const num       = Number(str);
  const isNumeric = str.trim() !== "" && !isNaN(num);

  const m = fmt.match(/^%([+,]*)(?:\.(\d+))?([fegdis]?)$/i);
  if (!m) return str;

  const flags     = m[1] ?? "";
  const precision = m[2] !== undefined ? parseInt(m[2]) : undefined;
  const spec      = (m[3] || "f").toLowerCase();
  const thousands = flags.includes(",");
  const showSign  = flags.includes("+");

  if (spec === "s" || !isNumeric) return str;

  let result: string;

  if (spec === "d" || spec === "i") {
    const intVal = Math.round(num);
    const abs    = String(Math.abs(intVal));
    result = (intVal < 0 ? "-" : "") + (thousands ? addThousandsSep(abs) : abs);

  } else if (spec === "f") {
    const p     = precision ?? 6;
    const fixed = num.toFixed(p);
    if (thousands) {
      const neg                = num < 0;
      const [intPart, decPart] = fixed.replace(/^-/, "").split(".");
      result = (neg ? "-" : "") + addThousandsSep(intPart) +
               (decPart !== undefined ? "." + decPart : "");
    } else {
      result = fixed;
    }

  } else if (spec === "e") {
    result = num.toExponential(precision ?? 6);

  } else if (spec === "g") {
    const p = precision ?? 6;
    result  = parseFloat(num.toPrecision(p || 1)).toString();
    if (thousands) {
      const neg                = num < 0;
      const [intPart, decPart] = result.replace(/^-/, "").split(".");
      result = (neg ? "-" : "") + addThousandsSep(intPart) +
               (decPart !== undefined ? "." + decPart : "");
    }

  } else {
    return str;
  }

  if (showSign && num >= 0 && !result.startsWith("+")) {
    result = "+" + result;
  }

  return result;
}

// ── Cell type → CSS class ────────────────────────────────────────────────────

function cellTypeClass(cell: XLSX.CellObject | undefined): string {
  if (!cell) return "excel-cell-str"; // empty cell: left-align like a string
  switch (cell.t) {
    case "n": return "excel-cell-num";   // number → right-align
    case "b": return "excel-cell-bool";  // boolean → centre
    default:  return "excel-cell-str";   // string / error / date text → left-align
  }
}

// ── Smart header detection ───────────────────────────────────────────────────

/**
 * Determine whether to render the first row as <th> header cells.
 *
 * headers: true  → always
 * headers: false → never
 * headers: "auto" or undefined → auto-detect: use headers if every cell in
 *   the first row is a string type ("s"). This catches the common pattern of
 *   column-label rows in Excel spreadsheets.
 */
function shouldUseHeaders(
  config: BlockConfig,
  ws: XLSX.WorkSheet,
  r1: number,
  c1: number,
  c2: number
): boolean {
  if (config.headers === true)  return true;
  if (config.headers === false) return false;

  // "auto" or undefined: inspect first row
  for (let c = c1; c <= c2; c++) {
    const cell = ws[a1FromRC(r1, c)];
    if (!cell || cell.t !== "s") return false;
  }
  return true;
}

// ── State renderers ──────────────────────────────────────────────────────────

/** Fill el with a "Loading…" placeholder. Called synchronously before any async work. */
export function renderSpinner(el: HTMLElement): void {
  el.empty();
  const div = el.createDiv({ cls: "excel-loading" });
  div.createSpan({ text: "Loading…" });
}

/**
 * Replace el's content with a styled error badge.
 * Safe to call at any point — always clears el first.
 */
export function renderError(el: HTMLElement, message: string): void {
  el.empty();
  const badge = el.createDiv({ cls: "excel-error" });
  badge.createSpan({ cls: "excel-error-icon", text: "⚠ " });
  badge.createSpan({ cls: "excel-error-message", text: message });
}

/**
 * Replace an inline <code> element with a styled error badge.
 * Used by the Markdown post-processor for inline xl: references.
 * Uses standard DOM APIs because the element may not be an Obsidian-managed el.
 */
export function replaceCodeWithError(codeEl: HTMLElement, message: string): void {
  const badge = document.createElement("span");
  badge.className = "excel-error";

  const icon = document.createElement("span");
  icon.className = "excel-error-icon";
  icon.textContent = "⚠ ";

  const msg = document.createElement("span");
  msg.className = "excel-error-message";
  msg.textContent = message;

  badge.appendChild(icon);
  badge.appendChild(msg);
  codeEl.replaceWith(badge);
}

// ── Block-mode rendering (fenced code blocks) ────────────────────────────────

/**
 * Main entry point for fenced block rendering.
 * Clears the spinner from el, then dispatches to renderCell or renderTable.
 */
export function renderBlock(
  el: HTMLElement,
  config: BlockConfig,
  wb: XLSX.WorkBook,
  sheetName: string
): void {
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    renderError(el, `Sheet not found: "${sheetName}"`);
    return;
  }

  el.empty(); // clear the spinner

  if (config.cell) {
    renderCell(el, config, ws);
  } else if (config.range) {
    renderTable(el, config, ws);
  }
}

function renderCell(
  el: HTMLElement,
  config: BlockConfig,
  ws: XLSX.WorkSheet
): void {
  const { r, c } = parseA1(config.cell!);
  const addr     = a1FromRC(r, c);
  const cell     = ws[addr];
  const raw      = cell ? String(cell.v ?? "") : "";
  const text     = config.format ? applyFormat(raw, config.format) : raw;

  const span = el.createSpan({ cls: "excel-cell" });
  span.textContent = text;
  span.setAttr("data-excel-a1", addr);
}

function renderTable(
  el: HTMLElement,
  config: BlockConfig,
  ws: XLSX.WorkSheet
): void {
  const { r1, c1, r2, c2 } = parseRange(config.range!);
  const useHeaders = shouldUseHeaders(config, ws, r1, c1, c2);

  const wrapper = el.createDiv({ cls: "excel-table-wrapper" });
  const table   = wrapper.createEl("table", { cls: "excel-range" });

  for (let r = r1; r <= r2; r++) {
    const isHeaderRow = useHeaders && r === r1;
    const tr = table.createEl("tr");

    for (let c = c1; c <= c2; c++) {
      const addr     = a1FromRC(r, c);
      const cell     = ws[addr];
      const raw      = cell ? String(cell.v ?? "") : "";
      const text     = config.format ? applyFormat(raw, config.format) : raw;
      const typeClass = cellTypeClass(cell);

      const td = isHeaderRow
        ? tr.createEl("th", { cls: typeClass })
        : tr.createEl("td", { cls: typeClass });

      td.textContent = text;
      td.setAttr("data-excel-a1", addr);
    }
  }
}

// ── Inline-mode rendering (post-processor for `xl:...` spans) ───────────────

/**
 * Replace an inline <code> element with the rendered Excel value or table.
 * Uses codeEl.replaceWith() — never modifies any parent container.
 */
export function renderInlineResult(
  codeEl: HTMLElement,
  config: BlockConfig,
  wb: XLSX.WorkBook,
  sheetName: string
): void {
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    replaceCodeWithError(codeEl, `Sheet not found: "${sheetName}"`);
    return;
  }

  if (config.cell) {
    renderInlineCell(codeEl, config, ws);
  } else if (config.range) {
    renderInlineTable(codeEl, config, ws);
  }
}

function renderInlineCell(
  codeEl: HTMLElement,
  config: BlockConfig,
  ws: XLSX.WorkSheet
): void {
  const { r, c } = parseA1(config.cell!);
  const addr     = a1FromRC(r, c);
  const cell     = ws[addr];
  const raw      = cell ? String(cell.v ?? "") : "";
  const text     = config.format ? applyFormat(raw, config.format) : raw;

  const span = document.createElement("span");
  span.className = "excel-inline-cell";
  span.setAttribute("data-excel-a1", addr);
  span.textContent = text;

  codeEl.replaceWith(span);
}

function renderInlineTable(
  codeEl: HTMLElement,
  config: BlockConfig,
  ws: XLSX.WorkSheet
): void {
  const { r1, c1, r2, c2 } = parseRange(config.range!);
  const useHeaders = shouldUseHeaders(config, ws, r1, c1, c2);

  const wrapper = document.createElement("div");
  wrapper.className = "excel-table-wrapper";

  const table = document.createElement("table");
  table.className = "excel-range";

  for (let r = r1; r <= r2; r++) {
    const isHeaderRow = useHeaders && r === r1;
    const tr = document.createElement("tr");

    for (let c = c1; c <= c2; c++) {
      const addr      = a1FromRC(r, c);
      const cell      = ws[addr];
      const raw       = cell ? String(cell.v ?? "") : "";
      const text      = config.format ? applyFormat(raw, config.format) : raw;
      const typeClass = cellTypeClass(cell);

      const td = document.createElement(isHeaderRow ? "th" : "td");
      td.className = typeClass;
      td.setAttribute("data-excel-a1", addr);
      td.textContent = text;
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  wrapper.appendChild(table);
  codeEl.replaceWith(wrapper);
}
