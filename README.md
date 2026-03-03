# XLSX Cells v2 — Obsidian Plugin

Renders live Excel cell values and ranges directly inside Obsidian notes via fenced code blocks and inline backtick spans. Values refresh automatically whenever the source `.xlsx` file is modified.

---

## Quick-start

1. Place your `.xlsx` file anywhere inside the vault.
2. *(Optional)* Set note-level defaults in frontmatter:

```yaml
---
excel_file: _external/MyData.xlsx
excel_sheet: Sheet1
excel_enabled: false   # set to false to disable all plugin processing for this note
---
```

3. Use a fenced code block or inline span (see below).

---

## Fenced Code Blocks

Use a fenced ` ```excel ``` ` block to embed a cell value or a table range.

~~~markdown
```excel
file:   _external/Budget.xlsx
sheet:  January
cell:   C5
format: %.2f
```
~~~

### Keys

| Key | Required | Description |
|-----|----------|-------------|
| `file` | No | Vault-relative path to the `.xlsx` file. Falls back to note frontmatter → plugin settings. |
| `sheet` | No | Sheet tab name. Falls back to note frontmatter → plugin settings. |
| `cell` | Yes* | A1-style cell reference, e.g. `C5`. Mutually exclusive with `range`. |
| `range` | Yes* | A1:A1 range, e.g. `A1:E11`. Mutually exclusive with `cell`. |
| `headers` | No | `true`, `false`, or `auto` (default). Controls whether the first row renders as `<th>` header cells. |
| `format` | No | printf-style format string, e.g. `%.2f`. See *Format Strings* below. |

*Exactly one of `cell` or `range` must be provided.

Lines starting with `#` are treated as comments and are ignored.

### Examples

Single cell:

~~~markdown
```excel
cell: B3
```
~~~

Range with headers auto-detected:

~~~markdown
```excel
file:  _external/Sales.xlsx
sheet: Q1
range: A1:D10
```
~~~

Range with explicit headers and formatting:

~~~markdown
```excel
range:   A1:F20
headers: true
format:  %,.2f
```
~~~

---

## Inline Spans

Use a backtick span starting with `xl:` to embed a value inline within a sentence.

```markdown
Revenue this quarter: `xl:Sheet1:B5:%.2f`
```

### Syntax

All parts are separated by colons. The parser reads from right to left:

| Pattern | Description |
|---------|-------------|
| `` `xl:C5` `` | Single cell, all defaults |
| `` `xl:C5:%.2f` `` | Single cell with format |
| `` `xl:Sheet1:C5` `` | Cell on explicit sheet |
| `` `xl:Budget.xlsx:Sheet1:C5` `` | Fully specified cell |
| `` `xl:A1:D5` `` | Range, all defaults |
| `` `xl:Sheet1:A1:D5` `` | Range on explicit sheet |
| `` `xl:Budget.xlsx:Sheet1:A1:D5:%.2f` `` | Fully specified range with format |

**Parsing rules:**
- If the last part starts with `%` it is extracted as the format string.
- If the last two parts are both valid A1 references (e.g. `A1` and `D5`), a range is created.
- If only the last part is a valid A1 reference, it is a single cell.
- Remaining left parts: none → use defaults; one part → sheet; two parts → file + sheet.

---

## Disabling the Plugin Per Note

Add `excel_enabled: false` to a note's frontmatter to turn off all plugin processing for that note:

```yaml
---
excel_enabled: false
---
```

When disabled:
- Fenced ` ```excel ``` ` blocks render as plain `<pre><code>` text (Obsidian's built-in code block styling), so the source is still readable.
- Inline `` `xl:...` `` spans are left completely untouched.

This is useful for README notes or documentation notes that contain example blocks you don't want the plugin to evaluate.

Only an explicit boolean `false` disables the plugin; omitting the key, setting it to `true`, or any other value leaves the plugin active.

---

## Defaults Cascade

File and sheet are resolved through a three-level cascade:

```
block / inline key  →  note frontmatter  →  plugin settings
```

1. **Block/inline key** — `file:` / `sheet:` specified directly in the code block or `xl:` reference.
2. **Note frontmatter** — `excel_file`, `excel_sheet`, and `excel_enabled` YAML keys at the top of the note.
3. **Plugin settings** — *Default file* and *Default sheet* configured in Obsidian's Settings UI.

---

## Plugin Settings

Navigate to **Settings → Community Plugins → XLSX Cells v2** to configure:

| Setting | Description |
|---------|-------------|
| Default file | Vault-relative path to the fallback `.xlsx` file (e.g. `_external/Budget.xlsx`). |
| Default sheet | Fallback sheet tab name (e.g. `Sheet1`). |

Settings are persisted to `.obsidian/plugins/xlsx-cells-v2/data.json` via Obsidian's standard `loadData()` / `saveData()` API.

---

## Auto-refresh

When an `.xlsx` file inside the vault is modified:

1. The plugin debounces 500ms to absorb rapid saves from Excel and other applications.
2. The cached workbook for that file is invalidated.
3. Only the notes that reference the changed file are re-rendered — not every open note.

---

## Format String Reference

The optional `format` key / suffix uses **C printf-style** syntax.

```
%[flags][.precision]specifier
```

### Flags (combinable)

| Flag | Effect | Example |
|------|--------|---------|
| `,` | Thousands separator | `%,.2f` → `1,234.56` |
| `+` | Always show sign | `%+.2f` → `+3.14` |

### Precision

| Syntax | Effect |
|--------|--------|
| `.N` | N digits after the decimal point (for `f`, `e`) |
| `.N` | N significant figures (for `g`) |

### Specifiers

| Specifier | Description | Example input | `%.2f` output |
|-----------|-------------|---------------|----------------|
| `f` | Fixed-point decimal | `3.14159` | `3.14` |
| `e` | Scientific notation | `314159` | `3.14e+5` |
| `g` | Shortest of `f` / `e` | `0.000314` | `3.1e-4` |
| `d` / `i` | Integer (rounded) | `3.7` | `4` |
| `s` | String pass-through (no-op) | `hello` | `hello` |

Non-numeric cell values always pass through unchanged, regardless of specifier.

### Quick examples

| Format string | Input | Output |
|---------------|-------|--------|
| `%.2f` | `3.14159` | `3.14` |
| `%.0f` | `3.7` | `4` |
| `%,.2f` | `12345.6` | `12,345.60` |
| `%+.2f` | `3.14` | `+3.14` |
| `%+,.0f` | `12345` | `+12,345` |
| `%.4e` | `314159` | `3.1416e+5` |
| `%.3g` | `3.14159` | `3.14` |
| `%d` | `3.9` | `4` |

---

## Headers (Range Mode)

The `headers` key controls whether the first row of a range renders as `<th>` cells:

| Value | Behaviour |
|-------|-----------|
| `true` / `yes` | Always use the first row as header cells. |
| `false` / `no` | Never use header cells. |
| `auto` *(default)* | Auto-detect: uses headers when every cell in the first row has a string type. |

---

## Styling

The plugin emits these CSS classes which can be targeted with an Obsidian CSS snippet (`.obsidian/snippets/xlsx-cells.css`):

| Class | Element | Description |
|-------|---------|-------------|
| `.excel-cell` | `<span>` | Block-mode single-cell value |
| `.excel-inline-cell` | `<span>` | Inline single-cell value |
| `.excel-range` | `<table>` | Range table |
| `.excel-table-wrapper` | `<div>` | Block wrapper around the table |
| `.excel-loading` | `<div>` | Loading placeholder shown before data arrives |
| `.excel-error` | `<div>` / `<span>` | Error badge |
| `.excel-cell-num` | `<td>` / `<th>` | Number cell (right-aligned by default) |
| `.excel-cell-str` | `<td>` / `<th>` | String / empty cell (left-aligned by default) |
| `.excel-cell-bool` | `<td>` / `<th>` | Boolean cell (centred by default) |

Example snippet:

```css
.excel-cell,
.excel-inline-cell {
  font-variant-numeric: tabular-nums;
  color: var(--text-accent);
}

.excel-table-wrapper {
  overflow-x: auto;
  margin: 0.5em 0;
}

.excel-range {
  border-collapse: collapse;
}

.excel-range td,
.excel-range th {
  border: 1px solid var(--background-modifier-border);
  padding: 4px 8px;
}

.excel-range th {
  background: var(--background-secondary);
  font-weight: 600;
}

.excel-cell-num { text-align: right; }
.excel-cell-str { text-align: left; }
.excel-cell-bool { text-align: center; }
```

---

## Building

```bash
npm install        # install dependencies
npm run build      # one-time production build → main.js
npm run dev        # watch mode (incremental builds with inline source maps)
```

Requires Node.js. The build uses [esbuild](https://esbuild.github.io/) and bundles `src/main.ts` into `main.js` (CommonJS, ES2020 target).

---

## Architecture

The plugin is split into five source files under `src/`:

| File | Responsibility |
|------|---------------|
| `src/main.ts` | Plugin entry point. Registers the `excel` code-block processor, the `xl:` inline post-processor, the file-change watcher, and settings. |
| `src/parser.ts` | Pure TypeScript parser for fenced block content and inline `xl:` references. No external dependencies. |
| `src/renderer.ts` | DOM rendering — cells, tables, spinners, and error badges. |
| `src/settings.ts` | Obsidian settings interface, defaults, and settings tab UI. |
| `src/workbook.ts` | `SafeWorkbookManager` — workbook caching, in-flight request deduplication, stale-cache fallback, and debounced reload. |

---

## Extending the Plugin

### Adding a new format specifier

1. Add the character to the specifier class in `applyFormat`'s regex in `src/renderer.ts`:
   ```typescript
   const m = fmt.match(/^%([+,]*)(?:\.(\d+))?([fegdis?]?)$/i);
   //                                                    ^ add new letter here
   ```
2. Add a corresponding `else if (spec === "?")` branch that returns a string.
3. Document it in this README.

### Adding plugin settings

New user-configurable values should be added to `XlsxCellsSettings` in `src/settings.ts`, given a default in `DEFAULT_SETTINGS`, and exposed via a new `Setting` in `XlsxCellsSettingTab.display()`.
