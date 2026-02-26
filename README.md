# XLSX Cells — Obsidian Plugin

Renders live Excel cell values (and ranges) directly inside Obsidian notes.
The values refresh automatically whenever the source `.xlsx` file is modified.

---

## Quick-start

1. Place your `.xlsx` file anywhere inside the vault.
2. In your note's frontmatter, declare the defaults:

```yaml
---
excel_file: _external/MyData.xlsx
excel_sheet: Sheet1
---
```

3. Anywhere in the note body, write a call expression (see below).

---

## Call Reference

### `=excel(...)` — single cell

Returns the value of one cell, rendered inline as a `<span>`.

| Signature | Example |
|-----------|---------|
| `=excel(col, row)` | `=excel(C, 5)` |
| `=excel(col, row, fmt)` | `=excel(C, 5, %.2f)` |
| `=excel("file", col, row)` | `=excel("Budget.xlsx", C, 5)` |
| `=excel("file", col, row, fmt)` | `=excel("Budget.xlsx", C, 5, %.2f)` |
| `=excel("file", "sheet", col, row)` | `=excel("Budget.xlsx", "Jan", C, 5)` |
| `=excel("file", "sheet", col, row, fmt)` | `=excel("Budget.xlsx", "Jan", C, 5, %.2f)` |

- Argument order is **col, row** — matching Excel's natural A1 notation (column letter first).
- **col** accepts a column letter (`A`, `G`, `AA` …) or a plain 1-based integer.
- **row** accepts a row number or a letter (treated as the same letter-to-int conversion).
- `R1`/`C1` style tokens are accepted for numeric refs (`R27`, `C2`).
- Use either commas or semicolons as argument separators.
- `"file"` can be a vault-relative path **or** a configured alias (see *Aliases* below).
- `"sheet"` is the sheet tab name.
- When `"file"` and/or `"sheet"` are omitted the note's frontmatter defaults
  (`excel_file`, `excel_sheet`) are used, falling back to the hard-coded
  `fallbackFile` / `fallbackSheet` values in `main.ts`.

---

### `=exceltable(...)` — cell range as a table

Renders a rectangular cell range as an HTML `<table>` inside a
`<div class="excel-table-wrapper">`.

| Signature                                           | Example                                              |
| --------------------------------------------------- | ---------------------------------------------------- |
| `=exceltable(c1, r1, c2, r2)`                       | 'exceltable(A, 1, D, 5)'                             |
| `=exceltable(c1, r1, c2, r2, fmt)`                  | 'exceltable(A, 1, D, 5, %.2f)'                       |
| `=exceltable(c1, r1, c2, r2, h)`                    | 'exceltable(A, 1, D, 5, h)'                          |
| `=exceltable(c1, r1, c2, r2, hfmt)`                 | `exceltable(A, 1, D, 5, h%.2f)`                      |
| `=exceltable("file", c1, r1, c2, r2)`               | `exceltable("Budget.xlsx", A, 1, D, 5)`              |
| `=exceltable("file", "sheet", c1, r1, c2, r2)`      | `exceltable("Budget.xlsx", "Jan", A, 1, D, 5)`       |
| `=exceltable("file", "sheet", c1, r1, c2, r2, fmt)` | `exceltable("Budget.xlsx", "Jan", A, 1, D, 5, %.2f)` |

- Argument order is **col, row** per corner — matching Excel's A1 notation.
  `=exceltable(A, 38, G, 44)` selects the same range as `A38:G44` in Excel.
- **col** and **row** each accept a column letter (`A`, `G`, `AA` …) or a plain integer.
- `R1`/`C1` style tokens are accepted for numeric refs (`R27`, `C2`).
- Reversed corners (c1 > c2 or r1 > r2) are normalised automatically.
- Use either commas or semicolons as argument separators.
- The optional **`h` prefix** on the format string makes the first row render
  as `<th>` header cells.  `h` alone = header row only; `h%.2f` = header + formatting.
- Format strings do **not** need quotes: `h%.2f` and `"h%.2f"` are both accepted.
- Best placed on its own paragraph line; the wrapper `<div>` is a block element.

---

## Format String Reference

The optional format argument uses **C printf-style** syntax.
Quotes are **optional** — both forms are accepted:

```
=excel(C, 5, %.2f)       ← unquoted (recommended)
=excel(C, 5, "%.2f")     ← quoted   (also fine)
```

Format syntax:

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

| Specifier | Description | Example input | `"%.2f"` output |
|-----------|-------------|---------------|-----------------|
| `f` | Fixed-point decimal | `3.14159` | `3.14` |
| `e` | Scientific notation | `314159` | `3.14e+5` |
| `g` | Shortest of `f` / `e` | `0.000314` | `0.00` (`.2g` → `3.1e-4`) |
| `d` / `i` | Integer (rounded) | `3.7` | `4` |
| `s` | String pass-through (no-op) | `"hello"` | `hello` |

**Non-numeric cell values always pass through unchanged**, regardless of specifier.

### Quick examples

| Format string | Input | Output |
|---------------|-------|--------|
| `"%.2f"` | `3.14159` | `3.14` |
| `"%.0f"` | `3.7` | `4` |
| `"%,.2f"` | `12345.6` | `12,345.60` |
| `"%+.2f"` | `3.14` | `+3.14` |
| `"%+,.0f"` | `12345` | `+12,345` |
| `"%.4e"` | `314159` | `3.1416e+5` |
| `"%.3g"` | `3.14159` | `3.14` |
| `"%d"` | `3.9` | `4` |

---

## Note-level Defaults (Frontmatter)

```yaml
---
excel_file: _external/Budget.xlsx   # vault-relative path or alias
excel_sheet: January                # sheet tab name
---
```

Any `=excel(...)` or `=exceltable(...)` call that omits the file and/or sheet
arguments inherits these values.

---

## Aliases

Aliases map a short name to a vault-relative path.  They are currently
hard-coded in `main.ts` inside `aliasToPath`:

```typescript
private aliasToPath: Record<string, string> = {
  "Expenses Actual": "_external/Expenses Actual.xlsx",
};
```

Use an alias in calls: `=excel("Expenses Actual", 5, 3)`.

To add a new alias, update the map and rebuild (see *Building* below).

---

## Styling

The plugin emits these CSS classes you can target with an Obsidian CSS snippet:

| Class | Element | Description |
|-------|---------|-------------|
| `.excel-cell` | `<span>` | Single-cell inline value |
| `.excel-range` | `<table>` | Range table |
| `.excel-table-wrapper` | `<div>` | Block wrapper around the table |

Example snippet (`/.obsidian/snippets/xlsx-cells.css`):

```css
.excel-cell {
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
```

---

## Building

```bash
cd .obsidian/plugins/xlsx-cells
node build.js
```

Requires Node.js. Dependencies (`xlsx`, `esbuild`, TypeScript types) are
installed with `npm install`.

---

## Adding New Features

The plugin is intentionally small — everything lives in `main.ts`.

### Adding a new format specifier

1. Add the character to the specifier character class in `applyFormat`'s regex:
   ```typescript
   const m = fmt.match(/^%([+,]*)(?:\.(\d+))?([fegdis?]?)$/i);
   //                                                    ^ add new letter here
   ```
2. Add a `else if (spec === "?")` branch in `applyFormat` that produces a string.
3. Document it in this README.

### Adding a new call type

1. Define a new regex constant in `onload()` (model it on `reCell` or `reTable`).
   - Use `([A-Za-z]+|\d+)` for any position arg that should accept letter or integer refs.
   - Use `(["'][^"']*["']|[h%][^\s)]*)` for an optional unquoted-or-quoted format arg.
2. Collect its matches into `allMatches` with a new `type` tag.
3. Handle the new type in the `for (const { type, m } of allMatches)` loop.
   - Use `parseRef(m[N])` (not `Number(m[N])`) for any position arg so letter refs work.
4. Add a corresponding private method (model it on `readCellValue` /
   `readRangeAsTable`) that returns a DOM node.
5. Document the call syntax in this README.

### Adding plugin settings

Obsidian's `Plugin.loadData()` / `saveData()` API stores JSON in
`.obsidian/plugins/xlsx-cells/data.json`.  Move `fallbackFile`,
`fallbackSheet`, and `aliasToPath` into a settings object loaded in `onload()`
to make them user-configurable without a rebuild.
