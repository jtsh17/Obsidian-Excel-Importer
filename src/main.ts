import { MarkdownRenderChild, MarkdownView, Plugin, TFile } from "obsidian";
import {
  DEFAULT_SETTINGS,
  XlsxCellsSettings,
  XlsxCellsSettingTab,
} from "./settings";
import { SafeWorkbookManager } from "./workbook";
import { parseBlockContent, parseInline } from "./parser";
import {
  renderBlock,
  renderError,
  renderInlineResult,
  renderSpinner,
  replaceCodeWithError,
} from "./renderer";

export default class XlsxCellsPlugin extends Plugin {
  settings!: XlsxCellsSettings;
  workbooks!: SafeWorkbookManager;

  /**
   * Maps an xlsx vault path → Set of note source paths that reference it.
   * Used to limit view rerenders to only notes affected by a changed file,
   * rather than refreshing every open markdown leaf.
   */
  private dependencyMap = new Map<string, Set<string>>();

  async onload(): Promise<void> {
    // ── 1. Settings ─────────────────────────────────────────────────────────
    await this.loadSettings();

    // ── 2. WorkbookManager ──────────────────────────────────────────────────
    this.workbooks = new SafeWorkbookManager(this.app);

    // ── 3. Settings tab ─────────────────────────────────────────────────────
    this.addSettingTab(new XlsxCellsSettingTab(this.app, this));

    // ── 4. Fenced code block processor  (```excel ... ```) ──────────────────
    //
    // Obsidian calls this handler for every ```excel block in both reading
    // mode and live preview, so no separate CodeMirror extension is needed.
    this.registerMarkdownCodeBlockProcessor(
      "excel",
      async (source, el, ctx) => {
        // Per-note disable: render source as plain code, no processing or errors
        if (this.isNoteDisabled(ctx.sourcePath)) {
          const pre  = el.createEl("pre");
          const code = pre.createEl("code");
          code.textContent = source;
          return;
        }

        // Register with Obsidian's component tree for proper lifecycle management
        ctx.addChild(new MarkdownRenderChild(el));

        // Show loading placeholder immediately (synchronous)
        renderSpinner(el);

        // Parse the block content
        const result = parseBlockContent(source);
        if (!result.ok) {
          renderError(el, result.error);
          return;
        }

        const cfg = result.config;

        // Resolve file and sheet via cascading fallback:
        //   block key  →  note frontmatter  →  plugin settings
        const noteDefaults = this.getNoteDefaults(ctx.sourcePath);
        const filePath  = cfg.file  ?? noteDefaults.file  ?? this.settings.defaultFile;
        const sheetName = cfg.sheet ?? noteDefaults.sheet ?? this.settings.defaultSheet;

        if (!filePath) {
          renderError(
            el,
            "No file specified. Add 'file:' to the block, set 'excel_file' in note frontmatter, or configure a default in plugin settings."
          );
          return;
        }
        if (!sheetName) {
          renderError(
            el,
            "No sheet specified. Add 'sheet:' to the block, set 'excel_sheet' in note frontmatter, or configure a default in plugin settings."
          );
          return;
        }

        // Track dependency before loading (file-change watcher uses this)
        this.registerDependency(filePath, ctx.sourcePath);

        // Load workbook (cached after first load; deduplicates concurrent requests)
        const wb = await this.workbooks.getWorkbook(filePath);
        if (!wb) {
          renderError(el, `File not found or could not be read: "${filePath}"`);
          return;
        }

        // Render synchronously now that we have the data
        renderBlock(el, cfg, wb, sheetName);
      }
    );

    // ── 5. Inline span post-processor  (`xl:...`) ───────────────────────────
    //
    // Runs on rendered HTML in both reading mode and live preview.
    // Scans for <code> elements whose text starts with "xl:" and replaces them
    // with rendered values or tables.
    this.registerMarkdownPostProcessor((el, ctx) => {
      // Per-note disable: leave all xl: spans completely untouched
      if (this.isNoteDisabled(ctx.sourcePath)) return;

      const codeEls = el.querySelectorAll("code");

      codeEls.forEach((code) => {
        const text = code.textContent?.trim() ?? "";

        // Fast-exit: most code spans will not be xl: references
        if (!text.startsWith("xl:")) return;

        const result = parseInline(text.slice(3));
        if (!result.ok) {
          replaceCodeWithError(code as HTMLElement, result.error);
          return;
        }

        const cfg = result.config;
        const noteDefaults = this.getNoteDefaults(ctx.sourcePath);
        const filePath  = cfg.file  ?? noteDefaults.file  ?? this.settings.defaultFile;
        const sheetName = cfg.sheet ?? noteDefaults.sheet ?? this.settings.defaultSheet;

        if (!filePath || !sheetName) {
          replaceCodeWithError(
            code as HTMLElement,
            "Missing file or sheet. Add them to the xl: reference or configure defaults."
          );
          return;
        }

        this.registerDependency(filePath, ctx.sourcePath);

        // Show a minimal placeholder while the async load runs
        code.textContent = "…";

        this.workbooks.getWorkbook(filePath).then((wb) => {
          // Guard: the element may have been removed if the note re-rendered
          if (!(code as HTMLElement).isConnected) return;

          if (!wb) {
            replaceCodeWithError(
              code as HTMLElement,
              `File not found or could not be read: "${filePath}"`
            );
            return;
          }

          renderInlineResult(code as HTMLElement, cfg, wb, sheetName);
        });
      });
    });

    // ── 6. xlsx file-change watcher ─────────────────────────────────────────
    //
    // When an xlsx file in the vault changes:
    //  - Debounce 500ms to absorb rapid saves from Excel / other apps
    //  - Invalidate the cached workbook
    //  - Rerender only the notes that reference the changed file
    this.registerEvent(
      this.app.vault.on("modify", (f) => {
        if (!(f instanceof TFile)) return;
        if (f.extension !== "xlsx") return;

        // Skip if no open note currently references this file
        if (!this.dependencyMap.has(f.path)) return;

        this.workbooks.scheduleReload(f.path, () =>
          this.refreshDependentViews(f.path)
        );
      })
    );
  }

  onunload(): void {
    this.workbooks.destroy();
    this.dependencyMap.clear();
  }

  // ── Settings helpers ───────────────────────────────────────────────────────

  async loadSettings(): Promise<void> {
    this.settings = Object.assign(
      {},
      DEFAULT_SETTINGS,
      await this.loadData()
    );
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  // ── Note frontmatter defaults ──────────────────────────────────────────────

  /**
   * Extract excel_file / excel_sheet from the frontmatter of the note at
   * sourcePath. These act as per-note defaults, overridden by block-level
   * keys but overriding plugin settings defaults.
   *
   * Trims whitespace from values to tolerate typos like "excel_file:  foo.xlsx".
   */
  private getNoteDefaults(sourcePath: string): {
    file?: string;
    sheet?: string;
  } {
    const file = this.app.vault.getFileByPath(sourcePath);
    if (!file) return {};

    const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
    if (!fm) return {};

    return {
      file:  (fm.excel_file  as string | undefined)?.trim() || undefined,
      sheet: (fm.excel_sheet as string | undefined)?.trim() || undefined,
    };
  }

  /**
   * Returns true when the note at sourcePath has `excel_enabled: false` in its
   * frontmatter, disabling all plugin processing for that note.
   * Accepts both boolean false and the string "false" (Obsidian stores text
   * properties as strings even when the value looks like a boolean).
   */
  private isNoteDisabled(sourcePath: string): boolean {
    const file = this.app.vault.getFileByPath(sourcePath);
    if (!file) return false;

    const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
    if (!fm) return false;

    const val = fm.excel_enabled;
    return val === false || val === "false";
  }

  // ── Dependency tracking ────────────────────────────────────────────────────

  /**
   * Record that the note at notePath references the xlsx file at xlsxPath.
   * This is used to limit file-change rerenders to only relevant notes.
   */
  private registerDependency(xlsxPath: string, notePath: string): void {
    let notes = this.dependencyMap.get(xlsxPath);
    if (!notes) {
      notes = new Set();
      this.dependencyMap.set(xlsxPath, notes);
    }
    notes.add(notePath);
  }

  /**
   * Rerender only the markdown leaves whose note path is in dependencyMap
   * for the given xlsx file. This avoids the cost of refreshing every open
   * leaf whenever any xlsx file changes.
   */
  private refreshDependentViews(xlsxPath: string): void {
    const notePaths = this.dependencyMap.get(xlsxPath);
    if (!notePaths || notePaths.size === 0) return;

    for (const leaf of this.app.workspace.getLeavesOfType("markdown")) {
      const view = leaf.view;
      if (
        view instanceof MarkdownView &&
        notePaths.has(view.file?.path ?? "")
      ) {
        view.previewMode?.rerender(true);
      }
    }
  }
}
