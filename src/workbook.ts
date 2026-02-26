import { App, TFile } from "obsidian";
import * as XLSX from "xlsx";

/**
 * SafeWorkbookManager — handles all xlsx file I/O with:
 *
 *  • Caching: once loaded, a workbook is served from memory until invalidated.
 *  • Inflight deduplication: if N callers request the same file at once, only
 *    one disk read occurs; all callers share the single in-flight Promise.
 *  • Stale-cache fallback: if a reload fails (e.g. file partially written by
 *    Excel mid-save), the last successfully-parsed workbook is returned instead
 *    of null. The user sees old data rather than an error flash.
 *  • Debounced reload: scheduleReload() waits 500ms after the last call before
 *    clearing the cache entry and invoking the refresh callback. This collapses
 *    the burst of vault "modify" events that applications like Excel emit during
 *    a single save into one reload operation.
 */
export class SafeWorkbookManager {
  private readonly app: App;
  private cache        = new Map<string, XLSX.WorkBook>();
  private inflight     = new Map<string, Promise<XLSX.WorkBook | null>>();
  private reloadTimers = new Map<string, ReturnType<typeof setTimeout>>();

  constructor(app: App) {
    this.app = app;
  }

  /**
   * Return the cached or freshly-loaded workbook for the given vault-relative
   * path. Returns null (never throws) if the file cannot be found or parsed.
   *
   * On parse failure, falls back to the last successfully-parsed version so
   * that in-progress saves by external apps do not produce error badges.
   */
  async getWorkbook(vaultPath: string): Promise<XLSX.WorkBook | null> {
    if (!vaultPath) return null;

    const file = this.app.vault.getFileByPath(vaultPath);
    if (!(file instanceof TFile)) {
      console.warn(`XLSX Cells v2: file not found in vault: "${vaultPath}"`);
      return null;
    }

    // Fast path — return cached workbook without any I/O
    const cached = this.cache.get(file.path);
    if (cached) return cached;

    // Deduplication — return the existing in-flight promise if one is running
    const existing = this.inflight.get(file.path);
    if (existing) return existing;

    // Capture stale workbook reference before the async load starts so the
    // closure can fall back to it if the new parse fails.
    const staleWb = this.cache.get(file.path) ?? null;

    const promise = (async (): Promise<XLSX.WorkBook | null> => {
      try {
        const raw = await this.app.vault.readBinary(file);
        // xlsx "array" type expects a Uint8Array / array-like, not a raw ArrayBuffer
        const wb = XLSX.read(new Uint8Array(raw), { type: "array" });
        this.cache.set(file.path, wb);
        return wb;
      } catch (e) {
        // File may be partially written (Excel mid-save). Return last-good
        // workbook so the UI stays stable during the write window.
        console.warn(
          `XLSX Cells v2: failed to read "${file.path}" — ` +
            (staleWb ? "using cached version." : "no cached version available."),
          e
        );
        return staleWb;
      } finally {
        this.inflight.delete(file.path);
      }
    })();

    this.inflight.set(file.path, promise);
    return promise;
  }

  /**
   * Schedule a cache invalidation + refresh callback after a 500ms quiet
   * period. Multiple rapid calls for the same path reset the timer, so only
   * the final "modify" event in a burst triggers the actual reload.
   */
  scheduleReload(vaultPath: string, onReload: () => void): void {
    const existing = this.reloadTimers.get(vaultPath);
    if (existing !== undefined) clearTimeout(existing);

    const timer = setTimeout(() => {
      this.reloadTimers.delete(vaultPath);
      this.invalidate(vaultPath);
      onReload();
    }, 500);

    this.reloadTimers.set(vaultPath, timer);
  }

  /** Immediately remove the cached workbook for a path (e.g. after debounce fires). */
  invalidate(vaultPath: string): void {
    this.cache.delete(vaultPath);
  }

  /** Release all state and cancel pending reload timers. Called from plugin.onunload(). */
  destroy(): void {
    for (const timer of this.reloadTimers.values()) {
      clearTimeout(timer);
    }
    this.cache.clear();
    this.inflight.clear();
    this.reloadTimers.clear();
  }
}
