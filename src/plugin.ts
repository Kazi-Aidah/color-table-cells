import { Plugin, Menu, Notice, debounce } from "obsidian";
import { debugLog, debugWarn, DEFAULT_SETTINGS } from "./constants";
import { normalizeCellData, getCellText, evaluateMatch } from "./utils";
import { ColorPickerMenu } from "./color-picker";
import { ColorTableSettingTab } from "./settings";
import { AdvancedRuleModal } from "./modals/advanced-rule-modal";
import { ChangelogModal } from "./modals/changelog-modal";
import type { PluginSettings, ColorSnapshot, CellColorData, CellDataStore } from "./types";

export default class TableColorPlugin extends Plugin {
  settings!: PluginSettings;
  cellData: CellDataStore = {};
  undoStack: ColorSnapshot[] = [];
  redoStack: ColorSnapshot[] = [];
  maxStackSize = 50;
  applyColorsToAllEditors!: () => void;
  statusBarRefresh?: HTMLElement;
  _ribbonRefreshIcon?: HTMLElement;
  _settingsTab?: ColorTableSettingTab;
  _settingsWasOpen = false;
  _appliedContainers!: Map<Element, number>;
  _containerObservers!: Map<Element, unknown>;
  _livePreviewObserver?: MutationObserver;
  _tablePreRenderer?: MutationObserver;
  _readingViewObserver?: MutationObserver;
  _globalObserver?: MutationObserver | null;
  _lastApplyCall = 0;
  _settingsMigrated = false;
  _readingModeChecker?: number;

  async onload(): Promise<void> {
    await this.loadSettings();
    if (this.settings.showStatusRefreshIcon) this.createStatusBarIcon();
    if (this.settings.showRibbonRefreshIcon && !this._ribbonRefreshIcon) {
      this._ribbonRefreshIcon = this.addRibbonIcon("table", "Refresh table colors", () => this.hardRefreshTableColors());
    }
    this._registerCommands();
    this._setupLivePreview();
    if (this.settings?.persistUndoHistory) await this.loadUndoRedoStacks();
    const rawSaved = (await this.loadData()) || {};
    this._appliedContainers = new Map();
    this._containerObservers = new Map();
    this.cellData = normalizeCellData(rawSaved) || {};
    try {
      const normalized = { settings: this.settings, cellData: this.cellData };
      if (JSON.stringify(rawSaved) !== JSON.stringify(normalized)) await this.saveData(normalized);
    } catch (e) { debugWarn("Migration error:", e); }
    if (!this._settingsTab) {
      this._settingsTab = new ColorTableSettingTab(this.app, this);
      try { this.addSettingTab(this._settingsTab); } catch { /* already added */ }
    }
    this.registerDomEvent(document, "click", () => {
      const sc = document.querySelector(".vertical-tabs-container, .settings");
      if ((!sc || !(sc as HTMLElement).offsetParent) && this._settingsWasOpen) {
        this._settingsWasOpen = false;
        window.setTimeout(() => this.applyColorsToActiveFile(), 100);
      }
    });
    const appSetting = (this.app as unknown as { setting?: { open?: () => void } }).setting;
    if (appSetting) {
      const orig = appSetting.open?.bind(appSetting) || (() => {});
      appSetting.open = () => { this._settingsWasOpen = true; return orig(); };
    }
    this.registerEvent(this.app.workspace.on("file-open", async () => {
      const attempt = (n = 0) => {
        window.setTimeout(() => {
          try {
            const toDelete: Element[] = [];
            this._appliedContainers?.forEach((_, c) => { if (!c?.isConnected) toDelete.push(c); });
            toDelete.forEach((k) => this._appliedContainers.delete(k));
          } catch { /* ignore */ }
          if (document.querySelectorAll("table").length === 0 && n < 2) { attempt(n + 1); return; }
          this.applyColorsToActiveFile();
        }, [100, 300, 800][Math.min(n, 2)]);
      };
      attempt();
    }));
    this.registerEvent(this.app.workspace.on("layout-change", async () => {
      window.setTimeout(() => this.applyColorsToActiveFile(), 100);
    }));
    this._setupObservers();
    if (this.settings.enableContextMenu) this._setupContextMenu();
    this.setupReadingViewScrollListener();
    this.startReadingModeTableChecker();
  }

  private _registerCommands(): void {
    this.addCommand({ id: "enable-live-preview-coloring", name: "Enable live preview table coloring",
      callback: async () => {
        this.settings.livePreviewColoring = true;
        await this.saveSettings();
        this.app.workspace.trigger("layout-change");
        window.setTimeout(() => this.applyColorsToAllEditors(), 0);
      },
    });
    this.addCommand({ id: "disable-live-preview-coloring", name: "Disable live preview table coloring",
      callback: async () => {
        this.settings.livePreviewColoring = false;
        await this.saveSettings();
        this.app.workspace.trigger("layout-change");
        document.querySelectorAll(".cm-content table td, .cm-content table th").forEach((c) => {
          (c as HTMLElement).style.backgroundColor = "";
          (c as HTMLElement).style.color = "";
        });
      },
    });
    this.addCommand({ id: "undo-color-change", name: "Undo last color change", callback: () => this.undo() });
    this.addCommand({ id: "redo-color-change", name: "Redo last color change", callback: () => this.redo() });
    this.addCommand({ id: "refresh-table-colors", name: "Refresh table colors", callback: () => this.hardRefreshTableColors() });
    this.addCommand({ id: "show-latest-release-notes", name: "Show latest release notes",
      callback: async () => {
        try { new ChangelogModal(this.app, this).open(); }
        catch { new Notice("Unable to open changelog modal."); }
      },
    });
    this.addCommand({ id: "add-advanced-rule", name: "Add advanced rule",
      callback: async () => {
        if (!Array.isArray(this.settings.advancedRules)) this.settings.advancedRules = [];
        this.settings.advancedRules.push({ logic: "any", conditions: [], target: "cell", color: null, bg: null });
        await this.saveSettings();
        new AdvancedRuleModal(this.app, this, this.settings.advancedRules.length - 1).open();
        document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
      },
    });
    this.addCommand({ id: "manage-coloring-rules", name: "Manage coloring rules",
      callback: () => {
        (this.app as unknown as { setting: { open: () => void; openTabById?: (id: string) => void } }).setting.open();
        window.setTimeout(() => (this.app as unknown as { setting: { openTabById?: (id: string) => void } }).setting.openTabById?.("color-table-cell"), 250);
      },
    });
  }

  private _setupLivePreview(): void {
    this.applyColorsToAllEditors = () => {
      if (!this.settings.livePreviewColoring) {
        document.querySelectorAll(".cm-content table td, .cm-content table th").forEach((c) => {
          (c as HTMLElement).style.backgroundColor = "";
          (c as HTMLElement).style.color = "";
        });
        return;
      }
      const file = this.app.workspace.getActiveFile();
      if (!file) return;
      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll("table")) as HTMLElement[];
      document.querySelectorAll(".cm-content").forEach((editorEl) => {
        Array.from(editorEl.querySelectorAll("table")).forEach((table, localIdx) => {
          const globalIdx = allDocTables.indexOf(table as HTMLElement);
          this.processSingleTable(table as HTMLElement, globalIdx >= 0 ? globalIdx : localIdx, file.path, noteData);
        });
      });
    };
    type EditorEl = HTMLElement & { _ctcObserver?: MutationObserver; _ctcScrollListener?: boolean; _ctcScrollHandler?: () => void };
    Array.from(document.querySelectorAll(".cm-content")).forEach((ed) => {
      const edEl = ed as EditorEl;
      if (edEl._ctcObserver) return;
      const obs = new MutationObserver(() => {
        const file = this.app.workspace.getActiveFile();
        if (file && this.settings.livePreviewColoring) this.applyColorsToContainer(edEl, file.path);
      });
      obs.observe(edEl, { childList: true, subtree: true });
      edEl._ctcObserver = obs;
      if (!edEl._ctcScrollListener) {
        const restore = () => {
          edEl.querySelectorAll("[data-ctc-bg], [data-ctc-color]").forEach((cell) => {
            const c = cell as HTMLElement;
            if (c.hasAttribute("data-ctc-bg")) c.style.backgroundColor = c.getAttribute("data-ctc-bg")!;
            if (c.hasAttribute("data-ctc-color")) c.style.color = c.getAttribute("data-ctc-color")!;
          });
        };
        edEl._ctcScrollHandler = debounce(restore, 50);
        edEl.addEventListener("scroll", edEl._ctcScrollHandler);
        edEl._ctcScrollListener = true;
      }
    });
    window.setTimeout(() => this.applyColorsToActiveFile(), 200);
    if (!this._livePreviewObserver) {
      this._livePreviewObserver = new MutationObserver(() => this.applyColorsToActiveFile());
      document.querySelectorAll(".cm-content").forEach((el) =>
        this._livePreviewObserver!.observe(el, { childList: true, subtree: true }),
      );
    }
    this.registerEvent(this.app.workspace.on("file-open", () => this.applyColorsToActiveFile()));
    this.registerEvent(this.app.workspace.on("layout-change", () => this.applyColorsToActiveFile()));
    this.registerDomEvent(document, "focusin", (e) => {
      if ((e.target as Element)?.closest?.(".cm-content table")) this.applyColorsToActiveFile();
    });
    this.registerDomEvent(document, "input", (e) => {
      if ((e.target as Element)?.closest?.(".cm-content table"))
        window.setTimeout(() => this.applyColorsToActiveFile(), 30);
    });
    this.registerDomEvent(document, "pointerdown", (e) => {
      if ((e.target as Element)?.closest?.("td, th") && (e.target as Element)?.closest?.(".cm-content table"))
        window.setTimeout(() => this.applyColorsToAllEditors(), 10);
    });
  }

  private _setupObservers(): void {
    this._tablePreRenderer = new MutationObserver((mutations) => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;
      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll("table")) as HTMLElement[];
      mutations.forEach((mutation) => {
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType !== Node.ELEMENT_NODE) return;
          const el = node as HTMLElement;
          const tables: HTMLElement[] = [];
          if (el.matches?.("table")) tables.push(el);
          el.querySelectorAll?.("table").forEach((t) => tables.push(t as HTMLElement));
          tables.forEach((table) => {
            if (table.closest(".markdown-preview-view") && !table.hasAttribute("data-ctc-processed")) {
              const idx = allDocTables.indexOf(table);
              if (idx >= 0) this.processSingleTable(table, idx, file.path, noteData);
            }
          });
        });
      });
    });
    this._tablePreRenderer.observe(document.body, { childList: true, subtree: true });

    this.registerMarkdownPostProcessor((el, ctx) => {
      if (!el.closest(".markdown-preview-view")) return;
      this.applyColorsToContainer(el, ctx.sourcePath);
      try {
        if (this._containerObservers.has(el)) return;
        let observer: MutationObserver | null = null;
        let debounceId: number | null = null;
        const safeDisconnect = () => {
          try { observer?.disconnect(); observer = null; } catch { /* ignore */ }
          if (debounceId) window.clearTimeout(debounceId);
          this._containerObservers.delete(el);
        };
        observer = new MutationObserver((mutations) => {
          const hasTable = mutations.some((m) =>
            Array.from(m.addedNodes).some((n) =>
              (n as Element).tagName === "TABLE" || (n as Element).querySelector?.("table"),
            ),
          );
          if (hasTable) {
            if (debounceId) window.clearTimeout(debounceId);
            debounceId = window.setTimeout(() => {
              if (el.isConnected) this.applyColorsToContainer(el, ctx.sourcePath);
              else safeDisconnect();
            }, 80);
          }
        });
        observer.observe(el, { childList: true, subtree: true });
        this._containerObservers.set(el, { observer, safeDisconnect });
        const checker = window.setInterval(() => { if (!el.isConnected) { safeDisconnect(); window.clearInterval(checker); } }, 2000);
        window.setTimeout(() => { window.clearInterval(checker); safeDisconnect(); }, 30000);
      } catch { /* ignore */ }
    });

    try {
      this._readingViewObserver = new MutationObserver((mutations) => {
        const file = this.app.workspace.getActiveFile();
        if (!file) return;
        const hasNewTable = mutations.some((m) =>
          Array.from(m.addedNodes).some((n) => {
            const el = n as Element;
            return (el.matches?.("table") || (el.querySelectorAll?.("table").length ?? 0) > 0) && el.closest?.(".markdown-preview-view");
          }),
        );
        if (hasNewTable) this.applyColorsToActiveFile();
      });
      this._readingViewObserver.observe(document.body, { childList: true, subtree: true });
    } catch (e) { debugWarn("Failed to setup reading view observer:", e); }

    try {
      this._globalObserver = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType !== Node.ELEMENT_NODE) return;
            const el = node as HTMLElement;
            if (el.matches?.("[data-ctc-bg], [data-ctc-color]")) this.restoreColorsFromAttributes(el);
            el.querySelectorAll?.("[data-ctc-bg], [data-ctc-color]").forEach((c) => this.restoreColorsFromAttributes(c as HTMLElement));
          });
        });
      });
      this._globalObserver.observe(document.body, { childList: true, subtree: true });
    } catch (e) { debugWarn("Failed to setup global observer:", e); this._globalObserver = null; }
  }

  private _setupContextMenu(): void {
    this.registerDomEvent(document, "contextmenu", (evt) => {
      const target = evt.target as Element;
      const cell = target?.closest("td, th") as HTMLElement | null;
      const tableEl = target?.closest("table") as HTMLElement | null;
      if (!cell || !tableEl) return;
      if (!cell.closest(".markdown-preview-view") || cell.closest(".cm-content")) return;
      const menu = new Menu();
      menu.addItem((item) => item.setTitle("Color cell text").setIcon("palette").onClick(() => this.pickColor(cell, tableEl, "color")));
      menu.addItem((item) => item.setTitle("Color cell background").setIcon("droplet").onClick(() => this.pickColor(cell, tableEl, "bg")));
      if (this.settings.showColorRowInMenu) {
        menu.addItem((item) => item.setTitle("Color row text").setIcon("palette").onClick(() => this.pickColorForRow(cell, tableEl, "color")));
        menu.addItem((item) => item.setTitle("Color row background").setIcon("droplet").onClick(() => this.pickColorForRow(cell, tableEl, "bg")));
      }
      if (this.settings.showColorColumnInMenu) {
        menu.addItem((item) => item.setTitle("Color column text").setIcon("palette").onClick(() => this.pickColorForColumn(cell, tableEl, "color")));
        menu.addItem((item) => item.setTitle("Color column background").setIcon("droplet").onClick(() => this.pickColorForColumn(cell, tableEl, "bg")));
      }
      menu.addSeparator();
      menu.addItem((item) => item.setTitle("Reset cell").setIcon("eraser").onClick(() => this.resetCell(cell, tableEl)));
      menu.addItem((item) => item.setTitle("Reset row").setIcon("eraser").onClick(() => this.resetRow(cell, tableEl)));
      menu.addItem((item) => item.setTitle("Reset column").setIcon("eraser").onClick(() => this.resetColumn(cell, tableEl)));
      if (this.settings.showUndoRedoInMenu) {
        menu.addSeparator();
        menu.addItem((item) => item.setTitle("Undo").setIcon("undo").onClick(() => this.undo()));
        menu.addItem((item) => item.setTitle("Redo").setIcon("redo").onClick(() => this.redo()));
      }
      menu.showAtMouseEvent(evt);
    });
  }

  createStatusBarIcon(): void {
    if (this.statusBarRefresh) return;
    this.statusBarRefresh = this.addStatusBarItem();
    this.statusBarRefresh.setText("⟳");
    this.statusBarRefresh.setAttr("title", "Refresh table colors");
    this.statusBarRefresh.addClass("ctc-refresh-table-color");
    this.statusBarRefresh.addEventListener("click", () => this.hardRefreshTableColors());
  }

  removeStatusBarIcon(): void {
    if (this.statusBarRefresh) {
      this.statusBarRefresh.remove();
      this.statusBarRefresh = undefined;
    }
  }

  hardRefreshTableColors(): void {
    document.querySelectorAll("table td, table th").forEach((cell) => {
      const el = cell as HTMLElement;
      if (!el.hasAttribute("data-ctc-manual")) {
        el.style.backgroundColor = "";
        el.style.color = "";
      }
    });
    this.applyColorsToActiveFile();
    if (this.settings.livePreviewColoring) window.setTimeout(() => this.applyColorsToAllEditors(), 50);
  }

  onunload(): void {
    try { this._livePreviewObserver?.disconnect(); } catch { /* ignore */ }
    try { this._tablePreRenderer?.disconnect(); } catch { /* ignore */ }
    try { this._readingViewObserver?.disconnect(); } catch { /* ignore */ }
    try { this._globalObserver?.disconnect(); } catch { /* ignore */ }
    try {
      this._containerObservers?.forEach((obs: unknown) => {
        (obs as { safeDisconnect?: () => void })?.safeDisconnect?.();
      });
    } catch { /* ignore */ }
    if (this._readingModeChecker) window.clearInterval(this._readingModeChecker);
  }

  async loadSettings(): Promise<void> {
    const data = await this.loadData();
    this.settings = Object.assign({}, DEFAULT_SETTINGS, data?.settings || {});
    // Migrate old "rules" format to "coloringRules"
    if (Array.isArray(data?.settings?.rules) && data.settings.rules.length > 0) {
      if (!Array.isArray(this.settings.coloringRules) || this.settings.coloringRules.length === 0) {
        this.settings.coloringRules = data.settings.rules.map((r: { regex?: boolean; match: string; color?: string; bg?: string }) => ({
          target: "cell", when: "theCell",
          match: r.regex ? "isRegex" : "contains",
          value: r.match, color: r.color || null, bg: r.bg || null,
        }));
        this._settingsMigrated = true;
      }
    }
    if (Array.isArray(this.settings.presetColors)) {
      this.settings.presetColors = this.settings.presetColors.map((pc) =>
        typeof pc === "string" ? { name: "", color: pc } : pc,
      );
    } else {
      this.settings.presetColors = [];
    }
  }

  async saveSettings(): Promise<void> {
    const dataToSave = { settings: this.settings, cellData: this.cellData };
    if (this._settingsMigrated) {
      delete (dataToSave.settings as unknown as Record<string, unknown>).rules;
      this._settingsMigrated = false;
    }
    await this.saveData(dataToSave);
  }

  async loadDataSettings(): Promise<void> { await this.loadSettings(); }
  async loadDataColors(): Promise<void> { const raw = (await this.loadData()) || {}; this.cellData = normalizeCellData(raw) || {}; }
  async saveDataSettings(): Promise<void> { await this.saveSettings(); }
  async saveDataColors(): Promise<void> {
    await this.saveData({ settings: this.settings, cellData: this.cellData });
  }

  async fetchAllReleases(): Promise<Array<{ name: string; tag_name: string; body: string; published_at: string }>> {
    const allReleases: Array<{ name: string; tag_name: string; body: string; published_at: string }> = [];
    let page = 1;
    let hasMore = true;
    while (hasMore) {
      const url = `https://api.github.com/repos/Kazi-Aidah/color-table-cells/releases?page=${page}&per_page=100`;
      try {
        let data = null;
        try {
          const res = await (await import("obsidian")).requestUrl({ url, headers: { Accept: "application/vnd.github.v3+json", "User-Agent": "Obsidian-Color-Table-Cells" } });
          data = res.json || (res.text ? JSON.parse(res.text) : null);
        } catch { /* fallback to fetch */ }
        if (!data) {
          const r = await fetch(url, { headers: { Accept: "application/vnd.github.v3+json" } });
          if (!r.ok) throw new Error("Network error");
          data = await r.json();
        }
        if (!Array.isArray(data) || data.length === 0) { hasMore = false; }
        else { allReleases.push(...data); if (data.length < 100) hasMore = false; else page++; }
      } catch { hasMore = false; }
    }
    return allReleases;
  }

  updateRecentColor(color: string): void {
    if (!color) return;
    const list = Array.isArray(this.settings.recentColors) ? [...this.settings.recentColors] : [];
    const idx = list.findIndex((c) => c.toUpperCase() === color.toUpperCase());
    if (idx !== -1) list.splice(idx, 1);
    list.unshift(color);
    this.settings.recentColors = list.slice(0, 10);
  }

  getGlobalTableIndex(tableEl: HTMLElement): number {
    const stored = tableEl.getAttribute("data-ctc-index");
    if (stored !== null) return parseInt(stored, 10);
    if (tableEl.closest(".markdown-preview-view")) {
      const view = tableEl.closest(".markdown-preview-view")!;
      return Array.from(view.querySelectorAll("table") as NodeListOf<HTMLElement>).indexOf(tableEl);
    }
    if (tableEl.closest(".cm-content")) {
      const editor = tableEl.closest(".cm-content") || tableEl.closest(".cm-editor");
      if (editor) return Array.from(editor.querySelectorAll("table") as NodeListOf<HTMLElement>).indexOf(tableEl);
    }
    return Array.from(document.querySelectorAll("table") as NodeListOf<HTMLElement>).indexOf(tableEl);
  }

  async pickColor(cell: HTMLElement, tableEl: HTMLElement, type: string): Promise<void> {
    const menu = new ColorPickerMenu(this, async (pickedColor) => {
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      const tableIndex = this.getGlobalTableIndex(tableEl);
      const rowIndex = Array.from(tableEl.querySelectorAll("tr") as NodeListOf<HTMLElement>).indexOf(cell.closest("tr") as HTMLElement);
      const colIndex = Array.from(cell.closest("tr")!.querySelectorAll("td, th")).indexOf(cell);
      const oldColors = this.cellData[fileId]?.[`table_${tableIndex}`]?.[`row_${rowIndex}`]?.[`col_${colIndex}`];
      if (!this.cellData[fileId]) this.cellData[fileId] = {};
      const noteData = this.cellData[fileId];
      const tableKey = `table_${tableIndex}`;
      if (!noteData[tableKey]) noteData[tableKey] = {};
      const rowKey = `row_${rowIndex}`;
      if (!noteData[tableKey][rowKey]) noteData[tableKey][rowKey] = {};
      const newColors = { ...(noteData[tableKey][rowKey][`col_${colIndex}`] || {}), [type]: pickedColor } as CellColorData;
      noteData[tableKey][rowKey][`col_${colIndex}`] = newColors;
      this.addToUndoStack(this.createSnapshot("cell_color", fileId, tableIndex, { row: rowIndex, col: colIndex }, oldColors, newColors));
      this.updateRecentColor(pickedColor);
      await this.saveDataColors();
    }, null, cell);
    menu._cell = cell;
    menu._type = type;
    menu.open();
  }

  async pickColorForRow(cell: HTMLElement, tableEl: HTMLElement, type: string): Promise<void> {
    const row = cell.closest("tr") as HTMLElement;
    const rowCells = Array.from(row.querySelectorAll("td, th")) as HTMLElement[];
    const menu = new ColorPickerMenu(this, async (pickedColor) => {
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      const tableIndex = this.getGlobalTableIndex(tableEl);
      const rowIndex = Array.from(tableEl.querySelectorAll("tr") as NodeListOf<HTMLElement>).indexOf(row);
      const oldColors = this.cellData[fileId]?.[`table_${tableIndex}`]?.[`row_${rowIndex}`];
      if (!this.cellData[fileId]) this.cellData[fileId] = {};
      const noteData = this.cellData[fileId];
      const tableKey = `table_${tableIndex}`;
      if (!noteData[tableKey]) noteData[tableKey] = {};
      const rowKey = `row_${rowIndex}`;
      if (!noteData[tableKey][rowKey]) noteData[tableKey][rowKey] = {};
      const newColors: Record<string, CellColorData> = {};
      rowCells.forEach((rowCell, colIndex) => {
        const colKey = `col_${colIndex}`;
        newColors[colKey] = { ...(noteData[tableKey][rowKey][colKey] || {}), [type]: pickedColor } as CellColorData;
        noteData[tableKey][rowKey][colKey] = newColors[colKey];
      });
      this.addToUndoStack(this.createSnapshot("row_color", fileId, tableIndex, { row: rowIndex }, oldColors, newColors));
      this.updateRecentColor(pickedColor);
      await this.saveDataColors();
    }, null, cell);
    menu._cells = rowCells;
    menu._type = type;
    menu.open();
  }

  async pickColorForColumn(cell: HTMLElement, tableEl: HTMLElement, type: string): Promise<void> {
    const colIndex = Array.from(cell.closest("tr")!.querySelectorAll("td, th")).indexOf(cell);
    const columnCells: HTMLElement[] = [];
    tableEl.querySelectorAll("tr").forEach((row) => {
      const cells = row.querySelectorAll("td, th");
      if (colIndex < cells.length) columnCells.push(cells[colIndex] as HTMLElement);
    });
    const menu = new ColorPickerMenu(this, async (pickedColor) => {
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      const tableIndex = this.getGlobalTableIndex(tableEl);
      if (!this.cellData[fileId]) this.cellData[fileId] = {};
      const noteData = this.cellData[fileId];
      const tableKey = `table_${tableIndex}`;
      if (!noteData[tableKey]) noteData[tableKey] = {};
      const newColors: Record<string, Record<string, CellColorData>> = {};
      tableEl.querySelectorAll("tr").forEach((row, rowIndex) => {
        const cells = row.querySelectorAll("td, th");
        if (colIndex < cells.length) {
          const rowKey = `row_${rowIndex}`;
          if (!noteData[tableKey][rowKey]) noteData[tableKey][rowKey] = {};
          const colKey = `col_${colIndex}`;
          if (!newColors[rowKey]) newColors[rowKey] = {};
          newColors[rowKey][colKey] = { ...(noteData[tableKey][rowKey][colKey] || {}), [type]: pickedColor } as CellColorData;
          noteData[tableKey][rowKey][colKey] = newColors[rowKey][colKey];
        }
      });
      this.addToUndoStack(this.createSnapshot("column_color", fileId, tableIndex, { col: colIndex }, {}, newColors));
      this.updateRecentColor(pickedColor);
      await this.saveDataColors();
    }, null, cell);
    menu._cells = columnCells;
    menu._type = type;
    menu.open();
  }

  async resetCell(cell: HTMLElement, tableEl: HTMLElement): Promise<void> {
    cell.style.backgroundColor = "";
    cell.style.color = "";
    cell.removeAttribute("data-ctc-bg");
    cell.removeAttribute("data-ctc-color");
    cell.removeAttribute("data-ctc-manual");
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;
    const tableIndex = this.getGlobalTableIndex(tableEl);
    const rowIndex = Array.from(tableEl.querySelectorAll("tr") as NodeListOf<HTMLElement>).indexOf(cell.closest("tr") as HTMLElement);
    const colIndex = Array.from(cell.closest("tr")!.querySelectorAll("td, th")).indexOf(cell);
    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (noteData?.[tableKey]?.[`row_${rowIndex}`]) {
      delete noteData[tableKey][`row_${rowIndex}`][`col_${colIndex}`];
      await this.saveDataColors();
    }
    window.setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async resetRow(cell: HTMLElement, tableEl: HTMLElement): Promise<void> {
    const row = cell.closest("tr") as HTMLElement;
    row.querySelectorAll("td, th").forEach((c) => {
      const el = c as HTMLElement;
      el.style.backgroundColor = "";
      el.style.color = "";
      el.removeAttribute("data-ctc-bg");
      el.removeAttribute("data-ctc-color");
      el.removeAttribute("data-ctc-manual");
    });
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;
    const tableIndex = this.getGlobalTableIndex(tableEl);
    const rowIndex = Array.from(tableEl.querySelectorAll("tr") as NodeListOf<HTMLElement>).indexOf(row);
    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (noteData?.[tableKey]) {
      delete noteData[tableKey][`row_${rowIndex}`];
      await this.saveDataColors();
    }
    window.setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async resetColumn(cell: HTMLElement, tableEl: HTMLElement): Promise<void> {
    const colIndex = Array.from(cell.closest("tr")!.querySelectorAll("td, th")).indexOf(cell);
    tableEl.querySelectorAll("tr").forEach((row) => {
      const cells = row.querySelectorAll("td, th");
      if (colIndex < cells.length) {
        const c = cells[colIndex] as HTMLElement;
        c.style.backgroundColor = "";
        c.style.color = "";
        c.removeAttribute("data-ctc-bg");
        c.removeAttribute("data-ctc-color");
        c.removeAttribute("data-ctc-manual");
      }
    });
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;
    const tableIndex = this.getGlobalTableIndex(tableEl);
    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (noteData?.[tableKey]) {
      Object.keys(noteData[tableKey]).forEach((rowKey) => {
        delete noteData[tableKey][rowKey][`col_${colIndex}`];
      });
      await this.saveDataColors();
    }
    window.setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  createSnapshot(
    type: ColorSnapshot["type"],
    fileId: string,
    tableIndex: number,
    coords: { row?: number; col?: number },
    oldColors: unknown,
    newColors: unknown,
  ): ColorSnapshot {
    return { type, fileId, tableIndex, coords, oldColors: oldColors as ColorSnapshot["oldColors"], newColors: newColors as ColorSnapshot["newColors"] };
  }

  addToUndoStack(snapshot: ColorSnapshot): void {
    this.undoStack.push(snapshot);
    if (this.undoStack.length > this.maxStackSize) this.undoStack.shift();
    this.redoStack = [];
    if (this.settings.persistUndoHistory) this.saveUndoRedoStacks();
  }

  async undo(): Promise<void> {
    if (!this.undoStack.length) { new Notice("Nothing to undo"); return; }
    const snapshot = this.undoStack.pop()!;
    this.redoStack.push(snapshot);
    await this._applySnapshot(snapshot, true);
    if (this.settings.persistUndoHistory) await this.saveUndoRedoStacks();
  }

  async redo(): Promise<void> {
    if (!this.redoStack.length) { new Notice("Nothing to redo"); return; }
    const snapshot = this.redoStack.pop()!;
    this.undoStack.push(snapshot);
    await this._applySnapshot(snapshot, false);
    if (this.settings.persistUndoHistory) await this.saveUndoRedoStacks();
  }

  private async _applySnapshot(snapshot: ColorSnapshot, isUndo: boolean): Promise<void> {
    const { fileId, tableIndex, coords, oldColors, newColors } = snapshot;
    if (!this.cellData[fileId]) this.cellData[fileId] = {};
    const tableKey = `table_${tableIndex}`;
    if (!this.cellData[fileId][tableKey]) this.cellData[fileId][tableKey] = {};
    const colors = isUndo ? oldColors : newColors;
    if (snapshot.type === "cell_color" && coords.row !== undefined && coords.col !== undefined) {
      const rowKey = `row_${coords.row}`;
      if (!this.cellData[fileId][tableKey][rowKey]) this.cellData[fileId][tableKey][rowKey] = {};
      if (colors) this.cellData[fileId][tableKey][rowKey][`col_${coords.col}`] = colors as CellColorData;
      else delete this.cellData[fileId][tableKey][rowKey][`col_${coords.col}`];
    }
    await this.saveDataColors();
    this.applyColorsToActiveFile();
  }

  async saveUndoRedoStacks(): Promise<void> {
    try {
      const data = (await this.loadData()) || {};
      data.undoStack = this.undoStack;
      data.redoStack = this.redoStack;
      await this.saveData(data);
    } catch (e) { debugWarn("Error saving undo/redo stacks:", e); }
  }

  async loadUndoRedoStacks(): Promise<void> {
    try {
      const data = (await this.loadData()) || {};
      this.undoStack = Array.isArray(data.undoStack) ? data.undoStack : [];
      this.redoStack = Array.isArray(data.redoStack) ? data.redoStack : [];
    } catch (e) { debugWarn("Error loading undo/redo stacks:", e); }
  }

  getCellText(cell: HTMLElement): string {
    return getCellText(cell);
  }

  isElementVisible(element: HTMLElement): boolean {
    if (!element?.isConnected) return false;
    const style = window.getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden") return false;
    const rect = element.getBoundingClientRect();
    return rect.width > 0 || rect.height > 0;
  }

  evaluateMatch(text: string, rule: { match: string; value: string | number | null }): boolean {
    return evaluateMatch(text, rule as Parameters<typeof evaluateMatch>[1]);
  }

  applyRulesToCell(cell: HTMLElement, cellText: string, colorData: CellColorData): void {
    if (colorData.bg) cell.style.backgroundColor = colorData.bg;
    if (colorData.color) cell.style.color = colorData.color;
  }

  applyColoringRulesToTable(tableEl: HTMLElement): void {
    const rules = Array.isArray(this.settings.coloringRules) ? this.settings.coloringRules : [];
    if (!rules.length) return;
    const rows = Array.from(tableEl.querySelectorAll("tr")) as HTMLElement[];
    const texts = rows.map((row) => Array.from(row.querySelectorAll("td, th")).map((c) => getCellText(c as HTMLElement)));
    const maxCols = Math.max(0, ...texts.map((r) => r.length));
    const headerRowIndex = rows.findIndex((r) => r.querySelector("th"));
    const hdr = headerRowIndex >= 0 ? headerRowIndex : 0;
    const firstDataRowIndex = rows.findIndex((r) => r.querySelector("td"));
    const fdr = firstDataRowIndex >= 0 ? firstDataRowIndex : 0;
    const getCell = (r: number, c: number) => {
      const row = rows[r];
      if (!row) return null;
      const cells = Array.from(row.querySelectorAll("td, th")) as HTMLElement[];
      return cells[c] || null;
    };
    const applyCellStyle = (cell: HTMLElement | null, rule: { bg?: string | null; color?: string | null }) => {
      if (!cell) return;
      if (cell.hasAttribute("data-ctc-manual")) return;
      const isHeader = cell.tagName === "TH";
      if (!isHeader && (cell.style.backgroundColor || cell.style.color)) return;
      if (rule.bg) cell.style.backgroundColor = rule.bg;
      if (rule.color) cell.style.color = rule.color;
    };
    for (const rule of rules) {
      if (!rule?.target || !rule.match) continue;
      if (rule.target === "cell") {
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < (texts[r]?.length || 0); c++) {
            if ((rule.when || "theCell") === "theCell" && this.evaluateMatch(texts[r][c], rule))
              applyCellStyle(getCell(r, c), rule);
          }
        }
      } else if (rule.target === "row") {
        const candidateRows = rule.when === "firstRow" ? [fdr] : Array.from({ length: rows.length }, (_, i) => i);
        for (const r of candidateRows) {
          const rowTexts = texts[r] || [];
          let cond = false;
          if (rule.when === "allCell") cond = rowTexts.length > 0 && rowTexts.every((t) => this.evaluateMatch(t, rule));
          else if (rule.when === "noCell") cond = rowTexts.every((t) => !this.evaluateMatch(t, rule));
          else cond = rowTexts.some((t) => this.evaluateMatch(t, rule));
          if (cond) Array.from(rows[r].querySelectorAll("td, th")).forEach((c) => applyCellStyle(c as HTMLElement, rule));
        }
      } else if (rule.target === "column") {
        for (let c = 0; c < maxCols; c++) {
          let cond = false;
          if (rule.when === "columnHeader") {
            cond = this.evaluateMatch(texts[hdr]?.[c] ?? "", rule);
          } else {
            const colTexts = Array.from({ length: rows.length }, (_, r) => texts[r]?.[c]).filter((t) => t !== undefined) as string[];
            if (rule.when === "allCell") cond = colTexts.length > 0 && colTexts.every((t) => this.evaluateMatch(t, rule));
            else if (rule.when === "noCell") cond = colTexts.every((t) => !this.evaluateMatch(t, rule));
            else cond = colTexts.some((t) => this.evaluateMatch(t, rule));
          }
          if (cond) for (let r = 0; r < rows.length; r++) applyCellStyle(getCell(r, c), rule);
        }
      }
    }
  }

  applyAdvancedRulesToTable(tableEl: HTMLElement): void {
    const adv = Array.isArray(this.settings.advancedRules) ? this.settings.advancedRules : [];
    if (!adv.length) return;
    const rows = Array.from(tableEl.querySelectorAll("tr")) as HTMLElement[];
    const texts = rows.map((row) => Array.from(row.querySelectorAll("td, th")).map((c) => getCellText(c as HTMLElement)));
    const maxCols = Math.max(0, ...texts.map((r) => r.length));
    const headerRowIndex = rows.findIndex((r) => r.querySelector("th"));
    const hdr = headerRowIndex >= 0 ? headerRowIndex : 0;
    const fdr = rows.findIndex((r) => r.querySelector("td"));
    const getCell = (r: number, c: number) => {
      const row = rows[r];
      if (!row) return null;
      return (Array.from(row.querySelectorAll("td, th")) as HTMLElement[])[c] || null;
    };
    const evalCond = (r: number, c: number, cond: { when: string; match: string; value: string }) => {
      if (cond.when === "columnHeader") return this.evaluateMatch(texts[hdr]?.[c] ?? "", cond);
      if (cond.when === "row") return (texts[r] || []).some((t) => this.evaluateMatch(t, cond));
      const cellMatch = this.evaluateMatch(texts[r]?.[c] ?? "", cond);
      if (cond.when === "allCell" || cond.when === "noCell") return cellMatch;
      return cellMatch;
    };
    const evalFlags = (flags: boolean[], logic: string) => {
      if (logic === "all") return flags.every(Boolean);
      if (logic === "none") return flags.every((f) => !f);
      return flags.some(Boolean);
    };
    for (const rule of adv) {
      const { logic = "any", target = "cell", color, bg } = rule;
      if (!bg && !color) continue;
      const conditions = Array.isArray(rule.conditions) ? rule.conditions : [];
      if (!conditions.length) continue;
      const applyCell = (cell: HTMLElement | null) => {
        if (!cell || cell.hasAttribute("data-ctc-manual")) return;
        if (cell.tagName !== "TH" && (cell.style.backgroundColor || cell.style.color)) return;
        if (bg) cell.style.backgroundColor = bg;
        if (color) cell.style.color = color;
      };
      if (target === "row") {
        const allHeader = conditions.every((c) => c.when === "columnHeader");
        const candidateRows = allHeader ? [hdr] : Array.from({ length: rows.length }, (_, i) => i);
        for (const r of candidateRows) {
          const flags = conditions.map((cond) => {
            if (cond.when === "columnHeader") {
              for (let c = 0; c < maxCols; c++) if (this.evaluateMatch(texts[hdr]?.[c] ?? "", cond)) return true;
              return false;
            }
            return (texts[r] || []).some((t) => this.evaluateMatch(t, cond));
          });
          if (evalFlags(flags, logic)) Array.from(rows[r].querySelectorAll("td, th")).forEach((c) => applyCell(c as HTMLElement));
        }
      } else if (target === "column") {
        for (let c = 0; c < maxCols; c++) {
          const flags = conditions.map((cond) => {
            if (cond.when === "columnHeader") return this.evaluateMatch(texts[hdr]?.[c] ?? "", cond);
            const colCells = Array.from({ length: rows.length }, (_, r) => this.evaluateMatch(texts[r]?.[c] ?? "", cond));
            if (cond.when === "allCell") return colCells.every(Boolean);
            if (cond.when === "noCell") return colCells.every((f) => !f);
            return colCells.some(Boolean);
          });
          if (evalFlags(flags, logic)) for (let r = 0; r < rows.length; r++) applyCell(getCell(r, c));
        }
      } else {
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < (texts[r]?.length || 0); c++) {
            const flags = conditions.map((cond) => evalCond(r, c, cond as { when: string; match: string; value: string }));
            if (evalFlags(flags, logic)) applyCell(getCell(r, c));
          }
        }
      }
    }
  }

  restoreColorsFromAttributes(element: HTMLElement): void {
    if (!element?.style) return;
    if (element.hasAttribute("data-ctc-bg")) {
      const bg = element.getAttribute("data-ctc-bg")!;
      if (bg && bg !== element.style.backgroundColor) element.style.backgroundColor = bg;
    }
    if (element.hasAttribute("data-ctc-color")) {
      const color = element.getAttribute("data-ctc-color")!;
      if (color && color !== element.style.color) element.style.color = color;
    }
  }

  applyColorsToContainer(container: HTMLElement, filePath: string): void {
    const hasClosest = typeof container.closest === "function";
    const inPreview = hasClosest && !!container.closest(".markdown-preview-view");
    const inEditor = hasClosest && !!(container.closest(".cm-content") || container.closest(".cm-editor") || container.closest(".cm-scroller"));
    if (!inPreview && (!this.settings.livePreviewColoring || !inEditor)) {
      let p = container?.parentElement;
      let found = false;
      while (p) {
        if (p.classList?.contains("markdown-preview-view")) { found = true; break; }
        if (this.settings.livePreviewColoring && (p.classList?.contains("cm-content") || p.classList?.contains("cm-editor"))) { found = true; break; }
        p = p.parentElement;
      }
      if (!found && !(this.settings.livePreviewColoring && container.classList?.contains("cm-content"))) return;
    }
    if (inEditor && !this.settings.livePreviewColoring) return;
    if (!inPreview && !inEditor && !container.classList?.contains("cm-content")) return;
    const now = Date.now();
    if (now - this._lastApplyCall < 100) return;
    this._lastApplyCall = now;
    let target = container;
    if (inPreview) {
      const readingView = container.closest(".markdown-preview-view") as HTMLElement;
      if (readingView) {
        const allTablesInView = Array.from(readingView.querySelectorAll("table"));
        if (allTablesInView.length > 0 && allTablesInView.some((t) => !container.contains(t))) target = readingView;
      }
    }
    const tables = Array.from(target.querySelectorAll("table")) as HTMLElement[];
    if (!tables.length) return;
    const noteData = this.cellData[filePath] || {};
    const allTables = Array.from(document.querySelectorAll("table")) as HTMLElement[];
    tables.forEach((tableEl) => {
      const globalIdx = allTables.indexOf(tableEl);
      const idx = globalIdx === -1 ? Array.from(target.querySelectorAll("table") as NodeListOf<HTMLElement>).indexOf(tableEl) : globalIdx;
      this.processSingleTable(tableEl, idx, filePath, noteData);
    });
    try {
      const prev = this._appliedContainers.get(target) || 0;
      const delays = [100, 200, 400, 800];
      if (prev < delays.length) {
        this._appliedContainers.set(target, prev + 1);
        window.setTimeout(() => { if (target.isConnected) this.applyColorsToContainer(target, filePath); }, delays[prev]);
      }
    } catch { /* ignore */ }
  }

  applyColorsToActiveFile(): void {
    const file = this.app.workspace.getActiveFile();
    if (!file) return;
    const noteData = this.cellData[file.path] || {};
    let previewViews: HTMLElement[] = [];
    try {
      const leaves = this.app.workspace.getLeavesOfType("markdown");
      const activeContainers = leaves
        .filter((l) => (l.view as unknown as { file?: { path: string } })?.file?.path === file.path)
        .map((l) => (l.view as unknown as { containerEl?: HTMLElement; contentEl?: HTMLElement })?.containerEl || (l.view as unknown as { contentEl?: HTMLElement })?.contentEl)
        .filter(Boolean) as HTMLElement[];
      activeContainers.forEach((container) => {
        previewViews.push(...Array.from(container.querySelectorAll(".markdown-preview-view")) as HTMLElement[]);
      });
    } catch {
      previewViews = Array.from(document.querySelectorAll(".markdown-preview-view")) as HTMLElement[];
    }
    let fileTableIndex = 0;
    const fileTableMap = new Map<HTMLElement, number>();
    previewViews.forEach((view) => {
      view.querySelectorAll("table").forEach((table) => {
        if (!fileTableMap.has(table as HTMLElement)) {
          fileTableMap.set(table as HTMLElement, fileTableIndex);
          (table as HTMLElement).setAttribute("data-ctc-index", String(fileTableIndex));
          fileTableIndex++;
        }
      });
    });
    if (this.settings.livePreviewColoring) {
      document.querySelectorAll(".cm-content table").forEach((table) => {
        if (!fileTableMap.has(table as HTMLElement)) {
          fileTableMap.set(table as HTMLElement, fileTableIndex);
          (table as HTMLElement).setAttribute("data-ctc-index", String(fileTableIndex));
          fileTableIndex++;
        }
      });
    }
    if (fileTableIndex === 0) { window.setTimeout(() => this.applyColorsToActiveFile(), 100); return; }
    previewViews.forEach((view) => {
      view.querySelectorAll("td, th").forEach((cell) => {
        if (!(cell as HTMLElement).hasAttribute("data-ctc-manual")) {
          (cell as HTMLElement).style.backgroundColor = "";
          (cell as HTMLElement).style.color = "";
        }
      });
    });
    if (this.settings.livePreviewColoring) {
      document.querySelectorAll(".cm-content table td, .cm-content table th").forEach((cell) => {
        if (!(cell as HTMLElement).hasAttribute("data-ctc-manual")) {
          (cell as HTMLElement).style.backgroundColor = "";
          (cell as HTMLElement).style.color = "";
        }
      });
    }
    previewViews.forEach((view) => {
      if (view.isConnected) {
        Array.from(view.querySelectorAll("table")).forEach((table) => {
          const idx = fileTableMap.get(table as HTMLElement);
          if (idx !== undefined) this.processSingleTable(table as HTMLElement, idx, file.path, noteData);
        });
      }
    });
    if (this.settings.livePreviewColoring) {
      document.querySelectorAll(".cm-content").forEach((editor) => {
        if ((editor as HTMLElement).isConnected) {
          Array.from(editor.querySelectorAll("table")).forEach((table) => {
            const idx = fileTableMap.get(table as HTMLElement);
            if (idx !== undefined) this.processSingleTable(table as HTMLElement, idx, file.path, noteData);
          });
        }
      });
    }
  }

  getTableSignature(table: HTMLElement): string {
    const rows = table.querySelectorAll("tr").length;
    const firstRow = table.querySelector("tr");
    const cols = firstRow ? firstRow.querySelectorAll("td, th").length : 0;
    return `${rows}x${cols}`;
  }

  getTableTextContent(table: HTMLElement): string {
    return Array.from(table.querySelectorAll("tr"))
      .map((row) => Array.from(row.querySelectorAll("td, th")).map((c) => getCellText(c as HTMLElement).trim()).join("|"))
      .join("\n");
  }

  processSingleTable(
    tableEl: HTMLElement,
    tableIndex: number,
    filePath: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    noteData: Record<string, any>,
  ): number {
    if (!tableEl.hasAttribute("data-ctc-processed")) {
      tableEl.setAttribute("data-ctc-processed", "true");
      tableEl.setAttribute("data-ctc-index", String(tableIndex));
      tableEl.setAttribute("data-ctc-file", filePath);
    }
    const tableKey = `table_${tableIndex}`;
    const tableColors = noteData[tableKey] || {};
    let coloredCount = 0;
    const tableId = `${filePath}:${tableIndex}`;
    const manualColorData: Record<string, CellColorData> = {};
    const cellsWithRuleColor = new Set<string>();

    // Store manual color data
    Array.from(tableEl.querySelectorAll("tr")).forEach((tr, rIdx) => {
      const rowKey = `row_${rIdx}`;
      const rowColors = tableColors[rowKey] || {};
      Array.from(tr.querySelectorAll("td, th")).forEach((cell, cIdx) => {
        const colorData = rowColors[`col_${cIdx}`];
        if (colorData) {
          const cellKey = `${rIdx}_${cIdx}`;
          manualColorData[cellKey] = colorData;
          const el = cell as HTMLElement;
          if (colorData.bg) el.setAttribute("data-ctc-bg", colorData.bg as string);
          if (colorData.color) el.setAttribute("data-ctc-color", colorData.color as string);
          el.setAttribute("data-ctc-manual", "true");
          el.setAttribute("data-ctc-table-id", tableId);
          el.setAttribute("data-ctc-row", String(rIdx));
          el.setAttribute("data-ctc-col", String(cIdx));
        }
      });
    });

    // Apply rules first
    this.applyColoringRulesToTable(tableEl);
    this.applyAdvancedRulesToTable(tableEl);

    // Track cells that got rule colors
    Array.from(tableEl.querySelectorAll("tr")).forEach((tr, rIdx) => {
      Array.from(tr.querySelectorAll("td, th")).forEach((cell, cIdx) => {
        const el = cell as HTMLElement;
        if (el.style.backgroundColor || el.style.color) cellsWithRuleColor.add(`${rIdx}_${cIdx}`);
      });
    });

    // Apply manual colors only to cells without rule colors
    Array.from(tableEl.querySelectorAll("tr")).forEach((tr, rIdx) => {
      Array.from(tr.querySelectorAll("td, th")).forEach((cell, cIdx) => {
        const cellKey = `${rIdx}_${cIdx}`;
        const colorData = manualColorData[cellKey];
        if (colorData && !cellsWithRuleColor.has(cellKey)) {
          coloredCount++;
          const el = cell as HTMLElement;
          if (colorData.bg) el.style.backgroundColor = colorData.bg as string;
          if (colorData.color) el.style.color = colorData.color as string;
        }
      });
    });

    tableEl.setAttribute("data-ctc-last-processed", String(Date.now()));
    return coloredCount;
  }

  setupReadingViewScrollListener(): void {
    const handleScroll = debounce(() => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;
      document.querySelectorAll(".markdown-preview-view").forEach((view) => {
        if ((view as HTMLElement).isConnected) this.applyColorsToContainer(view as HTMLElement, file.path);
      });
    }, 150);
    document.querySelectorAll(".markdown-preview-view").forEach((view) => {
      view.addEventListener("scroll", handleScroll, { passive: true });
    });
    this.registerDomEvent(document, "scroll", handleScroll, { capture: true });
  }

  startReadingModeTableChecker(): void {
    this._readingModeChecker = window.setInterval(() => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;
      const readingViews = document.querySelectorAll(".markdown-preview-view");
      if (!readingViews.length) return;
      readingViews.forEach((view) => {
        const tables = view.querySelectorAll("table");
        tables.forEach((table) => {
          if (!table.hasAttribute("data-ctc-processed")) {
            const allTables = Array.from(document.querySelectorAll("table") as NodeListOf<HTMLElement>);
            const idx = allTables.indexOf(table as HTMLElement);
            if (idx >= 0) {
              const noteData = this.cellData[file.path] || {};
              this.processSingleTable(table as HTMLElement, idx, file.path, noteData);
            }
          }
        });
      });
    }, 2000);
  }
}
