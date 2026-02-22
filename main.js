const {
  Plugin,
  PluginSettingTab,
  Setting,
  Menu,
  ButtonComponent,
  Modal,
  setIcon,
  debounce,
} = require("obsidian");

// Debug configuration - make it a getter so changes are reflected dynamically
let IS_DEVELOPMENT = false;
const debugLog = (...args) =>
  IS_DEVELOPMENT && console.log("[CTC-DEBUG]", ...args);
const debugWarn = (...args) =>
  IS_DEVELOPMENT && console.warn("[CTC-WARN]", ...args);

// Allow toggling debug mode from console: window.setDebugMode(true/false)
if (typeof window !== "undefined") {
  window.setDebugMode = (value) => {
    IS_DEVELOPMENT = value;
    console.log(`[CTC] Debug mode ${value ? "enabled" : "disabled"}`);
  };
}

module.exports = class TableColorPlugin extends Plugin {
  undoStack = [];
  redoStack = [];
  maxStackSize = 50;
  _changelogCommandRegistered = false;

  async onload() {
    // COMMAND PALETTE COMMANDS
    this.addCommand({
      id: "enable-live-preview-coloring",
      name: "Enable live preview table coloring",
      callback: async () => {
        IS_DEVELOPMENT && console.log("[CTC] Enabling live preview coloring");
        this.settings.livePreviewColoring = true;
        await this.saveSettings();
        if (
          this.app.workspace &&
          typeof this.app.workspace.trigger === "function"
        ) {
          this.app.workspace.trigger("layout-change");
        }
        if (typeof this.applyColorsToAllEditors === "function") {
          window.setTimeout(() => this.applyColorsToAllEditors(), 0);
        }
      },
    });
    this.addCommand({
      id: "disable-live-preview-coloring",
      name: "Disable live preview table coloring",
      callback: async () => {
        this.settings.livePreviewColoring = false;
        await this.saveSettings();
        if (
          this.app.workspace &&
          typeof this.app.workspace.trigger === "function"
        ) {
          this.app.workspace.trigger("layout-change");
        }
        // Remove colors from all .cm-content tables
        document.querySelectorAll(".cm-content table").forEach((table) => {
          table.querySelectorAll("td, th").forEach((cell) => {
            cell.style.backgroundColor = "";
            cell.style.color = "";
          });
        });
      },
    });
    this.addCommand({
      id: "undo-color-change",
      name: "Undo last color change",
      callback: () => this.undo(),
    });

    this.addCommand({
      id: "redo-color-change",
      name: "Redo last color change",
      callback: () => this.redo(),
    });

    this.addCommand({
      id: "add-cell-color-rule",
      name: "Add table cell color rule",
      callback: () => {
        // Open settings tab and scroll to rules section
        this.app.setting.open();
        window.setTimeout(() => {
          if (
            this.app.setting &&
            typeof this.app.setting.openTabById === "function"
          ) {
            this.app.setting.openTabById("color-table-cell");
          }
        }, 250);
      },
    });

    this.addCommand({
      id: "manage-coloring-rules",
      name: "Manage coloring rules",
      callback: () => {
        this.app.setting.open();
        window.setTimeout(() => {
          if (
            this.app.setting &&
            typeof this.app.setting.openTabById === "function"
          ) {
            this.app.setting.openTabById("color-table-cell");
          }
        }, 250);
      },
    });

    this.addCommand({
      id: "add-advanced-rule",
      name: "Add advanced rule",
      callback: async () => {
        if (!Array.isArray(this.settings.advancedRules))
          this.settings.advancedRules = [];
        this.settings.advancedRules.push({
          logic: "any",
          conditions: [],
          target: "cell",
          color: null,
          bg: null,
        });
        await this.saveSettings();
        const idx = this.settings.advancedRules.length - 1;
        new AdvancedRuleModal(this.app, this, idx).open();
        document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
      },
    });

    this.addCommand({
      id: "manage-advanced-rules",
      name: "Manage advanced rules",
      callback: () => {
        this.app.setting.open();
        window.setTimeout(() => {
          if (
            this.app.setting &&
            typeof this.app.setting.openTabById === "function"
          ) {
            this.app.setting.openTabById("color-table-cell");
            // Scroll to advanced rules section
            window.setTimeout(() => {
              const advHeading = document.querySelector(".ctc-cr-adv-heading");
              if (advHeading) {
                advHeading.scrollIntoView({
                  behavior: "smooth",
                  block: "start",
                });
              }
            }, 200);
          }
        }, 250);
      },
    });

    this.addCommand({
      id: "refresh-table-colors",
      name: "Refresh table colors",
      callback: () => {
        this.hardRefreshTableColors();
      },
    });

    // REGEX PATTERN TESTER COMMAND
    // this.addCommand({
    //   id: 'open-regex-tester',
    //   name: 'Open Regex Pattern Tester',
    //   callback: () => {
    //     new RegexTesterModal(this.app, this).open();
    //   }
    // });

    try {
      if (!this._changelogCommandRegistered) {
        this.addCommand({
          id: "show-latest-release-notes",
          name: "Show latest release notes",
          callback: async () => {
            try {
              new ReleaseNotesModal(this.app, this).open();
            } catch (e) {}
          },
        });
        this._changelogCommandRegistered = true;
      }
    } catch (e) {}

    // --- Live Preview Table Coloring logic ---
    this.applyColorsToAllEditors = () => {
      debugLog(
        "applyColorsToAllEditors called, livePreviewColoring:",
        this.settings.livePreviewColoring,
      );
      if (!this.settings.livePreviewColoring) {
        // Remove all colors if disabled
        document.querySelectorAll(".cm-content table").forEach((table) => {
          table.querySelectorAll("td, th").forEach((cell) => {
            cell.style.backgroundColor = "";
            cell.style.color = "";
          });
        });
        return;
      }

      const file = this.app.workspace.getActiveFile();
      if (!file) {
        debugWarn("No active file in applyColorsToAllEditors");
        return;
      }

      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll("table"));

      const editors = document.querySelectorAll(".cm-content");
      debugLog(
        `applyColorsToAllEditors: Found ${editors.length} editors for file: ${file.path}`,
      );

      editors.forEach((editorEl) => {
        // Get all tables in this editor
        const editorTables = Array.from(editorEl.querySelectorAll("table"));
        debugLog(
          `applyColorsToAllEditors: Found ${editorTables.length} tables in editor`,
        );

        // Process each table using processSingleTable (which handles clearing internally)
        editorTables.forEach((table, localIdx) => {
          const globalTableIndex = allDocTables.indexOf(table);
          // Use global index if found, otherwise use local index within editor
          const tableIdx = globalTableIndex >= 0 ? globalTableIndex : localIdx;
          debugLog(
            `applyColorsToAllEditors: Processing table ${tableIdx} (global: ${globalTableIndex}, local: ${localIdx})`,
          );
          this.processSingleTable(table, tableIdx, file.path, noteData);
        });
      });
    };

    // Observe editors to force coloring on large tables as they render
    const installEditorObservers = () => {
      const editors = Array.from(document.querySelectorAll(".cm-content"));
      editors.forEach((ed) => {
        if (ed._ctcObserver) return;
        const obs = new MutationObserver(() => {
          const file = this.app.workspace.getActiveFile();
          if (file && this.settings.livePreviewColoring) {
            this.applyColorsToContainer(ed, file.path);
          }
        });
        obs.observe(ed, { childList: true, subtree: true });
        ed._ctcObserver = obs;

        // Add scroll handler to restore colors from data attributes
        if (!ed._ctcScrollListener) {
          const restoreColorsFromData = () => {
            // Restore colors from data attributes to handle DOM recreation
            ed.querySelectorAll("[data-ctc-bg], [data-ctc-color]").forEach(
              (cell) => {
                if (cell.hasAttribute("data-ctc-bg")) {
                  cell.style.backgroundColor = cell.getAttribute("data-ctc-bg");
                }
                if (cell.hasAttribute("data-ctc-color")) {
                  cell.style.color = cell.getAttribute("data-ctc-color");
                }
              },
            );
          };
          const onScroll = debounce(restoreColorsFromData, 50);
          // Save handler for cleanup
          ed._ctcScrollHandler = onScroll;
          ed.addEventListener("scroll", ed._ctcScrollHandler);
          ed._ctcScrollListener = true;
        }
      });
    };
    installEditorObservers();

    const setupLivePreviewColoring = () => {
      // Initial application - use applyColorsToActiveFile which handles both modes properly
      window.setTimeout(() => this.applyColorsToActiveFile(), 200);
      // Observe DOM changes in editors
      if (!this._livePreviewObserver) {
        this._livePreviewObserver = new MutationObserver(() => {
          this.applyColorsToActiveFile();
        });
        document.querySelectorAll(".cm-content").forEach((editorEl) => {
          this._livePreviewObserver.observe(editorEl, {
            childList: true,
            subtree: true,
          });
        });
      }
      // Re-apply on file open/layout change
      this.registerEvent(
        this.app.workspace.on("file-open", () =>
          this.applyColorsToActiveFile(),
        ),
      );
      this.registerEvent(
        this.app.workspace.on("layout-change", () =>
          this.applyColorsToActiveFile(),
        ),
      );
      // Re-apply on cell focus/blur/input (to persist colors after editing)
      this.registerDomEvent(document, "focusin", (e) => {
        if (
          e.target &&
          e.target.closest &&
          e.target.closest(".cm-content table")
        ) {
          this.applyColorsToActiveFile();
        }
      });
      this.registerDomEvent(document, "input", (e) => {
        if (
          e.target &&
          e.target.closest &&
          e.target.closest(".cm-content table")
        ) {
          window.setTimeout(() => this.applyColorsToActiveFile(), 30);
        }
      });

      // Watch for style changes on colored cells and restore via fallback mechanism
      const colorRestorer = new MutationObserver((mutations) => {
        let needsReapply = false;
        mutations.forEach((mutation) => {
          if (
            mutation.target.closest &&
            mutation.target.closest(".cm-content table")
          ) {
            const cell = mutation.target.closest("td, th");
            if (cell && cell.hasAttribute("data-ctc-bg")) {
              // Color was applied before, check if it's still there
              if (
                !cell.style.backgroundColor ||
                cell.style.backgroundColor === ""
              ) {
                needsReapply = true;
              }
            }
          }
        });
        if (needsReapply) {
          this.applyColorsToAllEditors();
        }
      });

      // Re-apply colors immediately when clicking/selecting table cells
      // Using single delayed timer since rule-based rendering is more stable
      this.registerDomEvent(document, "pointerdown", (e) => {
        const cell = e.target?.closest && e.target.closest("td, th");
        const table =
          e.target?.closest && e.target.closest(".cm-content table");
        if (cell && table) {
          // Single reapplication after selection settles
          window.setTimeout(() => this.applyColorsToAllEditors(), 10);
        }
      });

      // Also watch existing tables for color loss (fallback restoration)
      this._colorRestorer = colorRestorer;
      document.querySelectorAll(".cm-content table").forEach((table) => {
        colorRestorer.observe(table, {
          attributes: true,
          attributeFilter: ["style"],
          subtree: true,
        });
      });
    };

    await this.loadSettings();
    if (
      typeof this.addStatusBarItem === "function" &&
      this.settings.showStatusRefreshIcon &&
      !this.statusBarRefresh
    ) {
      this.createStatusBarIcon();
    }
    if (
      typeof this.addRibbonIcon === "function" &&
      this.settings.showRibbonRefreshIcon &&
      !this._ribbonRefreshIcon
    ) {
      const iconEl = this.addRibbonIcon("table", "Refresh table colors", () => {
        this.hardRefreshTableColors();
      });
      this._ribbonRefreshIcon = iconEl;
    }
    setupLivePreviewColoring();
    if (this.settings?.persistUndoHistory) {
      await this.loadUndoRedoStacks();
    }
    const rawSaved = (await this.loadData()) || {};

    // Use Map instead of WeakMap for proper tracking and cleanup
    this._appliedContainers = new Map();

    // Register event to refresh colors when switching files
    this.registerEvent(
      this.app.workspace.on("file-open", async (file) => {
        // Wait for DOM to fully render before applying colors
        // Use escalating delays if no tables found
        const attemptColoring = (attempt = 0) => {
          const delays = [100, 300, 800];
          const delay = delays[Math.min(attempt, delays.length - 1)];

          window.setTimeout(() => {
            // Clear any cached containers from previous file
            try {
              if (
                this._appliedContainers &&
                typeof this._appliedContainers.forEach === "function"
              ) {
                const keysToDelete = [];
                this._appliedContainers.forEach((_, container) => {
                  try {
                    if (!container || !container.isConnected) {
                      keysToDelete.push(container);
                    }
                  } catch (e) {
                    keysToDelete.push(container);
                  }
                });
                keysToDelete.forEach((key) => {
                  try {
                    this._appliedContainers.delete(key);
                  } catch (e) {}
                });
              }
            } catch (e) {}

            // Check if we have tables in the DOM
            const tables = document.querySelectorAll("table");
            if (tables.length === 0 && attempt < 2) {
              // No tables found, retry
              debugLog(
                `No tables found on file-open (attempt ${attempt + 1}), retrying...`,
              );
              attemptColoring(attempt + 1);
              return;
            }

            // Reapply colors to new file
            if (typeof this.applyColorsToActiveFile === "function") {
              this.applyColorsToActiveFile();
            }
          }, delay);
        };

        attemptColoring();
      }),
    );

    // Reapply colors when layout changes (mode switches)
    this.registerEvent(
      this.app.workspace.on("layout-change", async () => {
        window.setTimeout(() => {
          if (typeof this.applyColorsToActiveFile === "function") {
            this.applyColorsToActiveFile();
          }
        }, 100);
      }),
    );

    const normalizeCellData = (obj) => {
      let cur = obj;
      const seen = new Set();
      try {
        while (cur && typeof cur === "object" && !Array.isArray(cur)) {
          // avoid infinite cycles
          if (seen.has(cur)) break;
          seen.add(cur);

          const keys = Object.keys(cur);
          // If there are keys other than metadata keys, assume this is the real cellData
          const nonMeta = keys.filter(
            (k) => k !== "settings" && k !== "cellData",
          );
          if (nonMeta.length > 0) return cur;

          // If only a single wrapper key, unwrap it
          if (keys.length === 1) {
            const k = keys[0];
            cur = cur[k];
            continue;
          }

          // If exactly the pair of keys 'settings' and 'cellData', prefer diving into 'cellData'
          if (
            keys.length === 2 &&
            keys.includes("settings") &&
            keys.includes("cellData")
          ) {
            if (cur.cellData && typeof cur.cellData === "object") {
              cur = cur.cellData;
              continue;
            }
            return {};
          }

          // Fallback: return current object
          return cur;
        }
        return {};
      } catch (e) {
        debugWarn("Error normalizing cell data:", e);
        return {};
      }
    };

    this.cellData = normalizeCellData(rawSaved) || {};

    try {
      const normalizedSave = {
        settings: this.settings,
        cellData: this.cellData,
      };
      const rawStr = JSON.stringify(rawSaved || {});
      const normStr = JSON.stringify(normalizedSave || {});
      if (rawStr !== normStr) {
        await this.saveData(normalizedSave);
        debugLog("color-table-cell: migrated and saved normalized plugin data");
      }
    } catch (e) {
      debugWarn("Error during settings migration:", e);
      // Continue without migration on error
    }

    if (!this._settingsTab) {
      this._settingsTab = new ColorTableSettingTab(this.app, this);
      try {
        this.addSettingTab(this._settingsTab);
      } catch (e) {
        /* ignore if already added */
      }
    }

    // Auto-refresh active document when settings are closed
    this.registerDomEvent(document, "click", (e) => {
      const settingsContainer = document.querySelector(
        ".vertical-tabs-container, .settings",
      );
      const isClosingSettings =
        !settingsContainer || !settingsContainer.offsetParent;
      if (isClosingSettings && this._settingsWasOpen) {
        this._settingsWasOpen = false;
        window.setTimeout(() => {
          if (typeof this.applyColorsToActiveFile === "function") {
            this.applyColorsToActiveFile();
          }
        }, 100);
      }
    });

    // Track when settings are opened
    const originalOpen =
      this.app.setting?.open?.bind(this.app.setting) || (() => {});
    if (this.app.setting) {
      this.app.setting.open = () => {
        this._settingsWasOpen = true;
        return originalOpen();
      };
    }

    // Track observers for cleanup
    this._containerObservers = new Map();

    // Add aggressive table pre-rendering observer for reading mode
    const tablePreRenderer = new MutationObserver((mutations) => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;

      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll("table"));

      // Look for newly added tables in reading mode
      mutations.forEach((mutation) => {
        if (mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              // Check if this node or its children contain tables
              const tables = [];
              if (node.matches && node.matches("table")) {
                tables.push(node);
              }
              if (node.querySelectorAll) {
                tables.push(...node.querySelectorAll("table"));
              }

              // Immediately color any new tables in reading mode
              tables.forEach((table) => {
                if (
                  table.closest(".markdown-preview-view") &&
                  !table.hasAttribute("data-ctc-processed")
                ) {
                  const globalTableIdx = allDocTables.indexOf(table);
                  if (globalTableIdx >= 0) {
                    this.processSingleTable(
                      table,
                      globalTableIdx,
                      file.path,
                      noteData,
                    );
                  }
                }
              });
            }
          });
        }
      });
    });

    // Observe the entire document body for table additions
    tablePreRenderer.observe(document.body, { childList: true, subtree: true });
    this._tablePreRenderer = tablePreRenderer;

    this.registerMarkdownPostProcessor((el, ctx) => {
      if (!el.closest(".markdown-preview-view")) return;
      const fileId = ctx.sourcePath;
      this.applyColorsToContainer(el, fileId);

      try {
        // Skip if we're already observing this element
        if (this._containerObservers.has(el)) {
          const existing = this._containerObservers.get(el);
          if (existing && existing.observer) {
            return; // Already observing this element
          }
        }

        let observer = null;
        let debounceId = null;
        let scrollListener = null;
        const scrollDebounceId = null; // kept for safeDisconnect compatibility (scrollListener now uses debounce())

        const safeDisconnect = () => {
          try {
            if (observer) {
              observer.disconnect();
              observer = null;
            }
          } catch (e) {}
          try {
            if (debounceId) {
              window.clearTimeout(debounceId);
              debounceId = null;
            }
          } catch (e) {}
          try {
            if (scrollListener && el.parentElement) {
              el.parentElement.removeEventListener("scroll", scrollListener);
              if (typeof scrollListener.cancel === "function")
                scrollListener.cancel();
            }
          } catch (e) {}
          try {
            if (scrollDebounceId) {
              window.clearTimeout(scrollDebounceId);
              scrollDebounceId = null;
            }
          } catch (e) {}

          // Remove from tracking map
          if (this._containerObservers.has(el)) {
            this._containerObservers.delete(el);
          }
        };

        const checkAndApply = () => {
          try {
            // Check if element is still connected to DOM
            if (!el.isConnected) {
              safeDisconnect();
              return;
            }

            if (
              el.querySelectorAll &&
              el.querySelectorAll("table").length > 0
            ) {
              this.applyColorsToContainer(el, fileId);
              // Don't disconnect immediately - keep observing for dynamic content
            }
          } catch (e) {
            /* ignore */
          }
        };

        observer = new MutationObserver((mutations) => {
          // Look specifically for added tables
          let hasTableAdded = false;
          mutations.forEach((mutation) => {
            if (
              mutation.type === "childList" &&
              mutation.addedNodes.length > 0
            ) {
              mutation.addedNodes.forEach((node) => {
                if (node.nodeType === 1) {
                  // Element node
                  // Check if this is a table or contains tables
                  if (
                    node.tagName === "TABLE" ||
                    (node.querySelector && node.querySelector("table"))
                  ) {
                    hasTableAdded = true;
                  }
                }
              });
            }
          });

          if (hasTableAdded) {
            debugLog(
              "MutationObserver: Table added to reading view, applying colors",
            );
            if (debounceId) window.clearTimeout(debounceId);
            debounceId = window.setTimeout(() => {
              checkAndApply();
            }, 80);
          }
        });

        observer.observe(el, { childList: true, subtree: true });

        // Add scroll listener to reapply colors on scroll (for lazy-loaded tables)
        try {
          const scrollContainer = el.parentElement || el;
          scrollListener = debounce(
            () => {
              // Check if element is still connected before applying
              if (el.isConnected) {
                debugLog(
                  "Scroll detected in reading view, checking for new tables",
                );
                this.applyColorsToContainer(el, fileId);
              } else {
                safeDisconnect();
              }
            },
            100,
            true,
          );
          scrollContainer.addEventListener("scroll", scrollListener, {
            passive: true,
          });
        } catch (e) {
          /* ignore if scroll listener fails */
        }

        // Store observer reference for cleanup
        this._containerObservers.set(el, {
          observer,
          debounceId,
          scrollListener,
          scrollDebounceId,
          safeDisconnect,
        });

        checkAndApply();

        // Check periodically if element is still connected
        const connectionChecker = window.setInterval(() => {
          if (!el.isConnected) {
            safeDisconnect();
            window.clearInterval(connectionChecker);
          }
        }, 2000);

        // Auto-cleanup after reasonable time
        window.setTimeout(() => {
          window.clearInterval(connectionChecker);
          safeDisconnect();
        }, 30000);
      } catch (e) {
        /* ignore if MutationObserver unsupported */
      }
    });

    // Also add a global observer specifically for reading mode tables
    try {
      const readingViewObserver = new MutationObserver((mutations) => {
        const file = this.app.workspace.getActiveFile();
        if (!file) return;

        mutations.forEach((mutation) => {
          if (mutation.type === "childList" && mutation.addedNodes.length > 0) {
            let hasNewTable = false;
            mutation.addedNodes.forEach((node) => {
              if (node.nodeType === 1) {
                // Check if a table was added to reading view
                if (
                  node.matches &&
                  node.matches("table") &&
                  node.closest(".markdown-preview-view")
                ) {
                  hasNewTable = true;
                } else if (node.querySelectorAll) {
                  const tables = node.querySelectorAll("table");
                  if (
                    tables.length > 0 &&
                    node.closest(".markdown-preview-view")
                  ) {
                    hasNewTable = true;
                  }
                }
              }
            });

            if (hasNewTable) {
              debugLog(
                "Global observer: New table detected in reading view, applying colors",
              );
              // Use applyColorsToActiveFile which handles both reading and live preview
              this.applyColorsToActiveFile();
            }
          }
        });
      });

      // Observe the entire document body for new reading view tables
      readingViewObserver.observe(document.body, {
        childList: true,
        subtree: true,
      });
      this._readingViewObserver = readingViewObserver;
    } catch (e) {
      debugWarn("Failed to setup reading view observer:", e);
    }

    // Setup reading view scroll listener for lazy-loaded tables
    this.setupReadingViewScrollListener();

    // Start periodic checker for reading mode tables
    this.startReadingModeTableChecker();

    // Enhanced global observer to restore colors from data attributes when DOM recreates cells
    try {
      const globalObserver = new MutationObserver((mutations) => {
        try {
          mutations.forEach((mutation) => {
            if (
              mutation.type === "childList" &&
              mutation.addedNodes.length > 0
            ) {
              mutation.addedNodes.forEach((node) => {
                try {
                  if (node.nodeType === 1) {
                    // Check for cells with data attributes
                    if (
                      node.matches &&
                      node.matches("[data-ctc-bg], [data-ctc-color]")
                    ) {
                      this.restoreColorsFromAttributes(node);
                    }
                    // Check descendants
                    if (node.querySelectorAll) {
                      node
                        .querySelectorAll("[data-ctc-bg], [data-ctc-color]")
                        .forEach((cell) => {
                          this.restoreColorsFromAttributes(cell);
                        });
                    }
                  }
                } catch (e) {}
              });
            }
          });
        } catch (e) {}
      });

      globalObserver.observe(document.body, { childList: true, subtree: true });
      this._globalObserver = globalObserver;
    } catch (e) {
      debugWarn("Failed to setup global observer:", e);
      this._globalObserver = null;
    }

    if (this.settings.enableContextMenu) {
      this.registerDomEvent(document, "contextmenu", (evt) => {
        const target = evt.target;
        const cell = target?.closest("td, th");
        const tableEl = target?.closest("table");
        if (!cell || !tableEl) return;
        const readingView = cell.closest(".markdown-preview-view");
        const livePreview = cell.closest(".cm-content");
        if (!readingView || livePreview) return;
        const menu = new Menu();
        menu.addItem((item) =>
          item
            .setTitle("Color cell text")
            .setIcon("palette")
            .onClick(() => this.pickColor(cell, tableEl, "color")),
        );
        menu.addItem((item) =>
          item
            .setTitle("Color cell background")
            .setIcon("droplet")
            .onClick(() => this.pickColor(cell, tableEl, "bg")),
        );
        menu.addSeparator();

        // Row coloring options - conditional on setting
        if (this.settings.showColorRowInMenu) {
          menu.addItem((item) =>
            item
              .setTitle("Color whole row text")
              .setIcon("rows-3")
              .onClick(() => this.pickColorForRow(cell, tableEl, "color")),
          );
          menu.addItem((item) =>
            item
              .setTitle("Color whole row background")
              .setIcon("droplet")
              .onClick(() => this.pickColorForRow(cell, tableEl, "bg")),
          );
          menu.addSeparator();
        }

        // Column coloring options - conditional on setting
        if (this.settings.showColorColumnInMenu) {
          menu.addItem((item) =>
            item
              .setTitle("Color whole column text")
              .setIcon("columns-3")
              .onClick(() => this.pickColorForColumn(cell, tableEl, "color")),
          );
          menu.addItem((item) =>
            item
              .setTitle("Color whole column background")
              .setIcon("droplet")
              .onClick(() => this.pickColorForColumn(cell, tableEl, "bg")),
          );
          menu.addSeparator();
        }

        menu.addItem((item) =>
          item
            .setTitle("Color multiple cells by rule")
            .setIcon("grid")
            .onClick(() => {
              this.app.setting.open();
              window.setTimeout(() => {
                if (
                  this.app.setting &&
                  typeof this.app.setting.openTabById === "function"
                ) {
                  this.app.setting.openTabById("color-table-cell");
                  // Scroll to Coloring Rules heading
                  window.setTimeout(() => {
                    const settingsContainer = document.querySelector(
                      ".vertical-tabs-container",
                    );
                    if (settingsContainer) {
                      const rulesHeading = Array.from(
                        settingsContainer.querySelectorAll("h3"),
                      ).find((el) => el.textContent.includes("Coloring Rules"));
                      if (rulesHeading) {
                        rulesHeading.scrollIntoView({
                          behavior: "smooth",
                          block: "start",
                        });
                      }
                    }
                  }, 100);
                }
              }, 250);
            }),
        );
        menu.addSeparator();

        // Undo/Redo options - conditional on setting
        if (this.settings.showUndoRedoInMenu) {
          menu.addItem((item) =>
            item
              .setTitle("Undo last color change")
              .setIcon("undo")
              .setDisabled(this.undoStack.length === 0)
              .onClick(() => this.undo()),
          );
          menu.addItem((item) =>
            item
              .setTitle("Redo last color change")
              .setIcon("redo")
              .setDisabled(this.redoStack.length === 0)
              .onClick(() => this.redo()),
          );
          menu.addSeparator();
        }

        menu.addItem((item) =>
          item
            .setTitle("Reset cell coloring")
            .setIcon("trash-2")
            .onClick(async () => this.resetCell(cell, tableEl)),
        );

        if (this.settings.showColorRowInMenu) {
          menu.addItem((item) =>
            item
              .setTitle("Remove row coloring")
              .setIcon("rows-3")
              .onClick(async () => this.resetRow(cell, tableEl)),
          );
        }

        if (this.settings.showColorColumnInMenu) {
          menu.addItem((item) =>
            item
              .setTitle("Remove column coloring")
              .setIcon("columns-3")
              .onClick(async () => this.resetColumn(cell, tableEl)),
          );
        }
        try {
          if (menu.containerEl && menu.containerEl.classList)
            menu.containerEl.classList.add("mod-shadow");
          if (menu.menuEl && menu.menuEl.classList)
            menu.menuEl.classList.add("mod-shadow");
        } catch (e) {
          /* ignore if properties not present */
        }

        menu.showAtMouseEvent(evt);
        evt.preventDefault();
      });
    }

    this.applyColorsToActiveFile();

    this.registerEvent(
      this.app.workspace.on("file-open", () => {
        this.applyColorsToActiveFile();
      }),
    );

    this.registerEvent(
      this.app.workspace.on("layout-change", () => {
        this.applyColorsToActiveFile();
      }),
    );

    // Also listen to active leaf changes to catch view switches
    this.registerEvent(
      this.app.workspace.on("active-leaf-change", () => {
        window.setTimeout(() => this.applyColorsToActiveFile(), 50);
      }),
    );
  }

  createStatusBarIcon() {
    if (!this.statusBarRefresh && typeof this.addStatusBarItem === "function") {
      const status = this.addStatusBarItem();
      this.statusBarRefresh = status;
      setIcon(status, "table");
      status.setAttribute("aria-label", "Refresh table colors");
      status.classList.add("ctc-refresh-table-color");

      status.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.hardRefreshTableColors();
      });
      // Ensure it's visible
      status.style.display = "";
      debugLog("[Status Bar] Icon created successfully");
    }
  }

  removeStatusBarIcon() {
    if (this.statusBarRefresh) {
      try {
        // Immediately hide the element
        this.statusBarRefresh.style.display = "none";

        // Try to remove it from DOM completely
        try {
          const parent = this.statusBarRefresh.parentElement;
          if (parent && this.statusBarRefresh) {
            parent.removeChild(this.statusBarRefresh);
          }
        } catch (e) {}

        this.statusBarRefresh = null;
        debugLog("[Status Bar] Icon removed successfully");
      } catch (e) {
        debugWarn("[Status Bar] Error removing icon:", e);
        this.statusBarRefresh = null;
      }
    }
  }

  hardRefreshTableColors() {
    debugLog("[Hard Refresh] Starting hard refresh of table colors");

    // STEP 1: Clear all table colors from DOM (both manual and rule-based)
    document.querySelectorAll("table td, table th").forEach((cell) => {
      cell.style.backgroundColor = "";
      cell.style.color = "";
      // Clear data attributes that might cache colors
      cell.removeAttribute("data-ctc-bg");
      cell.removeAttribute("data-ctc-color");
    });

    // STEP 2: Clear table processing markers to force reprocessing
    document.querySelectorAll("table").forEach((table) => {
      table.removeAttribute("data-ctc-processed");
      table.removeAttribute("data-ctc-index");
      table.removeAttribute("data-ctc-file");
      table.removeAttribute("data-ctc-last-processed");
    });

    // STEP 3: Reset internal cache for DOM tracking
    try {
      if (
        this._appliedContainers &&
        typeof this._appliedContainers.clear === "function"
      ) {
        this._appliedContainers.clear();
      } else {
        this._appliedContainers = new Map();
      }
    } catch (e) {
      this._appliedContainers = new Map();
    }

    // STEP 4: Reapply all colors from scratch
    debugLog("[Hard Refresh] Colors cleared, reapplying from scratch");

    if (typeof this.applyColorsToActiveFile === "function") {
      window.setTimeout(() => {
        this.applyColorsToActiveFile();
        debugLog("[Hard Refresh] Hard refresh complete");
      }, 50);
    }
  }

  onunload() {
    // Remove our settings tab when the plugin is
    // disabled/unloaded to avoid duplicate entries
    try {
      if (this._settingsTab && typeof this.removeSettingTab === "function") {
        this.removeSettingTab(this._settingsTab);
      }
    } catch (e) {}

    // Clean up live preview observer with proper error handling
    try {
      if (this._livePreviewObserver) {
        if (typeof this._livePreviewObserver.disconnect === "function") {
          this._livePreviewObserver.disconnect();
        }
        this._livePreviewObserver = null;
      }
    } catch (e) {
      debugWarn("Error cleaning up live preview observer:", e);
      this._livePreviewObserver = null;
    }

    // Clean up table pre-renderer
    try {
      if (this._tablePreRenderer) {
        if (typeof this._tablePreRenderer.disconnect === "function") {
          this._tablePreRenderer.disconnect();
        }
        this._tablePreRenderer = null;
      }
    } catch (e) {
      debugWarn("Error cleaning up table pre-renderer:", e);
      this._tablePreRenderer = null;
    }

    // Clean up reading view scroll observer
    try {
      if (this._readingViewScrollObserver) {
        if (typeof this._readingViewScrollObserver.disconnect === "function") {
          this._readingViewScrollObserver.disconnect();
        }
        this._readingViewScrollObserver = null;
      }
    } catch (e) {
      debugWarn("Error cleaning up reading view scroll observer:", e);
      this._readingViewScrollObserver = null;
    }

    // Clean up reading mode checker interval with validation
    try {
      if (this._readingModeChecker) {
        window.clearInterval(this._readingModeChecker);
        this._readingModeChecker = null;
      }
    } catch (e) {
      debugWarn("Error clearing reading mode checker interval:", e);
      this._readingModeChecker = null;
    }

    // Clean up reading view scroll listeners
    try {
      document.querySelectorAll(".markdown-preview-view").forEach((view) => {
        try {
          // Remove scroll listener if it was added
          if (view && view._ctcScrollListenerAdded && view._ctcScrollHandler) {
            view.removeEventListener("scroll", view._ctcScrollHandler);
            view._ctcScrollListenerAdded = false;
            view._ctcScrollHandler = null;
          }
        } catch (e) {}
      });
    } catch (e) {
      debugWarn("Error cleaning up reading view scroll listeners:", e);
    }

    // Clean up editor scroll listeners
    try {
      document.querySelectorAll(".cm-content").forEach((ed) => {
        try {
          // Disconnect observer
          if (ed._ctcObserver) {
            if (typeof ed._ctcObserver.disconnect === "function") {
              ed._ctcObserver.disconnect();
            }
            ed._ctcObserver = null;
          }
          // Remove scroll listener
          if (ed._ctcScrollListener && ed._ctcScrollHandler) {
            ed.removeEventListener("scroll", ed._ctcScrollHandler);
            ed._ctcScrollListener = false;
            ed._ctcScrollHandler = null;
          }
        } catch (e) {}
      });
    } catch (e) {
      debugWarn("Error cleaning up editor scroll listeners:", e);
    }

    // Clean up all container observers to prevent memory leaks
    try {
      if (this._containerObservers) {
        const entries = Array.from(this._containerObservers.entries());
        entries.forEach(([el, observerData]) => {
          try {
            if (observerData) {
              if (typeof observerData.safeDisconnect === "function") {
                observerData.safeDisconnect();
              } else if (typeof observerData.disconnect === "function") {
                observerData.disconnect();
              }
            }
          } catch (e) {}
        });
        this._containerObservers.clear();
        this._containerObservers = null;
      }
    } catch (e) {
      debugWarn("Error cleaning up container observers:", e);
      this._containerObservers = null;
    }

    // Clean up color restorer
    try {
      if (this._colorRestorer) {
        if (typeof this._colorRestorer.disconnect === "function") {
          this._colorRestorer.disconnect();
        }
        this._colorRestorer = null;
      }
    } catch (e) {
      debugWarn("Error cleaning up color restorer:", e);
      this._colorRestorer = null;
    }

    // Clean up global observer
    try {
      if (this._globalObserver) {
        if (typeof this._globalObserver.disconnect === "function") {
          this._globalObserver.disconnect();
        }
        this._globalObserver = null;
      }
    } catch (e) {
      debugWarn("Error cleaning up global observer:", e);
      this._globalObserver = null;
    }

    // Save undo/redo history if enabled
    try {
      if (
        this.settings?.persistUndoHistory &&
        typeof this.saveUndoRedoStacks === "function"
      ) {
        this.saveUndoRedoStacks();
      }
    } catch (e) {
      debugWarn("Error saving undo/redo stacks:", e);
    }
  }

  async loadSettings() {
    const data = await this.loadData();
    this.settings = Object.assign(
      {
        enableContextMenu: true,
        showColorRowInMenu: true,
        showColorColumnInMenu: true,
        showUndoRedoInMenu: true,
        coloringRules: [],
        coloringSort: "lastAdded",
        advancedRules: [],
        numericStrict: true,
        livePreviewColoring: false,
        persistUndoHistory: true,
        recentColors: [],
        presetColors: [],
        showStatusRefreshIcon: true,
        showRibbonRefreshIcon: true,
      },
      data?.settings || {},
    );

    // MIGRATION: Convert old "rules" format to new "coloringRules" format
    if (
      Array.isArray(data?.settings?.rules) &&
      data.settings.rules.length > 0
    ) {
      const oldRules = data.settings.rules;
      debugLog(`[Migration] Found ${oldRules.length} old rules to migrate`);

      // Convert each old rule to new coloringRules format
      const migratedRules = oldRules.map((oldRule) => {
        const newRule = {
          target: "cell", // All old rules were "color cell" by default
          when: "theCell", // Match the specific cell
          match: oldRule.regex ? "isRegex" : "contains", // regex or contains
          value: oldRule.match,
          color: oldRule.color || null,
          bg: oldRule.bg || null,
        };

        // If old rule had case sensitivity, note it was case-insensitive in old version
        // (old version didn't have when/target fields, so defaults apply)
        return newRule;
      });

      // Only migrate if coloringRules is empty or doesn't exist
      if (
        !Array.isArray(this.settings.coloringRules) ||
        this.settings.coloringRules.length === 0
      ) {
        this.settings.coloringRules = migratedRules;
        debugLog(
          `[Migration] Migrated ${migratedRules.length} rules to coloringRules`,
        );

        // Mark that migration happened so we know to save
        this._settingsMigrated = true;
      }
    }

    if (Array.isArray(this.settings.presetColors)) {
      this.settings.presetColors = this.settings.presetColors.map((pc) => {
        return typeof pc === "string" ? { name: "", color: pc } : pc;
      });
    } else {
      this.settings.presetColors = [];
    }
  }

  async saveSettings() {
    try {
      const dataToSave = {
        settings: this.settings,
        cellData: this.cellData,
      };

      // If migration happened, remove old 'rules' field to avoid confusion
      if (this._settingsMigrated) {
        // Old 'rules' won't be in this.settings anymore, but ensure it's not re-added
        delete dataToSave.settings.rules;
        this._settingsMigrated = false; // Reset migration flag
      }

      await this.saveData(dataToSave);
    } catch (error) {
      throw error;
    }
  }

  async fetchAllReleases() {
    const allReleases = [];
    let page = 1;
    let hasMorePages = true;
    while (hasMorePages) {
      const url = `https://api.github.com/repos/Kazi-Aidah/color-table-cells/releases?page=${page}&per_page=100`;
      try {
        let data = null;
        if (typeof requestUrl === "function") {
          try {
            const res = await requestUrl({
              url,
              headers: {
                Accept: "application/vnd.github.v3+json",
                "User-Agent": "Obsidian-Color-Table-Cells",
              },
            });
            data = res.json || (res.text ? JSON.parse(res.text) : null);
          } catch (e) {}
        }
        if (!data) {
          try {
            const r = await fetch(url, {
              headers: {
                Accept: "application/vnd.github.v3+json",
                "User-Agent": "Obsidian-Color-Table-Cells",
              },
            });
            if (!r.ok) throw new Error("Network error");
            data = await r.json();
          } catch (e) {
            hasMorePages = false;
            break;
          }
        }
        if (!Array.isArray(data) || data.length === 0) {
          hasMorePages = false;
        } else {
          allReleases.push(...data);
          if (data.length < 100) {
            hasMorePages = false;
          } else {
            page++;
          }
        }
      } catch (e) {
        hasMorePages = false;
      }
    }
    return allReleases;
  }

  updateRecentColor(color) {
    if (!color) return;
    const list = Array.isArray(this.settings.recentColors)
      ? [...this.settings.recentColors]
      : [];
    const existingIndex = list.findIndex(
      (c) => c.toUpperCase() === color.toUpperCase(),
    );
    if (existingIndex !== -1) list.splice(existingIndex, 1);
    list.unshift(color);
    this.settings.recentColors = list.slice(0, 10);
  }

  // Color Picker Menu Class
  createColorPickerMenu() {
    // Floating color picker menu
    class CustomColorPickerMenu {
      constructor(plugin, onPick, initialColor, anchorEl) {
        this.plugin = plugin;
        this.onPick = onPick;
        this.initialColor = initialColor || "#FFA500";
        this.anchorEl = anchorEl;
        const { h, s, v } = this.hexToHsv(this.initialColor);
        this.hue = h;
        this.sat = s;
        this.val = v;
        this.color = this.hsvToHex(this.hue, this.sat, this.val);
        this.menuEl = null;
        this._cells = []; // Support multiple cells for row/column coloring
      }
      open() {
        this.close();
        this.menuEl = document.createElement("div");
        this.menuEl.className = "ctc-color-picker-menu";

        const sbBox = document.createElement("canvas");
        sbBox.width = 210;
        sbBox.height = 120;
        sbBox.className = "ctc-color-picker-sb-box";
        this.menuEl.appendChild(sbBox);

        const sbCtx = sbBox.getContext("2d");
        const drawSB = (hue) => {
          for (let x = 0; x < sbBox.width; x++) {
            for (let y = 0; y < sbBox.height; y++) {
              const s = x / (sbBox.width - 1);
              const v = 1 - y / (sbBox.height - 1);
              sbCtx.fillStyle = this.hsvToHex(hue, s, v);
              sbCtx.fillRect(x, y, 1, 1);
            }
          }
        };
        drawSB(this.hue);

        const sbSelector = document.createElement("div");
        sbSelector.className = "ctc-color-picker-sb-selector";
        this.menuEl.appendChild(sbSelector);

        const hueBox = document.createElement("canvas");
        hueBox.width = 210;
        hueBox.height = 14;
        hueBox.className = "ctc-color-picker-hue-box";
        this.menuEl.appendChild(hueBox);

        const hueCtx = hueBox.getContext("2d");
        const grad = hueCtx.createLinearGradient(0, 0, hueBox.width, 0);
        for (let i = 0; i <= 360; i += 1) {
          grad.addColorStop(i / 360, `hsl(${i},100%,50%)`);
        }
        hueCtx.fillStyle = grad;
        hueCtx.fillRect(0, 0, hueBox.width, hueBox.height);

        const hueSelector = document.createElement("div");
        hueSelector.className = "ctc-color-picker-hue-selector";
        this.menuEl.appendChild(hueSelector);

        const hexRow = document.createElement("div");
        hexRow.className = "ctc-color-picker-hex-row";
        this.menuEl.appendChild(hexRow);
        const preview = document.createElement("div");
        preview.className = "ctc-color-picker-preview";
        preview.style.background = this.color;
        hexRow.appendChild(preview);
        const hexInput = document.createElement("input");
        hexInput.type = "text";
        hexInput.value = (this.color || "#FFA500").replace(/^#/, "");
        hexInput.maxLength = 7;
        hexInput.className = "ctc-color-picker-hex-input";
        let hexAnimFrame = null;
        hexInput.addEventListener("input", () => {
          hexInput.value = hexInput.value
            .replace(/[^0-9a-fA-F]/g, "")
            .slice(0, 6);
          if (hexInput.value.length === 6) {
            const hex = "#" + hexInput.value;
            const { h, s, v } = this.hexToHsv(hex);
            this.hue = h;
            this.sat = s;
            this.val = v;
            this.color = hex;
            preview.style.background = this.color;
            updateSelectors();
            drawSB(this.hue);
          }
        });
        hexRow.appendChild(hexInput);
        const pickBtn = document.createElement("button");
        pickBtn.className = "mod-ghost ctc-color-picker-icon-button";
        pickBtn.textContent = "";
        try {
          const ob = require("obsidian");
          if (ob && typeof ob.setIcon === "function") {
            ob.setIcon(pickBtn, "pipette");
          } else if (typeof window.setIcon === "function") {
            window.setIcon(pickBtn, "pipette");
          } else if (typeof setIcon === "function") {
            setIcon(pickBtn, "pipette");
          } else {
            const svg = document.createElementNS(
              "http://www.w3.org/2000/svg",
              "svg",
            );
            svg.setAttribute("viewBox", "0 0 24 24");
            svg.setAttribute("width", "20");
            svg.setAttribute("height", "20");
            const path = document.createElementNS(
              "http://www.w3.org/2000/svg",
              "path",
            );
            path.setAttribute(
              "d",
              "M13.5 6.5l4-4 4 4-4 4M4 20l6-6 3 3-6 6H4v-3z",
            );
            path.setAttribute("fill", "none");
            path.setAttribute("stroke", "currentColor");
            path.setAttribute("stroke-width", "2");
            path.setAttribute("stroke-linecap", "round");
            path.setAttribute("stroke-linejoin", "round");
            svg.appendChild(path);
            pickBtn.appendChild(svg);
          }
        } catch (e) {}
        pickBtn.title = "Pick color from screen";
        pickBtn.onclick = async () => {
          try {
            if ("EyeDropper" in window) {
              const eye = new window.EyeDropper();
              const result = await eye.open();
              const hex =
                result && result.sRGBHex ? result.sRGBHex.toUpperCase() : null;
              if (hex) {
                const { h, s, v } = this.hexToHsv(hex);
                this.hue = h;
                this.sat = s;
                this.val = v;
                this.color = hex;
                preview.style.background = this.color;
                drawSB(this.hue);
                updateSelectors();
              }
            } else {
              pickBtn.title = "Screen picker not supported in this environment";
            }
          } catch (e) {
            /* ignore user cancel or errors */
          }
        };
        hexRow.appendChild(pickBtn);
        const recentRow = document.createElement("div");
        recentRow.className = "ctc-color-picker-recents";
        const recents = Array.isArray(this.plugin.settings.recentColors)
          ? this.plugin.settings.recentColors
          : [];
        recents.slice(0, 10).forEach((rc) => {
          const sw = document.createElement("button");
          sw.className = "ctc-color-picker-recent-swatch";
          sw.style.background = rc;
          sw.addEventListener("click", () => {
            const { h, s, v } = this.hexToHsv(rc);
            this.hue = h;
            this.sat = s;
            this.val = v;
            this.color = rc;
            preview.style.background = this.color;
            drawSB(this.hue);
            updateSelectors();
          });
          recentRow.appendChild(sw);
        });
        this.menuEl.appendChild(recentRow);
        const updateSelectors = () => {
          const x = Math.round(this.sat * (sbBox.width - 1));
          const y = Math.round((1 - this.val) * (sbBox.height - 1));
          sbSelector.style.left = sbBox.offsetLeft + x + "px";
          sbSelector.style.top = sbBox.offsetTop + y + "px";
          const hx = Math.round((this.hue / 360) * (hueBox.width - 1));
          hueSelector.style.left = hueBox.offsetLeft + hx + "px";
          hueSelector.style.top = hueBox.offsetTop + 1 + "px";
          if (!document.activeElement || document.activeElement !== hexInput) {
            if (hexAnimFrame) cancelAnimationFrame(hexAnimFrame);
            const target = this.color.replace(/^#/, "").toUpperCase();
            if (hexInput.value !== target) {
              let i = 0;
              const animate = () => {
                if (hexInput.value !== target) {
                  hexInput.value = target.slice(0, i + 1);
                  i++;
                  if (i < target.length)
                    hexAnimFrame = requestAnimationFrame(animate);
                }
              };
              animate();
            }
          }
          if (this._cell && this._type) {
            if (this._type === "bg")
              this._cell.style.backgroundColor = this.color;
            else this._cell.style.color = this.color;
          }
          // Support for multiple cells (row/column)
          if (this._cells && this._cells.length > 0 && this._type) {
            this._cells.forEach((cell) => {
              if (this._type === "bg") cell.style.backgroundColor = this.color;
              else cell.style.color = this.color;
            });
          }
        };
        let sbDragging = false;
        sbBox.addEventListener("mousedown", (e) => {
          sbDragging = true;
          handleSB(e);
        });
        window.addEventListener("mousemove", this._sbMoveHandler);
        window.addEventListener("mouseup", this._sbUpHandler);
        const handleSB = (e) => {
          const rect = sbBox.getBoundingClientRect();
          let x = e.clientX - rect.left;
          let y = e.clientY - rect.top;
          x = Math.max(0, Math.min(sbBox.width - 1, x));
          y = Math.max(0, Math.min(sbBox.height - 1, y));
          this.sat = x / (sbBox.width - 1);
          this.val = 1 - y / (sbBox.height - 1);
          this.color = this.hsvToHex(this.hue, this.sat, this.val);
          preview.style.background = this.color;
          updateSelectors();
        };
        let hueDragging = false;
        let huePending = false;
        let hueX = 0;
        const updateHueSelector = () => {
          const newHue = (hueX / (hueBox.width - 1)) * 360;
          this.hue = newHue;
          this.color = this.hsvToHex(this.hue, this.sat, this.val);
          preview.style.background = this.color;
          hueSelector.style.left = hueBox.offsetLeft + hueX - 6 + "px";
          hueSelector.style.top = hueBox.offsetTop + 1 + "px";
          if (this._cell && this._type) {
            if (this._type === "bg")
              this._cell.style.backgroundColor = this.color;
            else this._cell.style.color = this.color;
          }
          // Support for multiple cells (row/column)
          if (this._cells && this._cells.length > 0 && this._type) {
            this._cells.forEach((cell) => {
              if (this._type === "bg") cell.style.backgroundColor = this.color;
              else cell.style.color = this.color;
            });
          }
          drawSB(this.hue);
        };
        const onHueMove = (e) => {
          const rect = hueBox.getBoundingClientRect();
          hueX = Math.max(0, Math.min(hueBox.width - 1, e.clientX - rect.left));
          if (!huePending) {
            huePending = true;
            requestAnimationFrame(() => {
              updateHueSelector();
              huePending = false;
            });
          }
        };
        hueBox.addEventListener("mousedown", (e) => {
          hueDragging = true;
          onHueMove(e);
        });
        window.addEventListener("mousemove", this._hueMoveHandler);
        window.addEventListener("mouseup", this._hueUpHandler);
        window.setTimeout(updateSelectors, 0);

        hexInput.addEventListener("keydown", (e) => {
          if (e.key === "Enter") applyBtn.click();
        });

        window.setTimeout(() => {
          let rect = this.anchorEl.getBoundingClientRect();
          let left = rect.left;
          let top = rect.bottom + 4;

          if (left + 280 > window.innerWidth) left = window.innerWidth - 280;
          if (top + 260 > window.innerHeight) top = rect.top - 260;
          this.menuEl.style.left = left + "px";
          this.menuEl.style.top = top + "px";
        }, 0);
        document.body.appendChild(this.menuEl);

        // Store event handlers for proper cleanup
        this._outsideHandler = (evt) => {
          if (this.menuEl && !this.menuEl.contains(evt.target)) this.close();
        };
        // Store all window event handlers for cleanup
        this._sbMoveHandler = (e) => {
          if (sbDragging) handleSB(e);
        };
        this._sbUpHandler = () => {
          sbDragging = false;
        };
        this._hueMoveHandler = (e) => {
          if (hueDragging) onHueMove(e);
        };
        this._hueUpHandler = () => {
          hueDragging = false;
        };

        window.setTimeout(
          () => document.addEventListener("mousedown", this._outsideHandler),
          10,
        );
      }
      close() {
        try {
          if (this.menuEl && this.menuEl.parentNode) {
            this.menuEl.parentNode.removeChild(this.menuEl);
            this.menuEl = null;
          }
        } catch (e) {}

        // Clean up event listeners
        try {
          if (this._outsideHandler)
            document.removeEventListener("mousedown", this._outsideHandler);
          if (this._sbMoveHandler)
            window.removeEventListener("mousemove", this._sbMoveHandler);
          if (this._sbUpHandler)
            window.removeEventListener("mouseup", this._sbUpHandler);
          if (this._hueMoveHandler)
            window.removeEventListener("mousemove", this._hueMoveHandler);
          if (this._hueUpHandler)
            window.removeEventListener("mouseup", this._hueUpHandler);
        } catch (e) {}

        // Clean up handler references
        this._outsideHandler = null;
        this._sbMoveHandler = null;
        this._sbUpHandler = null;
        this._hueMoveHandler = null;
        this._hueUpHandler = null;

        // Save the picked color
        try {
          if (typeof this.onPick === "function") {
            this.onPick(this.color);
          }
        } catch (e) {}
      }

      destroy() {
        // Explicit cleanup
        this.close();
        this.plugin = null;
      }

      hsvToHex(h, s, v) {
        let r, g, b;
        let i = Math.floor(h / 60);
        let f = h / 60 - i;
        let p = v * (1 - s);
        let q = v * (1 - f * s);
        let t = v * (1 - (1 - f) * s);
        switch (i % 6) {
          case 0:
            ((r = v), (g = t), (b = p));
            break;
          case 1:
            ((r = q), (g = v), (b = p));
            break;
          case 2:
            ((r = p), (g = v), (b = t));
            break;
          case 3:
            ((r = p), (g = q), (b = v));
            break;
          case 4:
            ((r = t), (g = p), (b = v));
            break;
          case 5:
            ((r = v), (g = p), (b = q));
            break;
        }
        return (
          "#" +
          [r, g, b]
            .map((x) =>
              Math.round(x * 255)
                .toString(16)
                .padStart(2, "0"),
            )
            .join("")
            .toUpperCase()
        );
      }
      hexToHsv(hex) {
        let r = 0,
          g = 0,
          b = 0;
        if (hex.length === 4) {
          r = parseInt(hex[1] + hex[1], 16);
          g = parseInt(hex[2] + hex[2], 16);
          b = parseInt(hex[3] + hex[3], 16);
        } else if (hex.length === 7) {
          r = parseInt(hex.substr(1, 2), 16);
          g = parseInt(hex.substr(3, 2), 16);
          b = parseInt(hex.substr(5, 2), 16);
        }
        r /= 255;
        g /= 255;
        b /= 255;
        let max = Math.max(r, g, b),
          min = Math.min(r, g, b);
        let h,
          s,
          v = max;
        let d = max - min;
        s = max === 0 ? 0 : d / max;
        if (max === min) h = 0;
        else {
          switch (max) {
            case r:
              h = (g - b) / d + (g < b ? 6 : 0);
              break;
            case g:
              h = (b - r) / d + 2;
              break;
            case b:
              h = (r - g) / d + 4;
              break;
          }
          h *= 60;
        }
        return { h, s, v };
      }
    }
    return CustomColorPickerMenu;
  }

  // Helper: Get global table index across entire document
  getGlobalTableIndex(tableEl) {
    // First, check if this table already has a stored index (from processSingleTable)
    const storedIndex = tableEl.getAttribute("data-ctc-index");
    if (storedIndex !== null) {
      debugLog(`Using stored table index: ${storedIndex}`);
      return parseInt(storedIndex, 10);
    }

    // Fallback: only count tables in the same view/file as this table
    const isInPreview = !!tableEl.closest(".markdown-preview-view");
    const isInEditor = !!tableEl.closest(".cm-content");

    if (isInPreview) {
      // In reading mode, get tables only from this preview view
      const previewView = tableEl.closest(".markdown-preview-view");
      if (previewView) {
        const tablesInThisView = Array.from(
          previewView.querySelectorAll("table"),
        );
        return tablesInThisView.indexOf(tableEl);
      }
    }

    if (isInEditor) {
      // In live preview, get tables only from this editor
      const editor =
        tableEl.closest(".cm-content") || tableEl.closest(".cm-editor");
      if (editor) {
        const tablesInThisEditor = Array.from(editor.querySelectorAll("table"));
        return tablesInThisEditor.indexOf(tableEl);
      }
    }

    // Last resort: use global index
    const allDocTables = Array.from(document.querySelectorAll("table"));
    return allDocTables.indexOf(tableEl);
  }

  async pickColor(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = null;
    // Use the cell as anchor for menu position
    const menu = new CustomColorPickerMenu(
      this,
      async (pickedColor) => {
        // Save color on close
        const fileId = this.app.workspace.getActiveFile()?.path;
        if (!fileId) return;
        // Use GLOBAL table index, not local
        const tableIndex = this.getGlobalTableIndex(tableEl);
        const rowIndex = Array.from(tableEl.querySelectorAll("tr")).indexOf(
          cell.closest("tr"),
        );
        const colIndex = Array.from(
          cell.closest("tr").querySelectorAll("td, th"),
        ).indexOf(cell);

        // Capture current state for undo
        const oldColors =
          this.cellData[fileId]?.[`table_${tableIndex}`]?.[`row_${rowIndex}`]?.[
            `col_${colIndex}`
          ];

        if (!this.cellData[fileId]) this.cellData[fileId] = {};
        const noteData = this.cellData[fileId];
        const tableKey = `table_${tableIndex}`;
        if (!noteData[tableKey]) noteData[tableKey] = {};
        const tableColors = noteData[tableKey];
        const rowKey = `row_${rowIndex}`;
        if (!tableColors[rowKey]) tableColors[rowKey] = {};
        const colKey = `col_${colIndex}`;

        const newColors = {
          ...tableColors[rowKey][colKey],
          [type]: pickedColor,
        };

        tableColors[rowKey][colKey] = newColors;

        // Add to undo stack
        const snapshot = this.createSnapshot(
          "cell_color",
          fileId,
          tableIndex,
          { row: rowIndex, col: colIndex },
          oldColors,
          newColors,
        );
        this.addToUndoStack(snapshot);
        this.updateRecentColor(pickedColor);

        await this.saveDataColors();
      },
      initialColor,
      cell,
    );
    menu._cell = cell;
    menu._type = type;
    menu.open();
  }

  async pickColorForRow(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = null;
    const row = cell.closest("tr");
    const rowCells = Array.from(row.querySelectorAll("td, th"));

    const menu = new CustomColorPickerMenu(
      this,
      async (pickedColor) => {
        // Save color on close for entire row
        const fileId = this.app.workspace.getActiveFile()?.path;
        if (!fileId) return;
        // Use GLOBAL table index, not local
        const tableIndex = this.getGlobalTableIndex(tableEl);
        const rowIndex = Array.from(tableEl.querySelectorAll("tr")).indexOf(
          row,
        );

        // Capture current state for undo
        const oldColors =
          this.cellData[fileId]?.[`table_${tableIndex}`]?.[`row_${rowIndex}`];

        if (!this.cellData[fileId]) this.cellData[fileId] = {};
        const noteData = this.cellData[fileId];
        const tableKey = `table_${tableIndex}`;
        if (!noteData[tableKey]) noteData[tableKey] = {};
        const tableColors = noteData[tableKey];
        const rowKey = `row_${rowIndex}`;
        if (!tableColors[rowKey]) tableColors[rowKey] = {};

        // Color all cells in the row
        const newColors = {};
        rowCells.forEach((rowCell, colIndex) => {
          const colKey = `col_${colIndex}`;
          newColors[colKey] = {
            ...tableColors[rowKey][colKey],
            [type]: pickedColor,
          };
          tableColors[rowKey][colKey] = newColors[colKey];
        });

        // Add to undo stack
        const snapshot = this.createSnapshot(
          "row_color",
          fileId,
          tableIndex,
          { row: rowIndex },
          oldColors,
          newColors,
        );
        this.addToUndoStack(snapshot);
        this.updateRecentColor(pickedColor);

        await this.saveDataColors();
      },
      initialColor,
      cell,
    );

    menu._cells = rowCells;
    menu._type = type;
    menu.open();
  }

  async pickColorForColumn(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = null;
    const colIndex = Array.from(
      cell.closest("tr").querySelectorAll("td, th"),
    ).indexOf(cell);

    // Collect all cells in the column for preview
    const columnCells = [];
    tableEl.querySelectorAll("tr").forEach((row) => {
      const cells = row.querySelectorAll("td, th");
      if (colIndex < cells.length) {
        columnCells.push(cells[colIndex]);
      }
    });

    const menu = new CustomColorPickerMenu(
      this,
      async (pickedColor) => {
        // Save color on close for entire column
        const fileId = this.app.workspace.getActiveFile()?.path;
        if (!fileId) return;
        // Use GLOBAL table index, not local
        const tableIndex = this.getGlobalTableIndex(tableEl);

        // Capture current state for undo
        const oldColors = {};
        if (this.cellData[fileId]?.[`table_${tableIndex}`]) {
          Object.entries(this.cellData[fileId][`table_${tableIndex}`]).forEach(
            ([rowKey, rowData]) => {
              if (rowKey.startsWith("row_") && rowData[`col_${colIndex}`]) {
                oldColors[rowKey] = { ...oldColors[rowKey] };
                oldColors[rowKey][`col_${colIndex}`] = {
                  ...rowData[`col_${colIndex}`],
                };
              }
            },
          );
        }

        if (!this.cellData[fileId]) this.cellData[fileId] = {};
        const noteData = this.cellData[fileId];
        const tableKey = `table_${tableIndex}`;
        if (!noteData[tableKey]) noteData[tableKey] = {};
        const tableColors = noteData[tableKey];

        // Color all cells in the column
        const newColors = {};
        tableEl.querySelectorAll("tr").forEach((row, rowIndex) => {
          const cells = row.querySelectorAll("td, th");
          if (colIndex < cells.length) {
            const rowKey = `row_${rowIndex}`;
            if (!tableColors[rowKey]) tableColors[rowKey] = {};
            const colKey = `col_${colIndex}`;
            newColors[rowKey] = { ...newColors[rowKey] };
            newColors[rowKey][colKey] = {
              ...tableColors[rowKey][colKey],
              [type]: pickedColor,
            };
            tableColors[rowKey][colKey] = newColors[rowKey][colKey];
          }
        });

        // Add to undo stack
        const snapshot = this.createSnapshot(
          "column_color",
          fileId,
          tableIndex,
          { col: colIndex },
          oldColors,
          newColors,
        );
        this.addToUndoStack(snapshot);
        this.updateRecentColor(pickedColor);

        await this.saveDataColors();
      },
      initialColor,
      cell,
    );

    menu._cells = columnCells;
    menu._type = type;
    menu.open();
  }

  async resetCell(cell, tableEl) {
    // Clear inline styles
    cell.style.backgroundColor = "";
    cell.style.color = "";

    // Remove data attributes that restore colors
    cell.removeAttribute("data-ctc-bg");
    cell.removeAttribute("data-ctc-color");
    cell.removeAttribute("data-ctc-manual");

    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

    const tableIndex = this.getGlobalTableIndex(tableEl);
    const rowIndex = Array.from(tableEl.querySelectorAll("tr")).indexOf(
      cell.closest("tr"),
    );
    const colIndex = Array.from(
      cell.closest("tr").querySelectorAll("td, th"),
    ).indexOf(cell);

    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (noteData?.[tableKey]?.[`row_${rowIndex}`]) {
      delete noteData[tableKey][`row_${rowIndex}`][`col_${colIndex}`];
      await this.saveDataColors();
    }

    // Force refresh to ensure no colors reappear
    window.setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async resetRow(cell, tableEl) {
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

    const tableIndex = this.getGlobalTableIndex(tableEl);
    const rowIndex = Array.from(tableEl.querySelectorAll("tr")).indexOf(
      cell.closest("tr"),
    );

    const noteData = this.cellData[fileId] || {};
    const tableKey = `table_${tableIndex}`;
    const tableColors = noteData[tableKey] || {};
    const rowKey = `row_${rowIndex}`;
    const oldColors = tableColors[rowKey]
      ? { ...tableColors[rowKey] }
      : undefined;

    if (!this.cellData[fileId]) this.cellData[fileId] = {};
    if (!this.cellData[fileId][tableKey]) this.cellData[fileId][tableKey] = {};
    delete this.cellData[fileId][tableKey][rowKey];

    // Clear all cells in the row
    cell
      .closest("tr")
      ?.querySelectorAll("td, th")
      ?.forEach((td) => {
        td.style.backgroundColor = "";
        td.style.color = "";
        // Remove data attributes
        td.removeAttribute("data-ctc-bg");
        td.removeAttribute("data-ctc-color");
        td.removeAttribute("data-ctc-manual");
      });

    const snapshot = this.createSnapshot(
      "row_color",
      fileId,
      tableIndex,
      { row: rowIndex },
      oldColors,
      undefined,
    );
    this.addToUndoStack(snapshot);

    await this.saveDataColors();

    // Force refresh to ensure no colors reappear
    window.setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async resetColumn(cell, tableEl) {
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

    const tableIndex = this.getGlobalTableIndex(tableEl);
    const colIndex = Array.from(
      cell.closest("tr").querySelectorAll("td, th"),
    ).indexOf(cell);

    if (!this.cellData[fileId]) this.cellData[fileId] = {};
    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (!noteData[tableKey]) noteData[tableKey] = {};
    const tableColors = noteData[tableKey];

    const oldColors = {};
    Object.entries(tableColors).forEach(([rk, rowData]) => {
      if (rk.startsWith("row_") && rowData[`col_${colIndex}`]) {
        oldColors[rk] = {
          [`col_${colIndex}`]: { ...rowData[`col_${colIndex}`] },
        };
        delete rowData[`col_${colIndex}`];
      }
    });

    // Clear all cells in the column and remove data attributes
    tableEl.querySelectorAll("tr").forEach((tr) => {
      const cells = tr.querySelectorAll("td, th");
      if (colIndex < cells.length) {
        const c = cells[colIndex];
        c.style.backgroundColor = "";
        c.style.color = "";
        // Remove data attributes
        c.removeAttribute("data-ctc-bg");
        c.removeAttribute("data-ctc-color");
        c.removeAttribute("data-ctc-manual");
      }
    });

    const snapshot = this.createSnapshot(
      "column_color",
      fileId,
      tableIndex,
      { col: colIndex },
      oldColors,
      undefined,
    );
    this.addToUndoStack(snapshot);

    await this.saveDataColors();

    // Force refresh to ensure no colors reappear
    window.setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async loadDataSettings() {
    return (await this.loadData())?.settings || {};
  }

  async loadDataColors() {
    return (await this.loadData())?.cellData || {};
  }

  async saveDataSettings() {
    await this.saveSettings();
    // Refresh any visible rendered notes
    // so rule changes take effect immediately
    this.applyColorsToActiveFile();
    if (
      this.settings.livePreviewColoring &&
      typeof this.applyColorsToAllEditors === "function"
    ) {
      window.setTimeout(() => this.applyColorsToAllEditors(), 10);
    }
  }

  async saveDataColors() {
    await this.saveData({ settings: this.settings, cellData: this.cellData });
    // Sync colors to both reading and live preview modes
    // Use setTimeout to ensure DOM is ready
    window.setTimeout(() => this.applyColorsToActiveFile(), 10);
    // Also explicitly apply to live preview editor if enabled
    if (
      this.settings.livePreviewColoring &&
      typeof this.applyColorsToAllEditors === "function"
    ) {
      window.setTimeout(() => this.applyColorsToAllEditors(), 20);
    }
  }

  // Create snapshot of current state for undo/redo
  createSnapshot(
    operationType,
    filePath,
    tableIndex,
    coordinates,
    oldColors,
    newColors,
  ) {
    return {
      timestamp: Date.now(),
      operationType,
      filePath,
      tableIndex,
      coordinates,
      oldColors,
      newColors,
    };
  }

  // Add operation to undo stack
  addToUndoStack(snapshot) {
    this.undoStack.push(snapshot);
    // Clear redo stack when new operation is performed
    this.redoStack = [];
    // Limit stack size
    if (this.undoStack.length > this.maxStackSize) {
      this.undoStack.shift();
    }
    if (this.settings?.persistUndoHistory) {
      this.saveUndoRedoStacks();
    }
  }

  // Undo last operation
  async undo() {
    if (this.undoStack.length === 0) return;

    const snapshot = this.undoStack.pop();
    this.redoStack.push(snapshot);

    // Apply the reverse operation
    const { filePath, tableIndex, coordinates, oldColors } = snapshot;

    if (!this.cellData[filePath]) this.cellData[filePath] = {};
    const noteData = this.cellData[filePath];
    const tableKey = `table_${tableIndex}`;

    if (!noteData[tableKey]) noteData[tableKey] = {};
    const tableColors = noteData[tableKey];

    // Restore old colors
    if (coordinates.row !== undefined && coordinates.col !== undefined) {
      // Single cell operation
      const rowKey = `row_${coordinates.row}`;
      if (!tableColors[rowKey]) tableColors[rowKey] = {};
      const colKey = `col_${coordinates.col}`;

      if (oldColors) {
        tableColors[rowKey][colKey] = { ...oldColors };
      } else {
        delete tableColors[rowKey][colKey];
      }
    } else if (coordinates.row !== undefined) {
      // Row operation
      const rowKey = `row_${coordinates.row}`;
      if (oldColors) {
        tableColors[rowKey] = { ...oldColors };
      } else {
        delete tableColors[rowKey];
      }
    } else if (coordinates.col !== undefined) {
      // Column operation - need to iterate through all rows
      for (const [rowKey, rowData] of Object.entries(tableColors)) {
        if (rowKey.startsWith("row_")) {
          const colKey = `col_${coordinates.col}`;
          if (oldColors && oldColors[rowKey]) {
            rowData[colKey] = { ...oldColors[rowKey][colKey] };
          } else {
            delete rowData[colKey];
          }
        }
      }
    }

    await this.saveDataColors();
    if (this.settings?.persistUndoHistory) {
      this.saveUndoRedoStacks();
    }
  }

  // Redo last undone operation
  async redo() {
    if (this.redoStack.length === 0) return;

    const snapshot = this.redoStack.pop();
    this.undoStack.push(snapshot);

    // Apply the original operation
    const { filePath, tableIndex, coordinates, newColors } = snapshot;

    if (!this.cellData[filePath]) this.cellData[filePath] = {};
    const noteData = this.cellData[filePath];
    const tableKey = `table_${tableIndex}`;

    if (!noteData[tableKey]) noteData[tableKey] = {};
    const tableColors = noteData[tableKey];

    // Apply new colors
    if (coordinates.row !== undefined && coordinates.col !== undefined) {
      // Single cell operation
      const rowKey = `row_${coordinates.row}`;
      if (!tableColors[rowKey]) tableColors[rowKey] = {};
      const colKey = `col_${coordinates.col}`;

      if (newColors) {
        tableColors[rowKey][colKey] = { ...newColors };
      } else {
        delete tableColors[rowKey][colKey];
      }
    } else if (coordinates.row !== undefined) {
      // Row operation
      const rowKey = `row_${coordinates.row}`;
      if (newColors) {
        tableColors[rowKey] = { ...newColors };
      } else {
        delete tableColors[rowKey];
      }
    } else if (coordinates.col !== undefined) {
      // Column operation - need to iterate through all rows
      for (const [rowKey, rowData] of Object.entries(tableColors)) {
        if (rowKey.startsWith("row_")) {
          const colKey = `col_${coordinates.col}`;
          if (newColors && newColors[rowKey]) {
            rowData[colKey] = { ...newColors[rowKey][colKey] };
          } else {
            delete rowData[colKey];
          }
        }
      }
    }

    await this.saveDataColors();
    if (this.settings?.persistUndoHistory) {
      this.saveUndoRedoStacks();
    }
  }

  async saveUndoRedoStacks() {
    try {
      localStorage.setItem(
        "table-color-undo-stack",
        JSON.stringify(this.undoStack),
      );
      localStorage.setItem(
        "table-color-redo-stack",
        JSON.stringify(this.redoStack),
      );
    } catch (e) {}
  }

  async loadUndoRedoStacks() {
    try {
      const u = localStorage.getItem("table-color-undo-stack");
      const r = localStorage.getItem("table-color-redo-stack");
      if (u) this.undoStack = JSON.parse(u) || [];
      if (r) this.redoStack = JSON.parse(r) || [];
    } catch (e) {}
  }

  // Apply rule and saved colors to a DOM container
  // Helper function to get visible text content from a cell, excluding editor markup and cursors
  getCellText(cell) {
    // For table cells in reading mode
    let text = "";

    // Get text from all child nodes, excluding editor elements
    const walkNodes = (node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        text += node.textContent;
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        // Skip CodeMirror cursor and editor elements
        if (
          node.classList &&
          (node.classList.contains("cm-cursor") ||
            node.classList.contains("cm-line"))
        ) {
          return; // Skip editor cursors
        }
        // For links and images, extract meaningful text
        if (node.tagName === "A") {
          // Extract link text (display text only, not the full href which may contain internal Obsidian paths)
          const linkText = node.textContent.trim();
          if (linkText) text += linkText;
          return; // Don't recurse into link, we got the text
        }
        if (node.tagName === "IMG") {
          // Extract image alt text (not the src which may contain internal paths)
          const alt = node.getAttribute("alt") || "";
          if (alt) text += alt;
          return; // Don't recurse into img
        }
        // Include br tags as newlines
        if (node.tagName === "BR") {
          text += "\n";
        } else {
          // Recurse into child nodes
          for (let child of node.childNodes) {
            walkNodes(child);
          }
        }
      }
    };

    for (let node of cell.childNodes) {
      walkNodes(node);
    }

    return text.trim();
  }

  // Helper method to check if element is visible in viewport
  isElementVisible(element) {
    if (!element || !element.getBoundingClientRect) return false;
    const rect = element.getBoundingClientRect();
    return (
      rect.top >= 0 &&
      rect.left >= 0 &&
      rect.bottom <=
        (window.innerHeight || document.documentElement.clientHeight) &&
      rect.right <= (window.innerWidth || document.documentElement.clientWidth)
    );
  }

  // Optimized method to apply all rules to a single cell
  applyRulesToCell(cell, cellText, colorData) {
    if (colorData) {
      if (colorData.bg) cell.style.backgroundColor = colorData.bg;
      if (colorData.color) cell.style.color = colorData.color;
    }
  }

  // Evaluate a value against a rule's MATCH type
  evaluateMatch(text, rule) {
    const val = rule.value ?? "";
    const t = (text ?? "").trim();
    const isEmpty = t.length === 0;

    const toNumber = (s) => {
      const cleaned = String(s).replace(/,/g, "");
      const n = parseFloat(cleaned);
      if (this.settings.numericStrict) {
        const ok = /^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$/.test(
          String(s).trim(),
        );
        return ok ? n : NaN;
      }
      return isNaN(n) ? NaN : n;
    };

    switch (rule.match) {
      case "is":
        return t.toLowerCase() === String(val).toLowerCase();
      case "isNot":
        return t.toLowerCase() !== String(val).toLowerCase();
      case "isRegex":
        try {
          const rx = new RegExp(String(val), "i");
          return rx.test(t);
        } catch {
          return false;
        }
      case "contains":
        return t.toLowerCase().includes(String(val).toLowerCase());
      case "notContains":
        return !t.toLowerCase().includes(String(val).toLowerCase());
      case "startsWith":
        return t.toLowerCase().startsWith(String(val).toLowerCase());
      case "endsWith":
        return t.toLowerCase().endsWith(String(val).toLowerCase());
      case "notStartsWith":
        return !t.toLowerCase().startsWith(String(val).toLowerCase());
      case "notEndsWith":
        return !t.toLowerCase().endsWith(String(val).toLowerCase());
      case "isEmpty":
        return isEmpty;
      case "isNotEmpty":
        return !isEmpty;
      case "eq": {
        const n = toNumber(t);
        const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n === v;
      }
      case "gt": {
        const n = toNumber(t);
        const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n > v;
      }
      case "lt": {
        const n = toNumber(t);
        const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n < v;
      }
      case "ge": {
        const n = toNumber(t);
        const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n >= v;
      }
      case "le": {
        const n = toNumber(t);
        const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n <= v;
      }
      default:
        return false;
    }
  }

  applyColoringRulesToTable(tableEl) {
    const rules = Array.isArray(this.settings.coloringRules)
      ? this.settings.coloringRules
      : [];
    if (!rules.length) return;

    const rows = Array.from(tableEl.querySelectorAll("tr"));
    const getCell = (r, c) => {
      const row = rows[r];
      if (!row) return null;
      const cells = Array.from(row.querySelectorAll("td, th"));
      return cells[c] || null;
    };
    const texts = rows.map((row) =>
      Array.from(row.querySelectorAll("td, th")).map((cell) =>
        this.getCellText(cell),
      ),
    );
    const maxCols = Math.max(0, ...texts.map((r) => r.length));

    const headerRowIndex = rows.findIndex((r) => r.querySelector("th"));
    const hdr = headerRowIndex >= 0 ? headerRowIndex : 0;
    const firstDataRowIndex = rows.findIndex((r) => r.querySelector("td"));
    const fdr = firstDataRowIndex >= 0 ? firstDataRowIndex : 0;
    debugLog(
      `[Coloring Rules Setup] Total rows=${rows.length}, maxCols=${maxCols}, headerRowIndex=${headerRowIndex}, hdr=${hdr}, fdr=${fdr}`,
    );
    debugLog(`[Header Row Text] Row ${hdr}:`, texts[hdr]);

    for (const rule of rules) {
      if (!rule || !rule.target || !rule.match) continue;
      const applyCellStyle = (cell) => {
        if (!cell) return;
        if (cell.hasAttribute("data-ctc-manual")) return;
        // For header cells (th), allow overriding existing colors from rules
        // For data cells (td), check if they already have colors
        const isHeaderCell = cell.tagName === "TH";
        if (!isHeaderCell && (cell.style.backgroundColor || cell.style.color))
          return;
        if (rule.bg) cell.style.backgroundColor = rule.bg;
        if (rule.color) cell.style.color = rule.color;
      };

      if (rule.target === "cell") {
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < (texts[r]?.length || 0); c++) {
            const text = texts[r][c];
            if ((rule.when || "theCell") === "theCell") {
              if (this.evaluateMatch(text, rule)) {
                applyCellStyle(getCell(r, c));
              }
            }
          }
        }
      } else if (rule.target === "row") {
        const candidateRows =
          rule.when === "firstRow"
            ? [fdr]
            : Array.from({ length: rows.length }, (_, i) => i);
        for (const r of candidateRows) {
          const rowTexts = texts[r] || [];
          let cond = false;
          if (rule.when === "allCell")
            cond =
              rowTexts.length > 0 &&
              rowTexts.every((t) => this.evaluateMatch(t, rule));
          else if (rule.when === "noCell")
            cond = rowTexts.every((t) => !this.evaluateMatch(t, rule));
          else cond = rowTexts.some((t) => this.evaluateMatch(t, rule)); // anyCell or firstRow default any
          if (cond) {
            const cells = Array.from(rows[r].querySelectorAll("td, th"));
            cells.forEach(applyCellStyle);
          }
        }
      } else if (rule.target === "column") {
        const candidateCols = Array.from({ length: maxCols }, (_, i) => i);
        for (const c of candidateCols) {
          let cond = false;
          if (rule.when === "columnHeader") {
            const text = texts[hdr]?.[c] ?? "";
            debugLog(
              `[Column Header Match] Rule: match=${rule.match}, value="${rule.value}", headerRow=${hdr}, col=${c}, text="${text}"`,
            );
            cond = this.evaluateMatch(text, rule);
            debugLog(`[Column Header Result] Column ${c}: match=${cond}`);
          } else {
            const colTexts = Array.from(
              { length: rows.length },
              (_, r) => texts[r]?.[c],
            ).filter((t) => t !== undefined);
            if (rule.when === "allCell")
              cond =
                colTexts.length > 0 &&
                colTexts.every((t) => this.evaluateMatch(t, rule));
            else if (rule.when === "noCell")
              cond = colTexts.every((t) => !this.evaluateMatch(t, rule));
            else cond = colTexts.some((t) => this.evaluateMatch(t, rule)); // anyCell default
          }
          if (cond) {
            for (let r = 0; r < rows.length; r++) applyCellStyle(getCell(r, c));
          }
        }
      }
    }
  }

  applyAdvancedRulesToTable(tableEl) {
    const adv = Array.isArray(this.settings.advancedRules)
      ? this.settings.advancedRules
      : [];
    if (!adv.length) return;
    const rows = Array.from(tableEl.querySelectorAll("tr"));
    const texts = rows.map((row) =>
      Array.from(row.querySelectorAll("td, th")).map((cell) =>
        this.getCellText(cell),
      ),
    );
    const maxCols = Math.max(0, ...texts.map((r) => r.length));
    const headerRowIndex = rows.findIndex((r) => r.querySelector("th"));
    const hdr = headerRowIndex >= 0 ? headerRowIndex : 0;
    const firstDataRowIndex = rows.findIndex((r) => r.querySelector("td"));
    const fdr = firstDataRowIndex >= 0 ? firstDataRowIndex : 0;
    debugLog(
      `[Advanced Rules Setup] Total rows=${rows.length}, maxCols=${maxCols}, headerRowIndex=${headerRowIndex}, hdr=${hdr}, fdr=${fdr}`,
    );
    debugLog(`[Header Row Text] Row ${hdr}:`, texts[hdr]);
    const getCell = (r, c) => {
      const row = rows[r];
      if (!row) return null;
      const cells = Array.from(row.querySelectorAll("td, th"));
      return cells[c] || null;
    };

    const evalCondCell = (r, c, cond) =>
      this.evaluateMatch(texts[r]?.[c] ?? "", {
        match: cond.match,
        value: cond.value,
      });
    const evalCondRow = (r, cond) => {
      const rowTexts = texts[r] || [];
      return rowTexts.some((t) =>
        this.evaluateMatch(t, { match: cond.match, value: cond.value }),
      );
    };
    const evalCondHeader = (c, cond) => {
      const headerText = texts[hdr]?.[c] ?? "";
      const result = this.evaluateMatch(headerText, {
        match: cond.match,
        value: cond.value,
      });
      debugLog(
        `[Advanced Column Header] Col ${c}, match=${cond.match}, value="${cond.value}", headerText="${headerText}", result=${result}`,
      );
      return result;
    };

    for (const rule of adv) {
      const logic = rule.logic || "any";
      const target = rule.target || "cell";
      const color = rule.color || null;
      const bg = rule.bg || null;
      if (!bg && !color) continue;
      const conditions = Array.isArray(rule.conditions) ? rule.conditions : [];
      if (!conditions.length) continue;

      if (target === "row") {
        const hasHeaderConds = conditions.some(
          (cond) => cond.when === "columnHeader",
        );
        const allHeaderConds =
          conditions.length > 0 &&
          conditions.every((cond) => cond.when === "columnHeader");
        let candidateRows =
          rule.when === "firstRow"
            ? [fdr]
            : Array.from({ length: rows.length }, (_, i) => i);
        if (allHeaderConds) {
          candidateRows = [hdr];
          debugLog(
            `Row target uses only columnHeader conditions; coloring header row index ${hdr}`,
          );
        }
        for (const r of candidateRows) {
          const flags = conditions.map((cond) => {
            if (cond.when === "columnHeader") {
              // For row target with column header condition, check if ANY column header matches
              for (let c = 0; c < maxCols; c++) {
                if (evalCondHeader(c, cond)) return true;
              }
              return false;
            } else if (cond.when === "row") {
              return evalCondRow(r, cond);
            } else {
              // anyCell, allCell, noCell for row target
              const cells = Array.from(rows[r].querySelectorAll("td, th"));
              const cellResults = cells.map((_, c) => evalCondCell(r, c, cond));
              if (cond.when === "allCell") return cellResults.every(Boolean);
              if (cond.when === "noCell") return cellResults.every((f) => !f);
              return cellResults.some(Boolean); // anyCell
            }
          });
          let ok = false;
          if (logic === "all") ok = flags.every(Boolean);
          else if (logic === "none") ok = flags.every((f) => !f);
          else ok = flags.some(Boolean);
          debugLog(
            `Row ${r}: flags=${JSON.stringify(flags)}, logic=${logic}, ok=${ok}, bg=${bg}, color=${color}`,
          );
          if (ok) {
            debugLog(`  -> Coloring row ${r} with bg=${bg} color=${color}`);
            Array.from(rows[r].querySelectorAll("td, th")).forEach((cell) => {
              if (cell.hasAttribute("data-ctc-manual")) return;
              // For header cells, allow overriding; for data cells, skip if already colored
              const isHeaderCell = cell.tagName === "TH";
              if (
                !isHeaderCell &&
                (cell.style.backgroundColor || cell.style.color)
              )
                return;
              if (bg) cell.style.backgroundColor = bg;
              if (color) cell.style.color = color;
            });
          }
        }
      } else if (target === "column") {
        for (let c = 0; c < maxCols; c++) {
          const flags = conditions.map((cond) => {
            if (cond.when === "columnHeader") return evalCondHeader(c, cond);
            if (cond.when === "row") return false;
            // anyCell, allCell, noCell for column
            const colCells = [];
            for (let r = 0; r < rows.length; r++) {
              colCells.push(evalCondCell(r, c, cond));
            }
            if (cond.when === "allCell") return colCells.every(Boolean);
            if (cond.when === "noCell") return colCells.every((f) => !f);
            return colCells.some(Boolean); // anyCell
          });
          let ok = false;
          if (logic === "all") ok = flags.every(Boolean);
          else if (logic === "none") ok = flags.every((f) => !f);
          else ok = flags.some(Boolean);
          if (ok) {
            for (let r = 0; r < rows.length; r++) {
              const cell = getCell(r, c);
              if (
                cell &&
                !cell.hasAttribute("data-ctc-manual") &&
                !cell.style.backgroundColor &&
                !cell.style.color
              ) {
                if (bg) cell.style.backgroundColor = bg;
                if (color) cell.style.color = color;
              }
            }
          }
        }
      } else {
        // target === 'cell' - color individual cells
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < (texts[r]?.length || 0); c++) {
            const flags = conditions.map((cond) => {
              if (cond.when === "columnHeader") {
                // For cell target: check if the column header matches
                return evalCondHeader(c, cond);
              } else if (cond.when === "row") {
                // For cell target: check if any cell in the row matches
                return evalCondRow(r, cond);
              } else if (
                cond.when === "anyCell" ||
                cond.when === "allCell" ||
                cond.when === "noCell"
              ) {
                // For cell target with these conditions, evaluate the specific cell
                const cellMatch = evalCondCell(r, c, cond);
                // Note: allCell and noCell don't make sense for individual cells, but we'll treat them as the cell match
                return cellMatch;
              } else {
                // Default: check the specific cell
                return evalCondCell(r, c, cond);
              }
            });
            let ok = false;
            if (logic === "all") ok = flags.every(Boolean);
            else if (logic === "none") ok = flags.every((f) => !f);
            else ok = flags.some(Boolean);
            if (ok) {
              const cell = getCell(r, c);
              if (
                cell &&
                !cell.hasAttribute("data-ctc-manual") &&
                !cell.style.backgroundColor &&
                !cell.style.color
              ) {
                if (bg) cell.style.backgroundColor = bg;
                if (color) cell.style.color = color;
              }
            }
          }
        }
      }
    }
  }

  restoreColorsFromAttributes(element) {
    if (!element || !element.style) return;

    if (element.hasAttribute("data-ctc-bg")) {
      const bg = element.getAttribute("data-ctc-bg");
      if (bg && bg !== element.style.backgroundColor) {
        element.style.backgroundColor = bg;
      }
    }
    if (element.hasAttribute("data-ctc-color")) {
      const color = element.getAttribute("data-ctc-color");
      if (color && color !== element.style.color) {
        element.style.color = color;
      }
    }
  }

  applyColorsToContainer(container, filePath) {
    // Only apply colors in Reading mode or Live Preview if enabled
    const hasClosest = typeof container.closest === "function";
    const inPreview =
      hasClosest && !!container.closest(".markdown-preview-view");
    const inEditor =
      hasClosest &&
      (container.closest(".cm-content") ||
        container.closest(".cm-editor") ||
        container.closest(".cm-scroller"));
    if (!inPreview && (!this.settings.livePreviewColoring || !inEditor)) {
      // Ensure container is within a preview/editor container
      let p = container && container.parentElement;
      let found = false;
      while (p) {
        if (p.classList && p.classList.contains("markdown-preview-view")) {
          found = true;
          break;
        }
        if (
          this.settings.livePreviewColoring &&
          p.classList &&
          (p.classList.contains("cm-content") ||
            p.classList.contains("cm-editor") ||
            p.classList.contains("cm-scroller"))
        ) {
          found = true;
          break;
        }
        p = p.parentElement;
      }
      if (!found) {
        // In Live Preview, allow direct .cm-content root
        if (
          this.settings.livePreviewColoring &&
          container.classList &&
          container.classList.contains("cm-content")
        ) {
          // ok
        } else {
          return;
        }
      }
    }
    if (inEditor && !this.settings.livePreviewColoring) {
      return;
    }
    // In Live Preview, always re-apply colors
    // In Reading mode, only apply if in preview
    if (
      !inPreview &&
      !inEditor &&
      !(container.classList && container.classList.contains("cm-content"))
    ) {
      return;
    }

    // Performance optimization: Debounce rapid calls
    const now = Date.now();
    const lastCall = this._lastApplyCall || 0;
    if (now - lastCall < 100) {
      // Debounce to 100ms
      return;
    }
    this._lastApplyCall = now;
    debugLog(`=== applyColorsToContainer called ===`);
    debugLog(`container:`, container);
    // Special handling for reading mode: always process all tables in the view
    if (inPreview) {
      // Check if this is a reading view container
      const readingView = container.closest(".markdown-preview-view");
      if (readingView) {
        // Process ALL tables in the reading view, not just in this container
        const allTablesInView = Array.from(
          readingView.querySelectorAll("table"),
        );
        debugLog(
          `Reading mode: found ${allTablesInView.length} total tables in view`,
        );

        // If we have tables but they're not all in our container, process the entire view
        if (
          allTablesInView.length > 0 &&
          allTablesInView.some((table) => !container.contains(table))
        ) {
          debugLog("Processing entire reading view for comprehensive coloring");
          container = readingView; // Process the entire reading view
        }
      }
    }

    const tables = Array.from(container.querySelectorAll("table"));
    if (!tables.length) return;
    const noteData = this.cellData[filePath] || {};

    debugLog(
      `applyColorsToContainer: filePath=${filePath}, found ${tables.length} tables, has noteData:`,
      Object.keys(noteData).length > 0,
      "inPreview:",
      inPreview,
      "inEditor:",
      inEditor,
    );
    // Get ALL tables in the correct scope - ALWAYS use global table list
    let allTables = Array.from(document.querySelectorAll("table"));

    debugLog(
      `All tables in document: ${allTables.length}, container tables: ${tables.length}, inPreview: ${inPreview}, inEditor: ${inEditor}`,
    );

    // Process each table
    let manualAppliedCount = 0;
    tables.forEach((tableEl) => {
      const globalTableIndex = allTables.indexOf(tableEl);
      if (globalTableIndex === -1) {
        debugWarn(`Could not find table in global list, using local index`);
        const containerTables = Array.from(container.querySelectorAll("table"));
        const fallbackIndex = containerTables.indexOf(tableEl);
        manualAppliedCount +=
          this.processSingleTable(tableEl, fallbackIndex, filePath, noteData) ||
          0;
      } else {
        manualAppliedCount +=
          this.processSingleTable(
            tableEl,
            globalTableIndex,
            filePath,
            noteData,
          ) || 0;
      }
    });
    debugLog(
      `applyColorsToContainer: manual colors applied to ${manualAppliedCount} cells in this container`,
    );
    // In Reading mode, retry applying colors with escalating delays if needed
    try {
      const prev = this._appliedContainers.get(container) || 0;
      const delays = [100, 200, 400, 800];
      if (prev < delays.length) {
        this._appliedContainers.set(container, prev + 1);
        window.setTimeout(() => {
          // Only retry if still connected to DOM
          if (container.isConnected) {
            this.applyColorsToContainer(container, filePath);
          }
        }, delays[prev]);
      }
    } catch (e) {}
  }
  applyColorsToActiveFile() {
    debugLog("applyColorsToActiveFile called");
    const file = this.app.workspace.getActiveFile();
    if (!file) {
      debugWarn("No active file found in applyColorsToActiveFile");
      return;
    }

    const noteData = this.cellData[file.path] || {};
    debugLog(
      `applyColorsToActiveFile: file=${file.path}, has noteData keys:`,
      Object.keys(noteData),
    );

    // Restrict reading mode application to views showing the active file
    let previewViews = [];
    try {
      const leaves =
        typeof this.app.workspace.getLeavesOfType === "function"
          ? this.app.workspace.getLeavesOfType("markdown")
          : [];
      const activeContainers = leaves
        .filter((l) => l.view && l.view.file && l.view.file.path === file.path)
        .map((l) => l.view && (l.view.containerEl || l.view.contentEl))
        .filter(Boolean);
      activeContainers.forEach((container) => {
        const views = Array.from(
          container.querySelectorAll(".markdown-preview-view"),
        );
        previewViews.push(...views);
      });
    } catch (e) {
      previewViews = Array.from(
        document.querySelectorAll(".markdown-preview-view"),
      );
    }
    debugLog(
      `applyColorsToActiveFile: found ${previewViews.length} preview views for active file`,
    );

    // Get all tables from the active file (both reading mode and live preview)
    // Use file-scoped indices to ensure consistency across modes
    let fileTableIndex = 0;
    const fileTableMap = new Map(); // Maps table element to its file-scoped index

    // First pass: scan reading mode tables
    previewViews.forEach((view) => {
      view.querySelectorAll("table").forEach((table) => {
        if (!fileTableMap.has(table)) {
          fileTableMap.set(table, fileTableIndex);
          table.setAttribute("data-ctc-index", String(fileTableIndex));
          fileTableIndex++;
        }
      });
    });

    // Second pass: scan live preview tables
    if (this.settings.livePreviewColoring) {
      document.querySelectorAll(".cm-content table").forEach((table) => {
        if (!fileTableMap.has(table)) {
          fileTableMap.set(table, fileTableIndex);
          table.setAttribute("data-ctc-index", String(fileTableIndex));
          fileTableIndex++;
        }
      });
    }

    // If no tables found yet, retry shortly
    if (fileTableIndex === 0) {
      debugLog(
        "No tables found in any view for active file, retrying after 100ms",
      );
      window.setTimeout(() => this.applyColorsToActiveFile(), 100);
      return;
    }

    // Clear all rule-based colors first (in both reading and live preview)
    // This ensures rule color changes take effect
    previewViews.forEach((view) => {
      view.querySelectorAll("td, th").forEach((cell) => {
        if (!cell.hasAttribute("data-ctc-manual")) {
          cell.style.backgroundColor = "";
          cell.style.color = "";
        }
      });
    });

    if (this.settings.livePreviewColoring) {
      document
        .querySelectorAll(".cm-content table td, .cm-content table th")
        .forEach((cell) => {
          if (!cell.hasAttribute("data-ctc-manual")) {
            cell.style.backgroundColor = "";
            cell.style.color = "";
          }
        });
    }

    // Apply to reading mode using file-scoped indices
    previewViews.forEach((view) => {
      if (view.isConnected) {
        const viewTables = Array.from(view.querySelectorAll("table"));
        let readingManualApplied = 0;
        viewTables.forEach((table) => {
          const tableIdx = fileTableMap.get(table);
          if (tableIdx !== undefined) {
            readingManualApplied +=
              this.processSingleTable(table, tableIdx, file.path, noteData) ||
              0;
          }
        });
        debugLog(
          `Reading mode: manual colors applied to ${readingManualApplied} cells in this view (${viewTables.length} tables)`,
        );
      }
    });

    // Apply to live preview if enabled - use same file-scoped indices for consistency
    if (this.settings.livePreviewColoring) {
      const cmEditors = document.querySelectorAll(".cm-content");
      debugLog(
        `applyColorsToActiveFile: Found ${cmEditors.length} live preview editors`,
      );

      cmEditors.forEach((editor, editorIdx) => {
        if (editor.isConnected) {
          const editorTables = Array.from(editor.querySelectorAll("table"));
          debugLog(`  Editor ${editorIdx}: ${editorTables.length} tables`);
          let editorManualApplied = 0;
          editorTables.forEach((table) => {
            const tableIdx = fileTableMap.get(table);
            if (tableIdx !== undefined) {
              editorManualApplied +=
                this.processSingleTable(table, tableIdx, file.path, noteData) ||
                0;
            }
          });
          debugLog(
            `  Editor ${editorIdx}: manual colors applied to ${editorManualApplied} cells`,
          );
        }
      });
    }
  }

  // Get a signature for a table based on its structure (row/col count)
  getTableSignature(table) {
    const rows = table.querySelectorAll("tr").length;
    const cols = table.querySelector("tr")
      ? table.querySelector("tr").querySelectorAll("td, th").length
      : 0;
    return `${rows}x${cols}`;
  }

  // Get table text content for matching tables between reading and live preview modes
  getTableTextContent(table) {
    const rows = Array.from(table.querySelectorAll("tr"));
    const text = rows
      .map((row) => {
        const cells = Array.from(row.querySelectorAll("td, th"));
        return cells.map((cell) => this.getCellText(cell).trim()).join("|");
      })
      .join("\n");
    debugLog(
      `  getTableTextContent: ${rows.length} rows, content hash: ${text.substring(0, 100)}`,
    );
    return text;
  }

  processSingleTable(tableEl, tableIndex, filePath, noteData) {
    debugLog(
      `Processing single table: index=${tableIndex}, has data-ctc-index: ${tableEl.getAttribute("data-ctc-index")}`,
    );
    const inLivePreview = !!tableEl.closest(".cm-content");
    const inReadingMode = !!tableEl.closest(".markdown-preview-view");
    debugLog(
      `Table context: ${inLivePreview ? "LivePreview" : inReadingMode ? "Reading" : "Unknown"}`,
    );

    // Mark table with unique data attribute for persistence
    if (!tableEl.hasAttribute("data-ctc-processed")) {
      tableEl.setAttribute("data-ctc-processed", "true");
      tableEl.setAttribute("data-ctc-index", tableIndex);
      tableEl.setAttribute("data-ctc-file", filePath);
    }

    const tableKey = `table_${tableIndex}`;
    const tableColors = noteData[tableKey] || {};

    debugLog(
      `Table index ${tableIndex}: key="${tableKey}", has colors:`,
      Object.keys(tableColors).length > 0,
    );

    let coloredCount = 0;
    const tableId = `${filePath}:${tableIndex}`;
    const manualColorData = {}; // Store manual color data temporarily
    const cellsWithRuleColor = new Set(); // Track cells that got colors from rules

    // NOTE: Clearing is now done in applyColorsToActiveFile before this function is called
    // STEP 1: Store manual color data (but don't apply yet)
    Array.from(tableEl.querySelectorAll("tr")).forEach((tr, rIdx) => {
      const rowKey = `row_${rIdx}`;
      const rowColors = tableColors[rowKey] || {};
      const cells = Array.from(tr.querySelectorAll("td, th"));

      cells.forEach((cell, cIdx) => {
        const colorData = rowColors[`col_${cIdx}`];
        if (colorData) {
          const cellKey = `${rIdx}_${cIdx}`;
          manualColorData[cellKey] = colorData;

          // Store attributes for reference
          if (colorData.bg) {
            cell.setAttribute("data-ctc-bg", colorData.bg);
          }
          if (colorData.color) {
            cell.setAttribute("data-ctc-color", colorData.color);
          }
          cell.setAttribute("data-ctc-manual", "true");
          cell.setAttribute("data-ctc-table-id", tableId);
          cell.setAttribute("data-ctc-row", String(rIdx));
          cell.setAttribute("data-ctc-col", String(cIdx));
        }
      });
    });

    debugLog(
      `Table index ${tableIndex}: found ${Object.keys(manualColorData).length} cells with manual colors (stored, not applied yet)`,
    );

    // STEP 2: Apply rules FIRST (rules have priority)
    this.applyColoringRulesToTable(tableEl);
    this.applyAdvancedRulesToTable(tableEl);

    // Track which cells got rule colors
    Array.from(tableEl.querySelectorAll("tr")).forEach((tr, rIdx) => {
      const cells = Array.from(tr.querySelectorAll("td, th"));
      cells.forEach((cell, cIdx) => {
        if (cell.style.backgroundColor || cell.style.color) {
          cellsWithRuleColor.add(`${rIdx}_${cIdx}`);
        }
      });
    });

    // STEP 3: Apply manual colors ONLY to cells that don't have rule colors
    Array.from(tableEl.querySelectorAll("tr")).forEach((tr, rIdx) => {
      const cells = Array.from(tr.querySelectorAll("td, th"));
      cells.forEach((cell, cIdx) => {
        const cellKey = `${rIdx}_${cIdx}`;
        const colorData = manualColorData[cellKey];

        // Only apply manual color if this cell doesn't have a rule color
        if (colorData && !cellsWithRuleColor.has(cellKey)) {
          coloredCount++;

          // Apply colors directly
          if (colorData.bg) {
            cell.style.backgroundColor = colorData.bg;
          }
          if (colorData.color) {
            cell.style.color = colorData.color;
          }
        }
      });
    });

    // Ensure colors persist by marking the table as processed
    tableEl.setAttribute("data-ctc-last-processed", Date.now());

    return coloredCount;
  }

  setupReadingViewScrollListener() {
    // Watch for scroll events in reading view to catch lazy-loaded tables
    const handleReadingViewScroll = debounce(
      () => {
        const file = this.app.workspace.getActiveFile();
        if (!file) return;

        // Check if we're in reading mode
        const readingViews = document.querySelectorAll(
          ".markdown-preview-view",
        );
        if (readingViews.length === 0) return;

        debugLog("Reading view scroll detected, checking for new tables");
        readingViews.forEach((view) => {
          if (view.isConnected) {
            // Apply colors to the entire reading view (not just container)
            this.applyColorsToContainer(view, file.path);
          }
        });
      },
      150,
      true,
    );

    // Add scroll listener to all reading views
    const addScrollListeners = () => {
      document.querySelectorAll(".markdown-preview-view").forEach((view) => {
        if (!view._ctcScrollListenerAdded) {
          // Save handler reference for cleanup
          view._ctcScrollHandler = handleReadingViewScroll;
          view.addEventListener("scroll", view._ctcScrollHandler, {
            passive: true,
          });
          view._ctcScrollListenerAdded = true;
          debugLog("Added scroll listener to reading view");
        }
      });
    };

    // Initial setup
    addScrollListeners();

    // Watch for new reading views being created
    const readingViewObserver = new MutationObserver(() => {
      addScrollListeners();
    });

    readingViewObserver.observe(document.body, {
      childList: true,
      subtree: true,
    });
    this._readingViewScrollObserver = readingViewObserver;
  }

  startReadingModeTableChecker() {
    // Stop any existing checker
    try {
      if (this._readingModeChecker) {
        window.clearInterval(this._readingModeChecker);
        this._readingModeChecker = null;
      }
    } catch (e) {
      debugWarn("Error clearing existing reading mode checker:", e);
      this._readingModeChecker = null;
    }

    this._readingModeChecker = window.setInterval(() => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;

      // Restrict to reading views of the active file
      let previewViews = [];
      try {
        const leaves =
          typeof this.app.workspace.getLeavesOfType === "function"
            ? this.app.workspace.getLeavesOfType("markdown")
            : [];
        const activeContainers = leaves
          .filter(
            (l) => l.view && l.view.file && l.view.file.path === file.path,
          )
          .map((l) => l.view && (l.view.containerEl || l.view.contentEl))
          .filter(Boolean);
        activeContainers.forEach((container) => {
          const views = Array.from(
            container.querySelectorAll(".markdown-preview-view"),
          );
          previewViews.push(...views);
        });
      } catch (e) {
        previewViews = Array.from(
          document.querySelectorAll(".markdown-preview-view"),
        );
      }
      if (previewViews.length === 0) return;

      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll("table"));

      previewViews.forEach((view) => {
        const viewTables = Array.from(view.querySelectorAll("table"));
        viewTables.forEach((table) => {
          if (!table.hasAttribute("data-ctc-processed")) {
            const globalTableIdx = allDocTables.indexOf(table);
            this.processSingleTable(table, globalTableIdx, file.path, noteData);
          }
        });
        // Ensure rules are applied to all tables in this view
        viewTables.forEach((table) => {
          this.applyColoringRulesToTable(table);
          this.applyAdvancedRulesToTable(table);
        });
      });
    }, 500); // Check every 500ms for new tables that need coloring

    // Clean up on unload
    this.register(() => {
      try {
        if (this._readingModeChecker) {
          window.clearInterval(this._readingModeChecker);
          this._readingModeChecker = null;
        }
      } catch (e) {
        debugWarn("Error during reading mode checker cleanup:", e);
      }
    });
  }
};

// Note: RegexTesterModal removed - command referencing it is commented out (line 156)
// Re-add this class if you uncomment the command

class ConditionRow {
  constructor(parent, index, initialData, onChange) {
    this.root = parent.createDiv({ cls: "ctc-num-edit-row ctc-pretty-flex" });
    this.typeSel = this.root.createEl("select");
    ["text", "numeric", "date", "regex", "empty"].forEach((t) => {
      const o = this.typeSel.createEl("option");
      o.value = t;
      o.text = t;
    });
    this.opSel = this.root.createEl("select");
    this.valInput = this.root.createEl("input", {
      type: "text",
      cls: "ctc-condition-val-input",
    });
    this.val2Input = this.root.createEl("input", {
      type: "text",
      cls: "ctc-condition-val-input",
    });
    this.val2Input.style.display = "none";
    this.caseChk = this.root.createEl("input", { type: "checkbox" });
    this.logicSel = this.root.createEl("select");
    ["AND", "OR"].forEach((t) => {
      const o = this.logicSel.createEl("option");
      o.value = t;
      o.text = t;
    });
    const delBtn = this.root.createEl("button", { cls: "mod-ghost" });
    delBtn.textContent = "Delete";
    const setOps = () => {
      const t = this.typeSel.value;
      const currentOp = this.opSel.value; // Remember current selection
      this.opSel.empty();
      let options = [];
      if (t === "text")
        options = ["contains", "equals", "startsWith", "endsWith"];
      else if (t === "numeric")
        options = ["gt", "ge", "eq", "le", "lt", "between"];
      else if (t === "date") options = ["before", "after", "between"];
      else if (t === "regex") options = ["matches"];
      else options = ["isEmpty"];

      options.forEach((op) => {
        const o = this.opSel.createEl("option");
        o.value = op;
        o.text = op;
        // Preserve selection if it exists in new options
        if (currentOp && op === currentOp) {
          o.selected = true;
        }
      });

      // If current selection wasn't preserved, select first option
      if (!this.opSel.value && options.length > 0) {
        this.opSel.value = options[0];
      }

      const op = this.opSel.value;
      const useTwo = op === "between";
      this.val2Input.style.display = useTwo ? "" : "none";
      if (t === "numeric") {
        this.valInput.type = "number";
        this.val2Input.type = "number";
      } else if (t === "date") {
        this.valInput.type = "date";
        this.val2Input.type = "date";
      } else {
        this.valInput.type = "text";
        this.val2Input.type = "text";
      }
      this.caseChk.style.display = t === "text" || t === "regex" ? "" : "none";
      if (typeof onChange === "function") onChange();
    };
    this.typeSel.onchange = setOps;
    this.opSel.onchange = setOps;
    delBtn.onclick = () => {
      this.root.remove();
      if (typeof onChange === "function") onChange();
    };
    if (initialData) {
      this.typeSel.value = initialData.type || "text";
      setOps();
      if (initialData.operator) this.opSel.value = initialData.operator;
      if (initialData.value != null) this.valInput.value = initialData.value;
      if (initialData.value2 != null) {
        this.val2Input.value = initialData.value2;
        this.val2Input.style.display = "";
      }
      this.caseChk.checked = !!initialData.caseSensitive;
      if (initialData.logic)
        this.logicSel.value = initialData.logic.toUpperCase();
    } else {
      setOps();
    }
  }
  getData() {
    const t = this.typeSel.value;
    const op = this.opSel.value;
    const v =
      t === "numeric"
        ? this.valInput.value !== ""
          ? Number(this.valInput.value)
          : null
        : this.valInput.value;
    const v2 =
      t === "numeric" || t === "date"
        ? this.val2Input.style.display === ""
          ? this.val2Input.value !== ""
            ? t === "numeric"
              ? Number(this.val2Input.value)
              : this.val2Input.value
            : null
          : null
        : null;
    return {
      type: t,
      operator: op,
      value: v,
      value2: v2,
      caseSensitive: this.caseChk.checked,
      logic: this.logicSel.value,
    };
  }
}

// Settings Tab
// Release Notes Modal - displays latest release notes from GitHub
class ReleaseNotesModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this._mdComp = null;
  }

  async onOpen() {
    const { contentEl } = this;
    contentEl.empty();

    try {
      this.modalEl.addClass("ctc-release-modal");
    } catch (e) {}

    const header = contentEl.createEl("div", { cls: "ctc-release-header" });

    header.createEl("h2", {
      text: "Color table cells",
      cls: "ctc-release-title",
    });

    const link = header.createEl("a", {
      text: "View on GitHub",
      cls: "ctc-release-link",
    });
    link.href = "https://github.com/Kazi-Aidah/color-table-cells/releases";
    link.target = "_blank";

    const body = contentEl.createDiv({ cls: "ctc-release-body" });

    body.createEl("div", {
      text: "Loading releases…",
      cls: "ctc-release-loading",
    });

    try {
      const releases = await this.plugin.fetchAllReleases();

      body.empty();

      if (!Array.isArray(releases) || releases.length === 0) {
        const noInfo = body.createEl("div", {
          text: "No release information available.",
          cls: "ctc-release-empty",
        });
        return;
      }

      for (const rel of releases) {
        const meta = body.createDiv({ cls: "ctc-release-meta" });

        meta.createEl("div", {
          text: rel.name || rel.tag_name || "Release",
          cls: "ctc-release-name",
        });

        try {
          const dateRaw =
            rel.published_at || rel.created_at || rel.release_date || null;
          if (dateRaw) {
            const dt = new Date(dateRaw);
            const monthNames = [
              "January",
              "February",
              "March",
              "April",
              "May",
              "June",
              "July",
              "August",
              "September",
              "October",
              "November",
              "December",
            ];
            const formatted = `${dt.getFullYear()} ${monthNames[dt.getMonth()]} ${String(dt.getDate()).padStart(2, "0")}`;
            meta.createEl("div", { text: formatted, cls: "ctc-release-date" });
          }
        } catch (e) {}

        const notes = body.createDiv({
          cls: "ctc-release-notes markdown-preview-view",
        });
        const md = rel.body || "No notes";

        try {
          const { MarkdownRenderer, Component } = require("obsidian");
          if (!this._mdComp) {
            try {
              this._mdComp = new Component();
            } catch (e) {
              this._mdComp = null;
            }
          }
          if (
            MarkdownRenderer &&
            typeof MarkdownRenderer.render === "function"
          ) {
            await MarkdownRenderer.render(
              this.plugin.app,
              md,
              notes,
              "",
              this._mdComp || void 0,
            );
          } else {
            throw new Error("MarkdownRenderer not available");
          }
        } catch (e) {
          const fallback = notes.createEl("pre", {
            cls: "ctc-release-notes-fallback",
          });
          fallback.textContent = md;
        }
      }
    } catch (error) {
      body.empty();
      body.createEl("div", {
        text: "Failed to load release notes.",
        cls: "ctc-release-error",
      });
      debugWarn("Error fetching release notes:", error);
    }
  }

  onClose() {
    try {
      if (this._mdComp && typeof this._mdComp.unload === "function") {
        this._mdComp.unload();
      }
    } catch (e) {}
    this._mdComp = null;
    try {
      this.contentEl.empty();
    } catch (e) {}
  }
}

class ColorTableSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.addClass("ctc-settings");

    new Setting(containerEl)
      .setName("Latest release notes")
      .setDesc("View the most recent plugin release notes")
      .addButton((btn) =>
        btn
          .setButtonText("Open changelog")
          .onClick(() => new ReleaseNotesModal(this.app, this.plugin).open()),
      );

    // Toggle for Live Preview Table Coloring
    new Setting(containerEl)
      .setName("Live preview table coloring")
      .setDesc(
        "Apply table coloring in live preview (editor) mode. Disabled by default; enabling may affect editor performance.",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.livePreviewColoring)
          .onChange(async (val) => {
            this.plugin.settings.livePreviewColoring = val;
            await this.plugin.saveSettings();
            // Force document refresh so table coloring updates immediately!
            if (
              this.plugin.app.workspace &&
              typeof this.plugin.app.workspace.trigger === "function"
            ) {
              this.plugin.app.workspace.trigger("layout-change");
            }
            // If disabling, remove colors from all .cm-content tabless
            if (!val) {
              document
                .querySelectorAll(".cm-content table")
                .forEach((table) => {
                  table.querySelectorAll("td, th").forEach((cell) => {
                    cell.style.backgroundColor = "";
                    cell.style.color = "";
                  });
                });
            } else {
              // If enabling, immediately apply colors to all editors
              if (typeof this.plugin.applyColorsToAllEditors === "function") {
                window.setTimeout(
                  () => this.plugin.applyColorsToAllEditors(),
                  0,
                );
              }
            }
          }),
      );

    // Toggle for strict numeric matching
    new Setting(containerEl)
      .setName("Strict numeric matching")
      .setDesc(
        "Only apply numerical rules to cells that are pure numbers (recommended; prevents dates and mixed text from being colored)",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.numericStrict)
          .onChange(async (val) => {
            this.plugin.settings.numericStrict = val;
            await this.plugin.saveSettings();
          }),
      );

    // Quick actions
    new Setting(containerEl).setName("Quick actions").setHeading();

    new Setting(containerEl)
      .setName("Show row coloring in right-click menu")
      .setDesc(
        "Display 'Color row' and 'Remove row' options in the context menu for table cells",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showColorRowInMenu)
          .onChange(async (val) => {
            this.plugin.settings.showColorRowInMenu = val;
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Show column coloring in right-click menu")
      .setDesc(
        "Display 'Color column' and 'Remove column' options in the context menu for table cells",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showColorColumnInMenu)
          .onChange(async (val) => {
            this.plugin.settings.showColorColumnInMenu = val;
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Show undo/redo in right-click menu")
      .setDesc("Display 'Undo' and 'Redo' color changes in the context menu")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showUndoRedoInMenu)
          .onChange(async (val) => {
            this.plugin.settings.showUndoRedoInMenu = val;
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("Show refresh icon in status bar")
      .setDesc("Toggle refresh table icon in the status bar")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showStatusRefreshIcon)
          .onChange(async (val) => {
            this.plugin.settings.showStatusRefreshIcon = val;
            await this.plugin.saveSettings();
            if (val) {
              this.plugin.createStatusBarIcon();
            } else {
              this.plugin.removeStatusBarIcon();
            }
          }),
      );

    new Setting(containerEl)
      .setName("Show refresh icon in ribbon")
      .setDesc("Toggle refresh table icon in the left ribbon")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showRibbonRefreshIcon)
          .onChange(async (val) => {
            this.plugin.settings.showRibbonRefreshIcon = val;
            await this.plugin.saveSettings();
            try {
              if (
                val &&
                !this.plugin._ribbonRefreshIcon &&
                typeof this.plugin.addRibbonIcon === "function"
              ) {
                const iconEl = this.plugin.addRibbonIcon(
                  "table",
                  "Refresh table colors",
                  () => {
                    document
                      .querySelectorAll(
                        ".markdown-preview-view table td, .markdown-preview-view table th",
                      )
                      .forEach((cell) => {
                        cell.style.backgroundColor = "";
                        cell.style.color = "";
                      });
                    this.plugin.applyColorsToActiveFile();
                    document
                      .querySelectorAll(
                        ".cm-content table td, .cm-content table th",
                      )
                      .forEach((cell) => {
                        cell.style.backgroundColor = "";
                        cell.style.color = "";
                      });
                    if (
                      this.plugin.settings.livePreviewColoring &&
                      typeof this.plugin.applyColorsToAllEditors === "function"
                    ) {
                      window.setTimeout(
                        () => this.plugin.applyColorsToAllEditors(),
                        10,
                      );
                    }
                  },
                );
                this.plugin._ribbonRefreshIcon = iconEl;
              } else if (!val && this.plugin._ribbonRefreshIcon) {
                try {
                  this.plugin._ribbonRefreshIcon.remove();
                } catch (e) {}
                this.plugin._ribbonRefreshIcon = null;
              }
            } catch (e) {}
          }),
      );

    // Coloring rules
    new Setting(containerEl).setName("Coloring rules").setHeading();
    const crSection = containerEl.createDiv({ cls: "cr-section" });

    // Search bar
    const searchContainer = crSection.createDiv({
      cls: "ctc-search-container ctc-pretty-flex",
    });
    const searchInput = searchContainer.createEl("input", {
      type: "text",
      cls: "ctc-search-input",
      placeholder: "Search rules...",
    });
    searchContainer.createDiv({ cls: "ctc-search-icon" });
    let searchTerm = "";
    searchInput.addEventListener("input", () => {
      searchTerm = searchInput.value.toLowerCase();
      renderRules();
    });

    const header = crSection.createDiv({
      cls: "ctc-cr-header-row ctc-cr-disabled-row ctc-pretty-flex",
    });
    try {
      header.classList.add("cr-sticky");
    } catch (e) {}
    const mkHeaderSel = (text) => {
      const s = header.createEl("select", {
        cls: "ctc-cr-select cr-header-select",
      });
      const o = s.createEl("option", { text, value: "" });
      o.selected = true;
      o.disabled = true;
      s.disabled = true;
      return s;
    };
    mkHeaderSel("TARGET");
    mkHeaderSel("WHEN");
    mkHeaderSel("MATCH");
    mkHeaderSel("VALUE");
    header.createEl("span", { text: "", cls: "cr-color-placeholder" });
    header.createEl("span", { text: "", cls: "cr-bg-placeholder" });
    header.createEl("span", { text: "", cls: "cr-x-placeholder" });

    const rulesContainer = crSection.createDiv({ cls: "cr-rules-container" });

    const TARGET_OPTIONS = [
      { label: "Color cell", value: "cell" },
      { label: "Color row", value: "row" },
      { label: "Color column", value: "column" },
    ];
    const WHEN_OPTIONS = [
      { label: "The cell", value: "theCell" },
      { label: "Any cell", value: "anyCell" },
      { label: "All cell", value: "allCell" },
      { label: "No cell", value: "noCell" },
      { label: "First row", value: "firstRow" },
      { label: "Column header", value: "columnHeader" },
    ];
    const MATCH_OPTIONS = [
      { label: "Is", value: "is" },
      { label: "Is not", value: "isNot" },
      { label: "Is regex", value: "isRegex" },
      { label: "Contains", value: "contains" },
      { label: "Does not contain", value: "notContains" },
      { label: "Starts with", value: "startsWith" },
      { label: "Ends with", value: "endsWith" },
      { label: "Does not start with", value: "notStartsWith" },
      { label: "Does not end with", value: "notEndsWith" },
      { label: "Is empty", value: "isEmpty" },
      { label: "Is not empty", value: "isNotEmpty" },
      { label: "Is equal to", value: "eq" },
      { label: "Is greater than", value: "gt" },
      { label: "Is less than", value: "lt" },
      { label: "Is greater than or equal to", value: "ge" },
      { label: "Is less than or equal to", value: "le" },
    ];

    const labelFor = (opts, val) =>
      (opts.find((o) => o.value === val)?.label || "").toLowerCase();
    const isRegexRule = (rule) => rule.match === "isRegex";
    const isNumericRule = (rule) =>
      ["eq", "gt", "lt", "ge", "le"].includes(rule.match);

    const renderRules = () => {
      rulesContainer.empty();
      let rules = Array.isArray(this.plugin.settings.coloringRules)
        ? [...this.plugin.settings.coloringRules]
        : [];
      // Filter
      if (searchTerm) {
        rules = rules.filter((r) => {
          const blob = [
            labelFor(TARGET_OPTIONS, r.target || ""),
            labelFor(WHEN_OPTIONS, r.when || ""),
            labelFor(MATCH_OPTIONS, r.match || ""),
            r.value != null ? String(r.value) : "",
          ]
            .join(" ")
            .toLowerCase();
          return blob.includes(searchTerm);
        });
      }
      // Sort
      const sortMode = this.plugin.settings.coloringSort || "lastAdded";
      if (sortMode === "az") {
        rules.sort((a, b) =>
          labelFor(MATCH_OPTIONS, a.match || "").localeCompare(
            labelFor(MATCH_OPTIONS, b.match || ""),
          ),
        );
      } else if (sortMode === "regexFirst") {
        rules.sort((a, b) => Number(isRegexRule(b)) - Number(isRegexRule(a)));
      } else if (sortMode === "numbersFirst") {
        rules.sort(
          (a, b) => Number(isNumericRule(b)) - Number(isNumericRule(a)),
        );
      } else if (sortMode === "mode") {
        const order = { cell: 0, row: 1, column: 2 };
        rules.sort(
          (a, b) => order[a.target || "cell"] - order[b.target || "cell"],
        );
      }
      rules.forEach((rule, idx) => {
        const row = rulesContainer.createDiv({
          cls: "ctc-cr-rule-row ctc-pretty-flex",
        });
        const originalIdx = this.plugin.settings.coloringRules.indexOf(rule);
        row.dataset.idx = String(originalIdx);

        const targetSel = row.createEl("select", { cls: "ctc-cr-select" });
        const tPh = targetSel.createEl("option", { text: "Target", value: "" });
        tPh.disabled = true;
        tPh.selected = !rule.target;
        TARGET_OPTIONS.forEach((opt) => {
          const o = targetSel.createEl("option");
          o.value = opt.value;
          o.text = opt.label;
          if (rule.target === opt.value) o.selected = true;
        });
        targetSel.addEventListener("change", async () => {
          rule.target = targetSel.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        const whenSel = row.createEl("select", { cls: "ctc-cr-select" });
        const wPh = whenSel.createEl("option", { text: "When", value: "" });
        wPh.disabled = true;
        wPh.selected = !rule.when;
        WHEN_OPTIONS.forEach((opt) => {
          const o = whenSel.createEl("option");
          o.value = opt.value;
          o.text = opt.label;
          if (rule.when === opt.value) o.selected = true;
        });
        whenSel.addEventListener("change", async () => {
          rule.when = whenSel.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        const matchSel = row.createEl("select", { cls: "ctc-cr-select" });
        const mPh = matchSel.createEl("option", { text: "Match", value: "" });
        mPh.disabled = true;
        mPh.selected = !rule.match;
        MATCH_OPTIONS.forEach((opt) => {
          const o = matchSel.createEl("option");
          o.value = opt.value;
          o.text = opt.label;
          if (rule.match === opt.value) o.selected = true;
        });
        matchSel.addEventListener("change", async () => {
          rule.match = matchSel.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
          renderRules();
        });

        const numericMatches = new Set(["eq", "gt", "lt", "ge", "le"]);
        const valueInput = row.createEl("input", {
          type: numericMatches.has(rule.match) ? "number" : "text",
          cls: "ctc-cr-value-input",
        });
        valueInput.placeholder = "Value";
        if (rule.value != null) valueInput.value = String(rule.value);
        valueInput.addEventListener("change", async () => {
          const v = valueInput.value;
          rule.value = numericMatches.has(rule.match)
            ? v === ""
              ? null
              : Number(v)
            : v;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        const colorPicker = row.createEl("input", {
          type: "color",
          cls: "ctc-cr-color-picker",
        });
        if (rule.color) colorPicker.value = rule.color;
        else colorPicker.value = "#000000";
        colorPicker.title = "Text color";
        colorPicker.addEventListener("change", async () => {
          rule.color = colorPicker.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        const bgPicker = row.createEl("input", {
          type: "color",
          cls: "ctc-cr-bg-picker",
        });
        if (rule.bg) bgPicker.value = rule.bg;
        else bgPicker.value = "#000000";
        bgPicker.title = "Background color";
        bgPicker.addEventListener("change", async () => {
          rule.bg = bgPicker.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        const delBtn = row.createEl("button", {
          cls: "mod-ghost ctc-cr-del-btn",
        });
        if (typeof window.setIcon === "function") {
          window.setIcon(delBtn, "x");
        } else if (typeof setIcon === "function") {
          setIcon(delBtn, "x");
        } else {
          delBtn.textContent = "×";
        }
        delBtn.addEventListener("click", async () => {
          const oi = this.plugin.settings.coloringRules.indexOf(rule);
          if (oi >= 0) {
            this.plugin.settings.coloringRules.splice(oi, 1);
            await this.plugin.saveSettings();
            this.plugin.applyColorsToActiveFile();
            renderRules();
          }
        });

        row.addEventListener("contextmenu", (evt) => {
          const menu = new Menu();
          menu.addItem((item) =>
            item
              .setTitle("Duplicate rule")
              .setIcon("copy")
              .onClick(async () => {
                const oi = this.plugin.settings.coloringRules.indexOf(rule);
                if (oi >= 0) {
                  const clone = JSON.parse(JSON.stringify(rule));
                  this.plugin.settings.coloringRules.splice(oi + 1, 0, clone);
                  await this.plugin.saveSettings();
                  this.plugin.applyColorsToActiveFile();
                  renderRules();
                }
              }),
          );
          const canMoveUp = originalIdx > 0;
          const canMoveDown =
            originalIdx >= 0 &&
            originalIdx < this.plugin.settings.coloringRules.length - 1;
          menu.addItem((item) =>
            item
              .setTitle("Move rule up")
              .setIcon("arrow-up")
              .setDisabled(!canMoveUp)
              .onClick(async () => {
                const oi = this.plugin.settings.coloringRules.indexOf(rule);
                if (oi > 0) {
                  const list = this.plugin.settings.coloringRules;
                  [list[oi - 1], list[oi]] = [list[oi], list[oi - 1]];
                  await this.plugin.saveSettings();
                  this.plugin.applyColorsToActiveFile();
                  renderRules();
                }
              }),
          );
          menu.addItem((item) =>
            item
              .setTitle("Move rule down")
              .setIcon("arrow-down")
              .setDisabled(!canMoveDown)
              .onClick(async () => {
                const oi = this.plugin.settings.coloringRules.indexOf(rule);
                const list = this.plugin.settings.coloringRules;
                if (oi >= 0 && oi < list.length - 1) {
                  [list[oi], list[oi + 1]] = [list[oi + 1], list[oi]];
                  await this.plugin.saveSettings();
                  this.plugin.applyColorsToActiveFile();
                  renderRules();
                }
              }),
          );
          menu.addSeparator();
          menu.addItem((item) =>
            item
              .setTitle("Reset text color")
              .setIcon("text")
              .onClick(async () => {
                rule.color = null;
                await this.plugin.saveSettings();
                this.plugin.applyColorsToActiveFile();
                renderRules();
              }),
          );
          menu.addItem((item) =>
            item
              .setTitle("Reset background color")
              .setIcon("rectangle-horizontal")
              .onClick(async () => {
                rule.bg = null;
                await this.plugin.saveSettings();
                this.plugin.applyColorsToActiveFile();
                renderRules();
              }),
          );
          try {
            if (menu.containerEl && menu.containerEl.classList)
              menu.containerEl.classList.add("mod-shadow");
            if (menu.menuEl && menu.menuEl.classList)
              menu.menuEl.classList.add("mod-shadow");
          } catch (e) {}
          menu.showAtMouseEvent(evt);
          evt.preventDefault();
        });
      });
    };

    const addRow = crSection.createDiv({ cls: "ctc-cr-add-row" });
    const sortSel = addRow.createEl("select", { cls: "ctc-cr-select" });
    const sortOptions = [
      { label: "Sort: last added", value: "lastAdded" },
      { label: "Sort: A–Z", value: "az" },
      { label: "Sort: regex first", value: "regexFirst" },
      { label: "Sort: numbers first", value: "numbersFirst" },
      { label: "Sort: mode", value: "mode" },
    ];
    sortOptions.forEach((opt) => {
      const o = sortSel.createEl("option");
      o.text = opt.label;
      o.value = opt.value;
      if (opt.value === (this.plugin.settings.coloringSort || "lastAdded"))
        o.selected = true;
    });
    sortSel.addEventListener("change", async () => {
      this.plugin.settings.coloringSort = sortSel.value;
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      renderRules();
    });

    const addBtn = addRow.createEl("button", {
      cls: "mod-cta ctc-cr-add-flex",
    });
    addBtn.textContent = "+ Add rule";
    addBtn.addEventListener("click", async () => {
      if (!Array.isArray(this.plugin.settings.coloringRules))
        this.plugin.settings.coloringRules = [];
      this.plugin.settings.coloringRules.push({
        target: "",
        when: "",
        match: "",
        value: null,
        color: null,
        bg: null,
      });
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      renderRules();
    });

    renderRules();

    const advHeading = new Setting(containerEl)
      .setName("Advanced rules")
      .setHeading();
    advHeading.settingEl.classList.add("ctc-cr-adv-heading");

    // Search bar for advanced rules
    const advSearchContainer = containerEl.createDiv({
      cls: "ctc-search-container ctc-pretty-flex",
    });
    const advSearchInput = advSearchContainer.createEl("input", {
      type: "text",
      cls: "ctc-search-input",
      placeholder: "Search advanced rules...",
    });
    advSearchContainer.createDiv({ cls: "ctc-search-icon" });
    let advSearchTerm = "";
    advSearchInput.addEventListener("input", () => {
      advSearchTerm = advSearchInput.value.toLowerCase();
      renderAdv();
    });

    const advList = containerEl.createDiv({ cls: "ctc-cr-adv-list" });
    const renderAdv = () => {
      advList.empty();
      let advRules = Array.isArray(this.plugin.settings.advancedRules)
        ? this.plugin.settings.advancedRules
        : [];

      const cap = (s) =>
        typeof s === "string" && s.length
          ? s.charAt(0).toUpperCase() + s.slice(1)
          : s;
      const verb = (m, v) => {
        if (m === "contains") return `contains ${v}`;
        if (m === "notContains") return `does not contain ${v}`;
        if (m === "is") return `is ${v}`;
        if (m === "isNot") return `is not ${v}`;
        if (m === "startsWith") return `starts with ${v}`;
        if (m === "endsWith") return `ends with ${v}`;
        if (m === "notStartsWith") return `does not start with ${v}`;
        if (m === "notEndsWith") return `does not end with ${v}`;
        if (m === "isEmpty") return `is empty`;
        if (m === "isNotEmpty") return `is not empty`;
        if (m === "isRegex") return `matches ${v}`;
        if (m === "eq") return `is equal to ${v}`;
        if (m === "gt") return `is greater than ${v}`;
        if (m === "lt") return `is less than ${v}`;
        if (m === "ge") return `is greater than or equal to ${v}`;
        if (m === "le") return `is less than or equal to ${v}`;
        return `${m} ${v}`;
      };
      const summaryForAdvRule = (ar) => {
        // Use custom name if provided
        if (ar.name && ar.name.trim()) {
          return ar.name;
        }
        // Auto-generate name
        const targetPhrase =
          ar.target === "row"
            ? "Color rows"
            : ar.target === "column"
              ? "Color columns"
              : "Color cells";
        const conds = Array.isArray(ar.conditions) ? ar.conditions : [];
        if (!conds.length) return "(empty)";

        // Build a readable string showing conditions
        const parts = conds
          .map((c, idx) => {
            const value = String(c.value || "").trim();
            // For header conditions, just show the value
            if (c.when === "columnHeader") {
              return value;
            }
            // For other conditions, show match + value unless it's a simple case
            if (c.match === "is" && value) {
              return value;
            }
            // Otherwise show verb(match, value)
            return verb(c.match, value);
          })
          .filter(Boolean);

        const condString = parts.length ? parts.join(", ") : "conditions";

        return `${targetPhrase} when ${condString}`;
      };

      // Filter by search term - search all condition values
      if (advSearchTerm) {
        advRules = advRules.filter((ar) => {
          const summary = summaryForAdvRule(ar).toLowerCase();
          const targetText = (ar.target || "cell").toLowerCase();
          const colorHex = (ar.color || ar.bg || "").toLowerCase();
          const allCondValues = Array.isArray(ar.conditions)
            ? ar.conditions
                .map((c) => String(c.value || "").toLowerCase())
                .join(" ")
            : "";
          return (
            summary.includes(advSearchTerm) ||
            targetText.includes(advSearchTerm) ||
            colorHex.includes(advSearchTerm) ||
            allCondValues.includes(advSearchTerm)
          );
        });
      }
      advRules.forEach((ar, idx) => {
        const originalIdx = this.plugin.settings.advancedRules.indexOf(ar);
        const row = advList.createDiv({
          cls: "ctc-cr-adv-row ctc-pretty-flex",
        });
        const drag = row.createEl("span", { cls: "ctc-drag-handle" });
        try {
          require("obsidian").setIcon(drag, "menu");
        } catch (e) {
          try {
            require("obsidian").setIcon(drag, "grip-vertical");
          } catch (e2) {
            drag.textContent = "≡";
          }
        }
        row.dataset.idx = String(originalIdx);
        drag.setAttribute("draggable", "true");
        drag.addEventListener("dragstart", (e) => {
          e.dataTransfer.effectAllowed = "move";
          e.dataTransfer.setData("text/plain", String(originalIdx));
          row.classList.add("dragging");
        });
        drag.addEventListener("dragend", () => {
          row.classList.remove("dragging");
          advList
            .querySelectorAll(".ctc-rule-over")
            .forEach((el) => el.classList.remove("ctc-rule-over"));
        });
        row.addEventListener("dragover", (e) => {
          e.preventDefault();
          row.classList.add("ctc-rule-over");
        });
        row.addEventListener("dragleave", () => {
          row.classList.remove("ctc-rule-over");
        });
        row.addEventListener("drop", async (e) => {
          e.preventDefault();
          const from = Number(e.dataTransfer.getData("text/plain"));
          const to = Number(row.dataset.idx);
          if (isNaN(from) || isNaN(to) || from === to) return;
          const list = this.plugin.settings.advancedRules;
          const [m] = list.splice(from, 1);
          list.splice(to, 0, m);
          await this.plugin.saveSettings();
          renderAdv();
        });

        const label = row.createEl("span", { cls: "ctc-cr-adv-label" });
        label.textContent = summaryForAdvRule(ar);

        const copyBtn = row.createEl("button", {
          cls: "mod-ghost cr-adv-copy",
        });
        copyBtn.setAttribute("aria-label", "Duplicate rule");
        copyBtn.setAttribute("title", "Duplicate rule");
        try {
          require("obsidian").setIcon(copyBtn, "copy");
        } catch (e) {}
        copyBtn.addEventListener("click", async () => {
          const ruleCopy = JSON.parse(JSON.stringify(ar));
          this.plugin.settings.advancedRules.splice(
            originalIdx + 1,
            0,
            ruleCopy,
          );
          await this.plugin.saveSettings();
          document.dispatchEvent(new Event("ctc-adv-rules-changed"));
        });

        const settingsBtn = row.createEl("button", {
          cls: "mod-ghost cr-adv-settings",
        });
        try {
          require("obsidian").setIcon(settingsBtn, "settings");
        } catch (e) {}
        settingsBtn.addEventListener("click", () => {
          new AdvancedRuleModal(this.app, this.plugin, originalIdx).open();
        });
      });
    };
    document.addEventListener("ctc-adv-rules-changed", renderAdv);
    const advActions = containerEl.createDiv({ cls: "ctc-cr-adv-actions" });
    const addAdvBtn = advActions.createEl("button", { cls: "mod-cta" });
    addAdvBtn.textContent = "Add advanced rule";
    addAdvBtn.addEventListener("click", async () => {
      if (!Array.isArray(this.plugin.settings.advancedRules))
        this.plugin.settings.advancedRules = [];
      this.plugin.settings.advancedRules.push({
        logic: "any",
        conditions: [],
        target: "cell",
        color: null,
        bg: null,
      });
      await this.plugin.saveSettings();
      renderAdv();
    });
    renderAdv();

    // Export/import
    new Setting(containerEl).setName("Data management").setHeading();

    const exportImportRow = containerEl.createDiv({
      cls: "ctc-cr-export-row ctc-pretty-flex",
    });
    const exportBtn = exportImportRow.createEl("button", {
      text: "Export settings",
    });
    exportBtn.addEventListener("click", async () => {
      try {
        const data = {
          settings: this.plugin.settings,
          cellData: this.plugin.cellData,
          exportDate: new Date().toISOString(),
        };
        const dataStr = JSON.stringify(data, null, 2);
        const blob = new Blob([dataStr], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `color-table-cells-backup-${new Date().getTime()}.json`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        window.setTimeout(
          () => new Notice("Settings exported successfully!"),
          500,
        );
      } catch (e) {
        new Notice("Failed to export settings: " + e.message);
      }
    });

    const importBtn = exportImportRow.createEl("button", {
      text: "Import settings",
    });
    importBtn.addEventListener("click", () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".json";
      input.addEventListener("change", async (e) => {
        try {
          const file = e.target.files?.[0];
          if (!file) return;
          const text = await file.text();
          const data = JSON.parse(text);

          if (data.settings) {
            this.plugin.settings = Object.assign(
              this.plugin.settings,
              data.settings,
            );
          }
          if (data.cellData) {
            this.plugin.cellData = data.cellData;
          }

          await this.plugin.saveSettings();
          this.display();
          window.setTimeout(
            () => new Notice("Settings imported successfully!"),
            500,
          );
        } catch (e) {
          new Notice("Failed to import settings: " + e.message);
        }
      });
      input.click();
    });

    // Danger zone
    const dangerHeading = new Setting(containerEl)
      .setName("Danger zone")
      .setHeading();
    dangerHeading.settingEl.classList.add("ctc-cr-danger-heading");
    const dangerZoneRow = containerEl.createDiv({
      cls: "ctc-cr-delete-container",
    });
    const deleteManualRow = dangerZoneRow.createDiv({
      cls: "ctc-cr-delete-row",
    });
    const deleteManualBtn = deleteManualRow.createEl("button", {
      text: "Delete all manual colors",
      cls: "mod-warning",
    });
    deleteManualBtn.addEventListener("click", () => {
      const modal = new Modal(this.app);
      new Setting(modal.contentEl)
        .setName("Delete all manual colors?")
        .setHeading();
      modal.contentEl.createEl("p", {
        text: "This will remove all manually colored cells (non-rule colors). This action cannot be undone.",
      });
      const btnRow = modal.contentEl.createDiv({
        cls: "ctc-modal-delete-buttons",
      });
      const cancelBtn = btnRow.createEl("button", {
        text: "Cancel",
        cls: "mod-ghost",
      });
      const confirmBtn = btnRow.createEl("button", {
        text: "Delete all",
        cls: "mod-warning",
      });
      cancelBtn.addEventListener("click", () => modal.close());
      confirmBtn.addEventListener("click", async () => {
        this.plugin.cellData = {};
        await this.plugin.saveData({
          settings: this.plugin.settings,
          cellData: this.plugin.cellData,
        });
        this.plugin.applyColorsToActiveFile();
        new Notice("All manual colors deleted");
        modal.close();
      });
      modal.open();
    });

    // Delete coloring rules
    const deleteRulesRow = dangerZoneRow.createDiv({
      cls: "ctc-cr-delete-row",
    });
    const deleteRulesBtn = deleteRulesRow.createEl("button", {
      text: "Delete all coloring rules",
      cls: "mod-warning",
    });
    deleteRulesBtn.addEventListener("click", () => {
      const modal = new Modal(this.app);
      new Setting(modal.contentEl)
        .setName("Delete all coloring rules?")
        .setHeading();
      modal.contentEl.createEl("p", {
        text: `This will remove all ${Array.isArray(this.plugin.settings.coloringRules) ? this.plugin.settings.coloringRules.length : 0} coloring rules. This action cannot be undone.`,
      });
      const btnRow = modal.contentEl.createDiv({
        cls: "ctc-modal-delete-buttons",
      });
      const cancelBtn = btnRow.createEl("button", {
        text: "Cancel",
        cls: "mod-ghost",
      });
      const confirmBtn = btnRow.createEl("button", {
        text: "Delete all",
        cls: "mod-warning",
      });
      cancelBtn.addEventListener("click", () => modal.close());
      confirmBtn.addEventListener("click", async () => {
        this.plugin.settings.coloringRules = [];
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        new Notice("All coloring rules deleted");
        modal.close();
        this.display();
      });
      modal.open();
    });

    // Delete advanced rules
    const deleteAdvRulesRow = dangerZoneRow.createDiv({
      cls: "ctc-cr-delete-row",
    });
    const deleteAdvRulesBtn = deleteAdvRulesRow.createEl("button", {
      text: "Delete all advanced rules",
      cls: "mod-warning",
    });
    deleteAdvRulesBtn.addEventListener("click", () => {
      const modal = new Modal(this.app);
      new Setting(modal.contentEl)
        .setName("Delete all advanced rules?")
        .setHeading();
      modal.contentEl.createEl("p", {
        text: `This will remove all ${Array.isArray(this.plugin.settings.advancedRules) ? this.plugin.settings.advancedRules.length : 0} advanced rules. This action cannot be undone.`,
      });
      const btnRow = modal.contentEl.createDiv({
        cls: "ctc-modal-delete-buttons",
      });
      const cancelBtn = btnRow.createEl("button", {
        text: "Cancel",
        cls: "mod-ghost",
      });
      const confirmBtn = btnRow.createEl("button", {
        text: "Delete all",
        cls: "mod-warning",
      });
      cancelBtn.addEventListener("click", () => modal.close());
      confirmBtn.addEventListener("click", async () => {
        this.plugin.settings.advancedRules = [];
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        if (
          this.plugin.settings.livePreviewColoring &&
          typeof this.plugin.applyColorsToAllEditors === "function"
        )
          window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        new Notice("All advanced rules deleted");
        modal.close();
        this.display();
      });
      modal.open();
    });
  }

  hide() {
    // Force refresh table colors when closing settings tab
    // Clear all rule-based colors first to ensure changes take effect
    document
      .querySelectorAll(
        ".markdown-preview-view table td, .markdown-preview-view table th",
      )
      .forEach((cell) => {
        if (!cell.hasAttribute("data-ctc-manual")) {
          cell.style.backgroundColor = "";
          cell.style.color = "";
        }
      });
    document
      .querySelectorAll(".cm-content table td, .cm-content table th")
      .forEach((cell) => {
        if (!cell.hasAttribute("data-ctc-manual")) {
          cell.style.backgroundColor = "";
          cell.style.color = "";
        }
      });

    // Now reapply colors with rule changes
    if (typeof this.plugin.applyColorsToActiveFile === "function") {
      window.setTimeout(() => this.plugin.applyColorsToActiveFile(), 50);
    }
    if (
      this.plugin.settings.livePreviewColoring &&
      typeof this.plugin.applyColorsToAllEditors === "function"
    ) {
      window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 100);
    }
  }
}

class AdvancedRuleModal extends Modal {
  constructor(app, plugin, index) {
    super(app);
    this.plugin = plugin;
    this.index = index;
  }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass("ctc-cr-adv-modal");
    const rule = this.plugin.settings.advancedRules?.[this.index] || {
      logic: "any",
      conditions: [],
      target: "cell",
      color: null,
      bg: null,
    };
    new Setting(contentEl)
      .setName("Advanced rules builder")
      .setHeading()
      .settingEl.addClass("ctc-cr-adv-modal-heading");
    const logicRow = contentEl.createDiv({ cls: "ctc-cr-adv-logic" });
    const logicLabel = logicRow.createEl("span", {
      cls: "ctc-cr-adv-logic-label",
    });
    logicLabel.textContent = "Conditions match";
    const logicButtons = logicRow.createDiv({
      cls: "ctc-cr-adv-logic-buttons",
    });
    const mkBtn = (txt, val) => {
      const b = logicButtons.createEl("button", {
        cls: "ctc-cr-adv-logic-btn mod-ghost",
      });
      b.textContent = txt;
      const applyActive = () => {
        ["any", "all", "none"].forEach((k) => {
          const q = Array.from(
            logicButtons.querySelectorAll(".ctc-cr-adv-logic-btn"),
          ).find((el) => el.textContent === k.toUpperCase());
          if (q) {
            q.classList.remove("mod-cta");
            q.classList.add("mod-ghost");
          }
        });
        b.classList.add("mod-cta");
        b.classList.remove("mod-ghost");
      };
      if (rule.logic === val) applyActive();
      b.addEventListener("click", async () => {
        rule.logic = val;
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        if (
          this.plugin.settings.livePreviewColoring &&
          typeof this.plugin.applyColorsToAllEditors === "function"
        )
          window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        applyActive();
      });
      return b;
    };
    mkBtn("ANY", "any");
    mkBtn("ALL", "all");
    mkBtn("NONE", "none");
    contentEl.createEl("h4", { text: "Conditions", cls: "ctc-cr-adv-h4" });
    const condsWrap = contentEl.createDiv({ cls: "ctc-cr-adv-conds-wrap" });
    const TARGET_OPTIONS = [
      { label: "Color cell", value: "cell" },
      { label: "Color row", value: "row" },
      { label: "Color column", value: "column" },
    ];
    const WHEN_OPTIONS = [
      { label: "Any cell", value: "anyCell" },
      { label: "All cell", value: "allCell" },
      { label: "No cell", value: "noCell" },
      { label: "First row", value: "firstRow" },
      { label: "Column header", value: "columnHeader" },
    ];
    const MATCH_OPTIONS = [
      { label: "Is", value: "is" },
      { label: "Is not", value: "isNot" },
      { label: "Is regex", value: "isRegex" },
      { label: "Contains", value: "contains" },
      { label: "Does not contain", value: "notContains" },
      { label: "Starts with", value: "startsWith" },
      { label: "Ends with", value: "endsWith" },
      { label: "Does not start with", value: "notStartsWith" },
      { label: "Does not end with", value: "notEndsWith" },
      { label: "Is empty", value: "isEmpty" },
      { label: "Is not empty", value: "isNotEmpty" },
      { label: "Is equal to", value: "eq" },
      { label: "Is greater than", value: "gt" },
      { label: "Is less than", value: "lt" },
      { label: "Is greater than or equal to", value: "ge" },
      { label: "Is less than or equal to", value: "le" },
    ];
    const renderConds = () => {
      condsWrap.empty();
      (rule.conditions || []).forEach((cond, ci) => {
        const row = condsWrap.createDiv({
          cls: "ctc-cr-adv-cond-row ctc-pretty-flex",
        });
        const whenSel = row.createEl("select", {
          cls: "ctc-cr-select ctc-cr-adv-cond-when",
        });
        WHEN_OPTIONS.forEach((opt) => {
          const o = whenSel.createEl("option");
          o.value = opt.value;
          o.text = opt.label;
          if (cond.when === opt.value) o.selected = true;
        });
        whenSel.addEventListener("change", async () => {
          cond.when = whenSel.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });
        const matchSel = row.createEl("select", {
          cls: "ctc-cr-select ctc-cr-adv-cond-match",
        });
        MATCH_OPTIONS.forEach((opt) => {
          const o = matchSel.createEl("option");
          o.value = opt.value;
          o.text = opt.label;
          if (cond.match === opt.value) o.selected = true;
        });
        matchSel.addEventListener("change", async () => {
          cond.match = matchSel.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
          const isNum = ["eq", "gt", "lt", "ge", "le"].includes(cond.match);
          valInput.type = isNum ? "number" : "text";
        });
        const isNum = ["eq", "gt", "lt", "ge", "le"].includes(cond.match);
        const valInput = row.createEl("input", {
          type: isNum ? "number" : "text",
          cls: "ctc-cr-value-input ctc-cr-adv-cond-value",
        });
        if (cond.value != null) valInput.value = String(cond.value);
        valInput.addEventListener("change", async () => {
          const v = valInput.value;
          cond.value = isNum ? (v === "" ? null : Number(v)) : v;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (
            this.plugin.settings.livePreviewColoring &&
            typeof this.plugin.applyColorsToAllEditors === "function"
          )
            window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        // Add per-condition delete button (X)
        const delBtn = row.createEl("button", {
          cls: "mod-ghost ctc-cr-adv-cond-del",
        });
        try {
          require("obsidian").setIcon(delBtn, "x");
        } catch (e) {
          delBtn.textContent = "×";
        }
        delBtn.addEventListener("click", async () => {
          if (Array.isArray(rule.conditions)) {
            rule.conditions.splice(ci, 1);
            await this.plugin.saveSettings();
            this.plugin.applyColorsToActiveFile();
            if (
              this.plugin.settings.livePreviewColoring &&
              typeof this.plugin.applyColorsToAllEditors === "function"
            )
              window.setTimeout(
                () => this.plugin.applyColorsToAllEditors(),
                10,
              );
            renderConds();
          }
        });
      });
    };
    renderConds();
    const addCondRow = contentEl.createDiv({ cls: "ctc-cr-adv-add-row" });
    const addCondBtn = addCondRow.createEl("button", {
      cls: "mod-ghost ctc-cr-adv-add-btn",
    });
    addCondBtn.textContent = "+ Add condition";
    addCondBtn.addEventListener("click", async () => {
      if (!Array.isArray(rule.conditions)) rule.conditions = [];
      rule.conditions.push({ when: "anyCell", match: "contains", value: "" });
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      renderConds();
    });

    contentEl.createEl("h4", { text: "Then color", cls: "ctc-cr-adv-h4" });

    const colorRow = contentEl.createDiv({
      cls: "ctc-cr-adv-color-row ctc-pretty-flex",
    });
    const targetSel = colorRow.createEl("select", {
      cls: "ctc-cr-select ctc-cr-adv-target",
    });
    TARGET_OPTIONS.forEach((opt) => {
      const o = targetSel.createEl("option");
      o.value = opt.value;
      o.text = opt.label;
      if (rule.target === opt.value) o.selected = true;
    });
    targetSel.addEventListener("change", async () => {
      rule.target = targetSel.value;
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
    });

    // Text color picker with reset button
    const textColorContainer = colorRow.createDiv({
      cls: "ctc-cr-adv-text-color-container",
    });
    const colorPicker = textColorContainer.createEl("input", {
      type: "color",
      cls: "ctc-cr-color-picker ctc-cr-adv-text-color ctc-cr-adv-color-input",
    });
    colorPicker.value = rule.color || "#000000";
    colorPicker.title = "Text color";
    colorPicker.addEventListener("change", async () => {
      rule.color = colorPicker.value;
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
    });
    const colorResetBtn = textColorContainer.createEl("button", {
      cls: "mod-ghost ctc-cr-adv-color-reset",
    });
    colorResetBtn.textContent = "Reset";
    colorResetBtn.title = "Reset text color to none";
    colorResetBtn.addEventListener("click", async () => {
      rule.color = null;
      colorPicker.value = "#000000";
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
    });

    // Background color picker with reset button
    const bgColorContainer = colorRow.createDiv({
      cls: "ctc-cr-adv-bg-color-container",
    });
    const bgPicker = bgColorContainer.createEl("input", {
      type: "color",
      cls: "ctc-cr-bg-picker ctc-cr-adv-bg-color ctc-cr-adv-bg-input",
    });
    bgPicker.value = rule.bg || "#000000";
    bgPicker.title = "Background color";
    bgPicker.addEventListener("change", async () => {
      rule.bg = bgPicker.value;
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
    });
    const bgResetBtn = bgColorContainer.createEl("button", {
      cls: "mod-ghost ctc-cr-adv-bg-reset",
    });
    bgResetBtn.textContent = "Reset";
    bgResetBtn.title = "Reset background color to none";
    bgResetBtn.addEventListener("click", async () => {
      rule.bg = null;
      bgPicker.value = "#000000";
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
    });

    // Custom Name Section - AFTER color swatches
    contentEl.createEl("h4", {
      text: "Rule name (optional)",
      cls: "ctc-cr-adv-h4",
    });
    const nameRow = contentEl.createDiv({ cls: "ctc-cr-adv-name-row" });
    const nameInput = nameRow.createEl("input", {
      type: "text",
      cls: "ctc-cr-adv-name-input",
    });
    nameInput.placeholder = "Leave empty to use automatic naming";
    if (rule.name) nameInput.value = rule.name;
    nameInput.addEventListener("change", async () => {
      rule.name = nameInput.value;
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
    });
    const actionsRow = contentEl.createDiv({ cls: "ctc-cr-adv-actions-row" });
    const deleteBtn = actionsRow.createEl("button", {
      cls: "ctc-cr-adv-delete",
    });
    deleteBtn.textContent = "Delete rule";
    deleteBtn.addEventListener("click", async () => {
      if (Array.isArray(this.plugin.settings.advancedRules)) {
        this.plugin.settings.advancedRules.splice(this.index, 1);
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        if (
          this.plugin.settings.livePreviewColoring &&
          typeof this.plugin.applyColorsToAllEditors === "function"
        )
          window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
      }
      this.close();
    });
    const saveBtn = actionsRow.createEl("button", {
      cls: "mod-cta ctc-cr-adv-save",
    });
    saveBtn.textContent = "Save rule";
    saveBtn.addEventListener("click", () => {
      this.plugin.applyColorsToActiveFile();
      if (
        this.plugin.settings.livePreviewColoring &&
        typeof this.plugin.applyColorsToAllEditors === "function"
      )
        window.setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
      this.close();
    });
  }
  onClose() {
    this.contentEl.empty();
  }
}
