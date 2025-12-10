const { Plugin, PluginSettingTab, Setting, Menu, ButtonComponent, Modal, setIcon } = require('obsidian');

// Debug configuration - make it a getter so changes are reflected dynamically
let IS_DEVELOPMENT = false;
const debugLog = (...args) => IS_DEVELOPMENT && console.log('[CTC-DEBUG]', ...args);
const debugWarn = (...args) => IS_DEVELOPMENT && console.warn('[CTC-WARN]', ...args);

// Allow toggling debug mode from console: window.setDebugMode(true/false)
if (typeof window !== 'undefined') {
  window.setDebugMode = (value) => {
    IS_DEVELOPMENT = value;
    console.log(`[CTC] Debug mode ${value ? 'enabled' : 'disabled'}`);
  };
}

module.exports = class TableColorPlugin extends Plugin {
  // Undo/redo stacks
  undoStack = [];
  redoStack = [];
  maxStackSize = 50;

  async onload() {
    // COMMAND PALETTE COMMANDS
    this.addCommand({
      id: 'enable-live-preview-coloring',
      name: 'Enable Live Preview Table Coloring',
      callback: async () => {
        this.settings.livePreviewColoring = true;
        await this.saveSettings();
        if (this.app.workspace && typeof this.app.workspace.trigger === 'function') {
          this.app.workspace.trigger('layout-change');
        }
        if (typeof this.applyColorsToAllEditors === 'function') {
          setTimeout(() => this.applyColorsToAllEditors(), 0);
        }
      }
    });
    this.addCommand({
      id: 'disable-live-preview-coloring',
      name: 'Disable Live Preview Table Coloring',
      callback: async () => {
        this.settings.livePreviewColoring = false;
        await this.saveSettings();
        if (this.app.workspace && typeof this.app.workspace.trigger === 'function') {
          this.app.workspace.trigger('layout-change');
        }
        // Remove colors from all .cm-content tables
        document.querySelectorAll('.cm-content table').forEach(table => {
          table.querySelectorAll('td, th').forEach(cell => {
            cell.style.backgroundColor = '';
            cell.style.color = '';
          });
        });
      }
    });
    this.addCommand({
      id: 'undo-color-change',
      name: 'Undo Last Color Change',
      callback: () => this.undo()
    });

    this.addCommand({
      id: 'redo-color-change',
      name: 'Redo Last Color Change',
      callback: () => this.redo()
    });

    this.addCommand({
      id: 'add-cell-color-rule',
      name: 'Add Table Cell Color Rule',
      callback: () => {
        // Open settings tab and scroll to rules section
        this.app.setting.open();
        setTimeout(() => {
          if (this.app.setting && typeof this.app.setting.openTabById === 'function') {
            this.app.setting.openTabById('color-table-cell');
          }
        }, 250);
      }
    });

    this.addCommand({
      id: 'manage-coloring-rules',
      name: 'Manage Coloring Rules',
      callback: () => {
        this.app.setting.open();
        setTimeout(() => {
          if (this.app.setting && typeof this.app.setting.openTabById === 'function') {
            this.app.setting.openTabById('color-table-cell');
          }
        }, 250);
      }
    });

    this.addCommand({
      id: 'add-advanced-rule',
      name: 'Add Advanced Rule',
      callback: async () => {
        if (!Array.isArray(this.settings.advancedRules)) this.settings.advancedRules = [];
        this.settings.advancedRules.push({ logic:'any', conditions:[], target:'cell', color:null, bg:null });
        await this.saveSettings();
        const idx = this.settings.advancedRules.length - 1;
        new AdvancedRuleModal(this.app, this, idx).open();
        document.dispatchEvent(new CustomEvent('ctc-adv-rules-changed'));
      }
    });

    this.addCommand({
      id: 'manage-advanced-rules',
      name: 'Manage Advanced Rules',
      callback: () => {
        this.app.setting.open();
        setTimeout(() => {
          if (this.app.setting && typeof this.app.setting.openTabById === 'function') {
            this.app.setting.openTabById('color-table-cell');
            // Scroll to advanced rules section
            setTimeout(() => {
              const advHeading = document.querySelector('.cr-adv-heading');
              if (advHeading) {
                advHeading.scrollIntoView({ behavior: 'smooth', block: 'start' });
              }
            }, 200);
          }
        }, 250);
      }
    });

    this.addCommand({
      id: 'refresh-table-colors',
      name: 'Refresh Table Colors',
      callback: () => {
        this.hardRefreshTableColors();
      }
    });

    // REGEX PATTERN TESTER COMMAND
    // this.addCommand({
    //   id: 'open-regex-tester',
    //   name: 'Open Regex Pattern Tester',
    //   callback: () => {
    //     new RegexTesterModal(this.app, this).open();
    //   }
    // });

    // --- Live Preview Table Coloring logic ---
    this.applyColorsToAllEditors = () => {
      debugLog('applyColorsToAllEditors called, livePreviewColoring:', this.settings.livePreviewColoring);
      if (!this.settings.livePreviewColoring) {
        // Remove all colors if disabled
        document.querySelectorAll('.cm-content table').forEach(table => {
          table.querySelectorAll('td, th').forEach(cell => {
            cell.style.backgroundColor = '';
            cell.style.color = '';
          });
        });
        return;
      }
      
      const file = this.app.workspace.getActiveFile();
      if (!file) {
        debugWarn('No active file in applyColorsToAllEditors');
        return;
      }
      
      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll('table'));
      
      const editors = document.querySelectorAll('.cm-content');
      debugLog(`applyColorsToAllEditors: Found ${editors.length} editors for file: ${file.path}`);
      
      editors.forEach(editorEl => {
        // Get all tables in this editor
        const editorTables = Array.from(editorEl.querySelectorAll('table'));
        debugLog(`applyColorsToAllEditors: Found ${editorTables.length} tables in editor`);
        
        // Process each table using processSingleTable (which handles clearing internally)
        editorTables.forEach((table, localIdx) => {
          const globalTableIndex = allDocTables.indexOf(table);
          // Use global index if found, otherwise use local index within editor
          const tableIdx = globalTableIndex >= 0 ? globalTableIndex : localIdx;
          debugLog(`applyColorsToAllEditors: Processing table ${tableIdx} (global: ${globalTableIndex}, local: ${localIdx})`);
          this.processSingleTable(table, tableIdx, file.path, noteData);
        });
      });
    };

    // Observe editors to force coloring on large tables as they render
    const installEditorObservers = () => {
      const editors = Array.from(document.querySelectorAll('.cm-content'));
      editors.forEach(ed => {
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
          let scrollTimeout;
          const onScroll = () => {
            clearTimeout(scrollTimeout);
            scrollTimeout = setTimeout(() => {
              // Restore colors from data attributes to handle DOM recreation
              ed.querySelectorAll('[data-ctc-bg], [data-ctc-color]').forEach(cell => {
                if (cell.hasAttribute('data-ctc-bg')) {
                  cell.style.backgroundColor = cell.getAttribute('data-ctc-bg');
                }
                if (cell.hasAttribute('data-ctc-color')) {
                  cell.style.color = cell.getAttribute('data-ctc-color');
                }
              });
            }, 50);
          };
          // Save handler for cleanup
          ed._ctcScrollHandler = onScroll;
          ed.addEventListener('scroll', ed._ctcScrollHandler);
          ed._ctcScrollListener = true;
        }
      });
    };
    installEditorObservers();

    const setupLivePreviewColoring = () => {
      // Initial application - use applyColorsToActiveFile which handles both modes properly
      setTimeout(() => this.applyColorsToActiveFile(), 200);
      // Observe DOM changes in editors
      if (!this._livePreviewObserver) {
        this._livePreviewObserver = new MutationObserver(() => {
          this.applyColorsToActiveFile();
        });
        document.querySelectorAll('.cm-content').forEach(editorEl => {
          this._livePreviewObserver.observe(editorEl, { childList: true, subtree: true });
        });
      }
      // Re-apply on file open/layout change
      this.registerEvent(this.app.workspace.on('file-open', () => this.applyColorsToActiveFile()));
      this.registerEvent(this.app.workspace.on('layout-change', () => this.applyColorsToActiveFile()));
      // Re-apply on cell focus/blur/input (to persist colors after editing)
      this.registerDomEvent(document, 'focusin', (e) => {
        if (e.target && e.target.closest && e.target.closest('.cm-content table')) {
          this.applyColorsToActiveFile();
        }
      });
      this.registerDomEvent(document, 'input', (e) => {
        if (e.target && e.target.closest && e.target.closest('.cm-content table')) {
          setTimeout(() => this.applyColorsToActiveFile(), 30);
        }
      });
      
      // Watch for style changes on colored cells and restore via fallback mechanism
      const colorRestorer = new MutationObserver((mutations) => {
        let needsReapply = false;
        mutations.forEach(mutation => {
          if (mutation.target.closest && mutation.target.closest('.cm-content table')) {
            const cell = mutation.target.closest('td, th');
            if (cell && cell.hasAttribute('data-ctc-bg')) {
              // Color was applied before, check if it's still there
              if (!cell.style.backgroundColor || cell.style.backgroundColor === '') {
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
      this.registerDomEvent(document, 'pointerdown', (e) => {
        const cell = e.target?.closest && e.target.closest('td, th');
        const table = e.target?.closest && e.target.closest('.cm-content table');
        if (cell && table) {
          // Single reapplication after selection settles
          setTimeout(() => this.applyColorsToAllEditors(), 10);
        }
      });

      // Also watch existing tables for color loss (fallback restoration)
      this._colorRestorer = colorRestorer;
      document.querySelectorAll('.cm-content table').forEach(table => {
        colorRestorer.observe(table, { 
          attributes: true, 
          attributeFilter: ['style'], 
          subtree: true 
        });
      });
    };

    await this.loadSettings();
    if (typeof this.addStatusBarItem === 'function' && this.settings.showStatusRefreshIcon && !this.statusBarRefresh) {
      this.createStatusBarIcon();
    }
    if (typeof this.addRibbonIcon === 'function' && this.settings.showRibbonRefreshIcon && !this._ribbonRefreshIcon) {
      const iconEl = this.addRibbonIcon('table', 'Refresh table colors', () => {
        this.hardRefreshTableColors();
      });
      this._ribbonRefreshIcon = iconEl;
    }
    setupLivePreviewColoring();
    if (this.settings?.persistUndoHistory) {
      await this.loadUndoRedoStacks();
    }
    const rawSaved = await this.loadData() || {};

    this._appliedContainers = new WeakMap();
    
    // Register event to refresh colors when switching files
    this.registerEvent(this.app.workspace.on('file-open', async (file) => {
      setTimeout(() => {
        // Clear any cached containers from previous file
        if (this._appliedContainers && typeof this._appliedContainers.forEach === 'function') {
          this._appliedContainers.forEach((_, container) => {
            if (!container.isConnected) {
              this._appliedContainers.delete(container);
            }
          });
        }
        
        // Reapply colors to new file
        if (typeof this.applyColorsToActiveFile === 'function') {
          this.applyColorsToActiveFile();
        }
      }, 50);
    }));
    
    // Reapply colors when layout changes (mode switches)
    this.registerEvent(this.app.workspace.on('layout-change', async () => {
      setTimeout(() => {
        if (typeof this.applyColorsToActiveFile === 'function') {
          this.applyColorsToActiveFile();
        }
      }, 100);
    }));

    const normalizeCellData = (obj) => {
      let cur = obj;
      const seen = new Set();
      while (cur && typeof cur === 'object' && !Array.isArray(cur)) {
        // avoid infinite cycles
        if (seen.has(cur)) break;
        seen.add(cur);

        const keys = Object.keys(cur);
        // If there are keys other than metadata keys, assume this is the real cellData
        const nonMeta = keys.filter(k => k !== 'settings' && k !== 'cellData');
        if (nonMeta.length > 0) return cur;

        // If only a single wrapper key, unwrap it
        if (keys.length === 1) {
          const k = keys[0];
          cur = cur[k];
          continue;
        }

        // If exactly the pair of keys 'settings' and 'cellData', prefer diving into 'cellData'
        if (keys.length === 2 && keys.includes('settings') && keys.includes('cellData')) {
          if (cur.cellData && typeof cur.cellData === 'object') {
            cur = cur.cellData;
            continue;
          }
          return {};
        }

        // Fallback: return current object
        return cur;
      }
      return {};
    };

    this.cellData = normalizeCellData(rawSaved) || {};

    try {
      const normalizedSave = { settings: this.settings, cellData: this.cellData };
      const rawStr = JSON.stringify(rawSaved || {});
      const normStr = JSON.stringify(normalizedSave || {});
      if (rawStr !== normStr) {
        await this.saveData(normalizedSave);
        debugLog('color-table-cell: migrated and saved normalized plugin data');
      }
    } catch (e) { /* ignore migration errors */ }

    if (!this._settingsTab) {
      this._settingsTab = new ColorTableSettingTab(this.app, this);
      try { this.addSettingTab(this._settingsTab); } catch (e) { /* ignore if already added */ }
    }

    // Auto-refresh active document when settings are closed
    this.registerDomEvent(document, 'click', (e) => {
      const settingsContainer = document.querySelector('.vertical-tabs-container, .settings');
      const isClosingSettings = !settingsContainer || !settingsContainer.offsetParent;
      if (isClosingSettings && this._settingsWasOpen) {
        this._settingsWasOpen = false;
        setTimeout(() => {
          if (typeof this.applyColorsToActiveFile === 'function') {
            this.applyColorsToActiveFile();
          }
        }, 100);
      }
    });
    
    // Track when settings are opened
    const originalOpen = this.app.setting?.open?.bind(this.app.setting) || (() => {});
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
      const allDocTables = Array.from(document.querySelectorAll('table'));
      
      // Look for newly added tables in reading mode
      mutations.forEach(mutation => {
        if (mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach(node => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              // Check if this node or its children contain tables
              const tables = [];
              if (node.matches && node.matches('table')) {
                tables.push(node);
              }
              if (node.querySelectorAll) {
                tables.push(...node.querySelectorAll('table'));
              }
              
              // Immediately color any new tables in reading mode
              tables.forEach(table => {
                if (table.closest('.markdown-preview-view') && !table.hasAttribute('data-ctc-processed')) {
                  const globalTableIdx = allDocTables.indexOf(table);
                  if (globalTableIdx >= 0) {
                    this.processSingleTable(table, globalTableIdx, file.path, noteData);
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
      if (!el.closest('.markdown-preview-view')) return;
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
        let scrollDebounceId = null;
        
        const safeDisconnect = () => {
          try { if (observer) { observer.disconnect(); observer = null; } } catch (e) { }
          try { if (debounceId) { clearTimeout(debounceId); debounceId = null; } } catch (e) { }
          try { if (scrollListener && el.parentElement) { el.parentElement.removeEventListener('scroll', scrollListener); } } catch (e) { }
          try { if (scrollDebounceId) { clearTimeout(scrollDebounceId); scrollDebounceId = null; } } catch (e) { }
          
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
            
            if (el.querySelectorAll && el.querySelectorAll('table').length > 0) {
              this.applyColorsToContainer(el, fileId);
              // Don't disconnect immediately - keep observing for dynamic content
            }
          } catch (e) { /* ignore */ }
        };

        observer = new MutationObserver((mutations) => {
          // Look specifically for added tables
          let hasTableAdded = false;
          mutations.forEach(mutation => {
            if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
              mutation.addedNodes.forEach(node => {
                if (node.nodeType === 1) { // Element node
                  // Check if this is a table or contains tables
                  if (node.tagName === 'TABLE' || (node.querySelector && node.querySelector('table'))) {
                    hasTableAdded = true;
                  }
                }
              });
            }
          });
          
          if (hasTableAdded) {
            debugLog('MutationObserver: Table added to reading view, applying colors');
            if (debounceId) clearTimeout(debounceId);
            debounceId = setTimeout(() => {
              checkAndApply();
            }, 80);
          }
        });

        observer.observe(el, { childList: true, subtree: true });

        // Add scroll listener to reapply colors on scroll (for lazy-loaded tables)
        try {
          const scrollContainer = el.parentElement || el;
          scrollListener = () => {
            if (scrollDebounceId) clearTimeout(scrollDebounceId);
            scrollDebounceId = setTimeout(() => {
              // Check if element is still connected before applying
              if (el.isConnected) {
                debugLog('Scroll detected in reading view, checking for new tables');
                this.applyColorsToContainer(el, fileId);
              } else {
                safeDisconnect();
              }
            }, 100);
          };
          scrollContainer.addEventListener('scroll', scrollListener, { passive: true });
        } catch (e) { /* ignore if scroll listener fails */ }

        // Store observer reference for cleanup
        this._containerObservers.set(el, {
          observer,
          debounceId,
          scrollListener,
          scrollDebounceId,
          safeDisconnect
        });

        checkAndApply();

        // Check periodically if element is still connected
        const connectionChecker = setInterval(() => {
          if (!el.isConnected) {
            safeDisconnect();
            clearInterval(connectionChecker);
          }
        }, 2000);
        
        // Auto-cleanup after reasonable time
        setTimeout(() => { 
          clearInterval(connectionChecker);
          safeDisconnect(); 
        }, 30000);
        
      } catch (e) { /* ignore if MutationObserver unsupported */ }
    });

    // Also add a global observer specifically for reading mode tables
    try {
      const readingViewObserver = new MutationObserver((mutations) => {
        const file = this.app.workspace.getActiveFile();
        if (!file) return;
        
        mutations.forEach(mutation => {
          if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
            let hasNewTable = false;
            mutation.addedNodes.forEach(node => {
              if (node.nodeType === 1) {
                // Check if a table was added to reading view
                if (node.matches && node.matches('table') && node.closest('.markdown-preview-view')) {
                  hasNewTable = true;
                } else if (node.querySelectorAll) {
                  const tables = node.querySelectorAll('table');
                  if (tables.length > 0 && node.closest('.markdown-preview-view')) {
                    hasNewTable = true;
                  }
                }
              }
            });
            
            if (hasNewTable) {
              debugLog('Global observer: New table detected in reading view, applying colors');
              // Use applyColorsToActiveFile which handles both reading and live preview
              this.applyColorsToActiveFile();
            }
          }
        });
      });
      
      // Observe the entire document body for new reading view tables
      readingViewObserver.observe(document.body, { childList: true, subtree: true });
      this._readingViewObserver = readingViewObserver;
    } catch (e) {
      debugWarn('Failed to setup reading view observer:', e);
    }

    // Setup reading view scroll listener for lazy-loaded tables
    this.setupReadingViewScrollListener();

    // Start periodic checker for reading mode tables
    this.startReadingModeTableChecker();

    // Enhanced global observer to restore colors from data attributes when DOM recreates cells
    try {
      const globalObserver = new MutationObserver((mutations) => {
        mutations.forEach(mutation => {
          if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
            mutation.addedNodes.forEach(node => {
              if (node.nodeType === 1) {
                // Check for cells with data attributes
                if (node.matches && node.matches('[data-ctc-bg], [data-ctc-color]')) {
                  this.restoreColorsFromAttributes(node);
                }
                // Check descendants
                if (node.querySelectorAll) {
                  node.querySelectorAll('[data-ctc-bg], [data-ctc-color]').forEach(cell => {
                    this.restoreColorsFromAttributes(cell);
                  });
                }
              }
            });
          }
        });
      });
      
      globalObserver.observe(document.body, { childList: true, subtree: true });
      this._globalObserver = globalObserver;
    } catch (e) {
      debugWarn('Failed to setup global observer:', e);
    }

    if (this.settings.enableContextMenu) {
      this.registerDomEvent(document, "contextmenu", (evt) => {
        const target = evt.target;
        const cell = target?.closest("td, th");
        const tableEl = target?.closest("table");
        if (!cell || !tableEl) return;
        const readingView = cell.closest('.markdown-preview-view');
        const livePreview = cell.closest('.cm-content');
        if (!readingView || livePreview) return;
        const menu = new Menu();
        menu.addItem(item =>
          item.setTitle("Color Cell Text")
              .setIcon('palette')
              .onClick(() => this.pickColor(cell, tableEl, "color"))
        );
        menu.addItem(item =>
          item.setTitle("Color Cell Background")
              .setIcon('droplet')
              .onClick(() => this.pickColor(cell, tableEl, "bg"))
        );
        menu.addSeparator();
        
        // Row coloring options - conditional on setting
        if (this.settings.showColorRowInMenu) {
          menu.addItem(item =>
            item.setTitle("Color Whole Row Text")
                .setIcon('rows-3')
                .onClick(() => this.pickColorForRow(cell, tableEl, "color"))
          );
          menu.addItem(item =>
            item.setTitle("Color Whole Row Background")
                .setIcon('droplet')
                .onClick(() => this.pickColorForRow(cell, tableEl, "bg"))
          );
          menu.addSeparator();
        }
        
        // Column coloring options - conditional on setting
        if (this.settings.showColorColumnInMenu) {
          menu.addItem(item =>
            item.setTitle("Color Whole Column Text")
                .setIcon('columns-3')
                .onClick(() => this.pickColorForColumn(cell, tableEl, "color"))
          );
          menu.addItem(item =>
            item.setTitle("Color Whole Column Background")
                .setIcon('droplet')
                .onClick(() => this.pickColorForColumn(cell, tableEl, "bg"))
          );
          menu.addSeparator();
        }
        
        menu.addItem(item =>
          item.setTitle("Color Multiple Cells by Rule")
              .setIcon('grid')
              .onClick(() => {
                this.app.setting.open();
                setTimeout(() => {
                  if (this.app.setting && typeof this.app.setting.openTabById === 'function') {
                    this.app.setting.openTabById('color-table-cell');
                    // Scroll to Coloring Rules heading
                    setTimeout(() => {
                      const settingsContainer = document.querySelector('.vertical-tabs-container');
                      if (settingsContainer) {
                        const rulesHeading = Array.from(settingsContainer.querySelectorAll('h3')).find(el => el.textContent.includes('Coloring Rules'));
                        if (rulesHeading) {
                          rulesHeading.scrollIntoView({ behavior: 'smooth', block: 'start' });
                        }
                      }
                    }, 100);
                  }
                }, 250);
              })
        );
        menu.addSeparator();
        
        // Undo/Redo options - conditional on setting
        if (this.settings.showUndoRedoInMenu) {
          menu.addItem(item =>
            item.setTitle("Undo Last Color Change")
                .setIcon('undo')
                .setDisabled(this.undoStack.length === 0)
                .onClick(() => this.undo())
          );
          menu.addItem(item =>
            item.setTitle("Redo Last Color Change")
                .setIcon('redo')
                .setDisabled(this.redoStack.length === 0)
                .onClick(() => this.redo())
          );
          menu.addSeparator();
        }
        
        menu.addItem(item =>
          item.setTitle("Reset Cell Coloring")
              .setIcon('trash-2')
              .onClick(async () => this.resetCell(cell, tableEl))
        );
        
        if (this.settings.showColorRowInMenu) {
          menu.addItem(item =>
            item.setTitle("Remove Row Coloring")
                .setIcon('rows-3')
                .onClick(async () => this.resetRow(cell, tableEl))
          );
        }
        
        if (this.settings.showColorColumnInMenu) {
          menu.addItem(item =>
            item.setTitle("Remove Column Coloring")
                .setIcon('columns-3')
                .onClick(async () => this.resetColumn(cell, tableEl))
          );
        }
        try {
          if (menu.containerEl && menu.containerEl.classList) menu.containerEl.classList.add('mod-shadow');
          if (menu.menuEl && menu.menuEl.classList) menu.menuEl.classList.add('mod-shadow');
        } catch (e) { /* ignore if properties not present */ }

        menu.showAtMouseEvent(evt);
        evt.preventDefault();
      });
    }

    this.applyColorsToActiveFile();

    this.registerEvent(
      this.app.workspace.on('file-open', () => {
        this.applyColorsToActiveFile();
      })
    );

    this.registerEvent(
      this.app.workspace.on('layout-change', () => {
        this.applyColorsToActiveFile();
      })
    );

    // Also listen to active leaf changes to catch view switches
    this.registerEvent(
      this.app.workspace.on('active-leaf-change', () => {
        setTimeout(() => this.applyColorsToActiveFile(), 50);
      })
    );
  }

  createStatusBarIcon() {
    if (!this.statusBarRefresh && typeof this.addStatusBarItem === 'function') {
      const status = this.addStatusBarItem();
      this.statusBarRefresh = status;
      setIcon(status, 'table');
      status.setAttribute('aria-label', 'Refresh Table Colors');
      status.classList.add('ctc-refresh-table-color');
      status.style.cursor = 'pointer';
      status.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.hardRefreshTableColors();
      });
      debugLog('[Status Bar] Icon created successfully');
    }
  }

  removeStatusBarIcon() {
    if (this.statusBarRefresh) {
      try {
        const container = this.statusBarRefresh.containerEl;
        if (container) {
          container.remove();
        }
        this.statusBarRefresh = null;
        debugLog('[Status Bar] Icon removed successfully');
      } catch (e) {
        debugWarn('[Status Bar] Error removing icon:', e);
      }
    }
  }

  hardRefreshTableColors() {
    debugLog('[Hard Refresh] Starting hard refresh of table colors');
    
    // STEP 1: Clear all table colors from DOM (both manual and rule-based)
    document.querySelectorAll('table td, table th').forEach(cell => {
      cell.style.backgroundColor = '';
      cell.style.color = '';
      // Clear data attributes that might cache colors
      cell.removeAttribute('data-ctc-bg');
      cell.removeAttribute('data-ctc-color');
    });
    
    // STEP 2: Clear table processing markers to force reprocessing
    document.querySelectorAll('table').forEach(table => {
      table.removeAttribute('data-ctc-processed');
      table.removeAttribute('data-ctc-index');
      table.removeAttribute('data-ctc-file');
      table.removeAttribute('data-ctc-last-processed');
    });
    
    // STEP 3: Reset internal cache for DOM tracking
    this._appliedContainers = new WeakMap();
    
    // STEP 4: Reapply all colors from scratch
    debugLog('[Hard Refresh] Colors cleared, reapplying from scratch');
    
    if (typeof this.applyColorsToActiveFile === 'function') {
      setTimeout(() => {
        this.applyColorsToActiveFile();
        debugLog('[Hard Refresh] Hard refresh complete');
      }, 50);
    }
  }

  onunload() {
    // Remove our settings tab when the plugin is 
    // disabled/unloaded to avoid duplicate entries
    try {
      if (this._settingsTab && typeof this.removeSettingTab === 'function') {
        this.removeSettingTab(this._settingsTab);
      }
    } catch (e) { }
    try {
      if (this._livePreviewObserver && typeof this._livePreviewObserver.disconnect === 'function') {
        this._livePreviewObserver.disconnect();
        this._livePreviewObserver = null;
      }
    } catch (e) { }
    
    // Clean up table pre-renderer
    try {
      if (this._tablePreRenderer && typeof this._tablePreRenderer.disconnect === 'function') {
        this._tablePreRenderer.disconnect();
        this._tablePreRenderer = null;
      }
    } catch (e) { }
    
    // Clean up reading view scroll observer
    try {
      if (this._readingViewScrollObserver && typeof this._readingViewScrollObserver.disconnect === 'function') {
        this._readingViewScrollObserver.disconnect();
        this._readingViewScrollObserver = null;
      }
    } catch (e) { }
    
    // Clean up reading mode checker interval
    try {
      if (this._readingModeChecker) {
        clearInterval(this._readingModeChecker);
        this._readingModeChecker = null;
      }
    } catch (e) { }
    
    // Clean up reading view scroll listeners
    try {
      document.querySelectorAll('.markdown-preview-view').forEach(view => {
        // Remove scroll listener if it was added
        if (view._ctcScrollListenerAdded && view._ctcScrollHandler) {
          view.removeEventListener('scroll', view._ctcScrollHandler);
          view._ctcScrollListenerAdded = false;
          view._ctcScrollHandler = null;
        }
      });
    } catch (e) { }
    
    // Clean up editor scroll listeners
    try {
      document.querySelectorAll('.cm-content').forEach(ed => {
        // Disconnect observer
        if (ed._ctcObserver) {
          ed._ctcObserver.disconnect();
          ed._ctcObserver = null;
        }
        // Remove scroll listener
        if (ed._ctcScrollListener && ed._ctcScrollHandler) {
          ed.removeEventListener('scroll', ed._ctcScrollHandler);
          ed._ctcScrollListener = false;
          ed._ctcScrollHandler = null;
        }
      });
    } catch (e) { }
    
    // Clean up all container observers to prevent memory leaks
    try {
      if (this._containerObservers) {
        for (const [el, observerData] of this._containerObservers.entries()) {
          try {
            if (observerData && typeof observerData.disconnect === 'function') {
              observerData.disconnect();
            }
          } catch (e) { /* ignore */ }
        }
        this._containerObservers.clear();
      }
    } catch (e) { }
    
    // Clean up color restorer
    try {
      if (this._colorRestorer && typeof this._colorRestorer.disconnect === 'function') {
        this._colorRestorer.disconnect();
        this._colorRestorer = null;
      }
    } catch (e) { }
    
    try {
      if (this.settings?.persistUndoHistory) {
        this.saveUndoRedoStacks();
      }
    } catch (e) { }
  }

  async loadSettings() {
    const data = await this.loadData();
    this.settings = Object.assign({


      enableContextMenu: true,
      showColorRowInMenu: true,
      showColorColumnInMenu: true,
      showUndoRedoInMenu: true,
      coloringRules: [],
      coloringSort: 'lastAdded',
      advancedRules: [],
      numericStrict: true,
      livePreviewColoring: false,
      persistUndoHistory: true,
      recentColors: [],
      presetColors: [],
      processAllCellsOnOpen: false, // Process all cells when file opens, not just visible ones
      showStatusRefreshIcon: false,
      showRibbonRefreshIcon: false
    }, data?.settings || {});
    if (Array.isArray(this.settings.presetColors)) {
      this.settings.presetColors = this.settings.presetColors.map(pc => {
        return typeof pc === 'string' ? { name: '', color: pc } : pc;
      });
    } else {
      this.settings.presetColors = [];
    }
  }

  async saveSettings() {
    try {
      const dataToSave = {
        settings: this.settings,
        cellData: this.cellData
      };
      await this.saveData(dataToSave);
    } catch (error) {
      throw error;
    }
    }

  updateRecentColor(color) {
    if (!color) return;
    const list = Array.isArray(this.settings.recentColors) ? [...this.settings.recentColors] : [];
    const existingIndex = list.findIndex(c => c.toUpperCase() === color.toUpperCase());
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
        this.initialColor = initialColor || '#FFA500';
        this.anchorEl = anchorEl;
        const {h, s, v} = this.hexToHsv(this.initialColor);
        this.hue = h;
        this.sat = s;
        this.val = v;
        this.color = this.hsvToHex(this.hue, this.sat, this.val);
        this.menuEl = null;
        this._cells = []; // Support multiple cells for row/column coloring
      }
      open() {
        this.close();
        this.menuEl = document.createElement('div');
        this.menuEl.className = 'custom-color-picker-menu';
        Object.assign(this.menuEl.style, {
          position: 'absolute',
          zIndex: 9999,
          background: 'var(--background-secondary, #232323)',
          borderRadius: '10px',
          boxShadow: '0 2px 8px #0007',
          padding: '12px 12px',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          width: '238px',
          minWidth: '0',
          userSelect: 'none',
          border: '1px solid var(--background-modifier-border, #333)',
        });

        const sbBox = document.createElement('canvas');
        sbBox.width = 210;
        sbBox.height = 120;
        Object.assign(sbBox.style, {
          borderRadius: '6px',
          cursor: 'crosshair',
          marginBottom: '8px',
          background: '#fff',
          position: 'relative',
          display: 'block',
          marginLeft: '0',
          marginRight: '0',
        });
        this.menuEl.appendChild(sbBox);

        const sbCtx = sbBox.getContext('2d');
        const drawSB = (hue) => {
          for (let x = 0; x < sbBox.width; x++) {
            for (let y = 0; y < sbBox.height; y++) {
              const s = x / (sbBox.width-1);
              const v = 1 - y / (sbBox.height-1);
              sbCtx.fillStyle = this.hsvToHex(hue, s, v);
              sbCtx.fillRect(x, y, 1, 1);
            }
          }
        };
        drawSB(this.hue);

        const sbSelector = document.createElement('div');
        Object.assign(sbSelector.style, {
          position: 'absolute',
          pointerEvents: 'none',
          width: '16px',
          height: '16px',
          border: '2px solid #fff',
          borderRadius: '50%',
          boxShadow: '0 0 4px #000a',
          transform: 'translate(-8px, -8px)',
          zIndex: 10,
          background: 'none',
          display: '',
        });
        this.menuEl.appendChild(sbSelector);

        const hueBox = document.createElement('canvas');
        hueBox.width = 210;
        hueBox.height = 14;
        Object.assign(hueBox.style, {
          borderRadius: '5px',
          margin: '6px 0 6px 0',
          display: 'block',
          marginLeft: '0',
          marginRight: '0',
        });
        this.menuEl.appendChild(hueBox);

        const hueCtx = hueBox.getContext('2d');
        const grad = hueCtx.createLinearGradient(0, 0, hueBox.width, 0);
        for (let i = 0; i <= 360; i += 1) {
          grad.addColorStop(i/360, `hsl(${i},100%,50%)`);
        }
        hueCtx.fillStyle = grad;
        hueCtx.fillRect(0, 0, hueBox.width, hueBox.height);

        const hueSelector = document.createElement('div');
        Object.assign(hueSelector.style, {
          position: 'absolute',
          width: '12px',
          height: '24px',
          transform: 'translate(0px, -6px)',
          borderRadius: '4px',
          border: '2px solid #fff',
          boxShadow: '0 0 4px #000a',
          left: '0px',
          top: '0px',
          zIndex: 11,
          background: 'none',
          display: ''
        });
        this.menuEl.appendChild(hueSelector);

        const hexRow = document.createElement('div');
        Object.assign(hexRow.style, {
          display: 'flex',
          alignItems: 'center',
          marginTop: '6px',
          width: '100%',
          justifyContent: 'space-between',
          gap: '6px',
        });
        this.menuEl.appendChild(hexRow);
        const preview = document.createElement('div');
        Object.assign(preview.style, {
          width: '36px',
          height: '28px',
          borderRadius: '5px',
          border: '2px solid transparent',
          background: this.color,
          marginRight: '0',
          marginLeft: '2px',
        });
        hexRow.appendChild(preview);
        const hexInput = document.createElement('input');
        hexInput.type = 'text';
        hexInput.value = (this.color || '#FFA500').replace(/^#/, '');
        hexInput.maxLength = 7;
        Object.assign(hexInput.style, {
          width: '100%',
          minWidth: '0',
          background: 'var(--background-primary, #181818)',
          color: 'var(--text-normal, #fff)',
          border: '2px solid var(--background-modifier-border, #333)',
          borderRadius: '5px',
          padding: '3px 4px',
          marginLeft: '0',
          marginRight: '0',
          outline: 'none',
          transition: 'color 0.15s, background 0.15s',
          textAlign: 'center',
        });
        let hexAnimFrame = null;
        hexInput.addEventListener('input', () => {
          hexInput.value = hexInput.value.replace(/[^0-9a-fA-F]/g, '').slice(0, 6);
          if (hexInput.value.length === 6) {
            const hex = '#' + hexInput.value;
            const {h, s, v} = this.hexToHsv(hex);
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
        const pickBtn = document.createElement('button');
        pickBtn.className = 'mod-ghost';
        pickBtn.textContent = '';
        Object.assign(pickBtn.style, { width: '36px', height: '32px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' });
        try {
          const ob = require('obsidian');
          if (ob && typeof ob.setIcon === 'function') {
            ob.setIcon(pickBtn, 'pipette');
            const svgEl = pickBtn.querySelector('svg');
            if (svgEl) { svgEl.style.width = '25px'; svgEl.style.height = '25px'; }
          } else if (typeof window.setIcon === 'function') {
            window.setIcon(pickBtn, 'pipette');
            const svgEl = pickBtn.querySelector('svg');
            if (svgEl) { svgEl.style.width = '25px'; svgEl.style.height = '25px'; }
          } else if (typeof setIcon === 'function') {
            setIcon(pickBtn, 'pipette');
            const svgEl = pickBtn.querySelector('svg');
            if (svgEl) { svgEl.style.width = '25px'; svgEl.style.height = '25px'; }
          } else {
            const svg = document.createElementNS('http://www.w3.org/2000/svg','svg');
            svg.setAttribute('viewBox','0 0 24 24');
            svg.setAttribute('width','20');
            svg.setAttribute('height','20');
            const path = document.createElementNS('http://www.w3.org/2000/svg','path');
            path.setAttribute('d','M13.5 6.5l4-4 4 4-4 4M4 20l6-6 3 3-6 6H4v-3z');
            path.setAttribute('fill','none');
            path.setAttribute('stroke','currentColor');
            path.setAttribute('stroke-width','2');
            path.setAttribute('stroke-linecap','round');
            path.setAttribute('stroke-linejoin','round');
            svg.appendChild(path);
            pickBtn.appendChild(svg);
          }
        } catch (e) { }
        pickBtn.title = 'Pick color from screen';
        pickBtn.onclick = async () => {
          try {
            if ('EyeDropper' in window) {
              const eye = new window.EyeDropper();
              const result = await eye.open();
              const hex = (result && result.sRGBHex) ? result.sRGBHex.toUpperCase() : null;
              if (hex) {
                const { h, s, v } = this.hexToHsv(hex);
                this.hue = h; this.sat = s; this.val = v; this.color = hex;
                preview.style.background = this.color;
                drawSB(this.hue);
                updateSelectors();
              }
            } else {
              pickBtn.title = 'Screen picker not supported in this environment';
            }
          } catch (e) { /* ignore user cancel or errors */ }
        };
        hexRow.appendChild(pickBtn);
        const recentRow = document.createElement('div');
        Object.assign(recentRow.style, { display: 'flex', flexWrap: 'wrap', gap: '6px', width: '100%', marginTop: '6px' });
        const recents = Array.isArray(this.plugin.settings.recentColors) ? this.plugin.settings.recentColors : [];
        recents.slice(0, 10).forEach(rc => {
          const sw = document.createElement('button');
          Object.assign(sw.style, { width: '22px', height: '22px', borderRadius: '4px', border: '1px solid var(--background-modifier-border, #333)', background: rc, cursor: 'pointer', padding: '0' });
          sw.addEventListener('click', () => {
            const { h, s, v } = this.hexToHsv(rc);
            this.hue = h; this.sat = s; this.val = v; this.color = rc;
            preview.style.background = this.color;
            drawSB(this.hue);
            updateSelectors();
          });
          recentRow.appendChild(sw);
        });
        this.menuEl.appendChild(recentRow);
        const updateSelectors = () => {
          const x = Math.round(this.sat * (sbBox.width-1));
          const y = Math.round((1-this.val) * (sbBox.height-1));
          sbSelector.style.left = (sbBox.offsetLeft + x) + 'px';
          sbSelector.style.top = (sbBox.offsetTop + y) + 'px';
          const hx = Math.round(this.hue/360 * (hueBox.width-1));
          hueSelector.style.left = (hueBox.offsetLeft + hx) + 'px';
          hueSelector.style.top = (hueBox.offsetTop + 1) + 'px';
          if (!document.activeElement || document.activeElement !== hexInput) {
            if (hexAnimFrame) cancelAnimationFrame(hexAnimFrame);
            const target = this.color.replace(/^#/, '').toUpperCase();
            if (hexInput.value !== target) {
              let i = 0;
              const animate = () => {
                if (hexInput.value !== target) {
                  hexInput.value = target.slice(0, i+1);
                  i++;
                  if (i < target.length) hexAnimFrame = requestAnimationFrame(animate);
                }
              };
              animate();
            }
          }
          if (this._cell && this._type) {
            if (this._type === 'bg') this._cell.style.backgroundColor = this.color;
            else this._cell.style.color = this.color;
          }
          // Support for multiple cells (row/column)
          if (this._cells && this._cells.length > 0 && this._type) {
            this._cells.forEach(cell => {
              if (this._type === 'bg') cell.style.backgroundColor = this.color;
              else cell.style.color = this.color;
            });
          }
        };
        let sbDragging = false;
        sbBox.addEventListener('mousedown', (e) => {
          sbDragging = true;
          handleSB(e);
        });
        window.addEventListener('mousemove', (e) => {
          if (sbDragging) handleSB(e);
        });
        window.addEventListener('mouseup', () => { sbDragging = false; });
        const handleSB = (e) => {
          const rect = sbBox.getBoundingClientRect();
          let x = e.clientX - rect.left;
          let y = e.clientY - rect.top;
          x = Math.max(0, Math.min(sbBox.width-1, x));
          y = Math.max(0, Math.min(sbBox.height-1, y));
          this.sat = x / (sbBox.width-1);
          this.val = 1 - y / (sbBox.height-1);
          this.color = this.hsvToHex(this.hue, this.sat, this.val);
          preview.style.background = this.color;
          updateSelectors();
        };
        let hueDragging = false;
        let huePending = false;
        let hueX = 0;
        const updateHueSelector = () => {
          const newHue = (hueX / (hueBox.width-1)) * 360;
          this.hue = newHue;
          this.color = this.hsvToHex(this.hue, this.sat, this.val);
          preview.style.background = this.color;
          hueSelector.style.left = (hueBox.offsetLeft + hueX - 6) + 'px';
          hueSelector.style.top = (hueBox.offsetTop + 1) + 'px';
          if (this._cell && this._type) {
            if (this._type === 'bg') this._cell.style.backgroundColor = this.color;
            else this._cell.style.color = this.color;
          }
          // Support for multiple cells (row/column)
          if (this._cells && this._cells.length > 0 && this._type) {
            this._cells.forEach(cell => {
              if (this._type === 'bg') cell.style.backgroundColor = this.color;
              else cell.style.color = this.color;
            });
          }
          drawSB(this.hue);
        };
        const onHueMove = (e) => {
          const rect = hueBox.getBoundingClientRect();
          hueX = Math.max(0, Math.min(hueBox.width-1, e.clientX - rect.left));
          if (!huePending) {
            huePending = true;
            requestAnimationFrame(() => {
              updateHueSelector();
              huePending = false;
            });
          }
        };
        hueBox.addEventListener('mousedown', (e) => {
          hueDragging = true;
          onHueMove(e);
        });
        window.addEventListener('mousemove', (e) => {
          if (hueDragging) onHueMove(e);
        });
        window.addEventListener('mouseup', () => {
          hueDragging = false;
        });
        setTimeout(updateSelectors, 0);

        hexInput.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') applyBtn.click();
        });

        setTimeout(() => {
          let rect = this.anchorEl.getBoundingClientRect();
          let left = rect.left;
          let top = rect.bottom + 4;

          if (left + 280 > window.innerWidth) left = window.innerWidth - 280;
          if (top + 260 > window.innerHeight) top = rect.top - 260;
          this.menuEl.style.left = left + 'px';
          this.menuEl.style.top = top + 'px';
        }, 0);
        document.body.appendChild(this.menuEl);

        this._outsideHandler = (evt) => {
          if (!this.menuEl.contains(evt.target)) this.close();
        };
        setTimeout(() => document.addEventListener('mousedown', this._outsideHandler), 10);
      }
      close() {
        if (this.menuEl && this.menuEl.parentNode) this.menuEl.parentNode.removeChild(this.menuEl);
        if (this._outsideHandler) document.removeEventListener('mousedown', this._outsideHandler);
        if (typeof this.onPick === 'function') {
          this.onPick(this.color); // <-- Save the picked color!
        }
      }

      hsvToHex(h, s, v) {
        let r, g, b;
        let i = Math.floor(h / 60);
        let f = h / 60 - i;
        let p = v * (1 - s);
        let q = v * (1 - f * s);
        let t = v * (1 - (1 - f) * s);
        switch (i % 6) {
          case 0: r = v, g = t, b = p; break;
          case 1: r = q, g = v, b = p; break;
          case 2: r = p, g = v, b = t; break;
          case 3: r = p, g = q, b = v; break;
          case 4: r = t, g = p, b = v; break;
          case 5: r = v, g = p, b = q; break;
        }
        return '#' + [r,g,b].map(x=>Math.round(x*255).toString(16).padStart(2,'0')).join('').toUpperCase();
      }
      hexToHsv(hex) {
        let r = 0, g = 0, b = 0;
        if (hex.length === 4) {
          r = parseInt(hex[1]+hex[1], 16);
          g = parseInt(hex[2]+hex[2], 16);
          b = parseInt(hex[3]+hex[3], 16);
        } else if (hex.length === 7) {
          r = parseInt(hex.substr(1,2), 16);
          g = parseInt(hex.substr(3,2), 16);
          b = parseInt(hex.substr(5,2), 16);
        }
        r /= 255; g /= 255; b /= 255;
        let max = Math.max(r, g, b), min = Math.min(r, g, b);
        let h, s, v = max;
        let d = max - min;
        s = max === 0 ? 0 : d / max;
        if (max === min) h = 0;
        else {
          switch(max){
            case r: h = (g-b)/d + (g<b?6:0); break;
            case g: h = (b-r)/d + 2; break;
            case b: h = (r-g)/d + 4; break;
          }
          h *= 60;
        }
        return {h, s, v};
      }
    }
    return CustomColorPickerMenu;
  }

  // Helper: Get global table index across entire document
  getGlobalTableIndex(tableEl) {
    const allDocTables = Array.from(document.querySelectorAll('table'));
    return allDocTables.indexOf(tableEl);
  }

  async pickColor(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = null;
    // Use the cell as anchor for menu position
    const menu = new CustomColorPickerMenu(this, async (pickedColor) => {
      // Save color on close
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      // Use GLOBAL table index, not local
      const tableIndex = this.getGlobalTableIndex(tableEl);
      const rowIndex = Array.from(tableEl.querySelectorAll('tr')).indexOf(cell.closest('tr'));
      const colIndex = Array.from(cell.closest('tr').querySelectorAll('td, th')).indexOf(cell);
      
      // Capture current state for undo
      const oldColors = this.cellData[fileId]?.[`table_${tableIndex}`]?.[`row_${rowIndex}`]?.[`col_${colIndex}`];
      
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
        [type]: pickedColor
      };
      
      tableColors[rowKey][colKey] = newColors;
      
      // Add to undo stack
      const snapshot = this.createSnapshot(
        'cell_color',
        fileId,
        tableIndex,
        { row: rowIndex, col: colIndex },
        oldColors,
        newColors
      );
      this.addToUndoStack(snapshot);
      this.updateRecentColor(pickedColor);
      
      await this.saveDataColors();
    }, initialColor, cell);
    menu._cell = cell;
    menu._type = type;
    menu.open();
  }

  async pickColorForRow(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = null;
    const row = cell.closest('tr');
    const rowCells = Array.from(row.querySelectorAll('td, th'));
    
    const menu = new CustomColorPickerMenu(this, async (pickedColor) => {
      // Save color on close for entire row
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      // Use GLOBAL table index, not local
      const tableIndex = this.getGlobalTableIndex(tableEl);
      const rowIndex = Array.from(tableEl.querySelectorAll('tr')).indexOf(row);
      
      // Capture current state for undo
      const oldColors = this.cellData[fileId]?.[`table_${tableIndex}`]?.[`row_${rowIndex}`];
      
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
          [type]: pickedColor
        };
        tableColors[rowKey][colKey] = newColors[colKey];
      });
      
      // Add to undo stack
      const snapshot = this.createSnapshot(
        'row_color',
        fileId,
        tableIndex,
        { row: rowIndex },
        oldColors,
        newColors
      );
      this.addToUndoStack(snapshot);
      this.updateRecentColor(pickedColor);
      
      await this.saveDataColors();
    }, initialColor, cell);
    
    menu._cells = rowCells;
    menu._type = type;
    menu.open();
  }

  async pickColorForColumn(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = null;
    const colIndex = Array.from(cell.closest('tr').querySelectorAll('td, th')).indexOf(cell);
    
    // Collect all cells in the column for preview
    const columnCells = [];
    tableEl.querySelectorAll('tr').forEach((row) => {
      const cells = row.querySelectorAll('td, th');
      if (colIndex < cells.length) {
        columnCells.push(cells[colIndex]);
      }
    });
    
    const menu = new CustomColorPickerMenu(this, async (pickedColor) => {
      // Save color on close for entire column
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      // Use GLOBAL table index, not local
      const tableIndex = this.getGlobalTableIndex(tableEl);
      
      // Capture current state for undo
      const oldColors = {};
      if (this.cellData[fileId]?.[`table_${tableIndex}`]) {
        Object.entries(this.cellData[fileId][`table_${tableIndex}`]).forEach(([rowKey, rowData]) => {
          if (rowKey.startsWith('row_') && rowData[`col_${colIndex}`]) {
            oldColors[rowKey] = { ...oldColors[rowKey] };
            oldColors[rowKey][`col_${colIndex}`] = { ...rowData[`col_${colIndex}`] };
          }
        });
      }
      
      if (!this.cellData[fileId]) this.cellData[fileId] = {};
      const noteData = this.cellData[fileId];
      const tableKey = `table_${tableIndex}`;
      if (!noteData[tableKey]) noteData[tableKey] = {};
      const tableColors = noteData[tableKey];
      
      // Color all cells in the column
      const newColors = {};
      tableEl.querySelectorAll('tr').forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('td, th');
        if (colIndex < cells.length) {
          const rowKey = `row_${rowIndex}`;
          if (!tableColors[rowKey]) tableColors[rowKey] = {};
          const colKey = `col_${colIndex}`;
          newColors[rowKey] = { ...newColors[rowKey] };
          newColors[rowKey][colKey] = {
            ...tableColors[rowKey][colKey],
            [type]: pickedColor
          };
          tableColors[rowKey][colKey] = newColors[rowKey][colKey];
        }
      });
      
      // Add to undo stack
      const snapshot = this.createSnapshot(
        'column_color',
        fileId,
        tableIndex,
        { col: colIndex },
        oldColors,
        newColors
      );
      this.addToUndoStack(snapshot);
      this.updateRecentColor(pickedColor);
      
      await this.saveDataColors();
    }, initialColor, cell);
    
    menu._cells = columnCells;
    menu._type = type;
    menu.open();
  }

  async resetCell(cell, tableEl) {
    // Clear inline styles
    cell.style.backgroundColor = "";
    cell.style.color = "";
    
    // Remove data attributes that restore colors
    cell.removeAttribute('data-ctc-bg');
    cell.removeAttribute('data-ctc-color');
    cell.removeAttribute('data-ctc-manual');

    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

    const tableIndex = this.getGlobalTableIndex(tableEl);
    const rowIndex = Array.from(tableEl.querySelectorAll('tr')).indexOf(cell.closest('tr'));
    const colIndex = Array.from(cell.closest('tr').querySelectorAll('td, th')).indexOf(cell);

    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (noteData?.[tableKey]?.[`row_${rowIndex}`]) {
      delete noteData[tableKey][`row_${rowIndex}`][`col_${colIndex}`];
      await this.saveDataColors();
    }
    
    // Force refresh to ensure no colors reappear
    setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async resetRow(cell, tableEl) {
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

    const tableIndex = this.getGlobalTableIndex(tableEl);
    const rowIndex = Array.from(tableEl.querySelectorAll('tr')).indexOf(cell.closest('tr'));

    const noteData = this.cellData[fileId] || {};
    const tableKey = `table_${tableIndex}`;
    const tableColors = noteData[tableKey] || {};
    const rowKey = `row_${rowIndex}`;
    const oldColors = tableColors[rowKey] ? { ...tableColors[rowKey] } : undefined;

    if (!this.cellData[fileId]) this.cellData[fileId] = {};
    if (!this.cellData[fileId][tableKey]) this.cellData[fileId][tableKey] = {};
    delete this.cellData[fileId][tableKey][rowKey];

    // Clear all cells in the row
    cell.closest('tr')?.querySelectorAll('td, th')?.forEach(td => { 
      td.style.backgroundColor = ""; 
      td.style.color = "";
      // Remove data attributes
      td.removeAttribute('data-ctc-bg');
      td.removeAttribute('data-ctc-color');
      td.removeAttribute('data-ctc-manual');
    });

    const snapshot = this.createSnapshot('row_color', fileId, tableIndex, { row: rowIndex }, oldColors, undefined);
    this.addToUndoStack(snapshot);

    await this.saveDataColors();
    
    // Force refresh to ensure no colors reappear
    setTimeout(() => this.applyColorsToActiveFile(), 50);
  }

  async resetColumn(cell, tableEl) {
    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

    const tableIndex = this.getGlobalTableIndex(tableEl);
    const colIndex = Array.from(cell.closest('tr').querySelectorAll('td, th')).indexOf(cell);

    if (!this.cellData[fileId]) this.cellData[fileId] = {};
    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (!noteData[tableKey]) noteData[tableKey] = {};
    const tableColors = noteData[tableKey];

    const oldColors = {};
    Object.entries(tableColors).forEach(([rk, rowData]) => {
      if (rk.startsWith('row_') && rowData[`col_${colIndex}`]) {
        oldColors[rk] = { [`col_${colIndex}`]: { ...rowData[`col_${colIndex}`] } };
        delete rowData[`col_${colIndex}`];
      }
    });

    // Clear all cells in the column and remove data attributes
    tableEl.querySelectorAll('tr').forEach(tr => {
      const cells = tr.querySelectorAll('td, th');
      if (colIndex < cells.length) { 
        const c = cells[colIndex]; 
        c.style.backgroundColor = ""; 
        c.style.color = "";
        // Remove data attributes
        c.removeAttribute('data-ctc-bg');
        c.removeAttribute('data-ctc-color');
        c.removeAttribute('data-ctc-manual');
      }
    });

    const snapshot = this.createSnapshot('column_color', fileId, tableIndex, { col: colIndex }, oldColors, undefined);
    this.addToUndoStack(snapshot);

    await this.saveDataColors();
    
    // Force refresh to ensure no colors reappear
    setTimeout(() => this.applyColorsToActiveFile(), 50);
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
    if (this.settings.livePreviewColoring && typeof this.applyColorsToAllEditors === 'function') {
      setTimeout(() => this.applyColorsToAllEditors(), 10);
    }
  }

  async saveDataColors() {
    await this.saveData({ settings: this.settings, cellData: this.cellData });
    // Sync colors to both reading and live preview modes
    // Use setTimeout to ensure DOM is ready
    setTimeout(() => this.applyColorsToActiveFile(), 10);
    // Also explicitly apply to live preview editor if enabled
    if (this.settings.livePreviewColoring && typeof this.applyColorsToAllEditors === 'function') {
      setTimeout(() => this.applyColorsToAllEditors(), 20);
    }
  }

  // Create snapshot of current state for undo/redo
  createSnapshot(operationType, filePath, tableIndex, coordinates, oldColors, newColors) {
    return {
      timestamp: Date.now(),
      operationType,
      filePath,
      tableIndex,
      coordinates,
      oldColors,
      newColors
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
        if (rowKey.startsWith('row_')) {
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
        if (rowKey.startsWith('row_')) {
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
      localStorage.setItem('table-color-undo-stack', JSON.stringify(this.undoStack));
      localStorage.setItem('table-color-redo-stack', JSON.stringify(this.redoStack));
    } catch (e) {}
  }

  async loadUndoRedoStacks() {
    try {
      const u = localStorage.getItem('table-color-undo-stack');
      const r = localStorage.getItem('table-color-redo-stack');
      if (u) this.undoStack = JSON.parse(u) || [];
      if (r) this.redoStack = JSON.parse(r) || [];
    } catch (e) {}
  }

  // Apply rule and saved colors to a DOM container
  // Helper function to get visible text content from a cell, excluding editor markup and cursors
  getCellText(cell) {
    // For table cells in reading mode
    let text = '';
    
    // Get text from all child nodes, excluding editor elements
    const walkNodes = (node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        text += node.textContent;
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        // Skip CodeMirror cursor and editor elements
        if (node.classList && (node.classList.contains('cm-cursor') || node.classList.contains('cm-line'))) {
          return; // Skip editor cursors
        }
        // Include br tags as newlines
        if (node.tagName === 'BR') {
          text += '\n';
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
      rect.bottom <= (window.innerHeight || document.documentElement.clientHeight) &&
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
    const val = rule.value ?? '';
    const t = (text ?? '').trim();
    const isEmpty = t.length === 0;

    const toNumber = (s) => {
      const cleaned = String(s).replace(/,/g, '');
      const n = parseFloat(cleaned);
      if (this.settings.numericStrict) {
        const ok = /^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$/.test(String(s).trim());
        return ok ? n : NaN;
      }
      return isNaN(n) ? NaN : n;
    };

    switch (rule.match) {
      case 'is':
        return t.toLowerCase() === String(val).toLowerCase();
      case 'isNot':
        return t.toLowerCase() !== String(val).toLowerCase();
      case 'isRegex':
        try {
          const rx = new RegExp(String(val), 'i');
          return rx.test(t);
        } catch { return false; }
      case 'contains':
        return t.toLowerCase().includes(String(val).toLowerCase());
      case 'notContains':
        return !t.toLowerCase().includes(String(val).toLowerCase());
      case 'startsWith':
        return t.toLowerCase().startsWith(String(val).toLowerCase());
      case 'endsWith':
        return t.toLowerCase().endsWith(String(val).toLowerCase());
      case 'notStartsWith':
        return !t.toLowerCase().startsWith(String(val).toLowerCase());
      case 'notEndsWith':
        return !t.toLowerCase().endsWith(String(val).toLowerCase());
      case 'isEmpty':
        return isEmpty;
      case 'isNotEmpty':
        return !isEmpty;
      case 'eq': {
        const n = toNumber(t); const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n === v;
      }
      case 'gt': {
        const n = toNumber(t); const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n > v;
      }
      case 'lt': {
        const n = toNumber(t); const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n < v;
      }
      case 'ge': {
        const n = toNumber(t); const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n >= v;
      }
      case 'le': {
        const n = toNumber(t); const v = toNumber(val);
        if (isNaN(n) || isNaN(v)) return false;
        return n <= v;
      }
      default:
        return false;
    }
  }

  applyColoringRulesToTable(tableEl) {
    const rules = Array.isArray(this.settings.coloringRules) ? this.settings.coloringRules : [];
    if (!rules.length) return;

    const rows = Array.from(tableEl.querySelectorAll('tr'));
    const getCell = (r, c) => {
      const row = rows[r];
      if (!row) return null;
      const cells = Array.from(row.querySelectorAll('td, th'));
      return cells[c] || null;
    };
    const texts = rows.map(row => Array.from(row.querySelectorAll('td, th')).map(cell => this.getCellText(cell)));
    const maxCols = Math.max(0, ...texts.map(r => r.length));

    const headerRowIndex = rows.findIndex(r => r.querySelector('th'));
    const hdr = headerRowIndex >= 0 ? headerRowIndex : 0;
    const firstDataRowIndex = rows.findIndex(r => r.querySelector('td'));
    const fdr = firstDataRowIndex >= 0 ? firstDataRowIndex : 0;

    for (const rule of rules) {
      if (!rule || !rule.target || !rule.match) continue;
      const applyCellStyle = (cell) => {
        if (!cell) return;
        if (cell.hasAttribute('data-ctc-manual')) return;
        // For header cells (th), allow overriding existing colors from rules
        // For data cells (td), check if they already have colors
        const isHeaderCell = cell.tagName === 'TH';
        if (!isHeaderCell && (cell.style.backgroundColor || cell.style.color)) return;
        if (rule.bg) cell.style.backgroundColor = rule.bg;
        if (rule.color) cell.style.color = rule.color;
      };

      if (rule.target === 'cell') {
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < (texts[r]?.length || 0); c++) {
            const text = texts[r][c];
            if ((rule.when || 'theCell') === 'theCell') {
              if (this.evaluateMatch(text, rule)) {
                applyCellStyle(getCell(r, c));
              }
            }
          }
        }
      } else if (rule.target === 'row') {
        const candidateRows = (rule.when === 'firstRow') ? [fdr] : Array.from({ length: rows.length }, (_, i) => i);
        for (const r of candidateRows) {
          const rowTexts = texts[r] || [];
          let cond = false;
          if (rule.when === 'allCell') cond = rowTexts.length > 0 && rowTexts.every(t => this.evaluateMatch(t, rule));
          else if (rule.when === 'noCell') cond = rowTexts.every(t => !this.evaluateMatch(t, rule));
          else cond = rowTexts.some(t => this.evaluateMatch(t, rule)); // anyCell or firstRow default any
          if (cond) {
            const cells = Array.from(rows[r].querySelectorAll('td, th'));
            cells.forEach(applyCellStyle);
          }
        }
      } else if (rule.target === 'column') {
        const candidateCols = Array.from({ length: maxCols }, (_, i) => i);
        for (const c of candidateCols) {
          let cond = false;
          if (rule.when === 'columnHeader') {
            const text = texts[hdr]?.[c] ?? '';
            cond = this.evaluateMatch(text, rule);
          } else {
            const colTexts = Array.from({ length: rows.length }, (_, r) => texts[r]?.[c]).filter(t => t !== undefined);
            if (rule.when === 'allCell') cond = colTexts.length > 0 && colTexts.every(t => this.evaluateMatch(t, rule));
            else if (rule.when === 'noCell') cond = colTexts.every(t => !this.evaluateMatch(t, rule));
            else cond = colTexts.some(t => this.evaluateMatch(t, rule)); // anyCell default
          }
          if (cond) {
            for (let r = 0; r < rows.length; r++) applyCellStyle(getCell(r, c));
          }
        }
      }
    }
  }

  applyAdvancedRulesToTable(tableEl) {
    const adv = Array.isArray(this.settings.advancedRules) ? this.settings.advancedRules : [];
    if (!adv.length) return;
    const rows = Array.from(tableEl.querySelectorAll('tr'));
    const texts = rows.map(row => Array.from(row.querySelectorAll('td, th')).map(cell => this.getCellText(cell)));
    const maxCols = Math.max(0, ...texts.map(r => r.length));
    const headerRowIndex = rows.findIndex(r => r.querySelector('th'));
    const hdr = headerRowIndex >= 0 ? headerRowIndex : 0;
    const firstDataRowIndex = rows.findIndex(r => r.querySelector('td'));
    const fdr = firstDataRowIndex >= 0 ? firstDataRowIndex : 0;
    const getCell = (r, c) => { const row = rows[r]; if (!row) return null; const cells = Array.from(row.querySelectorAll('td, th')); return cells[c] || null; };

    const evalCondCell = (r, c, cond) => this.evaluateMatch(texts[r]?.[c] ?? '', { match: cond.match, value: cond.value });
    const evalCondRow = (r, cond) => { const rowTexts = texts[r] || []; return rowTexts.some(t => this.evaluateMatch(t, { match: cond.match, value: cond.value })); };
    const evalCondHeader = (c, cond) => this.evaluateMatch(texts[hdr]?.[c] ?? '', { match: cond.match, value: cond.value });

    for (const rule of adv) {
      const logic = rule.logic || 'any';
      const target = rule.target || 'cell';
      const color = rule.color || null;
      const bg = rule.bg || null;
      if (!bg && !color) continue;
      const conditions = Array.isArray(rule.conditions) ? rule.conditions : [];
      if (!conditions.length) continue;

      if (target === 'row') {
        const hasHeaderConds = conditions.some(cond => cond.when === 'columnHeader');
        const allHeaderConds = conditions.length > 0 && conditions.every(cond => cond.when === 'columnHeader');
        let candidateRows = rule.when === 'firstRow' ? [fdr] : Array.from({ length: rows.length }, (_, i) => i);
        if (allHeaderConds) {
          candidateRows = [hdr];
          debugLog(`Row target uses only columnHeader conditions; coloring header row index ${hdr}`);
        }
        for (const r of candidateRows) {
          const flags = conditions.map(cond => {
            if (cond.when === 'columnHeader') {
              // For row target with column header condition, check if ANY column header matches
              for (let c = 0; c < maxCols; c++) { if (evalCondHeader(c, cond)) return true; }
              return false;
            } else if (cond.when === 'row') {
              return evalCondRow(r, cond);
            } else {
              // anyCell, allCell, noCell for row target
              const cells = Array.from(rows[r].querySelectorAll('td, th'));
              const cellResults = cells.map((_, c) => evalCondCell(r, c, cond));
              if (cond.when === 'allCell') return cellResults.every(Boolean);
              if (cond.when === 'noCell') return cellResults.every(f => !f);
              return cellResults.some(Boolean); // anyCell
            }
          });
          let ok = false;
          if (logic === 'all') ok = flags.every(Boolean);
          else if (logic === 'none') ok = flags.every(f => !f);
          else ok = flags.some(Boolean);
          debugLog(`Row ${r}: flags=${JSON.stringify(flags)}, logic=${logic}, ok=${ok}, bg=${bg}, color=${color}`);
          if (ok) {
            debugLog(`  -> Coloring row ${r} with bg=${bg} color=${color}`);
            Array.from(rows[r].querySelectorAll('td, th')).forEach(cell => { 
              if (cell.hasAttribute('data-ctc-manual')) return;
              // For header cells, allow overriding; for data cells, skip if already colored
              const isHeaderCell = cell.tagName === 'TH';
              if (!isHeaderCell && (cell.style.backgroundColor || cell.style.color)) return;
              if (bg) cell.style.backgroundColor = bg; 
              if (color) cell.style.color = color; 
            });
          }
        }
      } else if (target === 'column') {
        for (let c = 0; c < maxCols; c++) {
          const flags = conditions.map(cond => {
            if (cond.when === 'columnHeader') return evalCondHeader(c, cond);
            if (cond.when === 'row') return false;
            // anyCell, allCell, noCell for column
            const colCells = [];
            for (let r = 0; r < rows.length; r++) { 
              colCells.push(evalCondCell(r, c, cond));
            }
            if (cond.when === 'allCell') return colCells.every(Boolean);
            if (cond.when === 'noCell') return colCells.every(f => !f);
            return colCells.some(Boolean); // anyCell
          });
          let ok = false;
          if (logic === 'all') ok = flags.every(Boolean);
          else if (logic === 'none') ok = flags.every(f => !f);
          else ok = flags.some(Boolean);
          if (ok) { for (let r = 0; r < rows.length; r++) { const cell = getCell(r, c); if (cell && !cell.hasAttribute('data-ctc-manual') && !cell.style.backgroundColor && !cell.style.color) { if (bg) cell.style.backgroundColor = bg; if (color) cell.style.color = color; } } }
        }
      } else {
        // target === 'cell' - color individual cells
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < (texts[r]?.length || 0); c++) {
            const flags = conditions.map(cond => {
              if (cond.when === 'columnHeader') {
                // For cell target: check if the column header matches
                return evalCondHeader(c, cond);
              } else if (cond.when === 'row') {
                // For cell target: check if any cell in the row matches
                return evalCondRow(r, cond);
              } else if (cond.when === 'anyCell' || cond.when === 'allCell' || cond.when === 'noCell') {
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
            if (logic === 'all') ok = flags.every(Boolean);
            else if (logic === 'none') ok = flags.every(f => !f);
            else ok = flags.some(Boolean);
            if (ok) { 
              const cell = getCell(r, c); 
              if (cell && !cell.hasAttribute('data-ctc-manual') && !cell.style.backgroundColor && !cell.style.color) { 
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
    
    if (element.hasAttribute('data-ctc-bg')) {
      const bg = element.getAttribute('data-ctc-bg');
      if (bg && bg !== element.style.backgroundColor) {
        element.style.backgroundColor = bg;
      }
    }
    if (element.hasAttribute('data-ctc-color')) {
      const color = element.getAttribute('data-ctc-color');
      if (color && color !== element.style.color) {
        element.style.color = color;
      }
    }
  }

  applyColorsToContainer(container, filePath) {
    // Only apply colors in Reading mode or Live Preview if enabled
    const hasClosest = typeof container.closest === 'function';
    const inPreview = hasClosest && !!container.closest('.markdown-preview-view');
    const inEditor = hasClosest && (container.closest('.cm-content') || container.closest('.cm-editor') || container.closest('.cm-scroller'));
    if (!inPreview && (!this.settings.livePreviewColoring || !inEditor)) {
      // Ensure container is within a preview/editor container
      let p = container && container.parentElement;
      let found = false;
      while (p) {
        if (p.classList && p.classList.contains('markdown-preview-view')) {
          found = true;
          break;
        }
        if (this.settings.livePreviewColoring && (p.classList && (p.classList.contains('cm-content') || p.classList.contains('cm-editor') || p.classList.contains('cm-scroller')))) {
          found = true;
          break;
        }
        p = p.parentElement;
      }
      if (!found) {
        // In Live Preview, allow direct .cm-content root
        if (this.settings.livePreviewColoring && container.classList && container.classList.contains('cm-content')) {
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
    if (!inPreview && !inEditor && !(container.classList && container.classList.contains('cm-content'))) {
      return;
    }
    
    // Performance optimization: Debounce rapid calls
    const now = Date.now();
    const lastCall = this._lastApplyCall || 0;
    if (now - lastCall < 100) { // Debounce to 100ms
      return;
    }
    this._lastApplyCall = now;
    debugLog(`=== applyColorsToContainer called ===`);
    debugLog(`container:`, container);
    // Special handling for reading mode: always process all tables in the view
    if (inPreview) {
      // Check if this is a reading view container
      const readingView = container.closest('.markdown-preview-view');
      if (readingView) {
        // Process ALL tables in the reading view, not just in this container
        const allTablesInView = Array.from(readingView.querySelectorAll('table'));
        debugLog(`Reading mode: found ${allTablesInView.length} total tables in view`);
        
        // If we have tables but they're not all in our container, process the entire view
        if (allTablesInView.length > 0 && allTablesInView.some(table => !container.contains(table))) {
          debugLog('Processing entire reading view for comprehensive coloring');
          container = readingView; // Process the entire reading view
        }
      }
    }
    
    const tables = Array.from(container.querySelectorAll('table'));
    if (!tables.length) return;
    const noteData = this.cellData[filePath] || {};
    
    debugLog(`applyColorsToContainer: filePath=${filePath}, found ${tables.length} tables, has noteData:`, Object.keys(noteData).length > 0, 'inPreview:', inPreview, 'inEditor:', inEditor);
    // Get ALL tables in the correct scope - ALWAYS use global table list
    let allTables = Array.from(document.querySelectorAll('table'));

    debugLog(`All tables in document: ${allTables.length}, container tables: ${tables.length}, inPreview: ${inPreview}, inEditor: ${inEditor}`);

    // Process each table
    let manualAppliedCount = 0;
    tables.forEach((tableEl) => {
      const globalTableIndex = allTables.indexOf(tableEl);
      if (globalTableIndex === -1) {
        debugWarn(`Could not find table in global list, using local index`);
        const containerTables = Array.from(container.querySelectorAll('table'));
        const fallbackIndex = containerTables.indexOf(tableEl);
        manualAppliedCount += this.processSingleTable(tableEl, fallbackIndex, filePath, noteData) || 0;
      } else {
        manualAppliedCount += this.processSingleTable(tableEl, globalTableIndex, filePath, noteData) || 0;
      }
    });
    debugLog(`applyColorsToContainer: manual colors applied to ${manualAppliedCount} cells in this container`);
    // In Reading mode, retry applying colors with escalating delays if needed
    try {
      const prev = this._appliedContainers.get(container) || 0;
      const delays = [100, 200, 400, 800];
      if (prev < delays.length) {
        this._appliedContainers.set(container, prev + 1);
        setTimeout(() => {
          // Only retry if still connected to DOM
          if (container.isConnected) {
            this.applyColorsToContainer(container, filePath);
          }
        }, delays[prev]);
      }
    } catch (e) { }
  }
  applyColorsToActiveFile() {
    debugLog('applyColorsToActiveFile called');
    const file = this.app.workspace.getActiveFile();
    if (!file) {
      debugWarn('No active file found in applyColorsToActiveFile');
      return;
    }
    
    const noteData = this.cellData[file.path] || {};
    debugLog(`applyColorsToActiveFile: file=${file.path}, has noteData keys:`, Object.keys(noteData));
    
    // Restrict reading mode application to views showing the active file
    let previewViews = [];
    try {
      const leaves = typeof this.app.workspace.getLeavesOfType === 'function' ? this.app.workspace.getLeavesOfType('markdown') : [];
      const activeContainers = leaves
        .filter(l => l.view && l.view.file && l.view.file.path === file.path)
        .map(l => l.view && (l.view.containerEl || l.view.contentEl))
        .filter(Boolean);
      activeContainers.forEach(container => {
        const views = Array.from(container.querySelectorAll('.markdown-preview-view'));
        previewViews.push(...views);
      });
    } catch (e) {
      previewViews = Array.from(document.querySelectorAll('.markdown-preview-view'));
    }
    debugLog(`applyColorsToActiveFile: found ${previewViews.length} preview views for active file`);
    
    // Use GLOBAL table index consistently across saving and applying
    const allDocTables = Array.from(document.querySelectorAll('table'));
    
    // If no tables found yet, retry shortly
    if (previewViews.length > 0) {
      const totalTables = previewViews.reduce((acc, v) => acc + v.querySelectorAll('table').length, 0);
      if (totalTables === 0) {
        debugLog('No reading mode tables found in active views, retrying after 100ms');
        setTimeout(() => this.applyColorsToActiveFile(), 100);
        return;
      }
    }
    
    // Clear all rule-based colors first (in both reading and live preview)
    // This ensures rule color changes take effect
    previewViews.forEach(view => {
      view.querySelectorAll('td, th').forEach(cell => {
        if (!cell.hasAttribute('data-ctc-manual')) {
          cell.style.backgroundColor = '';
          cell.style.color = '';
        }
      });
    });
    
    if (this.settings.livePreviewColoring) {
      document.querySelectorAll('.cm-content table td, .cm-content table th').forEach(cell => {
        if (!cell.hasAttribute('data-ctc-manual')) {
          cell.style.backgroundColor = '';
          cell.style.color = '';
        }
      });
    }
    
    // Apply to reading mode using global indices
    previewViews.forEach(view => {
      if (view.isConnected) {
        const viewTables = Array.from(view.querySelectorAll('table'));
        let readingManualApplied = 0;
        viewTables.forEach((table) => {
          const globalIdx = allDocTables.indexOf(table);
          if (globalIdx >= 0) {
            readingManualApplied += this.processSingleTable(table, globalIdx, file.path, noteData) || 0;
          }
        });
        debugLog(`Reading mode: manual colors applied to ${readingManualApplied} cells in this view`);
      }
    });
    
    // Apply to live preview if enabled - use simple global indexing for consistency
    if (this.settings.livePreviewColoring) {
      const cmEditors = document.querySelectorAll('.cm-content');
      debugLog(`applyColorsToActiveFile: Found ${cmEditors.length} live preview editors`);
      
      cmEditors.forEach((editor, editorIdx) => {
        if (editor.isConnected) {
          const editorTables = Array.from(editor.querySelectorAll('table'));
          debugLog(`  Editor ${editorIdx}: ${editorTables.length} tables`);
          let editorManualApplied = 0;
          editorTables.forEach((table, localIdx) => {
            const globalIdx = allDocTables.indexOf(table);
            if (globalIdx >= 0) {
              editorManualApplied += this.processSingleTable(table, globalIdx, file.path, noteData) || 0;
            }
          });
          debugLog(`  Editor ${editorIdx}: manual colors applied to ${editorManualApplied} cells`);
        }
      });
    }
  }

  // Get a signature for a table based on its structure (row/col count)
  getTableSignature(table) {
    const rows = table.querySelectorAll('tr').length;
    const cols = table.querySelector('tr') ? table.querySelector('tr').querySelectorAll('td, th').length : 0;
    return `${rows}x${cols}`;
  }

  // Get table text content for matching tables between reading and live preview modes
  getTableTextContent(table) {
    const rows = Array.from(table.querySelectorAll('tr'));
    const text = rows.map(row => {
      const cells = Array.from(row.querySelectorAll('td, th'));
      return cells.map(cell => this.getCellText(cell).trim()).join('|');
    }).join('\n');
    debugLog(`  getTableTextContent: ${rows.length} rows, content hash: ${text.substring(0, 100)}`);
    return text;
  }

  processSingleTable(tableEl, tableIndex, filePath, noteData) {
    debugLog(`Processing single table: index=${tableIndex}, has data-ctc-index: ${tableEl.getAttribute('data-ctc-index')}`);
    const inLivePreview = !!tableEl.closest('.cm-content');
    const inReadingMode = !!tableEl.closest('.markdown-preview-view');
    debugLog(`Table context: ${inLivePreview ? 'LivePreview' : (inReadingMode ? 'Reading' : 'Unknown')}`);
    
    // Mark table with unique data attribute for persistence
    if (!tableEl.hasAttribute('data-ctc-processed')) {
      tableEl.setAttribute('data-ctc-processed', 'true');
      tableEl.setAttribute('data-ctc-index', tableIndex);
      tableEl.setAttribute('data-ctc-file', filePath);
    }
    
    const tableKey = `table_${tableIndex}`;
    const tableColors = noteData[tableKey] || {};
    
    debugLog(`Table index ${tableIndex}: key="${tableKey}", has colors:`, Object.keys(tableColors).length > 0);
    
    let coloredCount = 0;
    const tableId = `${filePath}:${tableIndex}`;
    const manualColoredCells = new Set(); // Track which cells have manual colors
    
    // NOTE: Clearing is now done in applyColorsToActiveFile before this function is called
    // STEP 1: Apply manual colors using rule-based rendering (direct style application)
    Array.from(tableEl.querySelectorAll('tr')).forEach((tr, rIdx) => {
      const rowKey = `row_${rIdx}`;
      const rowColors = tableColors[rowKey] || {};
      const cells = Array.from(tr.querySelectorAll('td, th'));
      
      cells.forEach((cell, cIdx) => {
        const colorData = rowColors[`col_${cIdx}`];
        if (colorData) {
          coloredCount++;
          manualColoredCells.add(cell);
          
          // Apply colors directly using rule-based approach (primary method)
          if (colorData.bg) {
            cell.style.backgroundColor = colorData.bg;
          }
          if (colorData.color) {
            cell.style.color = colorData.color;
          }
          
          // Store attributes as fallback for restoring colors if styles are cleared
          if (colorData.bg) {
            cell.setAttribute('data-ctc-bg', colorData.bg);
          }
          if (colorData.color) {
            cell.setAttribute('data-ctc-color', colorData.color);
          }
          
          // Mark as manual so rules skip it
          cell.setAttribute('data-ctc-manual', 'true');
          cell.setAttribute('data-ctc-table-id', tableId);
          cell.setAttribute('data-ctc-row', String(rIdx));
          cell.setAttribute('data-ctc-col', String(cIdx));
        }
      });
    });
    
    debugLog(`Table index ${tableIndex}: colored ${coloredCount} cells with manual colors`);
    
    // STEP 2: Apply rules after manual; rule methods will skip cells marked as manual
    this.applyColoringRulesToTable(tableEl);
    this.applyAdvancedRulesToTable(tableEl);
    
    // STEP 3: Fallback restoration - re-enforce manual colors if rules or selection cleared them
    // This is now only a safety net, not the primary method
    manualColoredCells.forEach(cell => {
      const bg = cell.getAttribute('data-ctc-bg');
      const color = cell.getAttribute('data-ctc-color');
      // Only restore if styles were unexpectedly cleared (fallback behavior)
      if (bg && !cell.style.backgroundColor) {
        cell.style.backgroundColor = bg;
      }
      if (color && !cell.style.color) {
        cell.style.color = color;
      }
    });
    
    // Ensure colors persist by marking the table as processed
    tableEl.setAttribute('data-ctc-last-processed', Date.now());
    
    return coloredCount;
  }

  setupReadingViewScrollListener() {
    // Watch for scroll events in reading view to catch lazy-loaded tables
    const handleReadingViewScroll = () => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;
      
      // Check if we're in reading mode
      const readingViews = document.querySelectorAll('.markdown-preview-view');
      if (readingViews.length === 0) return;
      
      // Debounce scroll events
      if (this._readingScrollTimeout) clearTimeout(this._readingScrollTimeout);
      this._readingScrollTimeout = setTimeout(() => {
        debugLog('Reading view scroll detected, checking for new tables');
        readingViews.forEach(view => {
          if (view.isConnected) {
            // Apply colors to the entire reading view (not just container)
            this.applyColorsToContainer(view, file.path);
          }
        });
      }, 150); // Slightly longer debounce for scroll
    };
    
    // Add scroll listener to all reading views
    const addScrollListeners = () => {
      document.querySelectorAll('.markdown-preview-view').forEach(view => {
        if (!view._ctcScrollListenerAdded) {
          // Save handler reference for cleanup
          view._ctcScrollHandler = handleReadingViewScroll;
          view.addEventListener('scroll', view._ctcScrollHandler, { passive: true });
          view._ctcScrollListenerAdded = true;
          debugLog('Added scroll listener to reading view');
        }
      });
    };
    
    // Initial setup
    addScrollListeners();
    
    // Watch for new reading views being created
    const readingViewObserver = new MutationObserver(() => {
      addScrollListeners();
    });
    
    readingViewObserver.observe(document.body, { childList: true, subtree: true });
    this._readingViewScrollObserver = readingViewObserver;
  }

  startReadingModeTableChecker() {
    // Stop any existing checker
    if (this._readingModeChecker) clearInterval(this._readingModeChecker);
    
    this._readingModeChecker = setInterval(() => {
      const file = this.app.workspace.getActiveFile();
      if (!file) return;
      
      // Restrict to reading views of the active file
      let previewViews = [];
      try {
        const leaves = typeof this.app.workspace.getLeavesOfType === 'function' ? this.app.workspace.getLeavesOfType('markdown') : [];
        const activeContainers = leaves
          .filter(l => l.view && l.view.file && l.view.file.path === file.path)
          .map(l => l.view && (l.view.containerEl || l.view.contentEl))
          .filter(Boolean);
        activeContainers.forEach(container => {
          const views = Array.from(container.querySelectorAll('.markdown-preview-view'));
          previewViews.push(...views);
        });
      } catch (e) {
        previewViews = Array.from(document.querySelectorAll('.markdown-preview-view'));
      }
      if (previewViews.length === 0) return;
      
      const noteData = this.cellData[file.path] || {};
      const allDocTables = Array.from(document.querySelectorAll('table'));
      
      previewViews.forEach(view => {
        const viewTables = Array.from(view.querySelectorAll('table'));
        viewTables.forEach(table => {
          if (!table.hasAttribute('data-ctc-processed')) {
            const globalTableIdx = allDocTables.indexOf(table);
            this.processSingleTable(table, globalTableIdx, file.path, noteData);
          }
        });
        // Ensure rules are applied to all tables in this view
        viewTables.forEach(table => {
          this.applyColoringRulesToTable(table);
          this.applyAdvancedRulesToTable(table);
        });
      });
    }, 500); // Check every 500ms for new tables that need coloring
    
    // Clean up on unload
    this.register(() => {
      if (this._readingModeChecker) clearInterval(this._readingModeChecker);
    });
  }
};

// Regex Tester Modal
class RegexTesterModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl('h2', { text: 'Regex Pattern Tester' });
    
    const patternDiv = contentEl.createDiv({ cls: 'regex-tester-section' });
    patternDiv.createEl('label', { text: 'Pattern:' });
    const patternInput = patternDiv.createEl('input', { type: 'text', cls: 'regex-pattern-input', attr: { placeholder: 'Enter regex pattern' } });
    
    const flagsDiv = contentEl.createDiv({ cls: 'regex-tester-section' });
    flagsDiv.createEl('label', { text: 'Flags:' });
    const flagsInput = flagsDiv.createEl('input', { type: 'text', value: 'gi', cls: 'regex-flags-input' });
    
    const testDiv = contentEl.createDiv({ cls: 'regex-tester-section' });
    testDiv.createEl('label', { text: 'Test Text:' });
    const testInput = testDiv.createEl('textarea', { cls: 'regex-test-input', attr: { placeholder: 'Enter text to test' } });
    
    const resultDiv = contentEl.createDiv({ cls: 'regex-test-result' });
    
    const testPattern = () => {
      resultDiv.empty();
      if (!patternInput.value.trim()) {
        resultDiv.createEl('div', { text: 'Enter a pattern to test', cls: 'test-info' });
        return;
      }
      
      try {
        const regex = new RegExp(patternInput.value, flagsInput.value);
        const matches = testInput.value.match(regex);
        
        if (matches && matches.length > 0) {
          const successDiv = resultDiv.createEl('div', { cls: 'test-success' });
          successDiv.createEl('div', { text: `✓ Found ${matches.length} match(es)` });
          
          const matchContainer = resultDiv.createDiv({ cls: 'matches-container' });
          matches.forEach((match, i) => {
            matchContainer.createEl('div', { 
              text: `[${i}]: "${match}"`,
              cls: 'match-item' 
            });
          });
        } else {
          resultDiv.createEl('div', { 
            text: '✗ No matches found',
            cls: 'test-failure'
          });
        }
      } catch (e) {
        resultDiv.createEl('div', { 
          text: `⚠ Error: ${e.message}`,
          cls: 'test-error'
        });
      }
    };
    
    const buttonContainer = contentEl.createDiv({ cls: 'regex-tester-buttons' });
    const testButton = buttonContainer.createEl('button', { text: 'Test Pattern', cls: 'mod-cta' });
    testButton.onclick = testPattern;
    
    const closeButton = buttonContainer.createEl('button', { text: 'Close', cls: 'mod-warning' });
    closeButton.onclick = () => this.close();
    
    // Test on input changes
    patternInput.oninput = testPattern;
    flagsInput.oninput = testPattern;
    testInput.oninput = testPattern;
  }

  onClose() {
    this.contentEl.empty();
  }
}

class ConditionRow {
  constructor(parent, index, initialData, onChange) {
    this.root = parent.createDiv({ cls: 'num-edit-row pretty-flex' });
    Object.assign(this.root.style, { display: 'flex', alignItems: 'center', gap: '8px', padding: '6px 8px', borderRadius: '6px', background: 'var(--background-secondary, #232323)' });
    this.typeSel = this.root.createEl('select');
    ['text','numeric','date','regex','empty'].forEach(t => { const o = this.typeSel.createEl('option'); o.value = t; o.text = t; });
    this.opSel = this.root.createEl('select');
    this.valInput = this.root.createEl('input', { type: 'text' });
    Object.assign(this.valInput.style, { flex: '1', padding: '6px 8px' });
    this.val2Input = this.root.createEl('input', { type: 'text' });
    this.val2Input.style.display = 'none';
    Object.assign(this.val2Input.style, { flex: '1', padding: '6px 8px' });
    this.caseChk = this.root.createEl('input', { type: 'checkbox' });
    this.logicSel = this.root.createEl('select');
    ;['AND','OR'].forEach(t => { const o = this.logicSel.createEl('option'); o.value = t; o.text = t; });
    const delBtn = this.root.createEl('button', { cls: 'mod-ghost' });
    delBtn.textContent = 'Delete';
    const setOps = () => {
      const t = this.typeSel.value;
      const currentOp = this.opSel.value; // Remember current selection
      this.opSel.empty();
      let options = [];
      if (t === 'text') options = ['contains','equals','startsWith','endsWith'];
      else if (t === 'numeric') options = ['gt','ge','eq','le','lt','between'];
      else if (t === 'date') options = ['before','after','between'];
      else if (t === 'regex') options = ['matches'];
      else options = ['isEmpty'];
      
      options.forEach(op => { 
        const o = this.opSel.createEl('option'); 
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
      const useTwo = op === 'between';
      this.val2Input.style.display = useTwo ? '' : 'none';
      if (t === 'numeric') { this.valInput.type = 'number'; this.val2Input.type = 'number'; }
      else if (t === 'date') { this.valInput.type = 'date'; this.val2Input.type = 'date'; }
      else { this.valInput.type = 'text'; this.val2Input.type = 'text'; }
      this.caseChk.style.display = (t === 'text' || t === 'regex') ? '' : 'none';
      if (typeof onChange === 'function') onChange();
    };
    this.typeSel.onchange = setOps;
    this.opSel.onchange = setOps;
    delBtn.onclick = () => { this.root.remove(); if (typeof onChange === 'function') onChange(); };
    if (initialData) {
      this.typeSel.value = initialData.type || 'text';
      setOps();
      if (initialData.operator) this.opSel.value = initialData.operator;
      if (initialData.value != null) this.valInput.value = initialData.value;
      if (initialData.value2 != null) { this.val2Input.value = initialData.value2; this.val2Input.style.display = ''; }
      this.caseChk.checked = !!initialData.caseSensitive;
      if (initialData.logic) this.logicSel.value = initialData.logic.toUpperCase();
    } else { setOps(); }
  }
  getData() {
    const t = this.typeSel.value;
    const op = this.opSel.value;
    const v = t === 'numeric' ? (this.valInput.value !== '' ? Number(this.valInput.value) : null) : this.valInput.value;
    const v2 = t === 'numeric' || t === 'date' ? (this.val2Input.style.display === '' ? (this.val2Input.value !== '' ? (t === 'numeric' ? Number(this.val2Input.value) : this.val2Input.value) : null) : null) : null;
    return { type: t, operator: op, value: v, value2: v2, caseSensitive: this.caseChk.checked, logic: this.logicSel.value };
  }
}



// Settings Tab
// Release Notes Modal - displays latest release notes from GitHub
class ReleaseNotesModal extends Modal {
  constructor(app) {
    super(app);
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    
    try {
      this.modalEl.style.maxWidth = '900px';
      this.modalEl.style.width = '900px';
      this.modalEl.style.padding = '25px';
    } catch (e) {}

    const header = contentEl.createEl('div');
    header.style.display = 'flex';
    header.style.alignItems = 'center';
    header.style.justifyContent = 'space-between';
    header.style.marginBottom = '0px';
    header.style.paddingBottom = '16px';
    header.style.borderBottom = '1px solid var(--divider-color)';

    const title = header.createEl('h2', { text: 'Color Table Cells' });
    title.style.margin = '0';
    title.style.fontSize = '1.5em';
    title.style.fontWeight = '600';

    const link = header.createEl('a', { text: 'View on GitHub' });
    link.href = 'https://github.com/Kazi-Aidah/color-table-cells/releases';
    link.target = '_blank';
    link.style.fontSize = '0.9em';
    link.style.opacity = '0.8';
    link.style.transition = 'opacity 0.2s';
    link.addEventListener('mouseenter', () => link.style.opacity = '1');
    link.addEventListener('mouseleave', () => link.style.opacity = '0.8');

    const body = contentEl.createDiv();
    body.style.maxHeight = '70vh';
    body.style.overflow = 'auto';

    const loading = body.createEl('div', { text: 'Loading release notes...' });
    loading.style.opacity = '0.7';
    loading.style.fontSize = '0.95em';
    loading.style.marginTop = '6px';

    // Fetch release notes from GitHub
    fetch('https://api.github.com/repos/Kazi-Aidah/color-table-cells/releases/latest')
      .then(response => response.json())
      .then(data => {
        body.empty();
        
        if (data.message === 'Not Found' || !data) {
          body.createEl('div', { text: 'No release information available.' });
          return;
        }

        const releaseName = body.createEl('div', { text: data.name || data.tag_name || 'Release' });
        releaseName.style.fontSize = '2em';
        releaseName.style.fontWeight = '900';
        releaseName.style.marginTop = '12px';
        releaseName.style.marginBottom = '12px';
        releaseName.style.color = 'var(--text-normal)';

        if (data.published_at) {
          const publishedDate = new Date(data.published_at);
          const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
          const formattedDate = `${publishedDate.getFullYear()} ${monthNames[publishedDate.getMonth()]} ${String(publishedDate.getDate()).padStart(2, '0')}`;
          const date = body.createEl('p', { text: `Released: ${formattedDate}` });
          date.style.color = 'var(--text-muted)';
          date.style.marginTop = '4px';
          date.style.marginBottom = '20px';
        }

        const notes = body.createEl('div');
        notes.style.marginTop = '16px';
        notes.style.lineHeight = '1.6';
        notes.style.fontSize = '0.95em';
        notes.style.padding = '12px !important';
        // notes.addClass('markdown-preview-view');

        const md = data.body || 'No notes';
        try {
          const { MarkdownRenderer } = require('obsidian');
          const comp = new require('obsidian').Component();
          MarkdownRenderer.render(this.app, md, notes, '', comp);
        } catch (e) {
          // Fallback: render markdown as plain text with line breaks
          const lines = md.split('\n');
          lines.forEach(line => {
            if (line.trim()) {
              notes.createEl('p', { text: line });
            }
          });
        }
      })
      .catch(error => {
        body.empty();
        body.createEl('div', { text: 'Failed to load release notes.' });
        debugWarn('Error fetching release notes:', error);
      });
  }

  onClose() {
    this.contentEl.empty();
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

    new Setting(containerEl).setName("Color Table Cells Settings").setHeading();
    
    new Setting(containerEl)
      .setName('Latest Release Notes')
      .setDesc('View the most recent plugin release notes')
      .addButton(btn => btn
        .setButtonText('Open Changelog')
        .onClick(() => new ReleaseNotesModal(this.app).open())
      );

    // Toggle for Live Preview Table Coloring
    new Setting(containerEl)
      .setName("Live Preview Table Coloring")
      .setDesc("Apply table coloring in Live Preview (editor) mode. Disabled by default; enabling may affect editor performance.")
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.livePreviewColoring)
              .onChange(async val => {
                this.plugin.settings.livePreviewColoring = val;
                await this.plugin.saveSettings();
                // Force document refresh so table coloring updates immediately!
                if (this.plugin.app.workspace && typeof this.plugin.app.workspace.trigger === 'function') {
                  this.plugin.app.workspace.trigger('layout-change');
                }
                // If disabling, remove colors from all .cm-content tabless
                if (!val) {
                  document.querySelectorAll('.cm-content table').forEach(table => {
                    table.querySelectorAll('td, th').forEach(cell => {
                      cell.style.backgroundColor = '';
                      cell.style.color = '';
                    });
                  });
                } else {
                  // If enabling, immediately apply colors to all editors
                  if (typeof this.plugin.applyColorsToAllEditors === 'function') {
                    setTimeout(() => this.plugin.applyColorsToAllEditors(), 0);
                  }
                }
              }));

    // Toggle for strict numeric matching
    new Setting(containerEl)
      .setName("Strict Numeric Matching")
      .setDesc("Only apply numerical rules to cells that are pure numbers (recommended, prevents dates and mixed text from being colored)")
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.numericStrict)
              .onChange(async val => {
                this.plugin.settings.numericStrict = val;
                await this.plugin.saveSettings();
              }));



    // Performance settings
    new Setting(containerEl)
      .setName("Process all cells on open")
      .setDesc("Process ALL table cells when a file opens, not just visible ones. May improve color consistency in long tables but can impact performance.")
      .addToggle(toggle =>
        toggle.setValue(this.plugin.settings.processAllCellsOnOpen)
              .onChange(async val => {
                this.plugin.settings.processAllCellsOnOpen = val;
                await this.plugin.saveSettings();
              }));

    // Quick Actions Section
    containerEl.createEl('h2', { text: 'Quick Actions' });
  
  new Setting(containerEl)
    .setName("Show Row Coloring in Right-Click Menu")
    .setDesc("Display 'Color Row' and 'Remove Row' options in the context menu for table cells")
    .addToggle(toggle =>
      toggle.setValue(this.plugin.settings.showColorRowInMenu)
            .onChange(async val => {
              this.plugin.settings.showColorRowInMenu = val;
              await this.plugin.saveSettings();
            }));

  new Setting(containerEl)
    .setName("Show Column Coloring in Right-Click Menu")
    .setDesc("Display 'Color Column' and 'Remove Column' options in the context menu for table cells")
    .addToggle(toggle =>
      toggle.setValue(this.plugin.settings.showColorColumnInMenu)
            .onChange(async val => {
              this.plugin.settings.showColorColumnInMenu = val;
              await this.plugin.saveSettings();
            }));

  new Setting(containerEl)
    .setName("Show Undo/Redo in Right-Click Menu")
    .setDesc("Display 'Undo' and 'Redo' color changes in the context menu")
    .addToggle(toggle =>
      toggle.setValue(this.plugin.settings.showUndoRedoInMenu)
            .onChange(async val => {
              this.plugin.settings.showUndoRedoInMenu = val;
              await this.plugin.saveSettings();
            }));

  new Setting(containerEl)
    .setName("Show Refresh Icon in Status Bar")
    .setDesc("Toggle refresh table icon in the status bar")
    .addToggle(toggle =>
      toggle.setValue(this.plugin.settings.showStatusRefreshIcon)
            .onChange(async val => {
              this.plugin.settings.showStatusRefreshIcon = val;
              await this.plugin.saveSettings();
              if (val) {
                this.plugin.createStatusBarIcon();
              } else {
                this.plugin.removeStatusBarIcon();
              }
            }));

  new Setting(containerEl)
    .setName("Show Refresh Icon in Ribbon")
    .setDesc("Toggle refresh table icon in the left ribbon")
    .addToggle(toggle =>
      toggle.setValue(this.plugin.settings.showRibbonRefreshIcon)
            .onChange(async val => {
              this.plugin.settings.showRibbonRefreshIcon = val;
              await this.plugin.saveSettings();
              try {
                if (val && !this.plugin._ribbonRefreshIcon && typeof this.plugin.addRibbonIcon === 'function') {
                  const iconEl = this.plugin.addRibbonIcon('table', 'Refresh table colors', () => {
                    document.querySelectorAll('.markdown-preview-view table td, .markdown-preview-view table th').forEach(cell => { cell.style.backgroundColor = ''; cell.style.color = ''; });
                    this.plugin.applyColorsToActiveFile();
                    document.querySelectorAll('.cm-content table td, .cm-content table th').forEach(cell => { cell.style.backgroundColor = ''; cell.style.color = ''; });
                    if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') { setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); }
                  });
                  this.plugin._ribbonRefreshIcon = iconEl;
                } else if (!val && this.plugin._ribbonRefreshIcon) {
                  try { this.plugin._ribbonRefreshIcon.remove(); } catch (e) {}
                  this.plugin._ribbonRefreshIcon = null;
                }
              } catch (e) {}
            }));


    // Coloring Rules
    containerEl.createEl('h3', { text: 'Coloring Rules' });
    const crSection = containerEl.createDiv({ cls: 'cr-section' });

    // Search bar
    const searchContainer = crSection.createDiv({ cls: 'ctc-search-container pretty-flex' });
    const searchInput = searchContainer.createEl('input', { type: 'text', cls: 'ctc-search-input', placeholder: 'Search rules...' });
    searchContainer.createDiv({ cls: 'ctc-search-icon' });
    let searchTerm = '';
    searchInput.addEventListener('input', () => { searchTerm = searchInput.value.toLowerCase(); renderRules(); });

    const header = crSection.createDiv({ cls: 'cr-header-row cr-disabled-row pretty-flex' });
    try { header.classList.add('cr-sticky'); } catch (e) {}
    const mkHeaderSel = (text) => {
      const s = header.createEl('select', { cls: 'cr-select cr-header-select' });
      const o = s.createEl('option', { text, value: '' });
      o.selected = true; o.disabled = true;
      s.disabled = true;
      return s;
    };
    mkHeaderSel('TARGET');
    mkHeaderSel('WHEN');
    mkHeaderSel('MATCH');
    mkHeaderSel('VALUE');
    header.createEl('span', { text: '', cls: 'cr-color-placeholder' });
    header.createEl('span', { text: '', cls: 'cr-bg-placeholder' });
    header.createEl('span', { text: '', cls: 'cr-x-placeholder' });

    const rulesContainer = crSection.createDiv({ cls: 'cr-rules-container' });

    const TARGET_OPTIONS = [
      { label: 'COLOR CELL', value: 'cell' },
      { label: 'COLOR ROW', value: 'row' },
      { label: 'COLOR COLUMN', value: 'column' },
    ];
    const WHEN_OPTIONS = [
      { label: 'THE CELL', value: 'theCell' },
      { label: 'ANY CELL', value: 'anyCell' },
      { label: 'ALL CELL', value: 'allCell' },
      { label: 'NO CELL', value: 'noCell' },
      { label: 'FIRST ROW', value: 'firstRow' },
      { label: 'COLUMN HEADER', value: 'columnHeader' },
    ];
    const MATCH_OPTIONS = [
      { label: 'IS', value: 'is' },
      { label: 'IS NOT', value: 'isNot' },
      { label: 'IS REGEX', value: 'isRegex' },
      { label: 'CONTAINS', value: 'contains' },
      { label: 'DOES NOT CONTAIN', value: 'notContains' },
      { label: 'STARTS WITH', value: 'startsWith' },
      { label: 'ENDS WITH', value: 'endsWith' },
      { label: 'DOES NOT START WITH', value: 'notStartsWith' },
      { label: 'DOES NOT END WITH', value: 'notEndsWith' },
      { label: 'IS EMPTY', value: 'isEmpty' },
      { label: 'IS NOT EMPTY', value: 'isNotEmpty' },
      { label: 'IS EQUAL TO', value: 'eq' },
      { label: 'IS GREATER THAN', value: 'gt' },
      { label: 'IS LESS THAN', value: 'lt' },
      { label: 'IS GREATER THAN & EQUAL TO', value: 'ge' },
      { label: 'IS LESS THAN & EQUAL TO', value: 'le' },
    ];

    const labelFor = (opts, val) => (opts.find(o => o.value === val)?.label || '').toLowerCase();
    const isRegexRule = (rule) => rule.match === 'isRegex';
    const isNumericRule = (rule) => ['eq','gt','lt','ge','le'].includes(rule.match);

    const renderRules = () => {
      rulesContainer.empty();
      let rules = Array.isArray(this.plugin.settings.coloringRules) ? [...this.plugin.settings.coloringRules] : [];
      // Filter
      if (searchTerm) {
        rules = rules.filter(r => {
          const blob = [
            labelFor(TARGET_OPTIONS, r.target || ''),
            labelFor(WHEN_OPTIONS, r.when || ''),
            labelFor(MATCH_OPTIONS, r.match || ''),
            r.value != null ? String(r.value) : ''
          ].join(' ').toLowerCase();
          return blob.includes(searchTerm);
        });
      }
      // Sort
      const sortMode = this.plugin.settings.coloringSort || 'lastAdded';
      if (sortMode === 'az') {
        rules.sort((a,b) => (labelFor(MATCH_OPTIONS, a.match||'')).localeCompare(labelFor(MATCH_OPTIONS, b.match||'')));
      } else if (sortMode === 'regexFirst') {
        rules.sort((a,b) => Number(isRegexRule(b)) - Number(isRegexRule(a)));
      } else if (sortMode === 'numbersFirst') {
        rules.sort((a,b) => Number(isNumericRule(b)) - Number(isNumericRule(a)));
      } else if (sortMode === 'mode') {
        const order = { cell:0, row:1, column:2 };
        rules.sort((a,b) => (order[a.target||'cell'] - order[b.target||'cell']));
      }
      rules.forEach((rule, idx) => {
        const row = rulesContainer.createDiv({ cls: 'cr-rule-row pretty-flex' });
        const originalIdx = this.plugin.settings.coloringRules.indexOf(rule);
        row.dataset.idx = String(originalIdx);

        const targetSel = row.createEl('select', { cls: 'cr-select' });
        const tPh = targetSel.createEl('option', { text: 'TARGET', value: '' }); tPh.disabled = true; tPh.selected = !rule.target;
        TARGET_OPTIONS.forEach(opt => { const o = targetSel.createEl('option'); o.value = opt.value; o.text = opt.label; if (rule.target === opt.value) o.selected = true; });
        targetSel.addEventListener('change', async () => { rule.target = targetSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });

        const whenSel = row.createEl('select', { cls: 'cr-select' });
        const wPh = whenSel.createEl('option', { text: 'WHEN', value: '' }); wPh.disabled = true; wPh.selected = !rule.when;
        WHEN_OPTIONS.forEach(opt => { const o = whenSel.createEl('option'); o.value = opt.value; o.text = opt.label; if (rule.when === opt.value) o.selected = true; });
        whenSel.addEventListener('change', async () => { rule.when = whenSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });

        const matchSel = row.createEl('select', { cls: 'cr-select' });
        const mPh = matchSel.createEl('option', { text: 'MATCH', value: '' }); mPh.disabled = true; mPh.selected = !rule.match;
        MATCH_OPTIONS.forEach(opt => { const o = matchSel.createEl('option'); o.value = opt.value; o.text = opt.label; if (rule.match === opt.value) o.selected = true; });
        matchSel.addEventListener('change', async () => { rule.match = matchSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); renderRules(); });

        const numericMatches = new Set(['eq','gt','lt','ge','le']);
        const valueInput = row.createEl('input', { type: numericMatches.has(rule.match) ? 'number' : 'text', cls: 'cr-value-input' });
        valueInput.placeholder = 'VALUE';
        if (rule.value != null) valueInput.value = String(rule.value);
        valueInput.addEventListener('change', async () => {
          const v = valueInput.value;
          rule.value = numericMatches.has(rule.match) ? (v === '' ? null : Number(v)) : v;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
          if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        });

        const colorPicker = row.createEl('input', { type: 'color', cls: 'cr-color-picker' });
        if (rule.color) colorPicker.value = rule.color; else colorPicker.value = '#000000';
        colorPicker.title = 'Text Color';
        colorPicker.addEventListener('change', async () => { rule.color = colorPicker.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });

        const bgPicker = row.createEl('input', { type: 'color', cls: 'cr-bg-picker' });
        if (rule.bg) bgPicker.value = rule.bg; else bgPicker.value = '#000000';
        bgPicker.title = 'Background Color';
        bgPicker.addEventListener('change', async () => { rule.bg = bgPicker.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });

        const delBtn = row.createEl('button', { cls: 'mod-ghost cr-del-btn' });
        if (typeof window.setIcon === 'function') {
          window.setIcon(delBtn, 'x');
        } else if (typeof setIcon === 'function') {
          setIcon(delBtn, 'x');
        } else {
          delBtn.textContent = '×';
        }
        delBtn.addEventListener('click', async () => {
          const oi = this.plugin.settings.coloringRules.indexOf(rule);
          if (oi >= 0) {
            this.plugin.settings.coloringRules.splice(oi, 1);
            await this.plugin.saveSettings();
            this.plugin.applyColorsToActiveFile();
            renderRules();
          }
        });

        row.addEventListener('contextmenu', (evt) => {
          const menu = new Menu();
          menu.addItem(item =>
            item.setTitle('Duplicate rule')
                .setIcon('copy')
                .onClick(async () => {
                  const oi = this.plugin.settings.coloringRules.indexOf(rule);
                  if (oi >= 0) {
                    const clone = JSON.parse(JSON.stringify(rule));
                    this.plugin.settings.coloringRules.splice(oi + 1, 0, clone);
                    await this.plugin.saveSettings();
                    this.plugin.applyColorsToActiveFile();
                    renderRules();
                  }
                })
          );
          const canMoveUp = originalIdx > 0;
          const canMoveDown = originalIdx >= 0 && originalIdx < (this.plugin.settings.coloringRules.length - 1);
          menu.addItem(item =>
            item.setTitle('Move rule up')
                .setIcon('arrow-up')
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
                })
          );
          menu.addItem(item =>
            item.setTitle('Move rule down')
                .setIcon('arrow-down')
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
                })
          );
          menu.addSeparator();
          menu.addItem(item =>
            item.setTitle('Reset text color')
                .setIcon('text')
                .onClick(async () => {
                  rule.color = null;
                  await this.plugin.saveSettings();
                  this.plugin.applyColorsToActiveFile();
                  renderRules();
                })
          );
          menu.addItem(item =>
            item.setTitle('Reset background color')
                .setIcon('droplet')
                .onClick(async () => {
                  rule.bg = null;
                  await this.plugin.saveSettings();
                  this.plugin.applyColorsToActiveFile();
                  renderRules();
                })
          );
          try {
            if (menu.containerEl && menu.containerEl.classList) menu.containerEl.classList.add('mod-shadow');
            if (menu.menuEl && menu.menuEl.classList) menu.menuEl.classList.add('mod-shadow');
          } catch (e) { }
          menu.showAtMouseEvent(evt);
          evt.preventDefault();
        });
      });
    };

    const addRow = crSection.createDiv({ cls: 'cr-add-row' });
    const sortSel = addRow.createEl('select', { cls: 'cr-select' });
    const sortOptions = [
      { label: 'Sort: Last Added', value: 'lastAdded' },
      { label: 'Sort: A-Z', value: 'az' },
      { label: 'Sort: Regex first', value: 'regexFirst' },
      { label: 'Sort: Numbers first', value: 'numbersFirst' },
      { label: 'Sort: Mode', value: 'mode' }
    ];
    sortOptions.forEach(opt => { const o = sortSel.createEl('option'); o.text = opt.label; o.value = opt.value; if (opt.value === (this.plugin.settings.coloringSort||'lastAdded')) o.selected = true; });
    sortSel.addEventListener('change', async () => { this.plugin.settings.coloringSort = sortSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); renderRules(); });

    const addBtn = addRow.createEl('button', { cls: 'mod-cta cr-add-flex' });
    addBtn.textContent = '+ Add Rule';
    addBtn.addEventListener('click', async () => {
      if (!Array.isArray(this.plugin.settings.coloringRules)) this.plugin.settings.coloringRules = [];
      this.plugin.settings.coloringRules.push({ target: '', when: '', match: '', value: null, color: null, bg: null });
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      renderRules();
    });

    renderRules();

    const advHeading = containerEl.createEl('h3', { text: 'Advanced Rules', cls: 'cr-adv-heading' });
    
    // Search bar for advanced rules
    const advSearchContainer = containerEl.createDiv({ cls: 'ctc-search-container pretty-flex' });
    const advSearchInput = advSearchContainer.createEl('input', { type: 'text', cls: 'ctc-search-input', placeholder: 'Search advanced rules...' });
    advSearchContainer.createDiv({ cls: 'ctc-search-icon' });
    let advSearchTerm = '';
    advSearchInput.addEventListener('input', () => { advSearchTerm = advSearchInput.value.toLowerCase(); renderAdv(); });
    
    const advList = containerEl.createDiv({ cls: 'cr-adv-list' });
    const renderAdv = () => {
      advList.empty();
      let advRules = Array.isArray(this.plugin.settings.advancedRules) ? this.plugin.settings.advancedRules : [];
      
      const cap = (s) => (typeof s === 'string' && s.length) ? s.charAt(0).toUpperCase() + s.slice(1) : s;
      const verb = (m, v) => {
        if (m === 'contains') return `contains ${v}`;
        if (m === 'notContains') return `does not contain ${v}`;
        if (m === 'is') return `is ${v}`;
        if (m === 'isNot') return `is not ${v}`;
        if (m === 'startsWith') return `starts with ${v}`;
        if (m === 'endsWith') return `ends with ${v}`;
        if (m === 'notStartsWith') return `does not start with ${v}`;
        if (m === 'notEndsWith') return `does not end with ${v}`;
        if (m === 'isEmpty') return `is empty`;
        if (m === 'isNotEmpty') return `is not empty`;
        if (m === 'isRegex') return `matches ${v}`;
        if (m === 'eq') return `is equal to ${v}`;
        if (m === 'gt') return `is greater than ${v}`;
        if (m === 'lt') return `is less than ${v}`;
        if (m === 'ge') return `is greater than or equal to ${v}`;
        if (m === 'le') return `is less than or equal to ${v}`;
        return `${m} ${v}`;
      };
      const summaryForAdvRule = (ar) => {
        // Use custom name if provided
        if (ar.name && ar.name.trim()) {
          return ar.name;
        }
        // Auto-generate name
        const targetPhrase = ar.target === 'row' ? 'Color rows' : (ar.target === 'column' ? 'Color columns' : 'Color cells');
        const conds = Array.isArray(ar.conditions) ? ar.conditions : [];
        if (!conds.length) return '(empty)';
        
        // Build a readable string showing conditions
        const parts = conds.map((c, idx) => {
          const value = String(c.value || '').trim();
          // For header conditions, just show the value
          if (c.when === 'columnHeader') {
            return value;
          }
          // For other conditions, show match + value unless it's a simple case
          if (c.match === 'is' && value) {
            return value;
          }
          // Otherwise show verb(match, value)
          return verb(c.match, value);
        }).filter(Boolean);
        
        const condString = parts.length ? parts.join(', ') : 'conditions';
        
        return `${targetPhrase} when ${condString}`;
      };
      
      // Filter by search term - search all condition values
      if (advSearchTerm) {
        advRules = advRules.filter((ar) => {
          const summary = summaryForAdvRule(ar).toLowerCase();
          const targetText = (ar.target || 'cell').toLowerCase();
          const colorHex = (ar.color || ar.bg || '').toLowerCase();
          const allCondValues = Array.isArray(ar.conditions) ? ar.conditions.map(c => String(c.value || '').toLowerCase()).join(' ') : '';
          return summary.includes(advSearchTerm) || targetText.includes(advSearchTerm) || colorHex.includes(advSearchTerm) || allCondValues.includes(advSearchTerm);
        });
      }
      advRules.forEach((ar, idx) => {
        const originalIdx = this.plugin.settings.advancedRules.indexOf(ar);
        const row = advList.createDiv({ cls: 'cr-adv-row pretty-flex' });
        const drag = row.createEl('span', { cls: 'drag-handle' });
        try { require('obsidian').setIcon(drag, 'menu'); } catch (e) { try { require('obsidian').setIcon(drag, 'grip-vertical'); } catch (e2) { drag.textContent = '≡'; } }
        row.dataset.idx = String(originalIdx);
        drag.setAttribute('draggable','true');
        drag.addEventListener('dragstart', (e) => { e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', String(originalIdx)); row.classList.add('dragging'); });
        drag.addEventListener('dragend', () => { row.classList.remove('dragging'); advList.querySelectorAll('.rule-over').forEach(el => el.classList.remove('rule-over')); });
        row.addEventListener('dragover', (e) => { e.preventDefault(); row.classList.add('rule-over'); });
        row.addEventListener('dragleave', () => { row.classList.remove('rule-over'); });
        row.addEventListener('drop', async (e) => { e.preventDefault(); const from = Number(e.dataTransfer.getData('text/plain')); const to = Number(row.dataset.idx); if (isNaN(from)||isNaN(to)||from===to) return; const list = this.plugin.settings.advancedRules; const [m] = list.splice(from,1); list.splice(to,0,m); await this.plugin.saveSettings(); renderAdv(); });

        const label = row.createEl('span', { cls: 'cr-adv-label' });
        label.textContent = summaryForAdvRule(ar);

        const copyBtn = row.createEl('button', { cls: 'mod-ghost cr-adv-copy' });
        copyBtn.setAttribute('aria-label', 'Duplicate rule');
        copyBtn.setAttribute('title', 'Duplicate rule');
        try { require('obsidian').setIcon(copyBtn, 'copy'); } catch (e) {}
        copyBtn.addEventListener('click', async () => { const ruleCopy = JSON.parse(JSON.stringify(ar)); this.plugin.settings.advancedRules.splice(originalIdx + 1, 0, ruleCopy); await this.plugin.saveSettings(); document.dispatchEvent(new Event('ctc-adv-rules-changed')); });

        const settingsBtn = row.createEl('button', { cls: 'mod-ghost cr-adv-settings' });
        try { require('obsidian').setIcon(settingsBtn, 'settings'); } catch (e) {}
        settingsBtn.addEventListener('click', () => { new AdvancedRuleModal(this.app, this.plugin, originalIdx).open(); });
      });
    };
    document.addEventListener('ctc-adv-rules-changed', renderAdv);
    const advActions = containerEl.createDiv({ cls: 'cr-adv-actions' });
    const addAdvBtn = advActions.createEl('button', { cls: 'mod-cta' });
    addAdvBtn.textContent = 'Add advanced rule';
    addAdvBtn.addEventListener('click', async () => { if (!Array.isArray(this.plugin.settings.advancedRules)) this.plugin.settings.advancedRules = []; this.plugin.settings.advancedRules.push({ logic:'any', conditions:[], target:'cell', color:null, bg:null }); await this.plugin.saveSettings(); renderAdv(); });
    renderAdv();
    
    // Export/Import Section
    containerEl.createEl('h2', { text: 'Data Management' });
    
    const exportImportRow = containerEl.createDiv({ cls: 'cr-export-row pretty-flex' });
    const exportBtn = exportImportRow.createEl('button', { text: 'Export Settings' });
    exportBtn.addEventListener('click', async () => {
      try {
        const data = {
          settings: this.plugin.settings,
          cellData: this.plugin.cellData,
          exportDate: new Date().toISOString()
        };
        const dataStr = JSON.stringify(data, null, 2);
        const blob = new Blob([dataStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `color-table-cells-backup-${new Date().getTime()}.json`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        setTimeout(() => new Notice('Settings exported successfully!'), 500);
      } catch (e) {
        new Notice('Failed to export settings: ' + e.message);
      }
    });
    
    const importBtn = exportImportRow.createEl('button', { text: 'Import Settings' });
    importBtn.addEventListener('click', () => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.json';
      input.addEventListener('change', async (e) => {
        try {
          const file = e.target.files?.[0];
          if (!file) return;
          const text = await file.text();
          const data = JSON.parse(text);
          
          if (data.settings) {
            this.plugin.settings = Object.assign(this.plugin.settings, data.settings);
          }
          if (data.cellData) {
            this.plugin.cellData = data.cellData;
          }
          
          await this.plugin.saveSettings();
          this.display();
          setTimeout(() => new Notice('Settings imported successfully!'), 500);
        } catch (e) {
          new Notice('Failed to import settings: ' + e.message);
        }
      });
      input.click();
    });

    // Delete manual coloring entries
    containerEl.createEl('h3', { text: 'Danger Zone', cls: 'cr-danger-heading' });
    const dangerZoneRow = containerEl.createDiv({ cls: 'cr-delete-container' });
    const deleteManualRow = dangerZoneRow.createDiv({ cls: 'cr-delete-row' });
    const deleteManualBtn = deleteManualRow.createEl('button', { text: 'Delete All Manual Colors', cls: 'mod-warning' });
    deleteManualBtn.addEventListener('click', () => {
      const modal = new Modal(this.app);
      modal.contentEl.createEl('h2', { text: 'Delete All Manual Colors?' });
      modal.contentEl.createEl('p', { text: 'This will remove all manually colored cells (non-rule colors). This action cannot be undone.' });
      const btnRow = modal.contentEl.createDiv({ cls: 'modal-delete-buttons' });
      const cancelBtn = btnRow.createEl('button', { text: 'Cancel', cls: 'mod-ghost' });
      const confirmBtn = btnRow.createEl('button', { text: 'Delete All', cls: 'mod-warning' });
      cancelBtn.addEventListener('click', () => modal.close());
      confirmBtn.addEventListener('click', async () => {
        this.plugin.cellData = {};
        await this.plugin.saveData({ settings: this.plugin.settings, cellData: this.plugin.cellData });
        this.plugin.applyColorsToActiveFile();
        new Notice('All manual colors deleted');
        modal.close();
      });
      modal.open();
    });

    // Delete coloring rules
    const deleteRulesRow = dangerZoneRow.createDiv({ cls: 'cr-delete-row' });
    const deleteRulesBtn = deleteRulesRow.createEl('button', { text: 'Delete All Coloring Rules', cls: 'mod-warning' });
    deleteRulesBtn.addEventListener('click', () => {
      const modal = new Modal(this.app);
      modal.contentEl.createEl('h2', { text: 'Delete All Coloring Rules?' });
      modal.contentEl.createEl('p', { text: `This will remove all ${Array.isArray(this.plugin.settings.coloringRules) ? this.plugin.settings.coloringRules.length : 0} coloring rules. This action cannot be undone.` });
      const btnRow = modal.contentEl.createDiv({ cls: 'modal-delete-buttons' });
      const cancelBtn = btnRow.createEl('button', { text: 'Cancel', cls: 'mod-ghost' });
      const confirmBtn = btnRow.createEl('button', { text: 'Delete All', cls: 'mod-warning' });
      cancelBtn.addEventListener('click', () => modal.close());
      confirmBtn.addEventListener('click', async () => {
        this.plugin.settings.coloringRules = [];
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        new Notice('All coloring rules deleted');
        modal.close();
        this.display();
      });
      modal.open();
    });

    // Delete advanced rules
    const deleteAdvRulesRow = dangerZoneRow.createDiv({ cls: 'cr-delete-row' });
    const deleteAdvRulesBtn = deleteAdvRulesRow.createEl('button', { text: 'Delete All Advanced Rules', cls: 'mod-warning' });
    deleteAdvRulesBtn.addEventListener('click', () => {
      const modal = new Modal(this.app);
      modal.contentEl.createEl('h2', { text: 'Delete All Advanced Rules?' });
      modal.contentEl.createEl('p', { text: `This will remove all ${Array.isArray(this.plugin.settings.advancedRules) ? this.plugin.settings.advancedRules.length : 0} advanced rules. This action cannot be undone.` });
      const btnRow = modal.contentEl.createDiv({ cls: 'modal-delete-buttons' });
      const cancelBtn = btnRow.createEl('button', { text: 'Cancel', cls: 'mod-ghost' });
      const confirmBtn = btnRow.createEl('button', { text: 'Delete All', cls: 'mod-warning' });
      cancelBtn.addEventListener('click', () => modal.close());
      confirmBtn.addEventListener('click', async () => {
        this.plugin.settings.advancedRules = [];
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        new Notice('All advanced rules deleted');
        modal.close();
        this.display();
      });
      modal.open();
    });
  }

  hide() {
    // Force refresh table colors when closing settings tab
    // Clear all rule-based colors first to ensure changes take effect
    document.querySelectorAll('.markdown-preview-view table td, .markdown-preview-view table th').forEach(cell => {
      if (!cell.hasAttribute('data-ctc-manual')) {
        cell.style.backgroundColor = '';
        cell.style.color = '';
      }
    });
    document.querySelectorAll('.cm-content table td, .cm-content table th').forEach(cell => {
      if (!cell.hasAttribute('data-ctc-manual')) {
        cell.style.backgroundColor = '';
        cell.style.color = '';
      }
    });
    
    // Now reapply colors with rule changes
    if (typeof this.plugin.applyColorsToActiveFile === 'function') {
      setTimeout(() => this.plugin.applyColorsToActiveFile(), 50);
    }
    if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') {
      setTimeout(() => this.plugin.applyColorsToAllEditors(), 100);
    }
  }
}

class AdvancedRuleModal extends Modal {
  constructor(app, plugin, index) { super(app); this.plugin = plugin; this.index = index; }
  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.addClass('cr-adv-modal');
    const rule = this.plugin.settings.advancedRules?.[this.index] || { logic:'any', conditions:[], target:'cell', color:null, bg:null };
    contentEl.createEl('h3', { text: 'Advanced Rules Builder', cls: 'cr-adv-modal-heading' });
    const logicRow = contentEl.createDiv({ cls: 'cr-adv-logic' });
    const logicLabel = logicRow.createEl('span', { cls: 'cr-adv-logic-label' });
    logicLabel.textContent = 'Conditions Match';
    const logicButtons = logicRow.createDiv({ cls: 'cr-adv-logic-buttons' });
    const mkBtn = (txt, val) => {
      const b = logicButtons.createEl('button', { cls: 'cr-adv-logic-btn mod-ghost' });
      b.textContent = txt;
      const applyActive = () => {
        ['any','all','none'].forEach(k => {
          const q = Array.from(logicButtons.querySelectorAll('.cr-adv-logic-btn')).find(el => el.textContent === k.toUpperCase());
          if (q) { q.classList.remove('mod-cta'); q.classList.add('mod-ghost'); }
        });
        b.classList.add('mod-cta');
        b.classList.remove('mod-ghost');
      };
      if (rule.logic === val) applyActive();
      b.addEventListener('click', async () => { rule.logic = val; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); applyActive(); });
      return b;
    };
    mkBtn('ANY','any'); mkBtn('ALL','all'); mkBtn('NONE','none');
    contentEl.createEl('h4', { text: 'Conditions', cls: 'cr-adv-h4' });
    const condsWrap = contentEl.createDiv({ cls: 'cr-adv-conds-wrap' });
    const TARGET_OPTIONS = [
      { label: 'COLOR CELL', value: 'cell' },
      { label: 'COLOR ROW', value: 'row' },
      { label: 'COLOR COLUMN', value: 'column' },
    ];
    const WHEN_OPTIONS = [
      { label: 'ANY CELL', value: 'anyCell' },
      { label: 'ALL CELL', value: 'allCell' },
      { label: 'NO CELL', value: 'noCell' },
      { label: 'FIRST ROW', value: 'firstRow' },
      { label: 'COLUMN HEADER', value: 'columnHeader' },
    ];
    const MATCH_OPTIONS = [
      { label: 'IS', value: 'is' },
      { label: 'IS NOT', value: 'isNot' },
      { label: 'IS REGEX', value: 'isRegex' },
      { label: 'CONTAINS', value: 'contains' },
      { label: 'DOES NOT CONTAIN', value: 'notContains' },
      { label: 'STARTS WITH', value: 'startsWith' },
      { label: 'ENDS WITH', value: 'endsWith' },
      { label: 'DOES NOT START WITH', value: 'notStartsWith' },
      { label: 'DOES NOT END WITH', value: 'notEndsWith' },
      { label: 'IS EMPTY', value: 'isEmpty' },
      { label: 'IS NOT EMPTY', value: 'isNotEmpty' },
      { label: 'IS EQUAL TO', value: 'eq' },
      { label: 'IS GREATER THAN', value: 'gt' },
      { label: 'IS LESS THAN', value: 'lt' },
      { label: 'IS GREATER THAN & EQUAL TO', value: 'ge' },
      { label: 'IS LESS THAN & EQUAL TO', value: 'le' },
    ];
    const renderConds = () => {
      condsWrap.empty();
      (rule.conditions||[]).forEach((cond, ci) => {
        const row = condsWrap.createDiv({ cls: 'cr-adv-cond-row pretty-flex' });
        const whenSel = row.createEl('select', { cls: 'cr-select cr-adv-cond-when' });
        WHEN_OPTIONS.forEach(opt => { const o = whenSel.createEl('option'); o.value = opt.value; o.text = opt.label; if (cond.when===opt.value) o.selected=true; });
        whenSel.addEventListener('change', async () => { cond.when = whenSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });
        const matchSel = row.createEl('select', { cls: 'cr-select cr-adv-cond-match' });
        MATCH_OPTIONS.forEach(opt => { const o = matchSel.createEl('option'); o.value = opt.value; o.text = opt.label; if (cond.match===opt.value) o.selected=true; });
        matchSel.addEventListener('change', async () => { cond.match = matchSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); const isNum = ['eq','gt','lt','ge','le'].includes(cond.match); valInput.type = isNum ? 'number' : 'text'; });
        const isNum = ['eq','gt','lt','ge','le'].includes(cond.match);
        const valInput = row.createEl('input', { type: isNum ? 'number' : 'text', cls: 'cr-value-input cr-adv-cond-value' });
        if (cond.value != null) valInput.value = String(cond.value);
        valInput.addEventListener('change', async () => { const v = valInput.value; cond.value = isNum ? (v === '' ? null : Number(v)) : v; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });

        // Add per-condition delete button (X)
        const delBtn = row.createEl('button', { cls: 'mod-ghost cr-adv-cond-del' });
        try { require('obsidian').setIcon(delBtn, 'x'); } catch (e) { delBtn.textContent = '×'; }
        delBtn.addEventListener('click', async () => {
          if (Array.isArray(rule.conditions)) {
            rule.conditions.splice(ci, 1);
            await this.plugin.saveSettings();
            this.plugin.applyColorsToActiveFile();
            if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
            renderConds();
          }
        });
      });
    };
    renderConds();
    const addCondRow = contentEl.createDiv({ cls: 'cr-adv-add-row' });
    const addCondBtn = addCondRow.createEl('button', { cls: 'mod-ghost cr-adv-add-btn' });
    addCondBtn.textContent = '+ Add Condition';
    addCondBtn.addEventListener('click', async () => {
      if (!Array.isArray(rule.conditions)) rule.conditions = [];
      rule.conditions.push({ when:'anyCell', match:'contains', value:'' });
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
      renderConds();
    });
    
    contentEl.createEl('h4', { text: 'Then Color', cls: 'cr-adv-h4' });
    
    const colorRow = contentEl.createDiv({ cls: 'cr-adv-color-row pretty-flex' });
    const targetSel = colorRow.createEl('select', { cls: 'cr-select cr-adv-target' });
    TARGET_OPTIONS.forEach(opt => { const o = targetSel.createEl('option'); o.value = opt.value; o.text = opt.label; if (rule.target === opt.value) o.selected = true; });
    targetSel.addEventListener('change', async () => { rule.target = targetSel.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });
    
    // Text color picker with reset button
    const textColorContainer = colorRow.createDiv({ style: 'display: flex; gap: 0.3em; align-items: center; margin-top: 3px;' });
    const colorPicker = textColorContainer.createEl('input', { type:'color', cls:'cr-color-picker cr-adv-text-color cr-adv-color-input' });
    colorPicker.value = rule.color || '#000000';
    colorPicker.title = 'Text Color';
    colorPicker.addEventListener('change', async () => { rule.color = colorPicker.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });
    const colorResetBtn = textColorContainer.createEl('button', { cls: 'mod-ghost cr-adv-color-reset', style: 'padding: 4px 8px; font-size: 12px; height: auto; margin: 0 6px;' });
    colorResetBtn.textContent = 'Reset';
    colorResetBtn.title = 'Reset text color to none';
    colorResetBtn.addEventListener('click', async () => { rule.color = null; colorPicker.value = '#000000'; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });
    
    // Background color picker with reset button
    const bgColorContainer = colorRow.createDiv({ style: 'display: flex; gap: 0.3em; align-items: center; margin-top: 3px;' });
    const bgPicker = bgColorContainer.createEl('input', { type:'color', cls:'cr-bg-picker cr-adv-bg-color cr-adv-bg-input' });
    bgPicker.value = rule.bg || '#000000';
    bgPicker.title = 'Background Color';
    bgPicker.addEventListener('change', async () => { rule.bg = bgPicker.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });
    const bgResetBtn = bgColorContainer.createEl('button', { cls: 'mod-ghost cr-adv-bg-reset', style: 'padding: 4px 8px; font-size: 12px; height: auto; margin: 0 6px;' });
    bgResetBtn.textContent = 'Reset';
    bgResetBtn.title = 'Reset background color to none';
    bgResetBtn.addEventListener('click', async () => { rule.bg = null; bgPicker.value = '#000000'; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); });
    
    // Custom Name Section - AFTER color swatches
    contentEl.createEl('h4', { text: 'Rule Name (Optional)', cls: 'cr-adv-h4' });
    const nameRow = contentEl.createDiv({ cls: 'cr-adv-name-row' });
    const nameInput = nameRow.createEl('input', { type: 'text', cls: 'cr-adv-name-input' });
    nameInput.style.width = '100%';
    nameInput.placeholder = 'Leave empty to use automatic naming';
    if (rule.name) nameInput.value = rule.name;
    nameInput.addEventListener('change', async () => { rule.name = nameInput.value; await this.plugin.saveSettings(); this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); document.dispatchEvent(new CustomEvent('ctc-adv-rules-changed')); });
    const actionsRow = contentEl.createDiv({ cls: 'cr-adv-actions-row' });
    const deleteBtn = actionsRow.createEl('button', { cls: 'cr-adv-delete' });
    deleteBtn.textContent = 'Delete Rule';
    deleteBtn.addEventListener('click', async () => {
      if (Array.isArray(this.plugin.settings.advancedRules)) {
        this.plugin.settings.advancedRules.splice(this.index, 1);
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10);
        document.dispatchEvent(new CustomEvent('ctc-adv-rules-changed'));
      }
      this.close();
    });
    const saveBtn = actionsRow.createEl('button', { cls: 'mod-cta cr-adv-save' });
    saveBtn.textContent = 'Save Rule';
    saveBtn.addEventListener('click', () => { this.plugin.applyColorsToActiveFile(); if (this.plugin.settings.livePreviewColoring && typeof this.plugin.applyColorsToAllEditors === 'function') setTimeout(() => this.plugin.applyColorsToAllEditors(), 10); document.dispatchEvent(new CustomEvent('ctc-adv-rules-changed')); this.close(); });
  }
  onClose() { this.contentEl.empty(); }
}

