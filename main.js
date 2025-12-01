const { Plugin, PluginSettingTab, Setting, Menu, ButtonComponent, Modal } = require('obsidian');

// Debug configuration
const debugMode = false;

// Debug logging helper - removed to follow guidelines
const debugLog = () => {};

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
      hotkeys: [
        {
          modifiers: ["Ctrl"],
          key: "z"
        }
      ],
      callback: () => this.undo()
    });

    this.addCommand({
      id: 'redo-color-change',
      name: 'Redo Last Color Change',
      hotkeys: [
        {
          modifiers: ["Ctrl"],
          key: "y"
        },
        {
          modifiers: ["Ctrl", "Shift"],
          key: "z"
        }
      ],
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
      id: 'open-regex-tester',
      name: 'Open Regex Pattern Tester',
      callback: () => {
        new RegexTesterModal(this.app, this).open();
      }
    });

    // --- Live Preview Table Coloring logic ---
    this.applyColorsToAllEditors = () => {
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
      document.querySelectorAll('.cm-content').forEach(editorEl => {
        if (editorEl.querySelector('table')) {
          const file = this.app.workspace.getActiveFile();
          if (file) {
            this.applyColorsToContainer(editorEl, file.path);
          }
        }
      });
    };

    const setupLivePreviewColoring = () => {
      // Initial application
      setTimeout(this.applyColorsToAllEditors, 200);
      // Observe DOM changes in editors
      if (!this._livePreviewObserver) {
        this._livePreviewObserver = new MutationObserver(() => {
          this.applyColorsToAllEditors();
        });
        document.querySelectorAll('.cm-content').forEach(editorEl => {
          this._livePreviewObserver.observe(editorEl, { childList: true, subtree: true });
        });
      }
      // Re-apply on file open/layout change
      this.registerEvent(this.app.workspace.on('file-open', this.applyColorsToAllEditors));
      this.registerEvent(this.app.workspace.on('layout-change', this.applyColorsToAllEditors));
      // Re-apply on cell focus/blur/input (to persist colors after editing)
      document.addEventListener('focusin', (e) => {
        if (e.target && e.target.closest && e.target.closest('.cm-content table')) {
          this.applyColorsToAllEditors();
        }
      });
      document.addEventListener('input', (e) => {
        if (e.target && e.target.closest && e.target.closest('.cm-content table')) {
          setTimeout(this.applyColorsToAllEditors, 30);
        }
      });
    };

    await this.loadSettings();
    // Setup Live Preview coloring if enabled (after settings loaded)
    setupLivePreviewColoring();
      // Plugin loaded    const rawSaved = await this.loadData() || {};

    this._appliedContainers = new WeakMap();

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

    this.registerMarkdownPostProcessor((el, ctx) => {
      if (!el.closest('.markdown-preview-view')) return;
      const fileId = ctx.sourcePath;
      this.applyColorsToContainer(el, fileId);

      try {
        let observer = null;
        let debounceId = null;
        let scrollListener = null;
        let scrollDebounceId = null;
        
        const safeDisconnect = () => {
          try { if (observer) { observer.disconnect(); observer = null; } } catch (e) { }
          try { if (debounceId) { clearTimeout(debounceId); debounceId = null; } } catch (e) { }
          try { if (scrollListener && el.parentElement) { el.parentElement.removeEventListener('scroll', scrollListener); } } catch (e) { }
          try { if (scrollDebounceId) { clearTimeout(scrollDebounceId); scrollDebounceId = null; } } catch (e) { }
        };

        const checkAndApply = () => {
          try {
            if (el.querySelectorAll && el.querySelectorAll('table').length > 0) {
              this.applyColorsToContainer(el, fileId);
              safeDisconnect();
            }
          } catch (e) { /* ignore */ }
        };

        observer = new MutationObserver((mutations) => {
          if (debounceId) clearTimeout(debounceId);
          debounceId = setTimeout(() => {
            checkAndApply();
          }, 80);
        });

        observer.observe(el, { childList: true, subtree: true });

        checkAndApply();

        // Add scroll listener to reapply colors on scroll
        try {
          const scrollContainer = el.parentElement || el;
          scrollListener = () => {
            if (scrollDebounceId) clearTimeout(scrollDebounceId);
            scrollDebounceId = setTimeout(() => {
              this.applyColorsToContainer(el, fileId);
            }, 100);
          };
          scrollContainer.addEventListener('scroll', scrollListener, { passive: true });
        } catch (e) { /* ignore if scroll listener fails */ }

        // Keep observer alive for longer to catch dynamic content
        setTimeout(() => { safeDisconnect(); }, 10000);
      } catch (e) { /* ignore if MutationObserver unsupported */ }
    });

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
        menu.addItem(item =>
          item.setTitle("Color Multiple Cells by Rule")
              .setIcon('grid')
              .onClick(() => {
                this.app.setting.open();
                setTimeout(() => {
                  if (this.app.setting && typeof this.app.setting.openTabById === 'function') {
                    this.app.setting.openTabById('color-table-cell');
                  }
                }, 250);
              })
        );
        menu.addSeparator();
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
        menu.addItem(item =>
          item.setTitle("Reset Cell Coloring")
              .setIcon('trash-2')
              .onClick(async () => this.resetCell(cell, tableEl))
        );
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

  onunload() {
    // Remove our settings tab when the plugin is 
    // disabled/unloaded to avoid duplicate entries
    try {
      if (this._settingsTab && typeof this.removeSettingTab === 'function') {
        this.removeSettingTab(this._settingsTab);
      }
    } catch (e) { }
    try {
      // no global preview observer to disconnect
    } catch (e) { }
  }

  async loadSettings() {
    const data = await this.loadData();
    this.settings = Object.assign({
      defaultBgColor: "#331717",
      defaultTextColor: "#ffb2b2",

      enableContextMenu: true,
      rules: [],
      numericRules: [], // { op: 'lt'|'gt'|'eq'|'le'|'ge', value: number, color: string, bg: string }
      numericStrict: true, // Only apply numeric rules to pure numbers
  livePreviewColoring: false, // Table coloring in Live Preview mode
    }, data?.settings || {});
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

  async pickColor(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = type === 'bg' ? this.settings.defaultBgColor : this.settings.defaultTextColor;
    // Use the cell as anchor for menu position
    const menu = new CustomColorPickerMenu(this, async (pickedColor) => {
      // Save color on close
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      const allTables = Array.from(tableEl.closest('.markdown-preview-section, .markdown-preview-view')?.querySelectorAll('table') || tableEl.ownerDocument.querySelectorAll('table'));
      const tableIndex = allTables.indexOf(tableEl);
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
      
      await this.saveDataColors();
    }, initialColor, cell);
    menu._cell = cell;
    menu._type = type;
    menu.open();
  }

  async pickColorForRow(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = type === 'bg' ? this.settings.defaultBgColor : this.settings.defaultTextColor;
    const row = cell.closest('tr');
    const rowCells = Array.from(row.querySelectorAll('td, th'));
    
    const menu = new CustomColorPickerMenu(this, async (pickedColor) => {
      // Save color on close for entire row
      const fileId = this.app.workspace.getActiveFile()?.path;
      if (!fileId) return;
      const allTables = Array.from(tableEl.closest('.markdown-preview-section, .markdown-preview-view')?.querySelectorAll('table') || tableEl.ownerDocument.querySelectorAll('table'));
      const tableIndex = allTables.indexOf(tableEl);
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
      
      await this.saveDataColors();
    }, initialColor, cell);
    
    menu._cells = rowCells;
    menu._type = type;
    menu.open();
  }

  async pickColorForColumn(cell, tableEl, type) {
    const CustomColorPickerMenu = this.createColorPickerMenu();
    const initialColor = type === 'bg' ? this.settings.defaultBgColor : this.settings.defaultTextColor;
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
      const allTables = Array.from(tableEl.closest('.markdown-preview-section, .markdown-preview-view')?.querySelectorAll('table') || tableEl.ownerDocument.querySelectorAll('table'));
      const tableIndex = allTables.indexOf(tableEl);
      
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
      
      await this.saveDataColors();
    }, initialColor, cell);
    
    menu._cells = columnCells;
    menu._type = type;
    menu.open();
  }

  async resetCell(cell, tableEl) {
    cell.style.backgroundColor = "";
    cell.style.color = "";

    const fileId = this.app.workspace.getActiveFile()?.path;
    if (!fileId) return;

  const allTables = Array.from(tableEl.closest('.markdown-preview-section, .markdown-preview-view')?.querySelectorAll('table') || tableEl.ownerDocument.querySelectorAll('table'));
  const tableIndex = allTables.indexOf(tableEl);
  const rowIndex = Array.from(tableEl.querySelectorAll('tr')).indexOf(cell.closest('tr'));
  const colIndex = Array.from(cell.closest('tr').querySelectorAll('td, th')).indexOf(cell);

    const noteData = this.cellData[fileId];
    const tableKey = `table_${tableIndex}`;
    if (noteData?.[tableKey]?.[`row_${rowIndex}`]) {
      delete noteData[tableKey][`row_${rowIndex}`][`col_${colIndex}`];
      await this.saveDataColors();
    }
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
  }

  async saveDataColors() {
  await this.saveData({ settings: this.settings, cellData: this.cellData });
  this.applyColorsToActiveFile();
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
    const noteData = this.cellData[filePath] || {};
    container.querySelectorAll("table").forEach((tableEl, tableIndex) => {
      const tableKey = `table_${tableIndex}`;
      const tableColors = noteData[tableKey] || {};
      tableEl.querySelectorAll("tr").forEach((tr, rowIndex) => {
        const rowKey = `row_${rowIndex}`;
        tr.querySelectorAll("td, th").forEach((cell, colIndex) => {
          const colKey = `col_${colIndex}`;
          const colorData = tableColors[rowKey]?.[colKey];
          
          // Get clean cell text without editor markup
          const cellText = this.getCellText(cell);
          
          // Preserve existing colors if colors are already applied
          const hasExistingStyles = cell.style.backgroundColor || cell.style.color;
          
          // Clear colors only if we're not in the middle of applying them
          if (!hasExistingStyles) {
            cell.style.backgroundColor = "";
            cell.style.color = "";
          }
          
          // Apply rule-based coloring first
          this.settings.rules.forEach(rule => {
            let matches = false;
            try {
              if (rule.regex) {
                // Use regex pattern matching
                const flags = rule.caseSensitive ? 'g' : 'gi';
                const regex = new RegExp(rule.match, flags);
                matches = regex.test(cellText);
              } else {
                // Use simple text matching
                if (rule.caseSensitive) {
                  matches = cellText.includes(rule.match);
                } else {
                  matches = cellText.toLowerCase().includes(rule.match.toLowerCase());
                }
              }
            } catch (e) {
              // Fallback to simple includes on regex error
              matches = cellText.includes(rule.match);
            }
            
            if (matches) {
              if (rule.bg) cell.style.backgroundColor = rule.bg;
              if (rule.color) cell.style.color = rule.color;
            }
          });
          
          if (this.settings.numericRules && this.settings.numericRules.length) {
            let text = cellText.trim();
            let isNumber = false;
            if (this.settings.numericStrict) {
              isNumber = /^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$/.test(text);
            } else {
              isNumber = !isNaN(parseFloat(text.replace(/,/g, '')));
            }
            if (isNumber) {
              const num = parseFloat(text.replace(/,/g, ''));
              for (const nRule of this.settings.numericRules) {
                let match = false;
                switch (nRule.op) {
                  case 'lt': match = num < nRule.value; break;
                  case 'gt': match = num > nRule.value; break;
                  case 'eq': match = num === nRule.value; break;
                  case 'le': match = num <= nRule.value; break;
                  case 'ge': match = num >= nRule.value; break;
                }
                if (match) {
                  if (nRule.bg) cell.style.backgroundColor = nRule.bg;
                  if (nRule.color) cell.style.color = nRule.color;
                  break;
                }
              }
            }
          }
          // Now apply single cell coloring so it overrides rule-based coloring
          if (colorData) {
            if (colorData.bg) cell.style.backgroundColor = colorData.bg;
            if (colorData.color) cell.style.color = colorData.color;
          }
        });
      });
    });
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
  // Refresh the currently active file's rendered view
  applyColorsToActiveFile() {
    const activeView = this.app.workspace.getActiveViewOfType(require('obsidian').MarkdownView);
    const filePath = this.app.workspace.getActiveFile()?.path;
    if (!activeView || !filePath) return;
    
    // Try to find preview element with retries, as it may not be rendered yet
    let attempts = 0;
    const maxAttempts = 8;
    const tryApplyColors = () => {
      attempts++;
      const previewEl = activeView.contentEl.querySelector('.markdown-preview-view');
      if (previewEl) {
        this.applyColorsToContainer(previewEl, filePath);
      } else if (attempts < maxAttempts) {
        // Retry with increasing delays
        const delay = attempts <= 3 ? 50 : (attempts <= 5 ? 100 : 200);
        setTimeout(tryApplyColors, delay);
      }
    };
    tryApplyColors();
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

// Settings Tab
class ColorTableSettingTab extends PluginSettingTab {

  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }


  display() {
    const { containerEl } = this;
    containerEl.empty();

  new Setting(containerEl).setName("Table Cell Color Settings").setHeading();

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

    new Setting(containerEl)
      .setName("Default Background Color")
      .addColorPicker(picker =>
        picker.setValue(this.plugin.settings.defaultBgColor)
              .onChange(async val => {
                this.plugin.settings.defaultBgColor = val;
                await this.plugin.saveSettings();
              }));

    new Setting(containerEl)
      .setName("Default Text Color")
      .addColorPicker(picker =>
        picker.setValue(this.plugin.settings.defaultTextColor)
              .onChange(async val => {
                this.plugin.settings.defaultTextColor = val;
                await this.plugin.saveSettings();
              }));

  // Rule-based coloring
  new Setting(containerEl).setName("Rule-based Cell Coloring").setHeading();

    // Add new rule input + color pickers + button
    const ruleDiv = containerEl.createDiv({ cls: "rule-input pretty-flex" });
    const input = ruleDiv.createEl("input", { type: "text", placeholder: "Text or Pattern to match" });
    
    // Regex presets dropdown
    const regexPresets = [
      { name: "Email", pattern: "\\b[\\w._%+-]+@[\\w.-]+\\.[A-Z]{2,}\\b" },
      { name: "URL", pattern: "https?:\\/\\/[^\\s]+" },
      { name: "Phone", pattern: "\\b\\d{3}[-.]?\\d{3}[-.]?\\d{4}\\b" },
      { name: "Number", pattern: "\\b\\d+\\b" },
      { name: "Currency", pattern: "\\$\\d+(?:\\.\\d{2})?" },
      { name: "Date", pattern: "\\b\\d{1,2}\\/\\d{1,2}\\/\\d{4}\\b" },
      { name: "IP Address", pattern: "\\b(?:\\d{1,3}\\.){3}\\d{1,3}\\b" }
    ];
    
    const presetSelect = ruleDiv.createEl("select", { cls: "regex-preset" });
    presetSelect.style.width = "120px";
    presetSelect.createEl("option", { text: "Presets...", value: "" });
    regexPresets.forEach(preset => {
      presetSelect.createEl("option", { text: preset.name, value: preset.pattern });
    });
    presetSelect.onchange = (e) => {
      if (e.target.value) {
        input.value = e.target.value;
        regexToggle.checked = true;
        presetSelect.value = "";
      }
    };
    
    // Regex toggle
    const regexToggle = ruleDiv.createEl("input", { type: "checkbox", cls: "regex-toggle" });
    regexToggle.title = "Use regex pattern";
    regexToggle.style.width = "20px";
    regexToggle.style.marginLeft = "4px";
    
    // Case sensitive toggle
    const caseToggle = ruleDiv.createEl("input", { type: "checkbox", cls: "case-toggle" });
    caseToggle.title = "Case sensitive";
    caseToggle.style.width = "20px";
    caseToggle.style.marginLeft = "8px";
    
    // Color pickers beside new rule btn
    const addRuleColorInput = ruleDiv.createEl("input", { type: "color", cls: "add-rule-color-picker" });
    addRuleColorInput.value = this.plugin.settings.defaultTextColor;
    addRuleColorInput.title = "Text Color";
    const addRuleBgInput = ruleDiv.createEl("input", { type: "color", cls: "add-rule-bg-picker" });
    addRuleBgInput.value = this.plugin.settings.defaultBgColor;
    addRuleBgInput.title = "Background Color";
  
  // Add Rule button
  const addBtn = ruleDiv.createEl('button', { cls: 'add-btn add-rule-btn', attr: { 'aria-label': 'Add Rule' } });
  addBtn.textContent = 'Add Rule';
  addBtn.onclick = async () => {
      const val = input.value.trim();
      if (!val) return;
      if (!this.plugin.settings) this.plugin.settings = { defaultBgColor: "#331717", defaultTextColor: "#ffb2b2", enableContextMenu: true, rules: [], numericRules: [] };
      if (!Array.isArray(this.plugin.settings.rules)) this.plugin.settings.rules = [];
      const newRule = { 
        match: val, 
        bg: addRuleBgInput.value, 
        color: addRuleColorInput.value,
        regex: regexToggle.checked,
        caseSensitive: caseToggle.checked
      };
      this.plugin.settings.rules.push(newRule);
      await this.plugin.saveSettings();
      input.value = "";
      regexToggle.checked = false;
      caseToggle.checked = false;
      presetSelect.value = "";
      this.display();
    };
    if (this.plugin.settings.rules?.length) {
      const rulesList = containerEl.createDiv({ cls: 'rules-list' });
      this.plugin.settings.rules.forEach((rule, i) => {
        const ruleSetting = new Setting(rulesList)
          .setName(rule.match)
          .addColorPicker(picker => picker.setValue(rule.color).onChange(async val => { rule.color = val; await this.plugin.saveDataSettings(); }))
          .addColorPicker(picker => picker.setValue(rule.bg).onChange(async val => { rule.bg = val; await this.plugin.saveDataSettings(); }));
        const nameEl = ruleSetting.nameEl;
        nameEl.empty();
        const dragHandle = nameEl.createEl('span', { cls: 'drag-handle' });
        dragHandle.title = 'Drag to reorder';
        if (typeof window.setIcon === 'function') {
          window.setIcon(dragHandle, 'grip-vertical');
        } else if (typeof setIcon === 'function') {
          setIcon(dragHandle, 'grip-vertical');
        } else {
          dragHandle.textContent = '≡';
        }
        
        // Display pattern with regex badge if applicable
        const patternContainer = nameEl.createDiv({ cls: 'rule-pattern-container' });
        const matchSpan = patternContainer.createEl('span', { text: rule.match, cls: 'rule-match' });
        matchSpan.style.cursor = 'pointer';
        matchSpan.title = 'Click to edit rule';
        
        // Add regex badge if rule uses regex
        if (rule.regex) {
          const regexBadge = patternContainer.createEl('span', { text: '.*', cls: 'regex-badge' });
          regexBadge.title = 'Regex pattern';
          regexBadge.style.marginLeft = '6px';
          regexBadge.style.fontSize = '0.75em';
          regexBadge.style.padding = '2px 4px';
          regexBadge.style.borderRadius = '3px';
          regexBadge.style.backgroundColor = 'var(--color-accent, #6366f1)';
          regexBadge.style.color = 'white';
          regexBadge.style.display = 'inline-block';
        }
        
        // Add case-sensitive badge if applicable
        if (rule.caseSensitive) {
          const caseBadge = patternContainer.createEl('span', { text: 'Aa', cls: 'case-badge' });
          caseBadge.title = 'Case sensitive';
          caseBadge.style.marginLeft = '6px';
          caseBadge.style.fontSize = '0.75em';
          caseBadge.style.padding = '2px 4px';
          caseBadge.style.borderRadius = '3px';
          caseBadge.style.backgroundColor = 'var(--background-modifier-border)';
          caseBadge.style.display = 'inline-block';
        }
        
        const startEditing = async () => {
          const editInput = nameEl.createEl('input', { type: 'text' });
          editInput.value = rule.match;
          patternContainer.style.display = 'none';
          const finish = async (save) => {
            const newVal = editInput.value.trim();
            if (save && newVal) {
              rule.match = newVal;
              await this.plugin.saveSettings();
            }
            editInput.remove();
            patternContainer.style.display = '';
          };
          editInput.addEventListener('blur', () => finish(true));
          editInput.addEventListener('keydown', async (e) => {
            if (e.key === 'Enter') {
              await finish(true);
            } else if (e.key === 'Escape') {
              await finish(false);
            }
          });
          editInput.focus();
          editInput.select();
        };
        matchSpan.addEventListener('click', startEditing);
        
        // Add regex and case-sensitive toggles to control panel
        const controlsPanel = ruleSetting.controlEl.createDiv({ cls: 'rule-toggles' });
        controlsPanel.style.marginRight = '12px';
        
        const regexCheckbox = controlsPanel.createEl('input', { type: 'checkbox', cls: 'rule-regex-checkbox' });
        regexCheckbox.checked = rule.regex || false;
        regexCheckbox.title = 'Use regex pattern';
        regexCheckbox.style.width = '18px';
        regexCheckbox.style.cursor = 'pointer';
        regexCheckbox.onchange = async (e) => {
          rule.regex = e.target.checked;
          await this.plugin.saveSettings();
          this.display();
        };
        
        const caseCheckbox = controlsPanel.createEl('input', { type: 'checkbox', cls: 'rule-case-checkbox' });
        caseCheckbox.checked = rule.caseSensitive || false;
        caseCheckbox.title = 'Case sensitive';
        caseCheckbox.style.width = '18px';
        caseCheckbox.style.marginLeft = '8px';
        caseCheckbox.style.cursor = 'pointer';
        caseCheckbox.onchange = async (e) => {
          rule.caseSensitive = e.target.checked;
          await this.plugin.saveSettings();
          this.display();
        };
        const rootEl = ruleSetting.settingEl;
        rootEl.dataset.ruleIndex = String(i);
        dragHandle.setAttribute('draggable', 'true');
        dragHandle.addEventListener('dragstart', (e) => {
          e.dataTransfer.setData('text/plain', String(i));
          rootEl.classList.add('dragging');
          if (e.dataTransfer.setDragImage) {
            e.dataTransfer.setDragImage(dragHandle, 8, 8);
          }
        });
        dragHandle.addEventListener('dragend', () => {
          rootEl.classList.remove('dragging');
          document.querySelectorAll('.rule-over').forEach(el => el.classList.remove('rule-over'));
        });
        rootEl.addEventListener('dragover', (e) => {
          e.preventDefault();
          rootEl.classList.add('rule-over');
        });
        rootEl.addEventListener('dragleave', () => {
          rootEl.classList.remove('rule-over');
        });
        rootEl.addEventListener('drop', async (e) => {
          e.preventDefault();
          const fromIndex = Number(e.dataTransfer.getData('text/plain'));
          const toIndex = Number(rootEl.dataset.ruleIndex);
          if (isNaN(fromIndex) || isNaN(toIndex) || fromIndex === toIndex) return;
          const rules = this.plugin.settings.rules;
          const [moved] = rules.splice(fromIndex, 1);
          rules.splice(toIndex, 0, moved);
          await this.plugin.saveDataSettings();
          this.display();
        });
        const controls = ruleSetting.controlEl.createDiv({ cls: 'rule-controls' });
        const delButton = new ButtonComponent(controls);
        delButton.setIcon('trash-2');
        delButton.setTooltip('Delete rule');
        delButton.buttonEl.classList.add('rule-delete-btn');
        delButton.onClick(async () => {
          this.plugin.settings.rules.splice(i, 1);
          await this.plugin.saveSettings();
          this.display();
        });
      });
    }

  // --- Numerical Rules Section ---
  new Setting(containerEl).setName("Numerical Rules").setHeading();
  const numRuleDiv = containerEl.createDiv({ cls: "num-rule-input pretty-flex num-rule-row" });
    // Operator select
    const opSelect = numRuleDiv.createEl("select", { cls: "num-op-select" });
    const opOptions = [
      { label: "<", value: "lt" },
      { label: "≤", value: "le" },
      { label: "=", value: "eq" },
      { label: "≥", value: "ge" },
      { label: ">", value: "gt" }
    ];
    opOptions.forEach(opt => {
      const o = opSelect.createEl("option");
      o.value = opt.value;
      o.text = opt.label;
    });
    // Value input
    const numInput = numRuleDiv.createEl("input", { type: "number", placeholder: "Value", cls: "num-value-input" });
    // Color pickers
    const colorInput = numRuleDiv.createEl("input", { type: "color", cls: "num-color-picker" });
    colorInput.value = this.plugin.settings.defaultTextColor;
    colorInput.title = "Text Color";
    const bgInput = numRuleDiv.createEl("input", { type: "color", cls: "num-bg-picker" });
    bgInput.value = this.plugin.settings.defaultBgColor;
    bgInput.title = "Background Color";
  // Add Numeric Rule button
  const addNumBtn = numRuleDiv.createEl('button', { cls: 'add-btn add-rule-btn', attr: { 'aria-label': 'Add Numeric Rule' } });
  addNumBtn.textContent = 'Add Rule';
  addNumBtn.onclick = async () => {
      const op = opSelect.value;
      const val = parseFloat(numInput.value);
      if (!op || isNaN(val)) return;
      if (!this.plugin.settings.numericRules) this.plugin.settings.numericRules = [];
      this.plugin.settings.numericRules.push({ op, value: val, color: colorInput.value, bg: bgInput.value });
      await this.plugin.saveSettings();
      numInput.value = "";
      this.display();
    };

    // --- Numerical Rules List ---
    if (this.plugin.settings.numericRules?.length) {
      const numRulesList = containerEl.createDiv({ cls: 'num-rules-list' });
      this.plugin.settings.numericRules.forEach((nRule, i) => {
        const nRuleSetting = new Setting(numRulesList);
        // Drag handle
        const nameEl = nRuleSetting.nameEl;
        nameEl.empty();
        const dragHandle = nameEl.createEl('span', { cls: 'drag-handle' });
        dragHandle.title = 'Drag to reorder';
        if (typeof window.setIcon === 'function') {
          window.setIcon(dragHandle, 'grip-vertical');
        } else if (typeof setIcon === 'function') {
          setIcon(dragHandle, 'grip-vertical');
        } else {
          dragHandle.textContent = '≡';
        }
        // Editable rule label
        const opMap = { lt: '<', le: '≤', eq: '=', ge: '≥', gt: '>' };
        const ruleLabel = nameEl.createEl('span', { text: `${opMap[nRule.op]} ${nRule.value}`, cls: 'num-rule-label' });
        ruleLabel.style.cursor = 'pointer';
        ruleLabel.title = 'Click to edit rule';
        const startEditing = () => {
          ruleLabel.style.display = 'none';
          const editDiv = nameEl.createDiv({ cls: 'num-edit-row pretty-flex' });
          const opEdit = editDiv.createEl('select', { cls: 'num-op-select' });
          opOptions.forEach(opt => {
            const o = opEdit.createEl('option');
            o.value = opt.value;
            o.text = opt.label;
            if (opt.value === nRule.op) o.selected = true;
          });
          const valEdit = editDiv.createEl('input', { type: 'number', value: nRule.value, cls: 'num-value-input' });
          valEdit.style.width = '70px';
          const saveEdit = () => {
            nRule.op = opEdit.value;
            nRule.value = parseFloat(valEdit.value);
            this.plugin.saveSettings();
            editDiv.remove();
            ruleLabel.textContent = `${opMap[nRule.op]} ${nRule.value}`;
            ruleLabel.style.display = '';
          };
          valEdit.addEventListener('keydown', e => { if (e.key === 'Enter') saveEdit(); });
          opEdit.addEventListener('change', saveEdit);
          valEdit.addEventListener('blur', saveEdit);
        };
        ruleLabel.addEventListener('click', startEditing);
        nRuleSetting.addColorPicker(picker => picker.setValue(nRule.color).onChange(async val => { nRule.color = val; await this.plugin.saveDataSettings(); }))
          .addColorPicker(picker => picker.setValue(nRule.bg).onChange(async val => { nRule.bg = val; await this.plugin.saveDataSettings(); }));
        const rootEl = nRuleSetting.settingEl;
        rootEl.dataset.numRuleIndex = String(i);
        dragHandle.setAttribute('draggable', 'true');
        dragHandle.addEventListener('dragstart', (e) => {
          e.dataTransfer.setData('text/plain', String(i));
          rootEl.classList.add('dragging');
          if (e.dataTransfer.setDragImage) {
            e.dataTransfer.setDragImage(dragHandle, 8, 8);
          }
        });
        dragHandle.addEventListener('dragend', () => {
          rootEl.classList.remove('dragging');
          document.querySelectorAll('.num-rule-over').forEach(el => el.classList.remove('num-rule-over'));
        });
        rootEl.addEventListener('dragover', (e) => {
          e.preventDefault();
          rootEl.classList.add('num-rule-over');
        });
        rootEl.addEventListener('dragleave', () => {
          rootEl.classList.remove('num-rule-over');
        });
        rootEl.addEventListener('drop', async (e) => {
          e.preventDefault();
          const fromIndex = Number(e.dataTransfer.getData('text/plain'));
          const toIndex = Number(rootEl.dataset.numRuleIndex);
          if (isNaN(fromIndex) || isNaN(toIndex) || fromIndex === toIndex) return;
          const rules = this.plugin.settings.numericRules;
          const [moved] = rules.splice(fromIndex, 1);
          rules.splice(toIndex, 0, moved);
          await this.plugin.saveDataSettings();
          this.display();
        });
        const controls = nRuleSetting.controlEl.createDiv({ cls: 'rule-controls' });
        const delButton = new ButtonComponent(controls);
        delButton.setIcon('trash-2');
        delButton.setTooltip('Delete rule');
        delButton.buttonEl.classList.add('rule-delete-btn');
        delButton.onClick(async () => {
          this.plugin.settings.numericRules.splice(i, 1);
          await this.plugin.saveSettings();
          this.display();
        });
      });
    }

if (!document.head.querySelector('style[data-color-table-cell-numeric]')) {
  style.setAttribute('data-color-table-cell-numeric', '');
  document.head.appendChild(style);
}
  }
}

// Modal to list rules and apply them to a specific table element
class RulesModal extends Modal {
  constructor(app, plugin, tableEl) {
    super(app);
    this.plugin = plugin;
    this.tableEl = tableEl;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl('h3', { text: 'Apply Rule to Table' });

    if (!this.plugin.settings.rules?.length) {
      contentEl.createEl('div', { text: 'No rules defined. Open plugin settings to add rules.' });
      const openBtn = contentEl.createEl('button', { text: 'Open Settings', cls: 'mod-cta' });
      openBtn.onclick = () => {
        this.close();
        // Open the settings
        this.app.commands.executeCommandById('app:open-settings');
      };
      return;
    }

    const list = contentEl.createEl('div', { cls: 'rule-list' });
    this.plugin.settings.rules.forEach((rule, idx) => {
      const row = list.createDiv({ cls: 'rule-row' });
      row.createEl('span', { text: rule.match, cls: 'rule-match' });
      row.createEl('span', { text: ` ` });
      const applyBtn = row.createEl('button', { text: 'Apply', cls: 'mod-cta' });
      applyBtn.onclick = async () => {
        // Apply to all cells in this.tableEl matching the rule
        this.tableEl.querySelectorAll('td, th').forEach(cell => {
          if (cell.textContent.includes(rule.match)) {
            if (rule.bg) cell.style.backgroundColor = rule.bg;
            if (rule.color) cell.style.color = rule.color;
          }
        });

        // Save into plugin.cellData for the active file
        const fileId = this.plugin.app.workspace.getActiveFile()?.path;
        if (fileId) {
          if (!this.plugin.cellData[fileId]) this.plugin.cellData[fileId] = {};
          const tableIndex = Array.from(this.tableEl.parentNode.querySelectorAll('table')).indexOf(this.tableEl);
          const tableKey = `table_${tableIndex}`;
          if (!this.plugin.cellData[fileId][tableKey]) this.plugin.cellData[fileId][tableKey] = {};
          const tableColors = this.plugin.cellData[fileId][tableKey];

          this.tableEl.querySelectorAll('tr').forEach((tr, rIdx) => {
            tr.querySelectorAll('td, th').forEach((cell, cIdx) => {
              if (cell.textContent.includes(rule.match)) {
                const rowKey = `row_${rIdx}`;
                if (!tableColors[rowKey]) tableColors[rowKey] = {};
                tableColors[rowKey][`col_${cIdx}`] = {
                  ...(tableColors[rowKey][`col_${cIdx}`] || {}),
                  bg: rule.bg,
                  color: rule.color
                };
              }
            });
          });

          await this.plugin.saveDataColors();
        }
      };
    });
  }

  onClose() {
    this.contentEl.empty();
  }
}