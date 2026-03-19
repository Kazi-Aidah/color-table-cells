import { App, PluginSettingTab, Setting, Modal, Notice, Menu, setIcon } from "obsidian";
import { TARGET_OPTIONS, WHEN_OPTIONS, MATCH_OPTIONS, SORT_OPTIONS } from "./constants";
import { AdvancedRuleModal } from "./modals/advanced-rule-modal";
import type { ColoringRule, ITableColorPlugin } from "./types";

export class ColorTableSettingTab extends PluginSettingTab {
  plugin: ITableColorPlugin;

  constructor(app: App, plugin: ITableColorPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.addClass("ctc-settings");

    // new Setting(containerEl).setName("Color Table Cells").setHeading();

    // Context menu settings
    new Setting(containerEl)
      .setName("Enable context menu")
      .setDesc("Show color options when right-clicking table cells.")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.enableContextMenu).onChange(async (v) => {
          this.plugin.settings.enableContextMenu = v;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Show 'Color row' in menu")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.showColorRowInMenu).onChange(async (v) => {
          this.plugin.settings.showColorRowInMenu = v;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Show 'Color column' in menu")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.showColorColumnInMenu).onChange(async (v) => {
          this.plugin.settings.showColorColumnInMenu = v;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Show undo/redo in menu")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.showUndoRedoInMenu).onChange(async (v) => {
          this.plugin.settings.showUndoRedoInMenu = v;
          await this.plugin.saveSettings();
        }),
      );

    // Live preview
    new Setting(containerEl)
      .setName("Live preview coloring")
      .setDesc("Apply coloring rules in live preview (editor) mode.")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.livePreviewColoring).onChange(async (v) => {
          this.plugin.settings.livePreviewColoring = v;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
        }),
      );

    // Persist undo history
    new Setting(containerEl)
      .setName("Persist undo history")
      .setDesc("Save undo/redo history between sessions.")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.persistUndoHistory).onChange(async (v) => {
          this.plugin.settings.persistUndoHistory = v;
          await this.plugin.saveSettings();
        }),
      );

    // Status bar icon
    new Setting(containerEl)
      .setName("Show status bar refresh icon")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.showStatusRefreshIcon).onChange(async (v) => {
          this.plugin.settings.showStatusRefreshIcon = v;
          await this.plugin.saveSettings();
          if (v) this.plugin.createStatusBarIcon();
          else this.plugin.removeStatusBarIcon();
        }),
      );

    // Ribbon icon
    new Setting(containerEl)
      .setName("Show ribbon refresh icon")
      .addToggle((t) =>
        t.setValue(this.plugin.settings.showRibbonRefreshIcon).onChange(async (v) => {
          this.plugin.settings.showRibbonRefreshIcon = v;
          await this.plugin.saveSettings();
        }),
      );

    // Release notes
    new Setting(containerEl)
      .setName("Release notes")
      .setDesc("View the latest release notes.")
      .setClass("ctc-settings-release-notes")
      .addButton((btn) =>
        btn.setButtonText("View release notes").onClick(() => {
          const { ChangelogModal } = require("./modals/changelog-modal");
          new ChangelogModal(this.app, this.plugin).open();
        }),
      );

    // ---- Coloring Rules Section ----
    new Setting(containerEl).setName("Coloring rules").setHeading();

    // Search bar
    const crSection = containerEl.createDiv();
    const searchContainer = crSection.createDiv({ cls: "ctc-search-container" });
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

    // Header row
    const headerRow = crSection.createDiv({ cls: "ctc-cr-header-row" });
    ["Target", "When", "Match", "Value", "Text", "BG", ""].forEach((h) => {
      headerRow.createEl("span", { text: h });
    });

    const rulesContainer = crSection.createDiv({ cls: "ctc-rules-list" });

    const labelFor = (opts: typeof TARGET_OPTIONS, val: string) =>
      (opts.find((o) => o.value === val)?.label || "").toLowerCase();
    const isRegexRule = (rule: ColoringRule) => rule.match === "isRegex";
    const isNumericRule = (rule: ColoringRule) =>
      ["eq", "gt", "lt", "ge", "le"].includes(rule.match);

    const renderRules = () => {
      rulesContainer.empty();
      let rules = Array.isArray(this.plugin.settings.coloringRules)
        ? [...this.plugin.settings.coloringRules]
        : [];

      if (searchTerm) {
        rules = rules.filter((r) => {
          const blob = [
            labelFor(TARGET_OPTIONS, r.target || ""),
            labelFor(WHEN_OPTIONS, r.when || ""),
            labelFor(MATCH_OPTIONS, r.match || ""),
            r.value != null ? String(r.value) : "",
          ].join(" ").toLowerCase();
          return blob.includes(searchTerm);
        });
      }

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
        rules.sort((a, b) => Number(isNumericRule(b)) - Number(isNumericRule(a)));
      } else if (sortMode === "mode") {
        const order: Record<string, number> = { cell: 0, row: 1, column: 2 };
        rules.sort((a, b) => order[a.target || "cell"] - order[b.target || "cell"]);
      }

      rules.forEach((rule) => {
        const row = rulesContainer.createDiv({ cls: "ctc-cr-rule-row ctc-pretty-flex" });
        const originalIdx = this.plugin.settings.coloringRules.indexOf(rule);
        row.dataset.idx = String(originalIdx);

        const makeSelect = (
          cls: string,
          opts: typeof TARGET_OPTIONS,
          current: string,
          placeholder: string,
          onChange: (v: string) => void,
        ) => {
          const sel = row.createEl("select", { cls });
          const ph = sel.createEl("option", { text: placeholder, value: "" });
          ph.disabled = true;
          ph.selected = !current;
          opts.forEach((opt) => {
            const o = sel.createEl("option");
            o.value = opt.value;
            o.text = opt.label;
            if (current === opt.value) o.selected = true;
          });
          sel.addEventListener("change", async () => {
            onChange(sel.value);
            await this.plugin.saveSettings();
            this.plugin.applyColorsToActiveFile();
          });
          return sel;
        };

        makeSelect("ctc-cr-select", TARGET_OPTIONS, rule.target || "", "Target", (v) => { rule.target = v; });
        makeSelect("ctc-cr-select", WHEN_OPTIONS, rule.when || "", "When", (v) => { rule.when = v; });
        makeSelect("ctc-cr-select", MATCH_OPTIONS, rule.match || "", "Match", (v) => {
          rule.match = v;
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
            ? v === "" ? null : Number(v)
            : v;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
        });

        const colorPicker = row.createEl("input", { type: "color", cls: "ctc-cr-color-picker" });
        colorPicker.value = rule.color || "#000000";
        colorPicker.title = "Text color";
        colorPicker.addEventListener("change", async () => {
          rule.color = colorPicker.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
        });

        const bgPicker = row.createEl("input", { type: "color", cls: "ctc-cr-bg-picker" });
        bgPicker.value = rule.bg || "#000000";
        bgPicker.title = "Background color";
        bgPicker.addEventListener("change", async () => {
          rule.bg = bgPicker.value;
          await this.plugin.saveSettings();
          this.plugin.applyColorsToActiveFile();
        });

        const delBtn = row.createEl("button", { cls: "mod-ghost ctc-cr-del-btn" });
        try { setIcon(delBtn, "x"); } catch { delBtn.textContent = "×"; }
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
            item.setTitle("Duplicate rule").setIcon("copy").onClick(async () => {
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
          const canMoveDown = originalIdx >= 0 && originalIdx < this.plugin.settings.coloringRules.length - 1;
          menu.addItem((item) =>
            item.setTitle("Move rule up").setIcon("arrow-up").setDisabled(!canMoveUp).onClick(async () => {
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
            item.setTitle("Move rule down").setIcon("arrow-down").setDisabled(!canMoveDown).onClick(async () => {
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
            item.setTitle("Reset text color").setIcon("text").onClick(async () => {
              rule.color = null;
              await this.plugin.saveSettings();
              this.plugin.applyColorsToActiveFile();
              renderRules();
            }),
          );
          menu.addItem((item) =>
            item.setTitle("Reset background color").setIcon("rectangle-horizontal").onClick(async () => {
              rule.bg = null;
              await this.plugin.saveSettings();
              this.plugin.applyColorsToActiveFile();
              renderRules();
            }),
          );
          menu.showAtMouseEvent(evt);
          evt.preventDefault();
        });
      });
    };

    // Sort + Add row
    const addRow = crSection.createDiv({ cls: "ctc-cr-add-row" });
    const sortSel = addRow.createEl("select", { cls: "ctc-cr-select" });
    SORT_OPTIONS.forEach((opt) => {
      const o = sortSel.createEl("option");
      o.text = opt.label;
      o.value = opt.value;
      if (opt.value === (this.plugin.settings.coloringSort || "lastAdded")) o.selected = true;
    });
    sortSel.addEventListener("change", async () => {
      this.plugin.settings.coloringSort = sortSel.value;
      await this.plugin.saveSettings();
      renderRules();
    });

    const addBtn = addRow.createEl("button", { cls: "mod-cta ctc-cr-add-flex" });
    addBtn.textContent = "+ Add rule";
    addBtn.addEventListener("click", async () => {
      if (!Array.isArray(this.plugin.settings.coloringRules))
        this.plugin.settings.coloringRules = [];
      this.plugin.settings.coloringRules.push({
        target: "", when: "", match: "", value: null, color: null, bg: null,
      });
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      renderRules();
    });

    renderRules();

    // ---- Advanced Rules Section ----
    const advHeading = new Setting(containerEl).setName("Advanced rules").setHeading();
    advHeading.settingEl.classList.add("ctc-cr-adv-heading");

    const advSearchContainer = containerEl.createDiv({ cls: "ctc-search-container ctc-pretty-flex" });
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

      const summaryForAdvRule = (ar: typeof advRules[0]) => {
        if (ar.name?.trim()) return ar.name;
        const targetPhrase = ar.target === "row" ? "Color rows" : ar.target === "column" ? "Color columns" : "Color cells";
        const conds = Array.isArray(ar.conditions) ? ar.conditions : [];
        if (!conds.length) return "(empty)";
        const parts = conds.map((c) => String(c.value || "").trim()).filter(Boolean);
        return `${targetPhrase} when ${parts.join(", ") || "conditions"}`;
      };

      if (advSearchTerm) {
        advRules = advRules.filter((ar) => {
          const summary = summaryForAdvRule(ar).toLowerCase();
          const allCondValues = Array.isArray(ar.conditions)
            ? ar.conditions.map((c) => String(c.value || "").toLowerCase()).join(" ")
            : "";
          return summary.includes(advSearchTerm) || allCondValues.includes(advSearchTerm);
        });
      }

      advRules.forEach((ar) => {
        const originalIdx = this.plugin.settings.advancedRules.indexOf(ar);
        const row = advList.createDiv({ cls: "ctc-cr-adv-row ctc-pretty-flex" });
        row.dataset.idx = String(originalIdx);

        const drag = row.createEl("span", { cls: "ctc-drag-handle" });
        try { setIcon(drag, "menu"); } catch { drag.textContent = "≡"; }
        drag.setAttribute("draggable", "true");
        drag.addEventListener("dragstart", (e) => {
          e.dataTransfer!.effectAllowed = "move";
          e.dataTransfer!.setData("text/plain", String(originalIdx));
          row.classList.add("dragging");
        });
        drag.addEventListener("dragend", () => {
          row.classList.remove("dragging");
          advList.querySelectorAll(".ctc-rule-over").forEach((el) => el.classList.remove("ctc-rule-over"));
        });
        row.addEventListener("dragover", (e) => { e.preventDefault(); row.classList.add("ctc-rule-over"); });
        row.addEventListener("dragleave", () => row.classList.remove("ctc-rule-over"));
        row.addEventListener("drop", async (e) => {
          e.preventDefault();
          const from = Number(e.dataTransfer!.getData("text/plain"));
          const to = Number(row.dataset.idx);
          if (isNaN(from) || isNaN(to) || from === to) return;
          const list = this.plugin.settings.advancedRules;
          const [m] = list.splice(from, 1);
          list.splice(to, 0, m);
          await this.plugin.saveSettings();
          renderAdv();
        });

        row.createEl("span", { text: summaryForAdvRule(ar), cls: "ctc-cr-adv-label" });

        const copyBtn = row.createEl("button", { cls: "mod-ghost" });
        copyBtn.title = "Duplicate rule";
        try { setIcon(copyBtn, "copy"); } catch { copyBtn.textContent = "⧉"; }
        copyBtn.addEventListener("click", async () => {
          const ruleCopy = JSON.parse(JSON.stringify(ar));
          this.plugin.settings.advancedRules.splice(originalIdx + 1, 0, ruleCopy);
          await this.plugin.saveSettings();
          document.dispatchEvent(new Event("ctc-adv-rules-changed"));
        });

        const settingsBtn = row.createEl("button", { cls: "mod-ghost" });
        try { setIcon(settingsBtn, "settings"); } catch { settingsBtn.textContent = "⚙"; }
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
        logic: "any", conditions: [], target: "cell", color: null, bg: null,
      });
      await this.plugin.saveSettings();
      renderAdv();
    });
    renderAdv();

    // ---- Data Management ----
    new Setting(containerEl).setName("Data management").setHeading();

    const exportImportRow = containerEl.createDiv({ cls: "ctc-cr-export-row ctc-pretty-flex" });

    const exportBtn = exportImportRow.createEl("button", { text: "Export settings" });
    exportBtn.addEventListener("click", async () => {
      try {
        const data = {
          settings: this.plugin.settings,
          cellData: this.plugin.cellData,
          exportDate: new Date().toISOString(),
        };
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `color-table-cells-backup-${Date.now()}.json`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        window.setTimeout(() => new Notice("Settings exported successfully!"), 500);
      } catch (e) {
        new Notice("Failed to export settings: " + (e as Error).message);
      }
    });

    const importBtn = exportImportRow.createEl("button", { text: "Import settings" });
    importBtn.addEventListener("click", () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".json";
      input.addEventListener("change", async (e) => {
        try {
          const file = (e.target as HTMLInputElement).files?.[0];
          if (!file) return;
          const text = await file.text();
          const data = JSON.parse(text);
          if (data.settings) Object.assign(this.plugin.settings, data.settings);
          if (data.cellData) this.plugin.cellData = data.cellData;
          await this.plugin.saveSettings();
          this.display();
          window.setTimeout(() => new Notice("Settings imported successfully!"), 500);
        } catch (e) {
          new Notice("Failed to import settings: " + (e as Error).message);
        }
      });
      input.click();
    });

    // ---- Danger Zone ----
    const dangerHeading = new Setting(containerEl).setName("Danger zone").setHeading();
    dangerHeading.settingEl.classList.add("ctc-cr-danger-heading");

    const dangerZoneRow = containerEl.createDiv({ cls: "ctc-cr-delete-container" });

    const makeDeleteBtn = (
      label: string,
      confirmTitle: string,
      confirmDesc: string,
      onConfirm: () => Promise<void>,
    ) => {
      const deleteRow = dangerZoneRow.createDiv({ cls: "ctc-cr-delete-row" });
      const btn = deleteRow.createEl("button", { text: label, cls: "mod-warning" });
      btn.addEventListener("click", () => {
        const modal = new Modal(this.app);
        const heading = new Setting(modal.contentEl)
          .setName(confirmTitle)
          .setDesc(confirmDesc)
          .setHeading();
        heading.settingEl.classList.add("ctc-modal-warning-heading");
        const btnRow = modal.contentEl.createDiv({ cls: "ctc-modal-delete-buttons" });
        const cancelBtn = btnRow.createEl("button", { text: "Cancel", cls: "mod-ghost" });
        const confirmBtn = btnRow.createEl("button", { text: "Delete all", cls: "mod-warning" });
        cancelBtn.addEventListener("click", () => modal.close());
        confirmBtn.addEventListener("click", async () => {
          await onConfirm();
          modal.close();
        });
        modal.open();
      });
    };

    makeDeleteBtn(
      "Delete all manual colors",
      "Delete all manual colors?",
      "This will remove all manually colored cells. This action cannot be undone.",
      async () => {
        this.plugin.cellData = {};
        await this.plugin.saveData({ settings: this.plugin.settings, cellData: {} });
        this.plugin.applyColorsToActiveFile();
        new Notice("All manual colors deleted");
      },
    );

    makeDeleteBtn(
      "Delete all coloring rules",
      "Delete all coloring rules?",
      `This will remove all ${this.plugin.settings.coloringRules?.length ?? 0} coloring rules. This action cannot be undone.`,
      async () => {
        this.plugin.settings.coloringRules = [];
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        new Notice("All coloring rules deleted");
        this.display();
      },
    );

    makeDeleteBtn(
      "Delete all advanced rules",
      "Delete all advanced rules?",
      `This will remove all ${this.plugin.settings.advancedRules?.length ?? 0} advanced rules. This action cannot be undone.`,
      async () => {
        this.plugin.settings.advancedRules = [];
        await this.plugin.saveSettings();
        this.plugin.applyColorsToActiveFile();
        new Notice("All advanced rules deleted");
        this.display();
      },
    );
  }

  hide(): void {
    // Clear rule-based colors so changes take effect on close
    document
      .querySelectorAll(".markdown-preview-view table td, .markdown-preview-view table th")
      .forEach((cell) => {
        const el = cell as HTMLElement;
        if (!el.hasAttribute("data-ctc-manual")) {
          el.style.backgroundColor = "";
          el.style.color = "";
        }
      });
    document
      .querySelectorAll(".cm-content table td, .cm-content table th")
      .forEach((cell) => {
        const el = cell as HTMLElement;
        if (!el.hasAttribute("data-ctc-manual")) {
          el.style.backgroundColor = "";
          el.style.color = "";
        }
      });
    window.setTimeout(() => this.plugin.applyColorsToActiveFile(), 100);
  }
}
