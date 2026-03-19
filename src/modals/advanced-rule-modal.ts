import { App, Modal, Setting, setIcon } from "obsidian";
import { WHEN_OPTIONS, MATCH_OPTIONS } from "../constants";
import type { AdvancedCondition, AdvancedRule } from "../types";
import type TableColorPlugin from "../plugin";

class ConditionRow {
  private containerEl: HTMLElement;
  private index: number;
  private data: AdvancedCondition;
  private onChange: () => void;

  constructor(
    parent: HTMLElement,
    index: number,
    initialData: AdvancedCondition,
    onChange: () => void,
  ) {
    this.containerEl = parent;
    this.index = index;
    this.data = { ...initialData };
    this.onChange = onChange;
    this.render();
  }

  private render(): void {
    const row = this.containerEl.createDiv({ cls: "ctc-cr-adv-cond-row" });

    const whenSel = row.createEl("select", { cls: "ctc-cr-adv-cond-when ctc-cr-select" });
    WHEN_OPTIONS.forEach((opt) => {
      const o = whenSel.createEl("option");
      o.value = opt.value;
      o.text = opt.label;
      if (this.data.when === opt.value) o.selected = true;
    });
    whenSel.addEventListener("change", () => {
      this.data.when = whenSel.value;
      this.onChange();
    });

    const matchSel = row.createEl("select", { cls: "ctc-cr-adv-cond-match ctc-cr-select" });
    MATCH_OPTIONS.forEach((opt) => {
      const o = matchSel.createEl("option");
      o.value = opt.value;
      o.text = opt.label;
      if (this.data.match === opt.value) o.selected = true;
    });
    matchSel.addEventListener("change", () => {
      this.data.match = matchSel.value;
      this.onChange();
    });

    const valueInput = row.createEl("input", {
      type: "text",
      cls: "ctc-cr-adv-cond-value ctc-condition-val-input",
    });
    valueInput.placeholder = "Value";
    valueInput.value = this.data.value || "";
    valueInput.addEventListener("input", () => {
      this.data.value = valueInput.value;
      this.onChange();
    });

    const delBtn = row.createEl("button", { cls: "mod-ghost" });
    try { setIcon(delBtn, "x"); } catch { delBtn.textContent = "×"; }
    delBtn.addEventListener("click", () => {
      row.remove();
      this.onChange();
    });
  }

  getData(): AdvancedCondition {
    return { ...this.data };
  }
}

export class AdvancedRuleModal extends Modal {
  private plugin: TableColorPlugin;
  private index: number;

  constructor(app: App, plugin: TableColorPlugin, index: number) {
    super(app);
    this.plugin = plugin;
    this.index = index;
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.addClass("ctc-cr-adv-modal");

    const rule: AdvancedRule = this.plugin.settings.advancedRules[this.index] || {
      logic: "any",
      conditions: [],
      target: "cell",
      color: null,
      bg: null,
    };

    const heading = new Setting(contentEl).setName("Advanced rule").setHeading();
    heading.settingEl.classList.add("ctc-cr-adv-modal-heading");

    // Name input
    const nameRow = contentEl.createDiv({ cls: "ctc-cr-adv-name-row" });
    const nameInput = nameRow.createEl("input", {
      type: "text",
      cls: "ctc-cr-adv-name-input",
      placeholder: "Rule name (optional)",
    });
    nameInput.value = rule.name || "";

    // Logic buttons
    const logicRow = contentEl.createDiv({ cls: "ctc-cr-adv-logic" });
    logicRow.createEl("span", { text: "Match:", cls: "ctc-cr-adv-logic-label" });
    const logicBtns = logicRow.createDiv({ cls: "ctc-cr-adv-logic-buttons" });

    let currentLogic = rule.logic || "any";
    const makeLogicBtn = (label: string, value: "any" | "all" | "none") => {
      const btn = logicBtns.createEl("button", {
        text: label,
        cls: `mod-ghost ctc-cr-adv-logic-btn${currentLogic === value ? " mod-cta" : ""}`,
      });
      btn.addEventListener("click", () => {
        currentLogic = value;
        logicBtns.querySelectorAll("button").forEach((b) => b.classList.remove("mod-cta"));
        btn.classList.add("mod-cta");
      });
    };
    makeLogicBtn("Any", "any");
    makeLogicBtn("All", "all");
    makeLogicBtn("None", "none");

    // Conditions
    contentEl.createEl("h4", { text: "Conditions", cls: "ctc-cr-adv-h4" });
    const condsWrap = contentEl.createDiv({ cls: "ctc-cr-adv-conds-wrap" });
    const condRows: ConditionRow[] = [];

    const addConditionRow = (data: AdvancedCondition) => {
      const cr = new ConditionRow(condsWrap, condRows.length, data, () => {});
      condRows.push(cr);
    };

    (rule.conditions || []).forEach((c) => addConditionRow(c));

    const addCondRow = contentEl.createDiv({ cls: "ctc-cr-adv-add-row" });
    const addCondBtn = addCondRow.createEl("button", {
      text: "+ Add condition",
      cls: "mod-ghost ctc-cr-adv-add-btn",
    });
    addCondBtn.addEventListener("click", () => {
      addConditionRow({ when: "theCell", match: "contains", value: "" });
    });

    // Target
    new Setting(contentEl).setName("Apply to").addDropdown((dd) => {
      dd.addOption("cell", "Cell");
      dd.addOption("row", "Row");
      dd.addOption("column", "Column");
      dd.setValue(rule.target || "cell");
      dd.onChange((v) => { rule.target = v as "cell" | "row" | "column"; });
    });

    // Colors
    const colorRow = contentEl.createDiv({ cls: "ctc-cr-adv-color-row" });

    const textColorContainer = colorRow.createDiv({ cls: "ctc-cr-adv-text-color-container" });
    textColorContainer.createEl("span", { text: "Text:" });
    const textColorPicker = textColorContainer.createEl("input", {
      type: "color",
      cls: "ctc-cr-adv-text-color",
    });
    textColorPicker.value = rule.color || "#ffffff";
    const textResetBtn = textColorContainer.createEl("button", { cls: "mod-ghost ctc-cr-adv-color-reset" });
    try { setIcon(textResetBtn, "x"); } catch { textResetBtn.textContent = "×"; }
    textResetBtn.title = "Reset text color";
    textResetBtn.addEventListener("click", () => { rule.color = null; });

    const bgColorContainer = colorRow.createDiv({ cls: "ctc-cr-adv-bg-color-container" });
    bgColorContainer.createEl("span", { text: "Background:" });
    const bgColorPicker = bgColorContainer.createEl("input", {
      type: "color",
      cls: "ctc-cr-adv-bg-color",
    });
    bgColorPicker.value = rule.bg || "#000000";
    const bgResetBtn = bgColorContainer.createEl("button", { cls: "mod-ghost ctc-cr-adv-bg-reset" });
    try { setIcon(bgResetBtn, "x"); } catch { bgResetBtn.textContent = "×"; }
    bgResetBtn.title = "Reset background color";
    bgResetBtn.addEventListener("click", () => { rule.bg = null; });

    // Actions
    const actionsRow = contentEl.createDiv({ cls: "ctc-cr-adv-actions-row" });

    const deleteBtn = actionsRow.createEl("button", {
      text: "Delete rule",
      cls: "mod-warning ctc-cr-adv-delete",
    });
    deleteBtn.addEventListener("click", async () => {
      this.plugin.settings.advancedRules.splice(this.index, 1);
      await this.plugin.saveSettings();
      document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
      this.close();
    });

    const saveBtn = actionsRow.createEl("button", {
      text: "Save",
      cls: "mod-cta ctc-cr-adv-save",
    });
    saveBtn.addEventListener("click", async () => {
      // Collect conditions from DOM
      const conditions: AdvancedCondition[] = [];
      condsWrap.querySelectorAll(".ctc-cr-adv-cond-row").forEach((rowEl) => {
        const selects = rowEl.querySelectorAll("select");
        const input = rowEl.querySelector("input") as HTMLInputElement | null;
        if (selects.length >= 2) {
          conditions.push({
            when: (selects[0] as HTMLSelectElement).value,
            match: (selects[1] as HTMLSelectElement).value,
            value: input?.value || "",
          });
        }
      });

      const updatedRule: AdvancedRule = {
        name: nameInput.value.trim() || undefined,
        logic: currentLogic,
        conditions,
        target: rule.target,
        color: textColorPicker.value || null,
        bg: bgColorPicker.value || null,
      };

      this.plugin.settings.advancedRules[this.index] = updatedRule;
      await this.plugin.saveSettings();
      this.plugin.applyColorsToActiveFile();
      document.dispatchEvent(new CustomEvent("ctc-adv-rules-changed"));
      this.close();
    });
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
