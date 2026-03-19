// All TypeScript interfaces and types for the Color Table Cells plugin
import type { Plugin } from "obsidian";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type CellDataStore = Record<string, Record<string, Record<string, any>>>;

export interface CellColorData {
  bg?: string | null;
  color?: string | null;
  [key: string]: string | null | undefined;
}

export interface CellCoordinates {
  row: number;
  col: number;
}

export interface ColorSnapshot {
  type: "cell_color" | "row_color" | "column_color";
  fileId: string;
  tableIndex: number;
  coords: { row?: number; col?: number };
  oldColors: CellColorData | Record<string, CellColorData> | null | undefined;
  newColors: CellColorData | Record<string, CellColorData>;
}

export interface AdvancedCondition {
  when: string;
  match: string;
  value: string;
}

export interface AdvancedRule {
  name?: string;
  logic: "any" | "all" | "none";
  conditions: AdvancedCondition[];
  target: "cell" | "row" | "column";
  color: string | null;
  bg: string | null;
}

export interface ColoringRule {
  target: string;
  when: string;
  match: string;
  value: string | number | null;
  color: string | null;
  bg: string | null;
}

export interface PresetColor {
  name: string;
  color: string;
}

export interface PluginSettings {
  enableContextMenu: boolean;
  showColorRowInMenu: boolean;
  showColorColumnInMenu: boolean;
  showUndoRedoInMenu: boolean;
  coloringRules: ColoringRule[];
  coloringSort: string;
  advancedRules: AdvancedRule[];
  numericStrict: boolean;
  livePreviewColoring: boolean;
  persistUndoHistory: boolean;
  recentColors: string[];
  presetColors: PresetColor[];
  showStatusRefreshIcon: boolean;
  showRibbonRefreshIcon: boolean;
  processAllCellsOnOpen?: boolean;
  defaultBgColor?: string;
  defaultTextColor?: string;
  skipBackgroundInAdvancedRules?: boolean;
}

export interface PluginData {
  settings: PluginSettings;
  cellData: Record<string, Record<string, Record<string, CellColorData>>>;
}

export interface ColorPickerOptions {
  plugin: { settings: PluginSettings };
  onPick: (color: string) => void;
  initialColor?: string | null;
  anchorEl: HTMLElement;
}

export interface SelectOption {
  label: string;
  value: string;
}

export interface ITableColorPlugin extends Plugin {
  settings: PluginSettings;
  cellData: CellDataStore;
  undoStack: ColorSnapshot[];
  redoStack: ColorSnapshot[];
  applyColorsToActiveFile(): void;
  saveSettings(): Promise<void>;
  createStatusBarIcon(): void;
  removeStatusBarIcon(): void;
  hardRefreshTableColors(): void;
}
