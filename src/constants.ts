import type { PluginSettings, SelectOption } from "./types";

// Debug configuration
export let IS_DEVELOPMENT = false;

export const debugLog = (...args: unknown[]): void => {
  IS_DEVELOPMENT && console.log("[CTC-DEBUG]", ...args);
};

export const debugWarn = (...args: unknown[]): void => {
  IS_DEVELOPMENT && console.warn("[CTC-WARN]", ...args);
};

// Allow toggling debug mode from console: window.setDebugMode(true/false)
if (typeof window !== "undefined") {
  (window as Window & { setDebugMode?: (v: boolean) => void }).setDebugMode = (
    value: boolean,
  ) => {
    IS_DEVELOPMENT = value;
    console.log(`[CTC] Debug mode ${value ? "enabled" : "disabled"}`);
  };
}

export const DEFAULT_SETTINGS: PluginSettings = {
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
};

export const TARGET_OPTIONS: SelectOption[] = [
  { label: "Color cell", value: "cell" },
  { label: "Color row", value: "row" },
  { label: "Color column", value: "column" },
];

export const WHEN_OPTIONS: SelectOption[] = [
  { label: "The cell", value: "theCell" },
  { label: "Any cell", value: "anyCell" },
  { label: "All cell", value: "allCell" },
  { label: "No cell", value: "noCell" },
  { label: "First row", value: "firstRow" },
  { label: "Column header", value: "columnHeader" },
];

export const MATCH_OPTIONS: SelectOption[] = [
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

export const SORT_OPTIONS: SelectOption[] = [
  { label: "Sort: last added", value: "lastAdded" },
  { label: "Sort: A–Z", value: "az" },
  { label: "Sort: regex first", value: "regexFirst" },
  { label: "Sort: numbers first", value: "numbersFirst" },
  { label: "Sort: mode", value: "mode" },
];
