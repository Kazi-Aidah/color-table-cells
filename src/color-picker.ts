import { setIcon } from "obsidian";
import { hexToHsv, hsvToHex } from "./utils";
import type { PluginSettings } from "./types";

interface ColorPickerPlugin {
  settings: PluginSettings;
}

export class ColorPickerMenu {
  private plugin: ColorPickerPlugin;
  private onPick: (color: string) => void;
  private initialColor: string;
  private anchorEl: HTMLElement;

  hue: number;
  sat: number;
  val: number;
  color: string;
  menuEl: HTMLElement | null = null;

  // Support single cell or multiple cells (row/column)
  _cell?: HTMLElement;
  _cells: HTMLElement[] = [];
  _type?: string;

  // Event handler references for cleanup
  private _outsideHandler: ((e: MouseEvent) => void) | null = null;
  private _sbMoveHandler: ((e: MouseEvent) => void) | null = null;
  private _sbUpHandler: (() => void) | null = null;
  private _hueMoveHandler: ((e: MouseEvent) => void) | null = null;
  private _hueUpHandler: (() => void) | null = null;

  constructor(
    plugin: ColorPickerPlugin,
    onPick: (color: string) => void,
    initialColor: string | null | undefined,
    anchorEl: HTMLElement,
  ) {
    this.plugin = plugin;
    this.onPick = onPick;
    this.initialColor = initialColor || "#FFA500";
    this.anchorEl = anchorEl;
    const { h, s, v } = hexToHsv(this.initialColor);
    this.hue = h;
    this.sat = s;
    this.val = v;
    this.color = hsvToHex(this.hue, this.sat, this.val);
  }

  open(): void {
    this.close();
    this.menuEl = document.createElement("div");
    this.menuEl.className = "ctc-color-picker-menu";

    const sbBox = document.createElement("canvas") as HTMLCanvasElement;
    sbBox.width = 210;
    sbBox.height = 120;
    sbBox.className = "ctc-color-picker-sb-box";
    this.menuEl.appendChild(sbBox);

    const sbCtx = sbBox.getContext("2d")!;
    const drawSB = (hue: number) => {
      for (let x = 0; x < sbBox.width; x++) {
        for (let y = 0; y < sbBox.height; y++) {
          const s = x / (sbBox.width - 1);
          const v = 1 - y / (sbBox.height - 1);
          sbCtx.fillStyle = hsvToHex(hue, s, v);
          sbCtx.fillRect(x, y, 1, 1);
        }
      }
    };
    drawSB(this.hue);

    const sbSelector = document.createElement("div");
    sbSelector.className = "ctc-color-picker-sb-selector";
    this.menuEl.appendChild(sbSelector);

    const hueBox = document.createElement("canvas") as HTMLCanvasElement;
    hueBox.width = 210;
    hueBox.height = 14;
    hueBox.className = "ctc-color-picker-hue-box";
    this.menuEl.appendChild(hueBox);

    const hueCtx = hueBox.getContext("2d")!;
    const grad = hueCtx.createLinearGradient(0, 0, hueBox.width, 0);
    for (let i = 0; i <= 360; i++) {
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
    let hexAnimFrame: number | null = null;

    hexInput.addEventListener("input", () => {
      hexInput.value = hexInput.value.replace(/[^0-9a-fA-F]/g, "").slice(0, 6);
      if (hexInput.value.length === 6) {
        const hex = "#" + hexInput.value;
        const { h, s, v } = hexToHsv(hex);
        this.hue = h; this.sat = s; this.val = v;
        this.color = hex;
        preview.style.background = this.color;
        updateSelectors();
        drawSB(this.hue);
      }
    });
    hexRow.appendChild(hexInput);

    const pickBtn = document.createElement("button");
    pickBtn.className = "mod-ghost ctc-color-picker-icon-button";
    try {
      setIcon(pickBtn, "pipette");
    } catch {
      pickBtn.textContent = "🎨";
    }
    pickBtn.title = "Pick color from screen";
    pickBtn.onclick = async () => {
      try {
        if ("EyeDropper" in window) {
          const eye = new (window as Window & { EyeDropper: new () => { open: () => Promise<{ sRGBHex: string }> } }).EyeDropper();
          const result = await eye.open();
          const hex = result?.sRGBHex?.toUpperCase() ?? null;
          if (hex) {
            const { h, s, v } = hexToHsv(hex);
            this.hue = h; this.sat = s; this.val = v;
            this.color = hex;
            preview.style.background = this.color;
            drawSB(this.hue);
            updateSelectors();
          }
        }
      } catch { /* user cancel */ }
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
        const { h, s, v } = hexToHsv(rc);
        this.hue = h; this.sat = s; this.val = v;
        this.color = rc;
        preview.style.background = this.color;
        drawSB(this.hue);
        updateSelectors();
      });
      recentRow.appendChild(sw);
    });
    this.menuEl.appendChild(recentRow);

    const applyColorToPreviewCells = () => {
      if (this._cell && this._type) {
        if (this._type === "bg") this._cell.style.backgroundColor = this.color;
        else this._cell.style.color = this.color;
      }
      if (this._cells.length > 0 && this._type) {
        this._cells.forEach((cell) => {
          if (this._type === "bg") cell.style.backgroundColor = this.color;
          else cell.style.color = this.color;
        });
      }
    };

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
              if (i < target.length) hexAnimFrame = requestAnimationFrame(animate);
            }
          };
          animate();
        }
      }
      applyColorToPreviewCells();
    };

    let sbDragging = false;
    const handleSB = (e: MouseEvent) => {
      const rect = sbBox.getBoundingClientRect();
      let x = e.clientX - rect.left;
      let y = e.clientY - rect.top;
      x = Math.max(0, Math.min(sbBox.width - 1, x));
      y = Math.max(0, Math.min(sbBox.height - 1, y));
      this.sat = x / (sbBox.width - 1);
      this.val = 1 - y / (sbBox.height - 1);
      this.color = hsvToHex(this.hue, this.sat, this.val);
      preview.style.background = this.color;
      updateSelectors();
    };
    sbBox.addEventListener("mousedown", (e) => { sbDragging = true; handleSB(e); });

    let hueDragging = false;
    let huePending = false;
    let hueX = 0;
    const updateHueSelector = () => {
      this.hue = (hueX / (hueBox.width - 1)) * 360;
      this.color = hsvToHex(this.hue, this.sat, this.val);
      preview.style.background = this.color;
      hueSelector.style.left = hueBox.offsetLeft + hueX - 6 + "px";
      hueSelector.style.top = hueBox.offsetTop + 1 + "px";
      applyColorToPreviewCells();
      drawSB(this.hue);
    };
    const onHueMove = (e: MouseEvent) => {
      const rect = hueBox.getBoundingClientRect();
      hueX = Math.max(0, Math.min(hueBox.width - 1, e.clientX - rect.left));
      if (!huePending) {
        huePending = true;
        requestAnimationFrame(() => { updateHueSelector(); huePending = false; });
      }
    };
    hueBox.addEventListener("mousedown", (e) => { hueDragging = true; onHueMove(e); });

    this._sbMoveHandler = (e) => { if (sbDragging) handleSB(e); };
    this._sbUpHandler = () => { sbDragging = false; };
    this._hueMoveHandler = (e) => { if (hueDragging) onHueMove(e); };
    this._hueUpHandler = () => { hueDragging = false; };

    window.addEventListener("mousemove", this._sbMoveHandler);
    window.addEventListener("mouseup", this._sbUpHandler);
    window.addEventListener("mousemove", this._hueMoveHandler);
    window.addEventListener("mouseup", this._hueUpHandler);

    hexInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") this.close();
    });

    window.setTimeout(() => {
      const rect = this.anchorEl.getBoundingClientRect();
      let left = rect.left;
      let top = rect.bottom + 4;
      if (left + 280 > window.innerWidth) left = window.innerWidth - 280;
      if (top + 260 > window.innerHeight) top = rect.top - 260;
      this.menuEl!.style.left = left + "px";
      this.menuEl!.style.top = top + "px";
    }, 0);

    document.body.appendChild(this.menuEl);

    this._outsideHandler = (evt: MouseEvent) => {
      if (this.menuEl && !this.menuEl.contains(evt.target as Node)) this.close();
    };
    window.setTimeout(
      () => document.addEventListener("mousedown", this._outsideHandler!),
      10,
    );

    window.setTimeout(updateSelectors, 0);
  }

  close(): void {
    try {
      if (this.menuEl?.parentNode) {
        this.menuEl.parentNode.removeChild(this.menuEl);
        this.menuEl = null;
      }
    } catch { /* ignore */ }

    try {
      if (this._outsideHandler) document.removeEventListener("mousedown", this._outsideHandler);
      if (this._sbMoveHandler) window.removeEventListener("mousemove", this._sbMoveHandler);
      if (this._sbUpHandler) window.removeEventListener("mouseup", this._sbUpHandler);
      if (this._hueMoveHandler) window.removeEventListener("mousemove", this._hueMoveHandler);
      if (this._hueUpHandler) window.removeEventListener("mouseup", this._hueUpHandler);
    } catch { /* ignore */ }

    this._outsideHandler = null;
    this._sbMoveHandler = null;
    this._sbUpHandler = null;
    this._hueMoveHandler = null;
    this._hueUpHandler = null;

    try {
      if (typeof this.onPick === "function") this.onPick(this.color);
    } catch { /* ignore */ }
  }

  destroy(): void {
    this.close();
  }
}
