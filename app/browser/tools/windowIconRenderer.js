"use strict";

const { nativeImage } = require("electron");
const TrayIconChooser = require("./trayIconChooser");
const {
  drawPresenceDot,
  isKnownPresence,
  wantsWindowStatus,
} = require("./presenceColors");

// The dot is laid out on the same 140-pixel grid as the tray overlay
// (trayIconRenderer.js) and scaled to the window icon, so both read alike.
// Top-right mirrors the bottom-right dot across the middle of the icon.
const GRID = 140;
const DOT = {
  "bottom-right": { x: 108, y: 104, radius: 24 },
  "top-right": { x: 108, y: GRID - 104, radius: 24 },
};

// X11 drops a window icon that does not fit in one request, which is why main
// caps it at 128 pixels (TrayIconChooser.windowImage); the composite keeps the
// same cap so the dotted icon is never larger than the plain one.
const MAX_SIZE = 128;

// Presence overlay on the main window icon (window.showStatus), which is what
// the taskbar or panel entry shows on Linux and Windows. Main has no canvas,
// so the icon is composited here, like the tray's, and sent on
// window-icon-update; null tells main to restore the plain icon.
class WindowIconRenderer {
  // Private: anything on `this` is reachable from page script, including via
  // the listeners below (GHSA-58j7-22f8-6xqp).
  #ipcRenderer = null;
  #appIcon = "";
  #appIconType = "default";
  #position = "bottom-right";
  #statusCode = null;
  #baseIconDataUrl = null;
  #size = 0;
  #rendered = new Map();
  #sequence = 0;

  init(config, ipcRenderer) {
    if (!wantsWindowStatus(config)) {
      return;
    }
    this.#ipcRenderer = ipcRenderer;
    this.#appIcon = config.appIcon ?? "";
    this.#appIconType = config.appIconType ?? "default";
    this.#position = config.window?.statusPosition === "top-right"
      ? "top-right"
      : "bottom-right";

    // The presence detector (mqttStatusMonitor.js) dispatches this event
    // whenever it reads a different status off the Teams UI.
    globalThis.addEventListener("user-status-changed-local", (event) => {
      void this.#updateStatus(event.detail?.status);
    });

    // The App Icon menu sets the window back to the plain icon every time,
    // even when the same file is picked again or Reset finds the default
    // already in place, and then sends this delta. So the dot is sent again
    // whenever the delta carries appIcon, and only a different path throws
    // the cached composites away. Read the delta, not config: preload's own
    // config-changed listener may not have run yet.
    ipcRenderer.on("config-changed", (_event, changes) => {
      if (!("appIcon" in changes)) return;
      const appIcon = changes.appIcon ?? "";
      if (appIcon !== this.#appIcon) {
        this.#appIcon = appIcon;
        this.#baseIconDataUrl = null;
        this.#rendered.clear();
      }
      if (isKnownPresence(this.#statusCode)) {
        void this.#send();
      }
    });

    // Main keeps the last icon across page loads, and the monitor reports
    // only statuses it finds, so a sign-in page or a reload would otherwise
    // leave the previous dot in place. Every page starts from the plain icon.
    void this.#send();
  }

  #updateStatus(status) {
    if (!Number.isInteger(status) || status === this.#statusCode) {
      return;
    }
    this.#statusCode = status;
    return this.#send();
  }

  async #send() {
    // A status change that lands while an earlier one is still rendering wins;
    // the earlier result is dropped, so the icon never steps back.
    const sequence = ++this.#sequence;
    const status = this.#statusCode;
    let icon = null;
    if (isKnownPresence(status)) {
      try {
        icon = await this.render(status);
      } catch (error) {
        console.error("[WINDOW_ICON] Render failed, using the plain icon", {
          message: error.message,
        });
        icon = null;
      }
    }
    if (sequence !== this.#sequence) return;
    this.#ipcRenderer.send("window-icon-update", icon);
  }

  #loadBaseIcon() {
    const chooser = new TrayIconChooser({
      appIcon: this.#appIcon,
      appIconType: this.#appIconType,
    });
    const image = nativeImage.createFromPath(chooser.getFile());
    const { width, height } = image.getSize();
    this.#size = Math.min(Math.max(width, height), MAX_SIZE);
    const dataUrl = image.toDataURL();
    this.#baseIconDataUrl = !dataUrl || dataUrl === "data:," || this.#size === 0
      ? null
      : dataUrl;
  }

  // Resolves the composited icon as a PNG data URL, or null when there is no
  // usable base icon (main then keeps the plain one). Each status is drawn
  // once per base icon.
  render(status) {
    if (this.#rendered.has(status)) {
      return Promise.resolve(this.#rendered.get(status));
    }
    if (this.#baseIconDataUrl === null) {
      this.#loadBaseIcon();
    }
    const baseIcon = this.#baseIconDataUrl;
    if (baseIcon === null) {
      return Promise.resolve(null);
    }
    const size = this.#size;
    return new Promise((resolve, reject) => {
      const image = new Image();
      image.onerror = () => resolve(null);
      image.onload = () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = size;
          canvas.height = size;
          const ctx = canvas.getContext("2d");
          ctx.drawImage(image, 0, 0, size, size);
          ctx.scale(size / GRID, size / GRID);
          drawPresenceDot(ctx, status, DOT[this.#position]);
          const dataUrl = canvas.toDataURL("image/png");
          // An App Icon change while this was drawing clears the cache; don't
          // put a composite of the old icon back into it.
          if (this.#baseIconDataUrl === baseIcon) {
            this.#rendered.set(status, dataUrl);
          }
          resolve(dataUrl);
        } catch (error) {
          reject(error);
        }
      };
      image.src = baseIcon;
    });
  }
}

module.exports = new WindowIconRenderer();
