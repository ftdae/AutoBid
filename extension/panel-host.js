(() => {
  const HOST_ID = "auto-bid-extension-panel-host";
  const MESSAGE_SOURCE = "auto-bid-panel";
  const SIZE_STORAGE_KEY = "autoBidPanelSize";
  const DEFAULT_SIZE = { width: 460, height: 680 };
  const MIN_SIZE = { width: 360, height: 420 };
  const EDGE_GAP = 16;

  const existing = document.getElementById(HOST_ID);
  if (existing) {
    existing.remove();
    return;
  }

  const host = document.createElement("div");
  host.id = HOST_ID;
  host.style.all = "initial";
  host.style.display = "block";
  host.style.position = "fixed";
  host.style.top = `${EDGE_GAP}px`;
  host.style.right = `${EDGE_GAP}px`;
  host.style.width = `${DEFAULT_SIZE.width}px`;
  host.style.height = `${DEFAULT_SIZE.height}px`;
  host.style.maxWidth = `calc(100vw - ${EDGE_GAP * 2}px)`;
  host.style.maxHeight = `calc(100vh - ${EDGE_GAP * 2}px)`;
  host.style.pointerEvents = "auto";
  host.style.contain = "layout style";
  host.style.zIndex = "2147483647";
  (document.body || document.documentElement).append(host);

  const shadow = host.attachShadow({ mode: "closed" });
  const style = document.createElement("style");
  style.textContent = `
    :host {
      all: initial;
      display: block;
      width: 100%;
      height: 100%;
    }

    .panel {
      position: relative;
      width: 100%;
      height: 100%;
      min-width: ${MIN_SIZE.width}px;
      min-height: ${MIN_SIZE.height}px;
      overflow: hidden;
      border: 1px solid rgba(22, 35, 30, .18);
      border-radius: 10px;
      background: #f6f8f7;
      box-shadow: 0 16px 48px rgba(12, 24, 18, .26), 0 2px 10px rgba(12, 24, 18, .18);
      color: #18211f;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    iframe {
      display: block;
      width: 100%;
      height: 100%;
      border: 0;
      background: #f6f8f7;
    }

    .handle {
      position: absolute;
      z-index: 2;
      background: transparent;
    }

    .handle-left {
      top: 0;
      bottom: 18px;
      left: 0;
      width: 10px;
      cursor: ew-resize;
    }

    .handle-bottom {
      right: 18px;
      bottom: 0;
      left: 10px;
      height: 10px;
      cursor: ns-resize;
    }

    .handle-corner {
      right: 0;
      bottom: 0;
      width: 22px;
      height: 22px;
      cursor: nesw-resize;
    }

    .handle-corner::after {
      content: "";
      position: absolute;
      right: 6px;
      bottom: 6px;
      width: 10px;
      height: 10px;
      border-right: 2px solid rgba(29, 107, 82, .55);
      border-bottom: 2px solid rgba(29, 107, 82, .55);
    }
  `;

  const panel = document.createElement("div");
  panel.className = "panel";

  const iframe = document.createElement("iframe");
  iframe.title = "Auto Bid";
  iframe.src = chrome.runtime.getURL("popup.html?surface=panel");

  const leftHandle = document.createElement("div");
  leftHandle.className = "handle handle-left";
  leftHandle.title = "Resize width";

  const bottomHandle = document.createElement("div");
  bottomHandle.className = "handle handle-bottom";
  bottomHandle.title = "Resize height";

  const cornerHandle = document.createElement("div");
  cornerHandle.className = "handle handle-corner";
  cornerHandle.title = "Resize panel";

  panel.append(iframe, leftHandle, bottomHandle, cornerHandle);
  shadow.append(style, panel);
  console.info("[AutoBid] panel:mounted", { iframe: iframe.src });

  loadSize().then((size) => applySize(panel, size));
  installResizeHandle(panel, leftHandle, "horizontal");
  installResizeHandle(panel, bottomHandle, "vertical");
  installResizeHandle(panel, cornerHandle, "both");

  window.addEventListener("message", (event) => {
    if (event.source !== iframe.contentWindow) return;
    const data = event.data || {};
    if (data.source === MESSAGE_SOURCE && data.type === "CLOSE_PANEL") {
      host.remove();
      return;
    }
    if (data.source === MESSAGE_SOURCE && data.type === "RELOAD_PANEL") {
      const url = new URL(iframe.src);
      url.searchParams.set("recovered", String(Date.now()));
      iframe.src = url.href;
    }
  });

  function installResizeHandle(panelElement, handle, mode) {
    handle.addEventListener("pointerdown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      handle.setPointerCapture?.(event.pointerId);

      const startX = event.clientX;
      const startY = event.clientY;
      const startWidth = panelElement.getBoundingClientRect().width;
      const startHeight = panelElement.getBoundingClientRect().height;

      const move = (moveEvent) => {
        const next = {
          width: startWidth,
          height: startHeight
        };

        if (mode === "horizontal" || mode === "both") {
          next.width = startWidth + (startX - moveEvent.clientX);
        }
        if (mode === "vertical" || mode === "both") {
          next.height = startHeight + (moveEvent.clientY - startY);
        }

        applySize(panelElement, next);
      };

      const up = () => {
        window.removeEventListener("pointermove", move, true);
        window.removeEventListener("pointerup", up, true);
        saveSize(readSize(panelElement));
      };

      window.addEventListener("pointermove", move, true);
      window.addEventListener("pointerup", up, true);
    });
  }

  function applySize(panelElement, size) {
    const next = clampSize(size);
    host.style.width = `${next.width}px`;
    host.style.height = `${next.height}px`;
    panelElement.style.width = "100%";
    panelElement.style.height = "100%";
  }

  function readSize(panelElement) {
    const rect = panelElement.getBoundingClientRect();
    return {
      width: rect.width,
      height: rect.height
    };
  }

  function clampSize(size) {
    const maxWidth = Math.max(MIN_SIZE.width, window.innerWidth - EDGE_GAP * 2);
    const maxHeight = Math.max(MIN_SIZE.height, window.innerHeight - EDGE_GAP * 2);
    return {
      width: Math.round(Math.min(Math.max(Number(size.width) || DEFAULT_SIZE.width, MIN_SIZE.width), maxWidth)),
      height: Math.round(Math.min(Math.max(Number(size.height) || DEFAULT_SIZE.height, MIN_SIZE.height), maxHeight))
    };
  }

  async function loadSize() {
    try {
      const data = await chrome.storage.local.get([SIZE_STORAGE_KEY]);
      return data[SIZE_STORAGE_KEY] || DEFAULT_SIZE;
    } catch (_error) {
      return DEFAULT_SIZE;
    }
  }

  async function saveSize(size) {
    try {
      await chrome.storage.local.set({ [SIZE_STORAGE_KEY]: clampSize(size) });
    } catch (_error) {
      // The panel still works if size persistence is unavailable.
    }
  }
})();
