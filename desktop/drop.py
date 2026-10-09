"""Translate native WebView2 file drops into React file-list events."""
import json
import logging
import os
import time
from pathlib import Path

from backend.reports import valid_pdf

logger = logging.getLogger("vghtpe")
DROP_EVENT = "vghtpe:pdf-drop"


def dropped_pdf_files(event):
    paths, rejected, seen = [], [], set()
    for item in event.get("dataTransfer", {}).get("files", []):
        raw = item.get("pywebviewFullPath")
        name = item.get("name") or "未命名檔案"
        if not isinstance(raw, str) or not raw:
            rejected.append(name)
            continue
        path = Path(raw)
        if not path.is_absolute() or path.suffix.lower() != ".pdf" or not valid_pdf(path):
            rejected.append(name)
            continue
        absolute = str(path.resolve())
        key = os.path.normcase(absolute)
        if key not in seen:
            seen.add(key)
            paths.append(absolute)
    return {"paths": paths, "rejected": rejected}


def bind_pdf_drop(window):
    from webview.dom import DOMEventHandler

    zone = window.dom.get_element("#pdf-drop-zone")
    if zone is None:
        raise RuntimeError("找不到 PDF 拖曳區域")
    if window.evaluate_js("Boolean(document.getElementById('pdf-drop-zone')?.dataset.nativeDropReady)"):
        return

    def on_drop(event):
        if not event.get("dataTransfer", {}).get("files"):
            return
        try:
            # The React listener checks busy/visibility again when the callback arrives.
            if not window.evaluate_js("""(() => {
                const zone = document.getElementById('pdf-drop-zone');
                return Boolean(zone && zone.offsetParent !== null && zone.getAttribute('aria-disabled') !== 'true');
            })()"""):
                return
            result = dropped_pdf_files(event)
        except Exception:
            logger.exception("Could not read dropped PDF files")
            result = {"paths": [], "rejected": [], "error": "無法讀取拖入的檔案，請改用「選擇 PDF」"}
        window.evaluate_js(
            f"window.dispatchEvent(new CustomEvent('{DROP_EVENT}', {{ detail: {json.dumps(result)} }}));"
        )

    zone.on("drop", DOMEventHandler(on_drop, prevent_default=True))
    window.evaluate_js("document.getElementById('pdf-drop-zone').dataset.nativeDropReady = 'true'")
    logger.info("Native PDF drop handler registered")


def bind_pdf_drop_when_ready(window):
    try:
        deadline = time.monotonic() + 15
        while time.monotonic() < deadline:
            if window.evaluate_js("Boolean(document.querySelector('[data-app-ready]'))"):
                bind_pdf_drop(window)
                return
            time.sleep(0.1)
        logger.error("Could not register native PDF drop: React did not become ready")
    except Exception:
        logger.exception("Could not register native PDF drop")
