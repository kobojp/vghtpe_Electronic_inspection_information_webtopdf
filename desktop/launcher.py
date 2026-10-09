"""Start the loopback API and a WebView2 desktop window."""
import argparse
import ctypes
import logging
import socket
import sys
import threading
import time
from pathlib import Path

import uvicorn

from backend import config
from backend.app import create_app
from desktop.drop import bind_pdf_drop_when_ready


class NativeDialogs:
    def __init__(self):
        self.window = None

    def select_pdfs(self):
        import webview
        return list(self.window.create_file_dialog(webview.FileDialog.OPEN, allow_multiple=True,
                    file_types=("PDF 檔案 (*.pdf)",)) or [])

    def select_folder(self):
        import webview
        paths = self.window.create_file_dialog(webview.FileDialog.FOLDER)
        return paths[0] if paths else None


def main():
    parser = argparse.ArgumentParser(description="臺北榮總水電消防報表桌面程式")
    parser.add_argument("--dev", action="store_true", help="載入 http://127.0.0.1:5173 的 Vite 開發介面")
    parser.add_argument("--headless", action="store_true", help="僅啟動 API，供開發與整合驗證")
    parser.add_argument("--port", type=int, default=None, help="開發驗證使用的本機連接埠")
    parser.add_argument("--smoke-test", action="store_true", help="以隱藏視窗驗證桌面啟動後自動結束")
    args = parser.parse_args()
    logger = config.configure_logging()
    if not args.dev:
        directory = config.resource_dir() / ("frontend_dist" if getattr(sys, "frozen", False) else "frontend/dist")
        if not (directory / "index.html").is_file():
            raise RuntimeError("找不到介面資源。請先在 frontend 執行 npm run build，再啟動程式。")
    dialogs = None if args.headless else NativeDialogs()
    app = create_app(dialogs=dialogs, dev=args.dev)
    server = uvicorn.Server(uvicorn.Config(app, host="127.0.0.1", log_config=None, access_log=False))
    listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    listener.bind(("127.0.0.1", args.port if args.port is not None else 8765 if args.dev else 0))
    listener.listen(128)
    port = listener.getsockname()[1]
    thread = threading.Thread(target=server.run, kwargs={"sockets": [listener]}, daemon=True, name="local-api")
    thread.start()
    smoke_result = {"ok": False}
    try:
        deadline = time.monotonic() + 20
        while not server.started:
            if not thread.is_alive() or time.monotonic() > deadline:
                raise RuntimeError("無法啟動本機服務，請查看 logs/error.log")
            time.sleep(0.05)
        logger.info("Local API started on port %s", port)
        if args.headless:
            if args.smoke_test:
                import urllib.request
                import json
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/bootstrap", timeout=5) as response:
                    catalog = json.load(response)
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=5) as response:
                    html = response.read().decode("utf-8")
                if not catalog["reports"] or "/assets/" not in html:
                    raise RuntimeError("打包資源或報表目錄驗證失敗")
                logger.info("Headless smoke test passed: %s reports and frontend assets", len(catalog["reports"]))
                return
            while thread.is_alive():
                time.sleep(0.2)
            return
        import webview
        if args.smoke_test:
            native_logger = logging.getLogger("pywebview")
            native_logger.setLevel(logging.DEBUG)
            for handler in logger.handlers:
                native_logger.addHandler(handler)
        url = "http://127.0.0.1:5173" if args.dev else f"http://127.0.0.1:{port}"
        dialogs.window = webview.create_window("臺北榮總｜水電消防報表", url, width=1280, height=900,
                                              min_size=(960, 720), hidden=args.smoke_test)
        dialogs.window.events.loaded += bind_pdf_drop_when_ready
        def inspect_window():
            try:
                if not dialogs.window.events.loaded.wait(30):
                    logger.error("WebView2 did not finish loading within 30 seconds")
                    return
                deadline = time.monotonic() + 15
                while time.monotonic() < deadline:
                    if dialogs.window.evaluate_js("Boolean(document.querySelector('[data-app-ready]'))"):
                        smoke_result["ok"] = True
                        logger.info("Desktop smoke test passed")
                        return
                    time.sleep(0.2)
            except Exception:
                logger.exception("Desktop smoke test failed")
            finally:
                if args.smoke_test or not smoke_result["ok"]:
                    dialogs.window.destroy()
        # Require the modern Windows renderer used by React.
        webview.start(func=inspect_window, gui="edgechromium", debug=args.dev,
                      storage_path=str(Path(config.get_settings_path()).parent / "webview"),
                      icon=str(config.resource_dir() / "app.ico"))
    finally:
        app.state.runtime.tasks.shutdown()
        server.should_exit = True
        thread.join(timeout=10)
        listener.close()
    if not smoke_result["ok"]:
        raise RuntimeError("桌面介面啟動驗證失敗，請查看 logs/error.log")


def run():
    try:
        main()
    except KeyboardInterrupt:
        return
    except Exception as error:
        logging.getLogger("vghtpe").exception("Application startup failed")
        if getattr(sys, "frozen", False) and "--smoke-test" not in sys.argv and "--headless" not in sys.argv:
            ctypes.windll.user32.MessageBoxW(None, f"程式無法啟動：{error}\n\n請確認 EXE 同層有 data.json，且已安裝 WebView2 Runtime。", "啟動失敗", 0x10)
        raise
