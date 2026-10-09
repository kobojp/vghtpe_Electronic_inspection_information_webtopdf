"""Application paths, persistent settings and error logging."""
import json
import logging
import os
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path


def application_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def resource_dir():
    return Path(getattr(sys, "_MEIPASS", application_dir()))


def get_data_file_path():
    return str(application_dir() / "data.json")


def get_settings_path():
    state = os.environ.get("VGHTPE_STATE_DIR")
    base = Path(state) if state else Path(os.environ.get("APPDATA", Path.home())) / "VghtpeReportDownloader"
    return str(base / "settings.json")


def default_settings():
    return {
        "output_folder": str(application_dir() / "水電消防報表"),
        "extract_folder": str(application_dir() / "報表合併pdf"),
        "open_folder_after_completion": False,
    }


def load_settings(settings_path=None):
    settings = default_settings()
    try:
        loaded = json.loads(Path(settings_path or get_settings_path()).read_text(encoding="utf-8"))
        if not isinstance(loaded, dict):
            return settings
        for key in ("output_folder", "extract_folder"):
            if isinstance(loaded.get(key), str) and Path(loaded[key]).is_absolute():
                settings[key] = loaded[key]
        if isinstance(loaded.get("open_folder_after_completion"), bool):
            settings["open_folder_after_completion"] = loaded["open_folder_after_completion"]
    except (OSError, ValueError):
        pass
    return settings


def save_settings(settings, settings_path=None):
    path = Path(settings_path or get_settings_path())
    path.parent.mkdir(parents=True, exist_ok=True)
    # A failed write must not damage the previous settings file.
    fd, temp = tempfile.mkstemp(prefix="settings-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(settings, stream, ensure_ascii=False, indent=2)
        os.replace(temp, path)
    finally:
        Path(temp).unlink(missing_ok=True)


def configure_logging(settings_path=None):
    directory = Path(settings_path or get_settings_path()).parent / "logs"
    directory.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("vghtpe")
    logger.setLevel(logging.INFO)
    filename = directory / "error.log"
    if not any(getattr(handler, "baseFilename", None) == str(filename.resolve()) for handler in logger.handlers):
        handler = RotatingFileHandler(filename, maxBytes=2_000_000, backupCount=2, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
    return logger


def open_folder(folder):
    path = Path(folder).resolve()
    path.mkdir(parents=True, exist_ok=True)
    os.startfile(str(path))
