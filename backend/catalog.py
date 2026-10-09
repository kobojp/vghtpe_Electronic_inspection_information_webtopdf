"""Editable external report catalog with revision checks and atomic backups."""
import hashlib
import json
import os
import re
import tempfile
import threading
from pathlib import Path

from backend.reports import REPORT_TYPE_MAP, catalog_reports


class CatalogConflict(ValueError):
    pass


def validate_entry(entry):
    name = entry.get("name")
    if not isinstance(name, str) or not name.strip() or name != name.strip() or len(name) > 180:
        raise ValueError("報表名稱須為 1 至 180 字，前後不可包含空白")
    if any(ord(char) < 32 or char in '<>:"/\\|?*' for char in name) or name.endswith((".", " ")):
        raise ValueError("報表名稱包含 Windows 檔名不允許的字元")
    stem = name.split(".")[0].upper()
    if stem in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        raise ValueError("報表名稱不可使用 Windows 保留檔名")
    if not isinstance(entry.get("api_1"), str) or not re.fullmatch(r"/[0-9]+/", entry["api_1"]):
        raise ValueError("網址參數一格式須為 /數字/，例如 /16/")
    # One original entry omits the leading slash. Preserve that supported format.
    if not isinstance(entry.get("api_2"), str) or not re.fullmatch(r"/?[0-9]+/[0-9]+", entry["api_2"]):
        raise ValueError("網址參數二格式須為 /數字/數字，例如 /16/32")


def atomic_write(path, content):
    fd, temporary = tempfile.mkstemp(prefix="catalog-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


class CatalogStore:
    def __init__(self, path):
        self.path = Path(path).resolve()
        self.backup_path = self.path.with_name(self.path.name + ".bak")
        self.lock = threading.RLock()

    def _read(self):
        content = self.path.read_bytes()
        data = json.loads(content.decode("utf-8-sig"))
        reports = catalog_reports(data)
        return content, data, reports, hashlib.sha256(content).hexdigest()

    def _snapshot(self, reports, revision):
        return {
            "reports": reports,
            "revision": revision,
            "data_path": str(self.path),
            "backup_path": str(self.backup_path),
        }

    def snapshot(self):
        with self.lock:
            _, _, reports, revision = self._read()
            return self._snapshot(reports, revision)

    def mutate(self, revision, entry=None, report_id=None, delete=False):
        with self.lock:
            before, data, reports, current = self._read()
            if revision != current:
                raise CatalogConflict("報表資料已變更，請重新載入清單後再操作")
            existing = next((report for report in reports if report["id"] == report_id), None)
            if report_id is not None and existing is None:
                raise KeyError("找不到此報表，請重新載入清單")
            if not delete:
                if entry["report_type"] not in REPORT_TYPE_MAP:
                    raise ValueError("請選擇現有的報表類型")
                validate_entry(entry)
                category = entry["report_type"][:2]
                if any(report["id"] != report_id and report["category"] == category
                       and report["name"].casefold() == entry["name"].casefold() for report in reports):
                    raise ValueError("同一大類已有相同報表名稱，請使用不同名稱，避免下載檔案互相覆蓋")
            if existing:
                old_key, index = report_id.rsplit(":", 1)
                old_entry = data[old_key][int(index)]
                new_key = None if delete else REPORT_TYPE_MAP[entry["report_type"]]
                if new_key == old_key:
                    data[old_key][int(index)] = {**old_entry, **{key: entry[key] for key in ("name", "api_1", "api_2")}}
                else:
                    del data[old_key][int(index)]
                    if not delete:
                        data[new_key].append({**old_entry, **{key: entry[key] for key in ("name", "api_1", "api_2")}})
            else:
                key = REPORT_TYPE_MAP[entry["report_type"]]
                data[key].append({key: entry[key] for key in ("name", "api_1", "api_2")})
            updated = catalog_reports(data)
            content = (json.dumps(data, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
            # Save the exact previous bytes. Backup failure leaves data.json untouched.
            atomic_write(self.backup_path, before)
            if self.path.read_bytes() != before:
                raise CatalogConflict("報表檔案已被其他程式修改，請重新載入後再操作")
            atomic_write(self.path, content)
            return self._snapshot(updated, hashlib.sha256(content).hexdigest())
