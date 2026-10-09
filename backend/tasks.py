"""One active task, with per-task cancellation and thread-safe snapshots."""
import logging
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime

logger = logging.getLogger("vghtpe")
ACTIVE_STATUSES = {"running", "cancelling"}


class TaskCancelled(Exception):
    pass


class TaskBusy(Exception):
    pass


@dataclass
class Task:
    id: str
    kind: str
    title: str
    total: int
    output_folder: str
    status: str = "running"
    processed: int = 0
    successful: int = 0
    skipped: int = 0
    failed: int = 0
    cancelled: int = 0
    fraction: float = 0
    current_item: str = ""
    message: str = "準備開始"
    logs: list = field(default_factory=list)
    cancel_event: threading.Event = field(default_factory=threading.Event)


class TaskContext:
    def __init__(self, manager, task):
        self.manager = manager
        self.task = task
        self.cancel_event = task.cancel_event

    def check_cancelled(self):
        if self.cancel_event.is_set():
            raise TaskCancelled()

    def log(self, message):
        with self.manager.lock:
            self.task.message = message
            self.task.logs.append({"time": datetime.now().strftime("%H:%M:%S"), "message": message})
            self.task.logs = self.task.logs[-200:]

    def item(self, name):
        self.check_cancelled()
        with self.manager.lock:
            self.task.current_item = name
            self.task.fraction = 0

    def progress(self, fraction):
        with self.manager.lock:
            self.task.fraction = max(0, min(0.99, fraction))

    def record(self, result):
        with self.manager.lock:
            attribute = {"downloaded": "successful", "extracted": "successful", "skipped": "skipped", "failed": "failed"}[result]
            setattr(self.task, attribute, getattr(self.task, attribute) + 1)
            self.task.processed += 1
            self.task.fraction = 0


class TaskManager:
    def __init__(self):
        self.lock = threading.RLock()
        self.task = None
        self.worker = None
        self.closed = False

    def start(self, kind, title, total, output_folder, work):
        with self.lock:
            if self.closed or (self.worker and self.worker.is_alive()):
                raise TaskBusy("目前已有任務執行中，請等待完成或取消結束")
            task = Task(uuid.uuid4().hex, kind, title, total, str(output_folder))
            self.task = task
            self.worker = threading.Thread(target=self._run, args=(task, work), daemon=True, name="report-task")
            self.worker.start()
            return self.snapshot()

    def _run(self, task, work):
        context = TaskContext(self, task)
        try:
            work(context)
            context.check_cancelled()
            with self.lock:
                task.status = "completed"
                task.fraction = 0
            context.log(f"任務完成：成功 {task.successful}、跳過 {task.skipped}、失敗 {task.failed}")
        except TaskCancelled:
            with self.lock:
                task.status = "cancelled"
                task.cancelled = max(0, task.total - task.processed)
                task.fraction = 0
            context.log("任務已取消")
        except Exception as error:
            logger.exception("Task %s failed", task.id)
            with self.lock:
                task.status = "failed"
                task.fraction = 0
            context.log(f"任務失敗：{error}")

    def snapshot(self):
        with self.lock:
            if self.task is None:
                return None
            task = self.task
            progress = 100.0 if task.status == "completed" else min(100, (task.processed + task.fraction) / max(1, task.total) * 100)
            return {
                **{key: getattr(task, key) for key in ("id", "kind", "title", "total", "output_folder", "status", "processed", "successful", "skipped", "failed", "cancelled", "current_item", "message")},
                "progress": round(progress, 1),
                "logs": [entry.copy() for entry in task.logs],
            }

    def cancel(self, task_id):
        with self.lock:
            if self.task is None or self.task.id != task_id:
                raise KeyError("找不到此任務")
            if self.task.status in ACTIVE_STATUSES:
                self.task.status = "cancelling"
                self.task.cancel_event.set()
            return self.snapshot()

    def shutdown(self):
        with self.lock:
            self.closed = True
            if self.task and self.task.status in ACTIVE_STATUSES:
                self.task.cancel_event.set()
                self.task.status = "cancelling"
            worker = self.worker
        if worker:
            worker.join(timeout=10)
