# batch_scheduler.py
import json
import os
import threading
import time
from datetime import datetime


SCHEDULE_DIR = os.path.join(os.path.expanduser("~"), ".excelai", "schedules")


class BatchJob:
    def __init__(
        self,
        name: str,
        commands: list[dict],
        filepath: str = "",
        schedule_time: str = "",
        repeat: bool = False,
    ):
        self.name = name
        self.commands = commands
        self.filepath = filepath
        self.schedule_time = schedule_time
        self.repeat = repeat
        self.created_at = datetime.now().isoformat()
        self.last_run = None
        self.run_count = 0

    def to_dict(self):
        return {
            "name": self.name,
            "commands": self.commands,
            "filepath": self.filepath,
            "schedule_time": self.schedule_time,
            "repeat": self.repeat,
            "created_at": self.created_at,
            "last_run": self.last_run,
            "run_count": self.run_count,
        }

    @staticmethod
    def from_dict(data: dict):
        job = BatchJob(
            name=data.get("name", "Untitled"),
            commands=data.get("commands", []),
            filepath=data.get("filepath", ""),
            schedule_time=data.get("schedule_time", ""),
            repeat=data.get("repeat", False),
        )
        job.created_at = data.get("created_at", "")
        job.last_run = data.get("last_run", None)
        job.run_count = data.get("run_count", 0)
        return job


class BatchScheduler:
    def __init__(self):
        self.jobs: list[BatchJob] = []
        self._running = False
        self._thread = None
        os.makedirs(SCHEDULE_DIR, exist_ok=True)
        self._load_jobs()

    def add_job(self, job: BatchJob):
        self.jobs.append(job)
        self._save_jobs()

    def remove_job(self, index: int):
        if 0 <= index < len(self.jobs):
            self.jobs.pop(index)
            self._save_jobs()

    def get_jobs(self) -> list[BatchJob]:
        return self.jobs

    def execute_job(self, job: BatchJob, excel_controller) -> tuple[bool, str]:
        from excel_controller import ExcelController

        try:
            if job.filepath and not excel_controller.is_connected():
                ok, msg = excel_controller.connect_or_open(filepath=job.filepath)
                if not ok:
                    return False, f"Could not open file: {msg}"

            results = []
            for cmd_item in job.commands:
                code = cmd_item.get("code", "")
                if code:
                    success, result = excel_controller.execute(code)
                    results.append(
                        f"{'OK' if success else 'FAIL'}: {cmd_item.get('command', 'unknown')}"
                    )

            job.last_run = datetime.now().isoformat()
            job.run_count += 1
            self._save_jobs()
            return True, "\n".join(results)
        except Exception as e:
            return False, str(e)

    def start_scheduler(self, excel_controller, callback=None):
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(
            target=self._scheduler_loop,
            args=(excel_controller, callback),
            daemon=True,
        )
        self._thread.start()

    def stop_scheduler(self):
        self._running = False

    def is_running(self):
        return self._running

    def _scheduler_loop(self, excel_controller, callback):
        while self._running:
            now = datetime.now()
            now_str = now.strftime("%H:%M")
            for job in self.jobs[:]:
                if job.schedule_time and job.schedule_time == now_str:
                    ok, msg = self.execute_job(job, excel_controller)
                    if callback:
                        callback(job, ok, msg)
                    if not job.repeat:
                        job.schedule_time = ""
                        self._save_jobs()
            time.sleep(30)

    def _save_jobs(self):
        try:
            data = [job.to_dict() for job in self.jobs]
            path = os.path.join(SCHEDULE_DIR, "jobs.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _load_jobs(self):
        try:
            path = os.path.join(SCHEDULE_DIR, "jobs.json")
            if not os.path.exists(path):
                return
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.jobs = [BatchJob.from_dict(d) for d in data]
        except Exception:
            self.jobs = []
