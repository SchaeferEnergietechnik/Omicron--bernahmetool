"""Windows worker for unattended Omicron OCC exports and Excel refreshes.

The worker deliberately has no interactive prompts. The GUI must resolve all
conflicts before it writes the job file. Progress is emitted as JSON Lines to
stdout so a separate GUI process can display it and request cancellation.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError, find_windows
import win32com.client


WINDOW_TITLE_RE = ".*OMICRON Control Center.*"
EXPORT_DIALOG_RE = ".*Datenexport.*"
EXCEL_SHEET = "Prüfprotokoll"
MACRO_PROTOCOL_NO = "Tabelle1.Protokollnummer_generieren_unsichtbar"
MACRO_TOGGLE_SECTIONS_CANDIDATES = [
    "BereicheEinOderAusblenden_Start",
    "Modul1.BereicheEinOderAusblenden_Start",
]
MACRO_HIDE_EMPTY_ROWS = "Tabelle7.ZeilenAusblendenWennLeer"
DEFAULT_EXPORT_DIRECTORY = Path(r"C:\Omicron_Datenexport")


class CancellationRequested(Exception):
    """Raised at safe boundaries after the GUI writes the cancellation file."""


class Worker:
    def __init__(self, job: dict[str, Any], cancel_file: Path | None):
        self.job = job
        self.cancel_file = cancel_file
        self.run_started_monotonic: float | None = None
        self.failures: list[dict[str, str]] = []
        self.skipped: list[dict[str, str]] = []

    def emit(self, event: str, **payload: Any) -> None:
        print(json.dumps({"timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"), "event": event, **payload}, ensure_ascii=False), flush=True)

    def check_cancelled(self) -> None:
        if self.cancel_file and self.cancel_file.exists():
            raise CancellationRequested()

    def wait(self, seconds: float) -> None:
        end_time = time.monotonic() + seconds
        while time.monotonic() < end_time:
            self.check_cancelled()
            time.sleep(min(0.5, max(0, end_time - time.monotonic())))

    def wait_for_main_window(self, timeout: int = 45):
        end_time = time.monotonic() + timeout
        last_error: Exception | None = None
        while time.monotonic() < end_time:
            self.check_cancelled()
            try:
                windows = find_windows(title_re=WINDOW_TITLE_RE)
                if windows:
                    app = Application(backend="uia").connect(handle=windows[0])
                    window = app.window(handle=windows[0])
                    window.wait("visible", timeout=5)
                    window.set_focus()
                    return window
            except Exception as error:  # pywinauto errors vary by Omicron version
                last_error = error
            self.wait(1)
        raise ElementNotFoundError(f"Kein OMICRON-Fenster gefunden: {last_error}")

    def open_export_dialog(self, window) -> bool:
        try:
            window.menu_select("Datei->Daten exportieren...")
            self.wait(1)
            return True
        except Exception:
            pass
        for candidate in (
            {"title": "Datei", "control_type": "TabItem"},
            {"title": "Datei", "control_type": "Button"},
            {"title": "Datei"},
        ):
            try:
                control = window.child_window(**candidate)
                if control.exists(timeout=2):
                    control.click_input()
                    self.wait(1)
                    break
            except Exception:
                continue
        else:
            return False
        for candidate in (
            {"title": "Daten exportieren...", "control_type": "MenuItem"},
            {"title": "Daten exportieren...", "control_type": "ListItem"},
            {"title": "Daten exportieren..."},
            {"title": "Daten exportieren"},
        ):
            try:
                control = window.child_window(**candidate)
                if control.exists(timeout=2):
                    control.click_input()
                    self.wait(1)
                    return True
            except Exception:
                continue
        return False

    def confirm_export_dialog(self, timeout: int = 30) -> bool:
        end_time = time.monotonic() + timeout
        while time.monotonic() < end_time:
            self.check_cancelled()
            try:
                windows = find_windows(title_re=EXPORT_DIALOG_RE)
                if windows:
                    app = Application(backend="uia").connect(handle=windows[0])
                    dialog = app.window(handle=windows[0])
                    dialog.wait("visible", timeout=5)
                    dialog.set_focus()
                    for title in ("OK", "Ok", "ok"):
                        button = dialog.child_window(title=title, control_type="Button")
                        if button.exists(timeout=1):
                            button.click_input()
                            self.wait(1)
                            return True
                    dialog.type_keys("{ENTER}")
                    self.wait(1)
                    return True
            except Exception:
                pass
            self.wait(1)
        return False

    def close_omicron(self, window) -> None:
        try:
            window.close()
            self.wait(2)
        except CancellationRequested:
            raise
        except Exception:
            return

        # Closing an OCC without changes can still show a save question. Never wait
        # for a user: discard instead, otherwise record a later file failure.
        try:
            for handle in find_windows(title_re=WINDOW_TITLE_RE):
                dialog = Application(backend="uia").connect(handle=handle).window(handle=handle)
                text = " ".join(item.window_text() for item in dialog.descendants() if item.window_text()).lower()
                if "speichern" in text:
                    for title in ("Nein", "NEIN", "No"):
                        button = dialog.child_window(title=title, control_type="Button")
                        if button.exists(timeout=1):
                            button.click_input()
                            return
        except Exception:
            return

    def terminate_mashup_loader(self) -> None:
        """Stoppt den Power-Query-Mashup-Loader vor einem neuen Verarbeitungslauf."""
        try:
            result = subprocess.run(
                ["taskkill", "/IM", "Microsoft.Mashup.Container.Loader.exe", "/F"],
                capture_output=True,
                text=True,
                shell=False,
            )
            if result.returncode == 0:
                self.emit("mashup_terminated")
            else:
                # Kein Fehlerfall: Prozess lief ggf. nicht.
                self.emit("mashup_not_running")
        except Exception as error:
            self.emit("mashup_termination_failed", message=str(error))

    def export_directory(self) -> Path:
        return Path(self.job.get("exportDirectory", DEFAULT_EXPORT_DIRECTORY))

    def csv_snapshot(self, directory: Path) -> dict[Path, tuple[int, int]]:
        try:
            return {
                file_path: (file_path.stat().st_mtime_ns, file_path.stat().st_size)
                for file_path in directory.glob("*.csv")
                if file_path.is_file()
            }
        except OSError:
            return {}

    def wait_for_export_output(
        self,
        directory: Path,
        before_export: dict[Path, tuple[int, int]],
        started_at_ns: int,
        timeout: int = 120,
    ) -> list[Path]:
        """Bestätigt den Export erst, wenn neue oder geänderte CSV-Dateien stabil sind."""
        self.emit("occ_export_waiting", exportDirectory=str(directory), timeoutSeconds=timeout)
        deadline = time.monotonic() + timeout
        stable_snapshot: dict[Path, tuple[int, int]] | None = None
        stable_since: float | None = None

        while time.monotonic() < deadline:
            self.check_cancelled()
            current = self.csv_snapshot(directory)
            changed_files = [
                path for path, signature in current.items()
                if before_export.get(path) != signature or signature[0] >= started_at_ns
            ]
            if changed_files:
                changed_snapshot = {path: current[path] for path in changed_files}
                if changed_snapshot == stable_snapshot:
                    if stable_since is not None and time.monotonic() - stable_since >= 2:
                        self.emit("occ_export_data_ready", csvFiles=[str(path) for path in changed_files])
                        return changed_files
                else:
                    stable_snapshot = changed_snapshot
                    stable_since = time.monotonic()
            else:
                stable_snapshot = None
                stable_since = None
            self.wait(1)

        raise TimeoutError(
            f"Keine neuen oder aktualisierten CSV-Dateien in {directory} nach {timeout} Sekunden erkannt"
        )

    def export_occ(self, occ_path: Path) -> None:
        if not occ_path.is_file():
            raise FileNotFoundError(occ_path)
        self.check_cancelled()
        os.startfile(str(occ_path))
        self.wait(5)
        window = self.wait_for_main_window()
        try:
            if not self.open_export_dialog(window):
                raise RuntimeError("Exportmenü konnte nicht geöffnet werden")
            self.wait(2)
            export_directory = self.export_directory()
            before_export = self.csv_snapshot(export_directory)
            export_started_at_ns = time.time_ns()
            if not self.confirm_export_dialog():
                raise RuntimeError("Exportdialog konnte nicht bestätigt werden")
            self.wait_for_export_output(export_directory, before_export, export_started_at_ns)
        finally:
            self.close_omicron(window)

    def elapsed_seconds(self) -> int:
        if self.run_started_monotonic is None:
            return 0
        return max(0, int(time.monotonic() - self.run_started_monotonic))

    def write_report_if_needed(self, succeeded: int, failed: int, skipped: int) -> str | None:
        report_path_value = self.job.get("reportPath")
        if not report_path_value:
            return None
        report_path = Path(report_path_value)
        payload = {
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "elapsedSeconds": self.elapsed_seconds(),
            "summary": {
                "succeeded": succeeded,
                "failed": failed,
                "skipped": skipped,
            },
            "failures": self.failures,
            "skippedItems": self.skipped,
        }
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return str(report_path)

    def sort_occ_paths(self, occ_paths: list[Path]) -> list[Path]:
        """Sortiert OCC-Dateien stabil: zuerst nicht-EZE (z. B. NAP), dann EZE."""
        def sort_key(path: Path) -> tuple[int, str]:
            name = path.name.lower()
            is_eze = "eze" in name
            return (1 if is_eze else 0, name)

        return sorted(occ_paths, key=sort_key)

    def resolve_excel_groups(self, item: dict[str, Any]) -> list[tuple[Path, list[Path]]]:
        """Liefert stabile Gruppen: [(excel_path, [occ_paths...])]."""
        mappings = item.get("mappings")
        if mappings:
            grouped: dict[Path, list[Path]] = {}
            for mapping in mappings:
                occ_value = mapping.get("occPath")
                excel_value = mapping.get("excelPath")
                if not occ_value or not excel_value:
                    raise RuntimeError("Ungültige manuelle Zuordnung: OCC- oder Excel-Pfad fehlt")
                excel_path = Path(excel_value)
                occ_path = Path(occ_value)
                grouped.setdefault(excel_path, []).append(occ_path)

            groups: list[tuple[Path, list[Path]]] = []
            for excel_path in sorted(grouped.keys(), key=lambda path: path.name.lower()):
                groups.append((excel_path, self.sort_occ_paths(grouped[excel_path])))
            return groups

        occ_paths = self.sort_occ_paths([Path(path) for path in item.get("occPaths", [])])
        excel_path = Path(item["excelPath"])
        return [(excel_path, occ_paths)]

    def run_macro(self, excel, workbook_name: str, macro_name: str) -> None:
        excel.Application.Run(f"'{workbook_name}'!{macro_name}")
        self.wait(2)

    def refresh_excel(self, excel_path: Path) -> None:
        if not excel_path.is_file():
            raise FileNotFoundError(excel_path)
        excel = None
        workbook = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = True
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(excel_path.resolve()))
            workbook.RefreshAll()
            timeout = time.monotonic() + 180
            while time.monotonic() < timeout:
                self.check_cancelled()
                if excel.CalculationState == 0:
                    break
                self.wait(1)
            try:
                excel.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            self.wait(5)
            workbook.Worksheets(EXCEL_SHEET).Activate()
            self.run_macro(excel, workbook.Name, MACRO_PROTOCOL_NO)
            last_error: Exception | None = None
            for macro_name in MACRO_TOGGLE_SECTIONS_CANDIDATES:
                try:
                    self.run_macro(excel, workbook.Name, macro_name)
                    break
                except Exception as error:
                    last_error = error
            else:
                raise RuntimeError(f"Bereichsmakro nicht verfügbar: {last_error}")
            self.run_macro(excel, workbook.Name, MACRO_HIDE_EMPTY_ROWS)
            self.check_cancelled()
            workbook.Save()
            workbook.Close(SaveChanges=True)
            workbook = None
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()

    def run(self) -> int:
        items = self.job.get("items", [])
        self.run_started_monotonic = time.monotonic()
        self.emit("run_started", itemCount=len(items))
        succeeded_count = 0
        failed_count = 0
        skipped_count = 0
        for index, item in enumerate(items, start=1):
            item_id = item.get("id", str(index))
            try:
                self.check_cancelled()
                if item.get("mappingStatus") != "eindeutig":
                    self.emit("item_skipped", itemId=item_id, reason="Zuordnung vor Start nicht eindeutig")
                    skipped_count += 1
                    self.skipped.append({"itemId": str(item_id), "reason": "Zuordnung vor Start nicht eindeutig"})
                    continue
                excel_groups = self.resolve_excel_groups(item)
                if not excel_groups:
                    self.emit("item_skipped", itemId=item_id, reason="Keine OCC-/Excel-Zuordnung vorhanden")
                    skipped_count += 1
                    self.skipped.append({"itemId": str(item_id), "reason": "Keine OCC-/Excel-Zuordnung vorhanden"})
                    continue
                self.emit("item_started", itemId=item_id, index=index, total=len(items))

                # Gewünschte Reihenfolge pro Ordner:
                # je Excel-Gruppe: zugehörige OCC exportieren -> Excel-Bearbeitung
                for excel_path, occ_paths in excel_groups:
                    if not occ_paths or not excel_path.is_file():
                        raise FileNotFoundError(f"OCC- oder Excel-Datei fehlt für Zuordnung: {excel_path}")

                    for occ_path in occ_paths:
                        self.terminate_mashup_loader()
                        self.emit("occ_started", itemId=item_id, occPath=str(occ_path), excelPath=str(excel_path))
                        self.export_occ(occ_path)
                        self.emit("occ_completed", itemId=item_id, occPath=str(occ_path), excelPath=str(excel_path))

                    self.emit("excel_started", itemId=item_id, excelPath=str(excel_path))
                    self.refresh_excel(excel_path)
                    self.emit("excel_completed", itemId=item_id, excelPath=str(excel_path))
                self.emit("item_completed", itemId=item_id)
                succeeded_count += 1
            except CancellationRequested:
                self.emit("run_cancelled", itemId=item_id, elapsedSeconds=self.elapsed_seconds())
                return 2
            except Exception as error:
                self.emit("item_failed", itemId=item_id, message=str(error))
                failed_count += 1
                self.failures.append({"itemId": str(item_id), "message": str(error)})

        report_path = None
        try:
            if failed_count > 0 or skipped_count > 0:
                report_path = self.write_report_if_needed(succeeded_count, failed_count, skipped_count)
                if report_path:
                    self.emit("run_report_written", reportPath=report_path)
        except Exception as error:
            self.emit("run_report_failed", message=str(error))

        self.emit(
            "run_completed",
            elapsedSeconds=self.elapsed_seconds(),
            succeededCount=succeeded_count,
            failedCount=failed_count,
            skippedCount=skipped_count,
            reportPath=report_path,
        )
        return 0 if failed_count == 0 else 1


def main() -> int:
    parser = argparse.ArgumentParser(description="Unbeaufsichtigter Omicron-OCC-Worker")
    parser.add_argument("job", type=Path, help="JSON-Datei mit vorab bestätigten Verarbeitungseinheiten")
    parser.add_argument("--cancel-file", type=Path, help="Vorhandene Datei löst kontrollierten Abbruch aus")
    args = parser.parse_args()
    with args.job.open(encoding="utf-8") as job_file:
        job = json.load(job_file)
    return Worker(job, args.cancel_file).run()


if __name__ == "__main__":
    sys.exit(main())