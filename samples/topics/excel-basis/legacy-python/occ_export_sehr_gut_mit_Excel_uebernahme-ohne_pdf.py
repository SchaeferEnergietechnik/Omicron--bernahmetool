from pathlib import Path
import os
import time
import sys
import ctypes
import subprocess
import glob

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

CSV_EXPORT_DIR = r"C:\Omicron_Datenexport"


def log(msg: str):
    timestamp = time.strftime("%H:%M:%S")
    print(f"[{timestamp}] {msg}", flush=True)


def show_done_message(success: int, failed: int, excel_ok: bool, excel_file: str | None):
    excel_info = excel_file if excel_file else "keine Excel-Datei gefunden"
    message = (
        f"Fertig.\n\n"
        f"OCC erfolgreich exportiert: {success}\n"
        f"OCC fehlgeschlagen: {failed}\n"
        f"Excel erfolgreich: {'Ja' if excel_ok else 'Nein'}\n"
        f"Excel-Datei: {excel_info}"
    )
    ctypes.windll.user32.MessageBoxW(0, message, "OCC Export", 0)


def cleanup_before_export():
    log("Vorbereitung: Beende Power Query / Mashup-Prozesse...")

    try:
        result = subprocess.run(
            ["taskkill", "/IM", "Microsoft.Mashup.Container.Loader.exe", "/F"],
            capture_output=True,
            text=True,
            shell=False
        )
        if result.returncode == 0:
            log("Mashup-Prozess erfolgreich beendet.")
        else:
            log("Kein laufender Mashup-Prozess gefunden oder Beenden nicht nötig.")
    except Exception as e:
        log(f"Fehler beim Beenden des Mashup-Prozesses: {e}")

    log(f"Vorbereitung: Lösche CSV-Dateien in {CSV_EXPORT_DIR} ...")
    try:
        deleted = 0
        for csv_file in glob.glob(os.path.join(CSV_EXPORT_DIR, "*.csv")):
            try:
                os.remove(csv_file)
                deleted += 1
            except Exception as e:
                log(f"CSV konnte nicht gelöscht werden: {csv_file} -> {e}")
        log(f"CSV-Bereinigung abgeschlossen. Gelöscht: {deleted}")
    except Exception as e:
        log(f"Fehler beim Löschen der CSV-Dateien: {e}")


def find_excel_file():
    xlsm_files = sorted(Path(".").glob("*.xlsm"))
    if not xlsm_files:
        return None
    return xlsm_files[0]


def wait_for_main_window(timeout=30):
    end_time = time.time() + timeout
    last_error = None

    while time.time() < end_time:
        try:
            windows = find_windows(title_re=WINDOW_TITLE_RE)
            if windows:
                app = Application(backend="uia").connect(handle=windows[0])
                win = app.window(handle=windows[0])
                win.wait("visible", timeout=5)
                win.set_focus()
                return app, win
        except Exception as e:
            last_error = e

        time.sleep(1)

    raise ElementNotFoundError(
        f"Kein OMICRON-Fenster gefunden. Letzter Fehler: {last_error}"
    )


def click_datei_menu(win):
    candidates = [
        {"title": "Datei", "control_type": "TabItem"},
        {"title": "Datei", "control_type": "Button"},
        {"title": "Datei"},
    ]

    for candidate in candidates:
        try:
            ctrl = win.child_window(**candidate)
            if ctrl.exists(timeout=2):
                ctrl.click_input()
                time.sleep(1)
                return True
        except Exception:
            pass

    try:
        win.set_focus()
        win.type_keys("%D")
        time.sleep(1)
        return True
    except Exception:
        return False


def click_export_menu_item(win):
    candidates = [
        {"title": "Daten exportieren...", "control_type": "MenuItem"},
        {"title": "Daten exportieren...", "control_type": "ListItem"},
        {"title": "Daten exportieren..."},
        {"title": "Daten exportieren"},
    ]

    for candidate in candidates:
        try:
            ctrl = win.child_window(**candidate)
            if ctrl.exists(timeout=2):
                ctrl.click_input()
                time.sleep(1)
                return True
        except Exception:
            pass

    try:
        win.type_keys("{DOWN 8}{ENTER}")
        time.sleep(1)
        return True
    except Exception:
        return False


def open_export_dialog(win):
    try:
        win.menu_select("Datei->Daten exportieren...")
        time.sleep(1)
        return True
    except Exception:
        pass

    if not click_datei_menu(win):
        return False

    if not click_export_menu_item(win):
        return False

    return True


def confirm_export_dialog(timeout=20):
    end_time = time.time() + timeout
    last_error = None

    while time.time() < end_time:
        try:
            windows = find_windows(title_re=EXPORT_DIALOG_RE)
            if windows:
                app = Application(backend="uia").connect(handle=windows[0])
                dlg = app.window(handle=windows[0])
                dlg.wait("visible", timeout=5)
                dlg.set_focus()

                for title in ["OK", "Ok", "ok"]:
                    try:
                        btn = dlg.child_window(title=title, control_type="Button")
                        if btn.exists(timeout=1):
                            btn.click_input()
                            time.sleep(1)
                            return True
                    except Exception:
                        pass

                try:
                    dlg.type_keys("{ENTER}")
                    time.sleep(1)
                    return True
                except Exception:
                    pass
        except Exception as e:
            last_error = e

        time.sleep(1)

    log(f"Exportdialog nicht gefunden oder nicht bedienbar. Letzter Fehler: {last_error}")
    return False


def handle_save_prompt(timeout=10):
    end_time = time.time() + timeout

    while time.time() < end_time:
        try:
            windows = find_windows(title_re=WINDOW_TITLE_RE)
            for handle in windows:
                try:
                    app = Application(backend="uia").connect(handle=handle)
                    dlg = app.window(handle=handle)

                    texts = []
                    try:
                        texts.append(dlg.window_text())
                    except Exception:
                        pass

                    try:
                        descendants = dlg.descendants()
                        for d in descendants:
                            try:
                                t = d.window_text()
                                if t:
                                    texts.append(t)
                            except Exception:
                                pass
                    except Exception:
                        pass

                    combined_text = " ".join(texts).lower()

                    if "speichern" in combined_text:
                        log("Speicherdialog erkannt -> klicke 'Nein'")
                        for title in ["Nein", "NEIN", "No"]:
                            try:
                                btn = dlg.child_window(title=title, control_type="Button")
                                if btn.exists(timeout=1):
                                    btn.click_input()
                                    time.sleep(1)
                                    return True
                            except Exception:
                                pass

                        try:
                            dlg.type_keys("%N")
                            time.sleep(1)
                            return True
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

        time.sleep(0.5)

    return False


def close_main_window(win):
    try:
        win.close()
        time.sleep(1)
        handle_save_prompt(timeout=8)
        time.sleep(1)
    except Exception:
        pass


def process_file(file_path: Path, index: int, total: int):
    log(f"[{index}/{total}] Bearbeite: {file_path.name}")

    os.startfile(str(file_path))
    time.sleep(5)

    app, win = wait_for_main_window()

    if not open_export_dialog(win):
        log(f"FEHLER: Exportmenü konnte nicht geöffnet werden: {file_path.name}")
        close_main_window(win)
        return False

    time.sleep(2)

    if not confirm_export_dialog():
        log(f"FEHLER: Exportdialog konnte nicht bestätigt werden: {file_path.name}")
        close_main_window(win)
        return False

    log(f"Export ausgelöst: {file_path.name}")
    time.sleep(3)

    close_main_window(win)
    log(f"Abgeschlossen: {file_path.name}")
    return True


def quote_macro(workbook_name: str, macro_name: str) -> str:
    return f"'{workbook_name}'!{macro_name}"


def wait_for_excel_refresh(excel, timeout=180):
    log("Warte auf Abschluss von Berechnung / Abfragen...")

    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            if excel.CalculationState == 0:
                break
        except Exception:
            pass
        time.sleep(1)

    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    # zusätzliche Sicherheitswartezeit für PQ / Verbindungen
    time.sleep(5)


def run_macro(excel, workbook_name: str, macro_name: str):
    full_macro = quote_macro(workbook_name, macro_name)
    log(f"Starte Makro: {full_macro}")
    excel.Application.Run(full_macro)
    time.sleep(2)


def run_toggle_macro(excel, workbook_name: str):
    last_error = None

    for macro_name in MACRO_TOGGLE_SECTIONS_CANDIDATES:
        try:
            run_macro(excel, workbook_name, macro_name)
            return
        except Exception as e:
            last_error = e
            log(f"Makrovariante fehlgeschlagen: {macro_name} -> {e}")

    raise RuntimeError(
        f"Kein aufrufbarer Makroname für BereicheEinOderAusblenden_Start gefunden. Letzter Fehler: {last_error}"
    )


def run_excel_workflow():
    excel_path = find_excel_file()
    if excel_path is None:
        log("Keine .xlsm-Datei im aktuellen Ordner gefunden.")
        return False, None

    log(f"Starte Excel: {excel_path.name}")

    excel = None
    wb = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(excel_path.resolve()))
        time.sleep(3)

        log("Aktualisiere Excel-Daten (RefreshAll)...")
        wb.RefreshAll()
        wait_for_excel_refresh(excel, timeout=180)

        # optional zusätzlich alle Verbindungen refreshen
        try:
            for conn in wb.Connections:
                try:
                    conn.Refresh()
                except Exception:
                    pass
        except Exception:
            pass

        time.sleep(5)

        log(f"Aktiviere Arbeitsblatt: {EXCEL_SHEET}")
        ws = wb.Worksheets(EXCEL_SHEET)
        ws.Activate()
        time.sleep(1)

        run_macro(excel, wb.Name, MACRO_PROTOCOL_NO)
        run_toggle_macro(excel, wb.Name)
        run_macro(excel, wb.Name, MACRO_HIDE_EMPTY_ROWS)

        log("Speichere Excel-Datei...")
        wb.Save()
        time.sleep(2)

        wb.Close(SaveChanges=True)
        excel.Quit()

        log("Excel-Verarbeitung abgeschlossen.")
        return True, excel_path.name

    except Exception as e:
        log(f"Excel-Fehler: {e}")

        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass

        return False, excel_path.name


def main():
    cleanup_before_export()

    occ_files = sorted(Path(".").glob("*.occ"))

    success = 0
    failed = 0
    total = len(occ_files)

    if total == 0:
        log("Keine .occ-Dateien im aktuellen Ordner gefunden.")
    else:
        log(f"{total} .occ-Datei(en) gefunden.")
        log("Starte OCC-Export...")

        for i, file_path in enumerate(occ_files, start=1):
            try:
                if process_file(file_path, i, total):
                    success += 1
                else:
                    failed += 1
                time.sleep(2)
            except ElementNotFoundError as e:
                failed += 1
                log(f"Fenster nicht gefunden bei {file_path.name}: {e}")
            except Exception as e:
                failed += 1
                log(f"Allgemeiner Fehler bei {file_path.name}: {e}")

    log("")
    log("Starte Excel-Nachverarbeitung...")
    excel_ok, excel_file = run_excel_workflow()

    log("")
    log("Fertig.")
    log(f"OCC erfolgreich: {success}")
    log(f"OCC fehlgeschlagen: {failed}")
    log(f"Excel erfolgreich: {'Ja' if excel_ok else 'Nein'}")
    log(f"Excel-Datei: {excel_file if excel_file else 'keine gefunden'}")

    show_done_message(success, failed, excel_ok, excel_file)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("Abgebrochen durch Benutzer.")
        sys.exit(1)