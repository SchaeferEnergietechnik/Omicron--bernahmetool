from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

V19G_DEFAULT = Path("samples/V19g_Übergeordneter_Entkupplungsschutz.xlsm")
V19M_DEFAULT = Path("samples/V19m_Übergeordneter_Entkupplungsschutz.xlsm")
MODULE_NAME = "Modul1"
REQUIRED_PROCEDURE = "Public Sub BereicheEinOderAusblenden_Start()"
W14_MARKER = 'Range("W14")'


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Uebertraegt das oeffentliche Bereichs-Makro aus V19g in V19m "
            "(Modul1) und erstellt vorher Sicherungen."
        )
    )
    parser.add_argument(
        "--v19g",
        type=Path,
        default=V19G_DEFAULT,
        help="Pfad zur V19g-Beispieldatei (.xlsm)",
    )
    parser.add_argument(
        "--v19m",
        type=Path,
        default=V19M_DEFAULT,
        help="Pfad zur V19m-Beispieldatei (.xlsm)",
    )
    parser.add_argument(
        "--backup-suffix",
        default=".bak_before_vba_migration",
        help="Suffix fuer Sicherungskopien (vor .xlsm)",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Excel sichtbar starten (Debug)",
    )
    return parser.parse_args()


def ensure_exists(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {path}")


def backup_file(path: Path, suffix: str) -> Path:
    backup_path = path.with_name(f"{path.stem}{suffix}{path.suffix}")
    shutil.copy2(path, backup_path)
    return backup_path


def get_module(vbproject, module_name: str):
    for component in vbproject.VBComponents:
        if component.Name == module_name:
            return component
    return None


def remove_standard_module_if_present(vbproject, module_name: str) -> None:
    component = get_module(vbproject, module_name)
    if component is None:
        return

    # 1 = vbext_ct_StdModule
    if int(component.Type) == 1:
        vbproject.VBComponents.Remove(component)
        return

    raise RuntimeError(
        f"Komponente '{module_name}' ist kein Standardmodul (Type={component.Type})."
    )


def verify_required_procedure(workbook, module_name: str, procedure_signature: str) -> None:
    component = get_module(workbook.VBProject, module_name)
    if component is None:
        raise RuntimeError(f"Modul '{module_name}' wurde nicht gefunden.")

    code_module = component.CodeModule
    code = code_module.Lines(1, code_module.CountOfLines)
    if procedure_signature not in code:
        raise RuntimeError(
            "Die erwartete Prozedur wurde nicht gefunden: "
            f"{procedure_signature}"
        )


def build_module_code_with_w14_logic(module_code: str) -> str:
    if W14_MARKER.lower() in module_code.lower():
        return module_code

    sub_name = "bereicheeinoderausblenden_start"
    lower_code = module_code.lower()
    start_index = lower_code.find(f"public sub {sub_name}()")
    if start_index == -1:
        raise RuntimeError("BereicheEinOderAusblenden_Start wurde in Modul1 nicht gefunden.")

    end_index = lower_code.find("\nend sub", start_index)
    if end_index == -1:
        raise RuntimeError("End Sub von BereicheEinOderAusblenden_Start wurde nicht gefunden.")

    w14_block = (
        "\n"
        "    If ThisWorkbook.Worksheets(\"Schutzpruef-Checkliste\").Range(\"W14\").Value = False Then\n"
        "        ThisWorkbook.Worksheets(\"Pruefprotokoll\").Rows(\"159:164\").EntireRow.Hidden = True\n"
        "    Else\n"
        "        ThisWorkbook.Worksheets(\"Pruefprotokoll\").Rows(\"159:164\").EntireRow.Hidden = False\n"
        "    End If\n"
    )

    return module_code[:end_index] + w14_block + module_code[end_index:]


def ensure_w14_logic(workbook, module_name: str) -> None:
    component = get_module(workbook.VBProject, module_name)
    if component is None:
        raise RuntimeError(f"Modul '{module_name}' wurde nicht gefunden.")

    code_module = component.CodeModule
    current_code = code_module.Lines(1, code_module.CountOfLines)
    updated_code = build_module_code_with_w14_logic(current_code)

    if updated_code == current_code:
        return

    if code_module.CountOfLines > 0:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(updated_code)


def transfer_module(v19g_path: Path, v19m_path: Path, visible: bool) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "Excel-COM ist nicht verfuegbar. Das Skript muss unter Windows mit pywin32 ausgefuehrt werden."
        ) from error

    pythoncom.CoInitialize()
    excel = None
    wb_g = None
    wb_m = None
    temp_bas: Path | None = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = bool(visible)
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        wb_g = excel.Workbooks.Open(str(v19g_path.resolve()))
        wb_m = excel.Workbooks.Open(str(v19m_path.resolve()))

        src_module = get_module(wb_g.VBProject, MODULE_NAME)
        if src_module is None:
            raise RuntimeError(f"In V19g fehlt das Modul '{MODULE_NAME}'.")

        with tempfile.TemporaryDirectory(prefix="vba_migration_") as tmp:
            temp_bas = Path(tmp) / f"{MODULE_NAME}.bas"
            src_module.Export(str(temp_bas))

            remove_standard_module_if_present(wb_m.VBProject, MODULE_NAME)
            wb_m.VBProject.VBComponents.Import(str(temp_bas))

            verify_required_procedure(wb_m, MODULE_NAME, REQUIRED_PROCEDURE)
            ensure_w14_logic(wb_m, MODULE_NAME)
            wb_m.Save()

        wb_g.Close(SaveChanges=False)
        wb_g = None
        wb_m.Close(SaveChanges=True)
        wb_m = None
    finally:
        if wb_g is not None:
            wb_g.Close(SaveChanges=False)
        if wb_m is not None:
            wb_m.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def main() -> int:
    args = parse_args()
    v19g_path = args.v19g
    v19m_path = args.v19m

    ensure_exists(v19g_path)
    ensure_exists(v19m_path)

    backup_g = backup_file(v19g_path, args.backup_suffix)
    backup_m = backup_file(v19m_path, args.backup_suffix)

    print(f"Sicherung erstellt: {backup_g}")
    print(f"Sicherung erstellt: {backup_m}")

    try:
        transfer_module(v19g_path=v19g_path, v19m_path=v19m_path, visible=args.visible)
    except Exception as error:
        print(f"Fehler bei der VBA-Uebertragung: {error}", file=sys.stderr)
        return 1

    print("VBA-Uebertragung erfolgreich abgeschlossen.")
    print(f"Geaenderte Datei: {v19m_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
