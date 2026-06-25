from __future__ import annotations

from pathlib import Path
from typing import Any
import tempfile


class EmbeddedWorkbookWriterUnavailable(RuntimeError):
    pass


def update_embedded_workbook(workbook_bytes: bytes, sheet_name: str, matrix: list[list[Any]]) -> bytes:
    return _update_embedded_workbook_with_excel(workbook_bytes, sheet_name, matrix)


def _update_embedded_workbook_with_excel(
    workbook_bytes: bytes,
    sheet_name: str,
    matrix: list[list[Any]],
) -> bytes:
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except Exception as exc:
        raise EmbeddedWorkbookWriterUnavailable(
            "A atualização de gráficos editáveis exige Microsoft Excel via COM. "
            "Este ambiente não tem pywin32/Excel disponível."
        ) from exc

    with tempfile.TemporaryDirectory(prefix="auto_ppt_embedded_") as tmp:
        path = Path(tmp) / "embedded.xlsx"
        path.write_bytes(workbook_bytes)
        excel = None
        workbook = None
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
            worksheet = _excel_sheet(workbook, sheet_name)
            _write_matrix_with_excel(worksheet, matrix)
            workbook.Save()
            workbook.Close(SaveChanges=True)
            workbook = None
            return path.read_bytes()
        except Exception as exc:
            raise EmbeddedWorkbookWriterUnavailable(
                "Não consegui salvar o workbook embutido pelo Microsoft Excel. "
                "Gere o PPT em uma sessão Windows com Excel instalado/licenciado e perfil de usuário ativo."
            ) from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


def _excel_sheet(workbook: Any, sheet_name: str) -> Any:
    if sheet_name:
        try:
            return workbook.Worksheets(sheet_name)
        except Exception:
            pass
    return workbook.Worksheets(1)


def _write_matrix_with_excel(worksheet: Any, matrix: list[list[Any]]) -> None:
    row_count = max(len(matrix), 1)
    col_count = max((len(row) for row in matrix), default=1)
    data = [
        [matrix[row][col] if row < len(matrix) and col < len(matrix[row]) else None for col in range(col_count)]
        for row in range(row_count)
    ]

    worksheet.UsedRange.ClearContents()
    target_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(row_count, col_count))
    target_range.Value = data

    if worksheet.ListObjects.Count:
        worksheet.ListObjects(1).Resize(target_range)
