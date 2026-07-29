import unittest
from io import BytesIO
from zipfile import ZipFile

import openpyxl

from ppt_automator.xlsx_parser import count_table_blocks, parse_xlsx_table
from worker.processor import _analysis_warnings


def _workbook_bytes(sheets: dict[str, list[list]]) -> bytes:
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets.items():
        worksheet = workbook.create_sheet(title=title)
        for row in rows:
            worksheet.append(row)
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


SIMPLE = [["Regiao", "2025"], ["Norte", 10], ["Sul", 20]]


class CountTableBlocksTests(unittest.TestCase):
    def test_single_table_counts_one(self) -> None:
        self.assertEqual(count_table_blocks([["a", "b"], [1, 2]]), 1)

    def test_empty_sheet_counts_zero(self) -> None:
        self.assertEqual(count_table_blocks([[None, None], ["", ""]]), 0)

    def test_blank_row_separates_two_tables(self) -> None:
        rows = [["a", "b"], [1, 2], [None, None], ["c", "d"], [3, 4]]
        self.assertEqual(count_table_blocks(rows), 2)

    def test_spacer_column_is_not_a_second_table(self) -> None:
        """Coluna vazia entre rotulos e dados e formatacao comum de planilha
        exportada. Tratar isso como segunda tabela dispararia aviso em quase
        todo arquivo real."""
        rows = [["Nivel", "Cristal", None, 15.9, 10.3], [None, "Bronze", None, 29.4, 17.8]]
        self.assertEqual(count_table_blocks(rows), 1)

    def test_lone_title_line_above_the_table_is_not_a_table(self) -> None:
        rows = [["Relatorio mensal"], [None, None], ["a", "b"], [1, 2]]
        self.assertEqual(count_table_blocks(rows), 1)

    def test_single_column_list_is_not_counted_as_table(self) -> None:
        rows = [["nota"], [None], ["a", "b"], [1, 2]]
        self.assertEqual(count_table_blocks(rows), 1)


class ParsedTableStructureTests(unittest.TestCase):
    def test_records_every_sheet_name(self) -> None:
        parsed = parse_xlsx_table(_workbook_bytes({"Base": SIMPLE, "Rascunho": SIMPLE}))
        self.assertEqual(parsed.sheet_names, ["Base", "Rascunho"])
        self.assertEqual(parsed.sheet_name, "Base")

    def test_flags_two_tables_in_the_same_sheet(self) -> None:
        rows = SIMPLE + [[None, None]] + [["Canal", "2025"], ["Loja", 5]]
        parsed = parse_xlsx_table(_workbook_bytes({"Base": rows}))
        self.assertEqual(parsed.table_blocks, 2)

    def test_single_table_single_sheet_is_not_flagged(self) -> None:
        parsed = parse_xlsx_table(_workbook_bytes({"Base": SIMPLE}))
        self.assertEqual(parsed.table_blocks, 1)
        self.assertEqual(parsed.sheet_names, ["Base"])


class AnalysisWarningTests(unittest.TestCase):
    def _sources(self, files: dict[str, bytes]):
        buffer = BytesIO()
        with ZipFile(buffer, "w") as archive:
            for name, data in files.items():
                archive.writestr(name, data)
        from ppt_automator.xlsx_parser import parse_datasource_zip

        return parse_datasource_zip(buffer.getvalue())

    def test_every_sheet_becomes_its_own_source(self) -> None:
        """Antes so a primeira aba era lida e as outras viravam aviso. Um
        workbook de relatorio real tem dezenas de abas alimentando dezenas de
        graficos, entao cada aba precisa virar uma fonte."""
        sources = self._sources({"a.xlsx": _workbook_bytes({"Base": SIMPLE, "Extra": SIMPLE})})
        self.assertEqual(len(sources), 2)
        self.assertEqual(
            sorted(source.file_name for source in sources),
            ["a.xlsx#Base", "a.xlsx#Extra"],
        )
        self.assertEqual(sorted(source.sheet_name for source in sources), ["Base", "Extra"])
        text = " ".join(_analysis_warnings([], sources, []))
        self.assertIn("aba por aba", text)

    def test_single_sheet_workbook_keeps_the_plain_name(self) -> None:
        """Sem sufixo quando ha uma aba so: mapeamentos ja aprendidos casam a
        fonte pelo nome do arquivo e nao podem ser invalidados."""
        sources = self._sources({"a.xlsx": _workbook_bytes({"Base": SIMPLE})})
        self.assertEqual([source.file_name for source in sources], ["a.xlsx"])

    def test_warns_about_two_tables_in_one_sheet(self) -> None:
        rows = SIMPLE + [[None, None]] + [["Canal", "2025"], ["Loja", 5]]
        sources = self._sources({"a.xlsx": _workbook_bytes({"Base": rows})})
        text = " ".join(_analysis_warnings([], sources, []))
        self.assertIn("mais de uma tabela", text)

    def test_warns_about_duplicate_file_names(self) -> None:
        book = _workbook_bytes({"Base": SIMPLE})
        sources = self._sources({"pasta1/dados.xlsx": book, "pasta2/dados.xlsx": book})
        text = " ".join(_analysis_warnings([], sources, []))
        self.assertIn("mesmo nome de arquivo", text)
        self.assertIn("dados.xlsx", text)

    def test_clean_workbook_produces_no_structure_warning(self) -> None:
        sources = self._sources({"a.xlsx": _workbook_bytes({"Base": SIMPLE})})
        text = " ".join(_analysis_warnings([], sources, []))
        self.assertNotIn("mais de uma aba", text)
        self.assertNotIn("mais de uma tabela", text)
        self.assertNotIn("mesmo nome de arquivo", text)


class UploadDuplicateNameTests(unittest.TestCase):
    """Renomear duplicata em silencio faria o mapeamento salvo reaplicar a
    planilha errada depois, porque ele procura a fonte pelo nome do arquivo."""

    def setUp(self) -> None:
        from web.main import _coalesce_datasources_to_zip

        self.coalesce = _coalesce_datasources_to_zip
        self.book = _workbook_bytes({"Base": SIMPLE})

    def test_rejects_two_files_with_the_same_name(self) -> None:
        with self.assertRaises(ValueError) as caught:
            self.coalesce([("dados.xlsx", self.book), ("dados.xlsx", self.book)])
        self.assertIn("mesmo nome", str(caught.exception))
        self.assertIn("dados.xlsx", str(caught.exception))

    def test_same_name_in_different_folders_still_collides(self) -> None:
        with self.assertRaises(ValueError):
            self.coalesce([("a/dados.xlsx", self.book), ("b/dados.xlsx", self.book)])

    def test_distinct_names_are_packed(self) -> None:
        data, label = self.coalesce([("um.xlsx", self.book), ("dois.xlsx", self.book)])
        with ZipFile(BytesIO(data)) as archive:
            self.assertEqual(sorted(archive.namelist()), ["dois.xlsx", "um.xlsx"])
        self.assertEqual(label, "2 planilhas")


if __name__ == "__main__":
    unittest.main()
