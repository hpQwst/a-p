from __future__ import annotations

import unittest

from ppt_automator.ppt_discovery import PptTarget
from ppt_automator.xlsx_parser import ParsedXlsxTable
from ppt_automator.learned_mapping import (
    mapping_entry_learning_fields,
    resolve_learned_matches,
)


def _target(shape: str, categories: list[str], series: list[str]) -> PptTarget:
    return PptTarget(
        slide_index=0,
        slide_number=3,
        slide_path="ppt/slides/slide3.xml",
        shape_name=shape,
        shape_id=shape,
        object_type="chart",
        left_in=0.0,
        top_in=0.0,
        width_in=5.0,
        height_in=3.0,
        expected_orientation="categories_rows_series_columns",
        expected_series=series,
        expected_categories=categories,
    )


def _source(name: str, categories: list[str], series: list[str]) -> ParsedXlsxTable:
    return ParsedXlsxTable(
        source_id="",
        file_name=name,
        sheet_name="Plan1",
        orientation="categories_rows_series_columns",
        categories=categories,
        series=series,
        values=[[1.0] * len(series) for _ in categories],
    )


def _entry(target: PptTarget, source: ParsedXlsxTable) -> dict:
    return {
        "target_id": target.target_id,
        "shape_name": target.shape_name,
        "target_aliases": [target.shape_name],
        "datasource": source.file_name,
        "datasource_basename": source.file_name.split("/")[-1],
        **mapping_entry_learning_fields(target, source),
    }


class LearnedMappingTests(unittest.TestCase):
    def test_source_recognized_after_rename_via_signature(self) -> None:
        # 1a execucao
        t = _target("chart1", ["Norte", "Sul"], ["Total"])
        s = _source("datasources/tab_slide3.xlsx", ["Norte", "Sul"], ["Total"])
        entry = _entry(t, s)

        # 2a execucao: MESMO deck, datasource com NOME diferente, mesmo conteudo
        s2 = _source("datasources/qualquer_nome_novo_slide3.xlsx", ["Norte", "Sul"], ["Total"])
        resolved = resolve_learned_matches({t.target_id: entry}, [t], [s2])

        self.assertIn(t.target_id, resolved)
        self.assertEqual(resolved[t.target_id]["datasource"], s2.file_name)

    def test_target_recognized_after_deck_rebuild_via_fingerprint(self) -> None:
        t = _target("111111", ["Norte", "Sul"], ["Total"])
        s = _source("datasources/tab_slide3.xlsx", ["Norte", "Sul"], ["Total"])
        entry = _entry(t, s)

        # deck recriado: shape mudou de numero, conteudo do Editar dados igual
        t_new = _target("999999", ["Norte", "Sul"], ["Total"])
        resolved = resolve_learned_matches({t.target_id: entry}, [t_new], [s])

        self.assertIn(t_new.target_id, resolved)
        self.assertEqual(resolved[t_new.target_id]["datasource"], s.file_name)

    def test_one_to_one_no_source_reuse(self) -> None:
        t1 = _target("c1", ["Norte", "Sul"], ["Total"])
        t2 = _target("c2", ["Leste", "Oeste"], ["Total"])
        s1 = _source("datasources/a_slide3.xlsx", ["Norte", "Sul"], ["Total"])
        s2 = _source("datasources/b_slide3.xlsx", ["Leste", "Oeste"], ["Total"])
        entries = {t1.target_id: _entry(t1, s1), t2.target_id: _entry(t2, s2)}

        resolved = resolve_learned_matches(entries, [t1, t2], [s1, s2])
        self.assertEqual(len(resolved), 2)
        picked = {resolved[t1.target_id]["datasource"], resolved[t2.target_id]["datasource"]}
        self.assertEqual(picked, {s1.file_name, s2.file_name})

    def test_saved_dynamic_range_is_returned_with_learned_match(self) -> None:
        target = _target("chart1", ["Jan/26", "Fev/26"], ["Total"])
        source = _source("datasources/mensal.xlsx", ["Jan/26", "Fev/26"], ["Total"])
        entry = {
            **_entry(target, source),
            "cell_range": "Base!A1:C2",
            "range_mode": "dynamic",
        }

        resolved = resolve_learned_matches({target.target_id: entry}, [target], [source])

        self.assertEqual(resolved[target.target_id]["cell_range"], "Base!A1:C2")
        self.assertEqual(resolved[target.target_id]["range_mode"], "dynamic")


if __name__ == "__main__":
    unittest.main()
