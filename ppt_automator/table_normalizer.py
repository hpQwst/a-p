from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable
import re
import unicodedata

from .ppt_discovery import PptTarget
from .xlsx_parser import ParsedXlsxTable


@dataclass(frozen=True)
class TransformPlan:
    target: PptTarget
    datasource: ParsedXlsxTable
    action: str
    orientation_xlsx: str
    orientation_ppt: str
    categories: list[str]
    series: list[str]
    values: list[list[Any]]
    confidence: float
    reason: str
    preserve_percentage_decimal: bool = False
    number_format: str = ""
    warnings: list[str] = field(default_factory=list)

    @property
    def target_id(self) -> str:
        return self.target.shape_name

    @property
    def object_type(self) -> str:
        return self.target.object_type


@dataclass(frozen=True)
class SourceMatchCandidate:
    source: ParsedXlsxTable
    score: float
    reason: str
    strong_id_match: bool = False


def build_transform_plans(
    targets: Iterable[PptTarget],
    sources: Iterable[ParsedXlsxTable],
) -> list[TransformPlan]:
    source_list = list(sources)
    plans: list[TransformPlan] = []
    for target in targets:
        if target.object_type not in {"chart", "table"}:
            continue
        source, score, reason = _best_source_for_target(target, source_list)
        if source is None:
            continue
        plans.append(normalize_to_target(target, source, confidence=score, match_reason=reason))
    return plans


def normalize_to_target(
    target: PptTarget,
    source: ParsedXlsxTable,
    confidence: float = 1.0,
    match_reason: str = "",
) -> TransformPlan:
    if target.object_type == "chart":
        return _normalize_chart(target, source, confidence, match_reason)
    if target.object_type == "table":
        return _normalize_table(target, source, confidence, match_reason)
    raise ValueError(f"Tipo de target nao suportado: {target.object_type}")


def _normalize_chart(
    target: PptTarget,
    source: ParsedXlsxTable,
    confidence: float,
    match_reason: str,
) -> TransformPlan:
    orientation_ppt = target.expected_orientation or "categories_rows_series_columns"
    target_rows, target_cols = _target_axes(target, source)
    source_rows, source_cols = _source_axes(source)
    axis_alignment = _best_axis_alignment(target_rows, target_cols, source_rows, source_cols)

    values = []
    for row_label in target_rows:
        output_row = []
        for col_label in target_cols:
            output_row.append(_aligned_value(source, axis_alignment, row_label, col_label))
        values.append(output_row)

    if orientation_ppt == "series_rows_categories_columns":
        series = target_rows
        categories = target_cols
    else:
        categories = target_rows
        series = target_cols

    action = "transpose" if axis_alignment["mode"] == "cross" else "align"
    warnings = []
    if any(value is None for row in values for value in row):
        warnings.append("Alguns valores nao foram encontrados no datasource.")
    reason = match_reason or "Datasource escolhido por compatibilidade estrutural."
    if action == "transpose":
        reason += " Os eixos do XLSX e do PPT estao cruzados, entao a matriz foi transposta."
    else:
        reason += " Os eixos do XLSX foram alinhados ao contrato do Editar dados do PPT."

    return TransformPlan(
        target=target,
        datasource=source,
        action=action,
        orientation_xlsx=source.orientation,
        orientation_ppt=orientation_ppt,
        categories=categories,
        series=series,
        values=values,
        confidence=confidence,
        reason=reason.strip(),
        preserve_percentage_decimal=_has_decimal_percentages(values),
        warnings=warnings,
    )


def _normalize_table(
    target: PptTarget,
    source: ParsedXlsxTable,
    confidence: float,
    match_reason: str,
) -> TransformPlan:
    categories = list(source.categories)
    if source.orientation == "categories_rows_series_columns":
        values = source.values[:1]
        series = source.series[: len(values[0])] if values else source.series
    else:
        values = [source.values[0]] if source.values else [[]]
        series = source.series[:1] or ["Valor"]
    number_format = "thousands_pt_br" if _looks_like_thousands(values) else ""
    return TransformPlan(
        target=target,
        datasource=source,
        action="fill_table_cells",
        orientation_xlsx=source.orientation,
        orientation_ppt="table_cells",
        categories=categories,
        series=series,
        values=values,
        confidence=confidence,
        reason=match_reason or "Tabela PowerPoint compativel com a matriz do XLSX.",
        number_format=number_format,
    )


def _best_source_for_target(
    target: PptTarget,
    sources: list[ParsedXlsxTable],
) -> tuple[ParsedXlsxTable | None, float, str]:
    candidates = source_match_candidates(target, sources)
    if not candidates:
        return None, 0.0, ""
    best = candidates[0]
    reason = best.reason
    if len(candidates) > 1 and best.score - candidates[1].score <= 0.08 and candidates[1].score >= 0.45:
        reason += f"; atenção: datasource parecido também encontrado ({candidates[1].source.file_name}, score {candidates[1].score:.0%})"
    threshold = 0.35 if best.strong_id_match or target.object_type == "table" else 0.45
    if best.score < threshold:
        return None, best.score, reason
    return best.source, best.score, reason


def source_match_candidates(
    target: PptTarget,
    sources: list[ParsedXlsxTable],
    limit: int | None = None,
) -> list[SourceMatchCandidate]:
    scored: list[SourceMatchCandidate] = []
    for source in sources:
        score = 0.0
        reasons = []
        strong_id_match = False
        if source.source_id and source.source_id == target.shape_name:
            score += 0.72
            reasons.append("nome do arquivo bate com o target")
            strong_id_match = True
        if source.metadata.get("graph_id") == target.shape_name or source.metadata.get("ppt_tag") == target.shape_name:
            score += 0.2
            reasons.append("metadado do XLSX bate com o target")
            strong_id_match = True
        filename_score = _filename_context_score(target, source)
        if filename_score >= 0.55:
            score += 0.18 * filename_score
            reasons.append(f"nome do arquivo/contexto {filename_score:.0%}")
        if target.object_type == "chart":
            cat_score = max(
                _coverage_score(target.expected_categories, source.categories),
                _coverage_score(target.expected_categories, source.series),
            )
            series_score = max(
                _coverage_score([s for s in target.expected_series if s], source.series),
                _coverage_score([s for s in target.expected_series if s], source.categories),
            )
            if min(cat_score, series_score) >= 0.45:
                score += 0.35 * cat_score + 0.3 * series_score
            else:
                score += 0.18 * cat_score + 0.16 * series_score
            if not strong_id_match and _requires_comparison_series(target.expected_series) and series_score < 0.8:
                score -= 0.25
                reasons.append("series de comparativo incompletas")
            reasons.append(f"categorias {cat_score:.0%}, series {series_score:.0%}")
        if target.object_type == "table" and target.table_cells:
            cell_count = max((len(row) for row in target.table_cells), default=0)
            if cell_count and len(source.categories) == cell_count:
                score += 0.18
                reasons.append("quantidade de colunas/celulas compativel")
        scored.append(
            SourceMatchCandidate(
                source=source,
                score=max(0.0, min(score, 1.0)),
                reason="; ".join(reasons),
                strong_id_match=strong_id_match,
            )
        )
    scored.sort(key=lambda item: item.score, reverse=True)
    return scored[:limit] if limit else scored


def _source_axes(source: ParsedXlsxTable) -> tuple[list[str], list[str]]:
    if source.orientation in {"series_rows_categories_columns", "single_series_row_categories_columns"}:
        return list(source.series), list(source.categories)
    return list(source.categories), list(source.series)


def _filename_context_score(target: PptTarget, source: ParsedXlsxTable) -> float:
    filename = Path(source.file_name).stem
    if not filename or len(_norm(filename)) <= 2:
        return 0.0
    target_texts = [
        target.shape_name,
        target.nearby_text,
        *target.expected_categories,
        *target.expected_series,
        *[cell for row in target.table_cells[:4] for cell in row[:8]],
    ]
    metadata_text = " ".join(str(value) for value in source.metadata.values())
    source_texts = [
        filename,
        metadata_text,
        *source.categories,
        *source.series,
    ]
    target_context = " ".join(str(value) for value in target_texts if _norm(value))
    source_context = " ".join(str(value) for value in source_texts if _norm(value))
    return max(
        _soft_text_score(filename, target_context),
        _soft_text_score(filename, target.nearby_text),
        _soft_text_score(target.nearby_text, source_context),
    )


def _requires_comparison_series(labels: list[str]) -> bool:
    return sum(1 for label in labels if "COMP" in _norm(label)) >= 2


def _target_axes(target: PptTarget, source: ParsedXlsxTable) -> tuple[list[str], list[str]]:
    if target.expected_orientation == "series_rows_categories_columns":
        rows = _fill_blank_labels(target.expected_series, [*source.series, *source.categories]) or _source_axes(source)[0]
        cols = _fill_blank_labels(target.expected_categories, [*source.categories, *source.series]) or _source_axes(source)[1]
        return rows, cols
    rows = _fill_blank_labels(target.expected_categories, [*source.categories, *source.series]) or _source_axes(source)[0]
    cols = _fill_blank_labels(target.expected_series, [*source.series, *source.categories]) or _source_axes(source)[1]
    return rows, cols


def _fill_blank_labels(labels: list[str], candidates: list[str]) -> list[str]:
    if not labels:
        return []
    output: list[str] = []
    used = {_norm(label) for label in labels if _norm(label)}
    candidate_pool = [candidate for candidate in candidates if _norm(candidate)]
    for label in labels:
        if _norm(label):
            output.append(label)
            continue
        remaining = [candidate for candidate in candidate_pool if _norm(candidate) not in used]
        nps = next((candidate for candidate in remaining if _norm(candidate) == "NPS"), "")
        replacement = nps or (remaining[0] if remaining else label)
        output.append(replacement)
        if _norm(replacement):
            used.add(_norm(replacement))
    return output


def _best_axis_alignment(
    target_rows: list[str],
    target_cols: list[str],
    source_rows: list[str],
    source_cols: list[str],
) -> dict[str, Any]:
    same_score = _coverage_score(target_rows, source_rows) + _coverage_score(target_cols, source_cols)
    cross_score = _coverage_score(target_rows, source_cols) + _coverage_score(target_cols, source_rows)
    if cross_score > same_score:
        return {
            "mode": "cross",
            "row_map": _label_map(target_rows, source_cols),
            "col_map": _label_map(target_cols, source_rows),
            "source_rows": source_rows,
            "source_cols": source_cols,
        }
    return {
        "mode": "same",
        "row_map": _label_map(target_rows, source_rows),
        "col_map": _label_map(target_cols, source_cols),
        "source_rows": source_rows,
        "source_cols": source_cols,
    }


def _aligned_value(source: ParsedXlsxTable, alignment: dict[str, Any], target_row: str, target_col: str) -> Any:
    row_match = alignment["row_map"].get(target_row)
    col_match = alignment["col_map"].get(target_col)
    if alignment["mode"] == "same":
        source_row_label = row_match
        source_col_label = col_match
    else:
        source_row_label = col_match
        source_col_label = row_match
    row_index = _label_index(source_row_label, alignment["source_rows"])
    col_index = _label_index(source_col_label, alignment["source_cols"])
    if row_index is None or col_index is None:
        return None
    try:
        return source.values[row_index][col_index]
    except IndexError:
        return None


def _label_map(targets: list[str], choices: list[str]) -> dict[str, str]:
    output: dict[str, str] = {}
    used: set[str] = set()
    for target in targets:
        ranked = sorted(
            choices,
            key=lambda choice: _soft_text_score(target, choice),
            reverse=True,
        )
        chosen = next((choice for choice in ranked if _norm(choice) not in used), ranked[0] if ranked else "")
        if chosen:
            used.add(_norm(chosen))
        output[target] = chosen
    return output


def _label_index(label: str | None, labels: list[str]) -> int | None:
    if not label:
        return None
    label_norm = _norm(label)
    for index, candidate in enumerate(labels):
        if _norm(candidate) == label_norm:
            return index
    return None


def _coverage_score(targets: list[str], choices: list[str]) -> float:
    required = [value for value in targets if _norm(value)]
    if not required:
        return 0.0
    return sum(1 for value in required if _best_text_score(value, choices) >= 0.68) / len(required)


def _best_match(value: str, choices: list[str]) -> str:
    if not choices:
        return ""
    return max(choices, key=lambda choice: _soft_text_score(value, choice))


def _best_text_score(value: str, choices: list[str]) -> float:
    return max((_soft_text_score(value, choice) for choice in choices), default=0.0)


def _soft_text_score(left: Any, right: Any) -> float:
    left_norm = _norm(left)
    right_norm = _norm(right)
    if not left_norm or not right_norm:
        return 0.0
    domain_score = _domain_text_score(left_norm, right_norm)
    if domain_score:
        return domain_score
    if left_norm == right_norm:
        return 1.0
    if left_norm in right_norm or right_norm in left_norm:
        return 0.9
    left_tokens = set(left_norm.split())
    right_tokens = set(right_norm.split())
    if left_tokens and right_tokens:
        return len(left_tokens & right_tokens) / max(len(left_tokens), len(right_tokens))
    return 0.0


def _domain_text_score(left_norm: str, right_norm: str) -> float:
    pair = " ".join([left_norm, right_norm])
    if "AUTO CONSUMO" in pair and "USO PROPRIO" in pair and ("NAO" in pair or "COMPRO" in pair):
        return 0.9
    if "NAO TEM AUTO CONSUMO" in pair and "NAO COMPRO" in pair:
        return 0.95
    percentage_patterns = [
        ("ATE 25", "ATE 25"),
        ("26 50", "26 50"),
        ("51 75", "51 75"),
        ("MAIS DE 75", "MAIS DE 75"),
    ]
    for left_pattern, right_pattern in percentage_patterns:
        if left_pattern in left_norm and right_pattern in right_norm:
            return 0.96
        if left_pattern in right_norm and right_pattern in left_norm:
            return 0.96
    return 0.0


def _has_decimal_percentages(values: list[list[Any]]) -> bool:
    numeric = [_to_number(value) for row in values for value in row]
    numeric = [value for value in numeric if value is not None]
    if not numeric:
        return False
    return sum(1 for value in numeric if -1 <= value <= 1 and value not in {0, 1}) >= 2


def _looks_like_thousands(values: list[list[Any]]) -> bool:
    numeric = [_to_number(value) for row in values for value in row]
    numeric = [value for value in numeric if value is not None]
    return bool(numeric) and sum(1 for value in numeric if abs(value) >= 1000 and float(value).is_integer()) >= len(numeric) * 0.8


def _to_number(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _norm(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    text = text.replace("+", " PLUS ")
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()
