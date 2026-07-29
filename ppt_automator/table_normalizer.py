from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable
import re
import unicodedata

from .ppt_discovery import PptTarget
from .xlsx_parser import ParsedXlsxTable


# Nivel 1 da politica de confianca do sistema: decide se existe ALGUM plano
# automatico para o target (abaixo disso o target fica "sem match" e so pode
# ser resolvido por IA ou override manual). Os outros 2 niveis (se um plano ja
# aceito ainda recebe segunda opiniao de IA, e se a revisao pesada por slide
# roda automaticamente) ficam documentados junto a _ai_review_confidence_floor()
# e _auto_slide_ai_confidence_floor() em web/main.py.
LOCAL_MATCH_THRESHOLD_STRONG_ID = 0.35
LOCAL_MATCH_THRESHOLD_DEFAULT = 0.45


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
    series_format_overrides: dict[str, str] = field(default_factory=dict)
    typed_edit_data: dict[str, Any] = field(default_factory=dict)
    table_matrix: list[list[Any]] = field(default_factory=list)
    table_header_rows: int = 0
    warnings: list[str] = field(default_factory=list)

    @property
    def target_id(self) -> str:
        return self.target.target_id

    @property
    def object_type(self) -> str:
        return self.target.object_type


@dataclass(frozen=True)
class SourceMatchCandidate:
    source: ParsedXlsxTable
    score: float
    reason: str
    strong_id_match: bool = False


@dataclass(frozen=True)
class _TextMatchFeatures:
    normalized: str
    tokens: frozenset[str]


@dataclass(frozen=True)
class _SourceMatchFeatures:
    obj_ids: frozenset[str]
    slide_hint: int | None
    filename: _TextMatchFeatures
    filename_context: _TextMatchFeatures
    semantic_contexts: tuple[_TextMatchFeatures, ...]
    categories: tuple[_TextMatchFeatures, ...]
    series: tuple[_TextMatchFeatures, ...]


@dataclass(frozen=True)
class _TargetMatchFeatures:
    keys: frozenset[str]
    filename_context: _TextMatchFeatures
    nearby_text: _TextMatchFeatures
    semantic_context: _TextMatchFeatures
    categories: tuple[_TextMatchFeatures, ...]
    series: tuple[_TextMatchFeatures, ...]
    comparison_series_required: bool
    table_cell_count: int


@dataclass(frozen=True)
class SourceMatchIndex:
    """Features imutaveis das fontes, calculadas uma vez por analise."""

    by_identity: dict[int, _SourceMatchFeatures]

    def features_for(self, source: ParsedXlsxTable) -> _SourceMatchFeatures:
        return self.by_identity.get(id(source)) or _source_match_features(source)


def build_source_match_index(sources: Iterable[ParsedXlsxTable]) -> SourceMatchIndex:
    return SourceMatchIndex(
        by_identity={
            id(source): _source_match_features(source)
            for source in sources
        }
    )


def build_transform_plans(
    targets: Iterable[PptTarget],
    sources: Iterable[ParsedXlsxTable],
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    source_match_index: SourceMatchIndex | None = None,
) -> list[TransformPlan]:
    """Match deterministico por SLIDE resolvido como atribuicao global 1:1.

    Em vez de cada target escolher gulosamente o melhor datasource (o que permite
    dois targets pegarem o mesmo arquivo e produz confianca baixa mesmo quando o
    vencedor e claro), montamos a matriz alvo x datasource do slide e resolvemos o
    casamento otimo (Hungarian). A confianca e calibrada pela MARGEM para o 2o
    melhor candidato: um vencedor folgado vira alta confianca e dispensa a IA.
    """
    source_list = list(sources)
    match_index = source_match_index or build_source_match_index(source_list)
    eligible = [target for target in targets if target.object_type in {"chart", "table"}]
    plans: list[TransformPlan] = []
    by_slide: dict[int, list[PptTarget]] = {}
    for target in eligible:
        by_slide.setdefault(target.slide_number, []).append(target)
    completed = 0
    total = len(eligible)
    for slide_number in sorted(by_slide):
        plans.extend(
            _assign_slide_plans(
                slide_number,
                by_slide[slide_number],
                source_list,
                source_match_index=match_index,
                progress_callback=progress_callback,
                completed_offset=completed,
                total_targets=total,
            )
        )
        completed += len(by_slide[slide_number])
    return plans


def _assign_slide_plans(
    slide_number: int,
    slide_targets: list[PptTarget],
    all_sources: list[ParsedXlsxTable],
    *,
    source_match_index: SourceMatchIndex,
    progress_callback: Callable[[dict[str, Any]], None] | None = None,
    completed_offset: int = 0,
    total_targets: int = 0,
) -> list[TransformPlan]:
    candidate_sources = _sources_for_slide(
        slide_number,
        all_sources,
        source_match_index=source_match_index,
    )
    if not candidate_sources:
        return []

    # matriz de scores alvo x datasource (0..1) + memoria de forca do id e razao
    score_matrix: list[list[float]] = []
    strong_matrix: list[list[bool]] = []
    reason_matrix: list[list[str]] = []
    for target_index, target in enumerate(slide_targets):
        candidates = {
            c.source.file_name: c
            for c in source_match_candidates(
                target,
                candidate_sources,
                source_match_index=source_match_index,
            )
        }
        row_scores, row_strong, row_reason = [], [], []
        for source in candidate_sources:
            cand = candidates.get(source.file_name)
            row_scores.append(cand.score if cand else 0.0)
            row_strong.append(bool(cand.strong_id_match) if cand else False)
            row_reason.append(cand.reason if cand else "")
        score_matrix.append(row_scores)
        strong_matrix.append(row_strong)
        reason_matrix.append(row_reason)
        _notify_progress(
            progress_callback,
            {
                "phase": "matching",
                "completed": completed_offset + target_index + 1,
                "total": total_targets or len(slide_targets),
                "slide": slide_number,
                "target_id": target.target_id,
                "message": f"Objeto {completed_offset + target_index + 1} de {total_targets or len(slide_targets)} analisado.",
            },
        )

    cost = [[1.0 - score for score in row] for row in score_matrix]
    assignment = _hungarian(cost)

    plans: list[TransformPlan] = []
    for row, col in enumerate(assignment):
        if col < 0:
            continue
        target = slide_targets[row]
        source = candidate_sources[col]
        score = score_matrix[row][col]
        strong = strong_matrix[row][col]
        second_best = max(
            (score_matrix[row][j] for j in range(len(candidate_sources)) if j != col),
            default=0.0,
        )
        threshold = LOCAL_MATCH_THRESHOLD_STRONG_ID if strong or target.object_type == "table" else LOCAL_MATCH_THRESHOLD_DEFAULT
        if score < threshold:
            continue
        if not _source_has_readable_data(source):
            continue
        confidence = _calibrated_confidence(score, second_best, strong)
        reason = reason_matrix[row][col] or "Datasource escolhido por compatibilidade estrutural."
        margin = score - second_best
        if not strong and margin >= _CLEAR_WINNER_MARGIN and score >= _CLEAR_WINNER_MIN_SCORE:
            reason += f"; vencedor claro no slide (margem {margin:.0%} para o 2o melhor)"
        plans.append(normalize_to_target(target, source, confidence=confidence, match_reason=reason))
    return plans


def _notify_progress(
    callback: Callable[[dict[str, Any]], None] | None,
    payload: dict[str, Any],
) -> None:
    if callback is None:
        return
    try:
        callback(payload)
    except Exception:
        # Observabilidade nao pode abortar uma analise valida.
        pass


def _sources_for_slide(
    slide_number: int,
    sources: list[ParsedXlsxTable],
    *,
    source_match_index: SourceMatchIndex | None = None,
) -> list[ParsedXlsxTable]:
    """Blocking por slide: usa o token 'slideN' do nome do arquivo para restringir
    candidatos. Datasources sem indicacao de slide continuam elegiveis para todos."""
    index = source_match_index or build_source_match_index(sources)
    hinted_here = [s for s in sources if index.features_for(s).slide_hint == slide_number]
    no_hint = [s for s in sources if index.features_for(s).slide_hint is None]
    scoped = hinted_here + no_hint
    return scoped or list(sources)


_SLIDE_HINT_RE = re.compile(r"slide[_\s-]*([0-9]{1,3})", re.IGNORECASE)


def _source_slide_hint(source: ParsedXlsxTable) -> int | None:
    haystack = " ".join([source.file_name, *(str(v) for v in source.metadata.values())])
    match = _SLIDE_HINT_RE.search(haystack)
    return int(match.group(1)) if match else None


# Calibracao de confianca (Camada 2). Um vencedor "folgado" dentro do slide vira
# alta confianca mesmo com score bruto moderado - e o que evita mandar para a IA
# um match determinnistico que ja esta correto (ex.: era 0.54 -> vira ~0.85).
_CLEAR_WINNER_MARGIN = 0.12
_CLEAR_WINNER_MIN_SCORE = 0.50


def _calibrated_confidence(score: float, second_best: float, strong: bool) -> float:
    if strong:
        return min(1.0, max(score, 0.9))
    margin = score - second_best
    if score >= _CLEAR_WINNER_MIN_SCORE and margin >= _CLEAR_WINNER_MARGIN:
        return min(1.0, score + margin + 0.15)
    return min(1.0, score)


def _hungarian(cost: list[list[float]]) -> list[int]:
    """Atribuicao 1:1 de custo minimo (Kuhn-Munkres, O(n^3)). Retorna, para cada
    linha, a coluna atribuida (ou -1).

    O algoritmo classico abaixo ja aceita n <= m. Antes a matriz era sempre
    preenchida ate max(n, m) x max(n, m): um slide com 1 objeto e 178 fontes
    fazia trabalho cubico para 178 linhas artificiais. So adicionamos colunas
    neutro-caras quando ha mais alvos do que fontes.
    """
    n = len(cost)
    if n == 0:
        return []
    original_m = len(cost[0]) if cost[0] else 0
    if original_m == 0:
        return [-1] * n
    pad = 10.0
    if n > original_m:
        matrix = [row + [pad] * (n - original_m) for row in cost]
    else:
        matrix = cost
    m = len(matrix[0])
    INF = float("inf")
    u = [0.0] * (n + 1)
    v = [0.0] * (m + 1)
    p = [0] * (m + 1)
    way = [0] * (m + 1)
    for i in range(1, n + 1):
        p[0] = i
        j0 = 0
        minv = [INF] * (m + 1)
        used = [False] * (m + 1)
        while True:
            used[j0] = True
            i0 = p[j0]
            delta = INF
            j1 = -1
            for j in range(1, m + 1):
                if not used[j]:
                    cur = matrix[i0 - 1][j - 1] - u[i0] - v[j]
                    if cur < minv[j]:
                        minv[j] = cur
                        way[j] = j0
                    if minv[j] < delta:
                        delta = minv[j]
                        j1 = j
            for j in range(m + 1):
                if used[j]:
                    u[p[j]] += delta
                    v[j] -= delta
                else:
                    minv[j] -= delta
            j0 = j1
            if p[j0] == 0:
                break
        while True:
            j1 = way[j0]
            p[j0] = p[j1]
            j0 = j1
            if j0 == 0:
                break
    ans = [-1] * n
    for j in range(1, m + 1):
        row = p[j] - 1
        col = j - 1
        if 0 <= row < n and 0 <= col < original_m:
            ans[row] = col
    return ans


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
    # O contrato do PPT decide apenas a orientacao fisica da matriz. Dentro de
    # cada eixo, a ordem e os labels exibidos devem vir do XLSX selecionado.
    target_rows = _target_axis_in_source_order(target_rows, axis_alignment, "row")
    target_cols = _target_axis_in_source_order(target_cols, axis_alignment, "col")

    values = []
    for row_label in target_rows:
        output_row = []
        for col_label in target_cols:
            output_row.append(_aligned_value(source, axis_alignment, row_label, col_label))
        values.append(output_row)

    output_rows = [_source_label_for_target_axis(axis_alignment, "row", label) for label in target_rows]
    output_cols = [_source_label_for_target_axis(axis_alignment, "col", label) for label in target_cols]
    if orientation_ppt == "series_rows_categories_columns":
        series = output_rows
        categories = output_cols
    else:
        categories = output_rows
        series = output_cols

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
        preserve_percentage_decimal=False,
        warnings=warnings,
    )


def _normalize_table(
    target: PptTarget,
    source: ParsedXlsxTable,
    confidence: float,
    match_reason: str,
) -> TransformPlan:
    if source.orientation == "key_value_rows":
        return _normalize_key_value_table(target, source, confidence, match_reason)

    if source.orientation == "categories_rows_series_columns":
        headers = list(source.series)
        row_labels = list(source.categories)
    else:
        headers = list(source.categories)
        row_labels = list(source.series)
    values = [list(row) for row in source.values]
    table_matrix, table_header_rows = _adaptive_table_matrix(
        target,
        headers=headers,
        row_labels=row_labels,
        values=values,
    )
    number_format = "thousands_pt_br" if _looks_like_thousands(values) else ""
    return TransformPlan(
        target=target,
        datasource=source,
        action="fill_table_cells",
        orientation_xlsx=source.orientation,
        orientation_ppt="table_cells",
        categories=headers,
        series=row_labels,
        values=values,
        confidence=confidence,
        reason=match_reason or "Tabela PowerPoint compativel com a matriz do XLSX.",
        number_format=number_format,
        table_matrix=table_matrix,
        table_header_rows=table_header_rows,
    )


def _normalize_key_value_table(
    target: PptTarget,
    source: ParsedXlsxTable,
    confidence: float,
    match_reason: str,
) -> TransformPlan:
    if target.table_cells and all(len(row) >= 2 and _norm(row[0]) for row in target.table_cells):
        template_only_rows: list[list[Any]] = []
        for row in target.table_cells:
            label = row[0]
            source_label = _best_match(label, source.categories)
            if _soft_text_score(label, source_label) < 0.68 and not _label_token_overlap(label, source_label):
                template_only_rows.append([label, row[1] if len(row) > 1 else ""])
        values = [
            *template_only_rows,
            *[
                [category, "" if row_values[0] is None else row_values[0]]
                for category, row_values in zip(source.categories, source.values)
            ],
        ]
        categories = ["", source.series[0] if source.series else "Valor"]
        series = [row[0] for row in values]
    else:
        categories = ["", source.series[0] if source.series else "Valor"]
        series = list(source.categories)
        values = [
            [category, "" if row_values[0] is None else row_values[0]]
            for category, row_values in zip(source.categories, source.values)
        ]
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
        reason=match_reason or "Tabela PowerPoint preenchida por linhas chave-valor do XLSX.",
        number_format=number_format,
        table_matrix=[list(row) for row in values],
        table_header_rows=0,
    )


def _adaptive_table_matrix(
    target: PptTarget,
    headers: list[str],
    row_labels: list[str],
    values: list[list[Any]],
) -> tuple[list[list[Any]], int]:
    """Monta a grade completa respeitando a estrutura atual e o XLSX.

    O PowerPoint decide se existe uma linha de cabecalho/coluna de rotulos. O
    XLSX decide a ordem e os nomes. Se surgirem varias linhas de atributos, uma
    coluna de rotulos e criada para que os nomes novos nao sejam descartados.
    """

    target_cells = [list(row) for row in (target.table_cells or [])]
    has_header = _target_has_header_row(target_cells, headers)
    data_width = max((len(row) for row in values), default=len(headers))
    has_labels = _target_has_label_column(
        target_cells,
        row_labels,
        data_width=data_width,
        has_header=has_header,
    )
    if len(row_labels) > 1:
        has_labels = True

    matrix: list[list[Any]] = []
    if has_header:
        matrix.append(([None] if has_labels else []) + list(headers))
    for index, row in enumerate(values):
        output_row = list(row)
        if has_labels:
            label = row_labels[index] if index < len(row_labels) else f"Linha {index + 1}"
            output_row.insert(0, label)
        matrix.append(output_row)
    return matrix, 1 if has_header else 0


def _target_has_header_row(target_cells: list[list[Any]], headers: list[str]) -> bool:
    if not target_cells or not headers:
        return False
    first_row = [_norm(value) for value in target_cells[0] if _norm(value)]
    normalized_headers = {_norm(value) for value in headers if _norm(value)}
    if first_row and normalized_headers:
        matches = sum(1 for value in first_row if value in normalized_headers)
        comparable = min(len(first_row), len(normalized_headers))
        if matches >= max(1, (comparable + 1) // 2):
            return True
    if len(target_cells) < 2:
        return False
    later_values = [
        value
        for row in target_cells[1:]
        for value in row
        if value is not None and str(value).strip() != ""
    ]
    return bool(first_row) and not later_values


def _target_has_label_column(
    target_cells: list[list[Any]],
    row_labels: list[str],
    data_width: int,
    has_header: bool,
) -> bool:
    if not target_cells:
        return False
    if max((len(row) for row in target_cells), default=0) == data_width + 1:
        return True
    data_rows = target_cells[1:] if has_header else target_cells
    first_column = [_norm(row[0]) for row in data_rows if row and _norm(row[0])]
    normalized_labels = {_norm(value) for value in row_labels if _norm(value)}
    return bool(first_column) and any(value in normalized_labels for value in first_column)


def _source_has_readable_data(source: ParsedXlsxTable) -> bool:
    return bool(source.values) and any(
        cell is not None and str(cell).strip() != "" for row in source.values for cell in row
    )


def source_match_candidates(
    target: PptTarget,
    sources: list[ParsedXlsxTable],
    limit: int | None = None,
    source_match_index: SourceMatchIndex | None = None,
) -> list[SourceMatchCandidate]:
    index = source_match_index or build_source_match_index(sources)
    target_features = _target_match_features(target)
    scored: list[SourceMatchCandidate] = []
    for source in sources:
        source_features = index.features_for(source)
        score = 0.0
        reasons = []
        strong_id_match = False
        target_keys = target_features.keys
        if source_features.obj_ids & target_keys:
            # Sinal deterministico mais forte: o XLSX carrega "obj<numero_do_shape>"
            # (em table_title/context_text), que casa exatamente com o shape do PPT.
            # Isso resolve o match sem IA para todo datasource devidamente marcado.
            score += 0.95
            reasons.append("id do objeto embutido no XLSX (obj<shape>) bate com o target")
            strong_id_match = True
        if source.source_id and source.source_id in target_keys:
            score += 0.72
            reasons.append("nome do arquivo bate com o target")
            strong_id_match = True
        if source.metadata.get("graph_id") in target_keys or source.metadata.get("ppt_tag") in target_keys:
            score += 0.2
            reasons.append("metadado do XLSX bate com o target")
            strong_id_match = True
        filename_score = _filename_context_score_prepared(target_features, source_features)
        if filename_score >= 0.55:
            score += 0.18 * filename_score
            reasons.append(f"nome do arquivo/contexto {filename_score:.0%}")
        semantic_context_score = _semantic_context_score_prepared(target_features, source_features)
        if semantic_context_score >= 0.55:
            reasons.append(f"contexto semantico {semantic_context_score:.0%}")
        if target.object_type == "chart":
            cat_score = max(
                _coverage_feature_score(target_features.categories, source_features.categories),
                _coverage_feature_score(target_features.categories, source_features.series),
            )
            series_score = max(
                _coverage_feature_score(target_features.series, source_features.series),
                _coverage_feature_score(target_features.series, source_features.categories),
            )
            if min(cat_score, series_score) >= 0.45:
                score += 0.35 * cat_score + 0.3 * series_score
            else:
                score += 0.18 * cat_score + 0.16 * series_score
            if not strong_id_match and target_features.comparison_series_required and series_score < 0.8:
                score -= 0.25
                reasons.append("series de comparativo incompletas")
            reasons.append(f"categorias {cat_score:.0%}, series {series_score:.0%}")
        if target.object_type == "table" and target.table_cells:
            cell_count = target_features.table_cell_count
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


_OBJ_ID_RE = re.compile(r"obj[_\s]*([0-9]{4,})", re.IGNORECASE)


def _source_obj_ids(source: ParsedXlsxTable) -> set[str]:
    """IDs de shape embutidos no XLSX no formato 'obj<numero>'.

    Datasources exportados carregam o id do objeto do PPT (ex.: 'obj3958478347')
    no titulo/contexto da tabela; esse numero e o proprio shape do grafico/tabela,
    entao serve como chave de match deterministica e barata (dispensa IA)."""
    ids: set[str] = set()
    for value in source.metadata.values():
        for match in _OBJ_ID_RE.findall(str(value)):
            ids.add(match)
    return ids


def _source_axes(source: ParsedXlsxTable) -> tuple[list[str], list[str]]:
    if source.orientation in {"series_rows_categories_columns", "single_series_row_categories_columns"}:
        return list(source.series), list(source.categories)
    return list(source.categories), list(source.series)


def _text_match_features(value: Any) -> _TextMatchFeatures:
    normalized = _norm(value)
    return _TextMatchFeatures(
        normalized=normalized,
        tokens=frozenset(normalized.split()) if normalized else frozenset(),
    )


def _joined_context(values: Iterable[Any]) -> str:
    return " ".join(str(value) for value in values if _norm(value))


def _source_match_features(source: ParsedXlsxTable) -> _SourceMatchFeatures:
    filename = Path(source.file_name).stem
    metadata_text = " ".join(str(value) for value in source.metadata.values())
    source_context = _joined_context(
        [
            filename,
            metadata_text,
            *source.categories,
            *source.series,
        ]
    )
    source_values = [
        str(source.metadata.get(key) or "")
        for key in ("table_title", "row_group_label", "context_text", "variable", "ppt_tag")
    ]
    return _SourceMatchFeatures(
        obj_ids=frozenset(_source_obj_ids(source)),
        slide_hint=_source_slide_hint(source),
        filename=_text_match_features(filename),
        filename_context=_text_match_features(source_context),
        semantic_contexts=tuple(_text_match_features(value) for value in _context_variants(source_values)),
        categories=tuple(_text_match_features(value) for value in source.categories),
        series=tuple(_text_match_features(value) for value in source.series),
    )


def _target_match_features(target: PptTarget) -> _TargetMatchFeatures:
    table_cells = [cell for row in target.table_cells[:4] for cell in row[:8]]
    filename_context = _joined_context(
        [
            target.target_id,
            target.shape_name,
            target.nearby_text,
            *target.expected_categories,
            *target.expected_series,
            *table_cells,
        ]
    )
    semantic_context = _joined_context(
        [
            target.nearby_text,
            target.slide_text,
            *target.expected_categories,
            *target.expected_series,
            *table_cells,
        ]
    )
    return _TargetMatchFeatures(
        keys=frozenset(
            str(key)
            for key in {target.target_id, target.shape_name}
            if key
        ),
        filename_context=_text_match_features(filename_context),
        nearby_text=_text_match_features(target.nearby_text),
        semantic_context=_text_match_features(semantic_context),
        categories=tuple(_text_match_features(value) for value in target.expected_categories),
        series=tuple(_text_match_features(value) for value in target.expected_series if value),
        comparison_series_required=_requires_comparison_series(target.expected_series),
        table_cell_count=max((len(row) for row in target.table_cells), default=0),
    )


def _soft_feature_score(left: _TextMatchFeatures, right: _TextMatchFeatures) -> float:
    left_norm = left.normalized
    right_norm = right.normalized
    if not left_norm or not right_norm:
        return 0.0
    domain_score = _domain_text_score(left_norm, right_norm)
    if domain_score:
        return domain_score
    if left_norm == right_norm:
        return 1.0
    if left_norm in right_norm or right_norm in left_norm:
        return 0.9
    if left.tokens and right.tokens:
        return len(left.tokens & right.tokens) / max(len(left.tokens), len(right.tokens))
    return 0.0


def _coverage_feature_score(
    targets: tuple[_TextMatchFeatures, ...],
    choices: tuple[_TextMatchFeatures, ...],
) -> float:
    required = [value for value in targets if value.normalized]
    if not required:
        return 0.0
    return sum(
        1
        for value in required
        if max((_soft_feature_score(value, choice) for choice in choices), default=0.0) >= 0.68
    ) / len(required)


def _filename_context_score_prepared(
    target: _TargetMatchFeatures,
    source: _SourceMatchFeatures,
) -> float:
    if not source.filename.normalized or len(source.filename.normalized) <= 2:
        return 0.0
    return max(
        _soft_feature_score(source.filename, target.filename_context),
        _soft_feature_score(source.filename, target.nearby_text),
        _soft_feature_score(target.nearby_text, source.filename_context),
    )


def _semantic_context_score_prepared(
    target: _TargetMatchFeatures,
    source: _SourceMatchFeatures,
) -> float:
    return max(
        (_soft_feature_score(target.semantic_context, value) for value in source.semantic_contexts),
        default=0.0,
    )


def _filename_context_score(target: PptTarget, source: ParsedXlsxTable) -> float:
    return _filename_context_score_prepared(
        _target_match_features(target),
        _source_match_features(source),
    )


def _semantic_context_score(target: PptTarget, source: ParsedXlsxTable) -> float:
    return _semantic_context_score_prepared(
        _target_match_features(target),
        _source_match_features(source),
    )


def _context_variants(values: list[str]) -> list[str]:
    variants: list[str] = []
    for value in values:
        text = str(value or "").strip()
        if not text:
            continue
        variants.append(text)
        for separator in (" - ", " | ", ":"):
            if separator in text:
                variants.extend(part.strip() for part in text.split(separator) if part.strip())
    return variants


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


def _source_label_for_target_axis(alignment: dict[str, Any], axis: str, target_label: str) -> str:
    mapping = alignment["row_map"] if axis == "row" else alignment["col_map"]
    return mapping.get(target_label) or target_label


def _target_axis_in_source_order(
    target_labels: list[str],
    alignment: dict[str, Any],
    axis: str,
) -> list[str]:
    mapping = alignment["row_map"] if axis == "row" else alignment["col_map"]
    if alignment["mode"] == "same":
        source_labels = alignment["source_rows"] if axis == "row" else alignment["source_cols"]
    else:
        source_labels = alignment["source_cols"] if axis == "row" else alignment["source_rows"]
    source_positions = {_norm(label): index for index, label in enumerate(source_labels)}

    def order_key(item: tuple[int, str]) -> tuple[int, int]:
        original_index, label = item
        mapped = mapping.get(label) or label
        source_index = source_positions.get(_norm(mapped), len(source_positions))
        return source_index, original_index

    return [label for _index, label in sorted(enumerate(target_labels), key=order_key)]


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


def _label_token_overlap(left: Any, right: Any) -> bool:
    left_tokens = {token for token in _norm(left).split() if len(token) >= 3}
    right_tokens = {token for token in _norm(right).split() if len(token) >= 3}
    return bool(left_tokens & right_tokens)


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
