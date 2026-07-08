from __future__ import annotations

from dataclasses import replace
from typing import Any, Iterable
import hashlib
import re
import unicodedata

from .ppt_discovery import PptTarget


def assign_slide_target_ids(targets: Iterable[PptTarget]) -> list[PptTarget]:
    grouped: dict[int, list[PptTarget]] = {}
    for target in targets:
        grouped.setdefault(target.slide_number, []).append(target)

    output: list[PptTarget] = []
    used_target_ids: set[str] = set()
    name_counts: dict[str, int] = {}
    for target in targets:
        if is_internal_target_id(target.shape_name):
            name_counts[target.shape_name] = name_counts.get(target.shape_name, 0) + 1

    for slide_number in sorted(grouped):
        slide_targets = sorted(grouped[slide_number], key=_target_sort_key)
        ordinal = 1
        for target in slide_targets:
            if is_internal_target_id(target.shape_name) and name_counts.get(target.shape_name) == 1:
                target_id = target.shape_name
            else:
                target_id = stable_target_id(target, ordinal)
                while target_id in used_target_ids:
                    ordinal += 1
                    target_id = stable_target_id(target, ordinal)
            used_target_ids.add(target_id)
            output.append(replace(target, target_key=target_id))
            ordinal += 1
    return sorted(output, key=lambda item: (item.slide_number, item.top_in, item.left_in, item.shape_id, item.shape_name))


def is_internal_target_id(value: str) -> bool:
    return bool(re.fullmatch(r"S\d{3}_T\d{3}_[A-Z0-9_]+", str(value or "")))


def stable_target_id(target: PptTarget, ordinal: int) -> str:
    target_type = re.sub(r"[^A-Za-z0-9]+", "_", target.object_type or "target").strip("_").upper()
    return f"S{target.slide_number:03d}_T{ordinal:03d}_{target_type}"


def visual_label(target_id: str) -> str:
    match = re.search(r"_T0*(\d+)_", target_id)
    if not match:
        return target_id
    return f"T{int(match.group(1))}"


def target_aliases(target: PptTarget) -> set[str]:
    aliases = {
        target.target_id,
        target.shape_name,
        target.shape_id,
    }
    if target.target_key:
        aliases.add(target.target_key)
    return {alias for alias in aliases if alias}


def _fp_norm(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch)).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _fp_join(values: Iterable[Any]) -> str:
    # Ordem-insensivel: o mesmo conjunto de rotulos gera a mesma assinatura,
    # mesmo que a ordem das linhas/colunas mude entre versoes do arquivo.
    return "|".join(sorted(token for token in (_fp_norm(v) for v in values) if token))


def target_fingerprint(target: PptTarget) -> str:
    """Impressao digital de CONTEUDO do objeto do PPT, estavel entre versoes do
    deck (independe do numero do shape e da posicao). Baseada no "esquema" do
    Editar dados: tipo + categorias + series + orientacao (ou cabecalhos, no caso
    de tabela). E a chave de aprendizado da Camada 4."""
    parts = [
        (target.object_type or "").lower(),
        _fp_join(target.expected_categories),
        _fp_join(target.expected_series),
    ]
    if target.object_type == "table" and target.table_cells:
        header = target.table_cells[0] if target.table_cells else []
        labels = [row[0] for row in target.table_cells[1:] if row]
        parts.append(_fp_join(header))
        parts.append(_fp_join(labels))
    raw = "||".join(parts)
    if not raw.strip("|"):
        return ""
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def source_signature(categories: Iterable[Any], series: Iterable[Any]) -> str:
    """Assinatura de CONTEUDO do datasource (categorias + series), estavel a
    renomeacoes do arquivo. Permite reencontrar o mesmo datasource na proxima
    execucao mesmo que o nome mude."""
    raw = "||".join([_fp_join(categories), _fp_join(series)])
    if not raw.strip("|"):
        return ""
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def label_set_overlap(left: Iterable[Any], right: Iterable[Any]) -> float:
    """Similaridade de Jaccard entre dois conjuntos de rotulos normalizados."""
    left_set = {token for token in (_fp_norm(v) for v in left) if token}
    right_set = {token for token in (_fp_norm(v) for v in right) if token}
    if not left_set or not right_set:
        return 0.0
    return len(left_set & right_set) / len(left_set | right_set)


def _target_sort_key(target: PptTarget) -> tuple[float, float, int, str, str]:
    try:
        shape_id = int(target.shape_id)
    except (TypeError, ValueError):
        shape_id = 0
    return (
        round(target.top_in or 0, 3),
        round(target.left_in or 0, 3),
        shape_id,
        target.object_type,
        target.shape_name,
    )
