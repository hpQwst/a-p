"""Camada 4 - aprendizado do mapeamento (match "perfeito na 2a vez").

A cada download bem-sucedido, o sistema guarda, para cada objeto do PPT, uma
impressao digital de CONTEUDO (target_fingerprint) e a assinatura de conteudo do
datasource usado (source_signature). Na proxima execucao este modulo reaplica o
que foi aprendido ANTES de qualquer heuristica de conteudo ou IA - e faz isso de
forma resistente a:

  * renomeacao dos datasources (casa por assinatura/overlap de rotulos, nao nome);
  * recriacao/edicao do deck (casa o target por fingerprint, nao pelo shape).

Assim o cliente so precisa jogar os XLSX + o PPT.
"""

from __future__ import annotations

from typing import Any, Iterable

from .ppt_discovery import PptTarget
from .xlsx_parser import ParsedXlsxTable
from .target_labeler import (
    label_set_overlap,
    source_signature,
    target_aliases,
    target_fingerprint,
)

_SOURCE_OVERLAP_FLOOR = 0.6


def resolve_learned_matches(
    entries: dict[str, Any],
    targets: Iterable[PptTarget],
    sources: Iterable[ParsedXlsxTable],
) -> dict[str, dict[str, Any]]:
    """Traduz as entradas do template salvo em matches para os targets/sources
    ATUAIS. Chave do retorno: target_id atual. 1:1 (um datasource por target)."""
    updatable = [target for target in targets if target.object_type in {"chart", "table"}]
    source_list = list(sources)

    alias_to_tid: dict[str, str] = {}
    fingerprint_to_tids: dict[str, list[str]] = {}
    for target in updatable:
        for alias in target_aliases(target):
            alias_to_tid.setdefault(str(alias), target.target_id)
        fingerprint = target_fingerprint(target)
        if fingerprint:
            fingerprint_to_tids.setdefault(fingerprint, []).append(target.target_id)

    used_sources: set[str] = set()
    resolved: dict[str, dict[str, Any]] = {}
    for raw_id, entry in (entries or {}).items():
        entry = entry or {}
        target_id = _resolve_target_id(str(raw_id), entry, alias_to_tid, fingerprint_to_tids)
        if not target_id or target_id in resolved:
            continue
        source = _resolve_source(entry, source_list, used_sources)
        if source is None:
            continue
        used_sources.add(source.file_name)
        resolved[target_id] = {
            "datasource": source.file_name,
            "confidence": 1.0,
            "reason": _reason_for(entry, source),
        }
    return resolved


def _resolve_target_id(
    raw_id: str,
    entry: dict[str, Any],
    alias_to_tid: dict[str, str],
    fingerprint_to_tids: dict[str, list[str]],
) -> str | None:
    # 1) id/alias direto (shape estavel entre execucoes)
    direct = alias_to_tid.get(raw_id)
    if direct:
        return direct
    for alias in entry.get("target_aliases") or []:
        if str(alias) in alias_to_tid:
            return alias_to_tid[str(alias)]
    # 2) fingerprint de conteudo (sobrevive a recriacao do deck) - so quando unico
    fingerprint = str(entry.get("target_fingerprint") or "")
    candidates = fingerprint_to_tids.get(fingerprint, [])
    if len(candidates) == 1:
        return candidates[0]
    return None


def _resolve_source(
    entry: dict[str, Any],
    sources: list[ParsedXlsxTable],
    used_sources: set[str],
) -> ParsedXlsxTable | None:
    available = [source for source in sources if source.file_name not in used_sources]
    if not available:
        return None

    # 1) nome do arquivo (basename) igual ao salvo
    basename = str(entry.get("datasource_basename") or "").strip().lower()
    if basename:
        for source in available:
            if source.file_name.split("/")[-1].lower() == basename:
                return source

    # 2) assinatura de conteudo identica (resistente a renome)
    signature = str(entry.get("source_signature") or "")
    if signature:
        for source in available:
            if source_signature(source.categories, source.series) == signature:
                return source

    # 3) melhor sobreposicao de rotulos (renome + pequena variacao de conteudo)
    saved_categories = entry.get("source_categories") or []
    saved_series = entry.get("source_series") or []
    if saved_categories or saved_series:
        best: ParsedXlsxTable | None = None
        best_score = 0.0
        for source in available:
            score = (
                label_set_overlap(saved_categories, source.categories)
                + label_set_overlap(saved_series, source.series)
            ) / 2
            if score > best_score:
                best_score = score
                best = source
        if best is not None and best_score >= _SOURCE_OVERLAP_FLOOR:
            return best
    return None


def _reason_for(entry: dict[str, Any], source: ParsedXlsxTable) -> str:
    saved_name = str(entry.get("datasource_basename") or entry.get("datasource") or "")
    if saved_name and source.file_name.split("/")[-1].lower() != saved_name.strip().lower():
        return (
            f"Mapeamento aprendido: datasource reconhecido por conteudo "
            f"(salvo como {saved_name}, agora {source.file_name.split('/')[-1]})."
        )
    return f"Mapeamento aprendido aplicou {source.file_name.split('/')[-1]}."


def mapping_entry_learning_fields(target: PptTarget, source: ParsedXlsxTable) -> dict[str, Any]:
    """Campos de aprendizado a persistir na entrada do template de mapeamento."""
    return {
        "target_fingerprint": target_fingerprint(target),
        "source_signature": source_signature(source.categories, source.series),
        "source_categories": list(source.categories),
        "source_series": list(source.series),
    }
