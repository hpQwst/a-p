from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import json
import os
import re
import shutil
import time
import uuid

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from ppt_automator import generate_updated_pptx
from ppt_automator.ai import ai_configured, format_ai_error
from ppt_automator.embedded_workbook_writer import EmbeddedWorkbookWriterUnavailable
from ppt_automator.ai_mapper import suggest_source_matches_with_ai
from ppt_automator.ai_transform import suggest_transform_diagnostics
from ppt_automator.table_normalizer import source_match_candidates
from ppt_automator.project_store import (
    SQUADS,
    create_project,
    create_run,
    ensure_store,
    load_project_bytes,
    load_project_json,
    list_projects,
    load_project,
    safe_filename,
    save_project_bytes,
    save_project_json,
)
from worker.processor import (
    AnalysisResult,
    analyze_files,
    apply_ai_source_matches_to_analysis,
    apply_ai_recommendations_to_analysis,
    parse_slide_selection,
)


APP_ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = APP_ROOT.parent
RUNTIME_ROOT = Path(os.getenv("AUTO_PPT_RUNTIME_ROOT", PROJECT_ROOT / "workspace_data" / "web_jobs")).resolve()

app = FastAPI(title="QWST Auto PPT")
app.mount("/static", StaticFiles(directory=APP_ROOT / "static"), name="static")
templates = Jinja2Templates(directory=APP_ROOT / "templates")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    ensure_store()
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "squads": _squad_labels(),
            "projects_by_squad": _projects_by_squad(),
            "project_cards_by_squad": _project_cards_by_squad(),
            "ai_available": ai_configured(PROJECT_ROOT),
        },
    )


@app.post("/preview", response_class=HTMLResponse)
async def preview(
    request: Request,
    project_ref: str = Form(""),
    squad: str = Form("squad1"),
    project_name: str = Form(""),
    pptx: UploadFile = File(...),
    datasources: UploadFile = File(...),
    mapping: UploadFile | None = File(None),
    use_ai: str = Form(""),
    slides_to_update: str = Form(""),
) -> HTMLResponse:
    try:
        project = _resolve_project(project_ref, squad, project_name)
        selected_slides = parse_slide_selection(slides_to_update)
        pptx_bytes = await pptx.read()
        datasource_bytes = await datasources.read()
        mapping_bytes = await mapping.read() if mapping and mapping.filename else b""
        _validate_upload(pptx, ".pptx", "Envie um arquivo PPTX.")
        _validate_upload(datasources, ".zip", "Envie um ZIP com os XLSX.")
        if mapping and mapping.filename:
            _validate_upload(mapping, ".xlsx", "A planilha de mapeamento precisa ser XLSX.")
    except Exception as exc:
        return _error_response(request, str(exc), status_code=400)

    job_id = uuid.uuid4().hex
    job_dir = _job_dir(job_id, create=True)
    (job_dir / "input.pptx").write_bytes(pptx_bytes)
    (job_dir / "datasources.zip").write_bytes(datasource_bytes)
    if mapping_bytes:
        (job_dir / "mapping.xlsx").write_bytes(mapping_bytes)

    _save_job_metadata(
        job_dir,
        {
            "job_id": job_id,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "project": {
                "squad": project.squad,
                "slug": project.slug,
                "name": project.name,
            },
            "files": {
                "pptx": pptx.filename or "modelo.pptx",
                "datasources": datasources.filename or "datasources.zip",
                "mapping": mapping.filename if mapping and mapping.filename else "",
            },
            "slides": {
                "raw": slides_to_update.strip(),
                "numbers": selected_slides,
            },
            "use_ai": bool(use_ai) or ai_configured(PROJECT_ROOT),
        },
    )
    _save_project_checkpoint(job_dir, status="in_progress")
    return _render_preview(request, job_id)


@app.get("/projects/{squad}/{slug}/preview", response_class=HTMLResponse)
async def resume_project_preview(request: Request, squad: str, slug: str) -> HTMLResponse:
    try:
        project = load_project(squad, slug)
        if project is None:
            raise ValueError("Projeto nao encontrado.")
        job_id = _restore_project_checkpoint(project)
    except Exception as exc:
        return _error_response(request, str(exc), status_code=400)
    return _render_preview(request, job_id, notice="Checkpoint do projeto carregado.", prefer_cache=True)


@app.post("/jobs/{job_id}/slides", response_class=HTMLResponse)
async def update_job_slides(
    request: Request,
    job_id: str,
    slides_to_add: str = Form(""),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        added_slides = parse_slide_selection(slides_to_add)
        if not added_slides:
            raise ValueError("Informe ao menos um slide. Ex.: 2 ou 2, 5-7.")
        metadata = _load_job_metadata(job_dir)
        metadata["use_ai"] = ai_configured(PROJECT_ROOT) or bool(metadata.get("use_ai"))
        current = set(_selected_slides_for_job(job_dir))
        if current:
            merged = sorted(current | set(added_slides))
        else:
            merged = sorted(set(added_slides))
        metadata["slides"] = {
            "raw": ", ".join(str(slide) for slide in merged),
            "numbers": merged,
        }
        _save_job_metadata(job_dir, metadata)
        _clear_render_cache(job_dir)
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc))
    return _render_preview(request, job_id, notice=f"Slides adicionados ao escopo: {', '.join(str(slide) for slide in added_slides)}.")


@app.post("/jobs/{job_id}/targets/{target_id}/override", response_class=HTMLResponse)
async def override_target_datasource(
    request: Request,
    job_id: str,
    target_id: str,
    datasource: UploadFile = File(...),
    cell_range: str = Form(""),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        _validate_target_id(target_id)
        _validate_cell_range(cell_range)
        _validate_upload(datasource, ".xlsx", "Envie um XLSX para substituir o datasource deste target.")
        data = await datasource.read()
        target_dir = job_dir / "overrides" / target_id
        target_dir.mkdir(parents=True, exist_ok=True)
        for existing in target_dir.glob("*.xlsx"):
            existing.unlink()
        filename = safe_filename(datasource.filename or f"{target_id}.xlsx")
        (target_dir / filename).write_bytes(data)
        range_path = target_dir / "range.txt"
        if cell_range.strip():
            range_path.write_text(cell_range.strip(), encoding="utf-8")
        elif range_path.exists():
            range_path.unlink()
        metadata = _load_job_metadata(job_dir)
        metadata["use_ai"] = ai_configured(PROJECT_ROOT) or bool(metadata.get("use_ai"))
        _save_job_metadata(job_dir, metadata)
        _clear_ai_cache(job_dir, target_id=target_id)
        _clear_render_cache(job_dir)
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc))
    range_notice = f" com range {cell_range.strip()}" if cell_range.strip() else ""
    return _render_preview(request, job_id, notice=f"Datasource {filename}{range_notice} aplicado ao target {target_id}.")


@app.get("/jobs/{job_id}/download")
async def download(job_id: str) -> Response:
    job_dir = _job_dir(job_id)
    pptx_path = job_dir / "input.pptx"
    datasource_path = job_dir / "datasources.zip"
    if not pptx_path.exists() or not datasource_path.exists():
        raise HTTPException(status_code=404, detail="Job nao encontrado.")

    manual_sources = _manual_sources_for_job(job_dir)
    selected_slides = _selected_slides_for_job(job_dir)
    pptx_bytes = pptx_path.read_bytes()
    datasource_bytes = datasource_path.read_bytes()
    analysis = analyze_files(
        pptx_bytes,
        datasource_bytes,
        manual_sources=manual_sources,
        slide_numbers=selected_slides,
    )
    ai_matches, _ai_match_status = _ai_source_matches_for_job(job_dir, analysis, allow_ai=False)
    analysis = apply_ai_source_matches_to_analysis(analysis, ai_matches)
    ai_diagnostics, _ai_status = _ai_diagnostics_for_job(job_dir, analysis, allow_ai=False)
    analysis = apply_ai_recommendations_to_analysis(analysis, ai_diagnostics)
    try:
        output = generate_updated_pptx(pptx_bytes, analysis.plans)
    except EmbeddedWorkbookWriterUnavailable as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    file_name = f"ppt_automatizado_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
    _save_project_run(job_dir, output, analysis, file_name)
    _save_project_checkpoint(job_dir, status="completed")
    return Response(
        output,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
    )


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


def _render_preview(
    request: Request,
    job_id: str,
    notice: str = "",
    error: str = "",
    prefer_cache: bool = False,
) -> HTMLResponse:
    job_dir = _job_dir(job_id)
    if prefer_cache:
        cached = _load_render_cache(job_dir)
        if cached:
            cached["notice"] = notice
            cached["error"] = error
            cached["ai_log_entries"] = _read_ai_logs(job_dir)
            return templates.TemplateResponse(request, "preview.html", cached)
    metadata = _load_job_metadata(job_dir)
    selected_slides = _selected_slides_for_job(job_dir)
    try:
        analysis = analyze_files(
            (job_dir / "input.pptx").read_bytes(),
            (job_dir / "datasources.zip").read_bytes(),
            manual_sources=_manual_sources_for_job(job_dir),
            slide_numbers=selected_slides,
        )
    except Exception as exc:
        return _error_response(request, f"Nao consegui analisar os arquivos: {exc}", status_code=400)

    ai_matches, ai_match_status = _ai_source_matches_for_job(job_dir, analysis)
    analysis = apply_ai_source_matches_to_analysis(analysis, ai_matches)
    ai_diagnostics, ai_diagnostic_status = _ai_diagnostics_for_job(job_dir, analysis)
    analysis = apply_ai_recommendations_to_analysis(analysis, ai_diagnostics)
    ai_status = _combine_ai_status(ai_match_status, ai_diagnostic_status)
    cards_by_slide = _cards_by_slide(
        analysis,
        _manual_source_names(job_dir),
        _manual_source_ranges(job_dir),
        ai_diagnostics,
    )
    context = {
        "job_id": job_id,
        "metadata": metadata,
        "squad": metadata["project"]["squad"].title(),
        "project_name": metadata["project"]["name"],
        "target_count": analysis.target_count,
        "source_count": analysis.source_count,
        "mapped_count": len(analysis.plans),
        "cards_by_slide": dict(sorted(cards_by_slide.items())),
        "ai_status": ai_status,
        "analysis_warnings": analysis.warnings,
        "slide_selection_label": _slide_selection_label(selected_slides),
        "notice": notice,
        "error": error,
        "ai_log_entries": _read_ai_logs(job_dir),
    }
    _save_render_cache(job_dir, context)
    _save_project_checkpoint(job_dir, status="in_progress")
    return templates.TemplateResponse(request, "preview.html", context)


def _cards_by_slide(
    analysis: AnalysisResult,
    manual_names: dict[str, str],
    manual_ranges: dict[str, str],
    ai_diagnostics: dict[str, dict],
) -> dict[int, list[dict]]:
    preview_by_target = {item.target: item for item in analysis.preview}
    plan_by_target = {plan.target_id: plan for plan in analysis.plans}
    cards: dict[int, list[dict]] = defaultdict(list)
    for target in analysis.targets:
        item = preview_by_target.get(target.shape_name)
        plan = plan_by_target.get(target.shape_name)
        cards[target.slide_number].append(
            {
                "slide": target.slide_number,
                "target": target.shape_name,
                "object_type": target.object_type,
                "nearby_text": target.nearby_text,
                "has_plan": item is not None,
                "datasource": item.datasource if item else "",
                "action": item.action if item else "aguardando_datasource",
                "reason": item.reason if item else "Nenhum datasource compativel foi escolhido automaticamente.",
                "confidence": item.confidence if item else None,
                "headers": item.headers if item else [],
                "rows": item.rows if item else [],
                "manual_file": manual_names.get(target.shape_name, ""),
                "manual_range": manual_ranges.get(target.shape_name, ""),
                "ppt_contract": _ppt_contract_for_target(target),
                "source_detected": _source_detected_for_plan(plan) if plan else None,
                "ai": ai_diagnostics.get(target.shape_name),
            }
        )
    return cards


def _ai_diagnostics_for_job(
    job_dir: Path,
    analysis: AnalysisResult,
    allow_ai: bool = True,
) -> tuple[dict[str, dict], dict[str, str]]:
    if not ai_configured(PROJECT_ROOT):
        return {}, {"state": "disabled", "message": "IA indisponivel: configure OPENAI_API_KEY no .env."}
    if not analysis.plans:
        return {}, {"state": "warn", "message": "IA nao tem planos para diagnosticar depois do mapeamento."}
    cache_path = job_dir / "ai_diagnostics.json"
    payload: dict[str, dict] = {}
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    missing_plans = [plan for plan in analysis.plans if plan.target_id not in payload]
    if not missing_plans:
        return payload, {"state": "cached", "message": "Diagnostico IA carregado do cache."}
    if not allow_ai:
        return payload, {
            "state": "cached" if payload else "warn",
            "message": f"Download usou cache de IA; {len(missing_plans)} target(s) sem diagnostico cached nao foram enviados para IA.",
        }
    try:
        sent_at = _now_iso()
        started = time.perf_counter()
        payload_summary = _diagnostic_payload_summary(missing_plans)
        diagnostics = suggest_transform_diagnostics(missing_plans, root=PROJECT_ROOT)
        duration_ms = round((time.perf_counter() - started) * 1000)
        payload.update(
            {
                item.target: {
                    "status": item.status,
                    "confidence": round(item.confidence * 100, 1),
                    "action": item.action,
                    "reason": item.reason,
                    "row_mapping": item.row_mapping,
                    "column_mapping": item.column_mapping,
                    "recommended_edit_data": item.recommended_edit_data,
                }
            }
            for item in diagnostics
        )
        for plan in missing_plans:
            payload.setdefault(
                plan.target_id,
                {
                    "status": "review",
                    "confidence": 0,
                    "action": "no_ai_response",
                    "reason": "A IA nao retornou diagnostico para este target nesta chamada.",
                    "row_mapping": {},
                    "column_mapping": {},
                    "recommended_edit_data": {"orientation": "", "headers": [], "rows": []},
                },
            )
        cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        _append_ai_log(
            job_dir,
            {
                "operation": "transform_diagnostics",
                "status": "ok",
                "sent_at": sent_at,
                "returned_at": _now_iso(),
                "duration_ms": duration_ms,
                "target_count": len(missing_plans),
                "returned_count": len(diagnostics),
                "payload_summary": payload_summary,
            },
        )
        return payload, {"state": "ok", "message": f"Diagnostico IA concluiu {len(diagnostics)} target(s) novo(s)."}
    except Exception as exc:
        _append_ai_log(
            job_dir,
            {
                "operation": "transform_diagnostics",
                "status": "error",
                "sent_at": sent_at if "sent_at" in locals() else _now_iso(),
                "returned_at": _now_iso(),
                "duration_ms": round((time.perf_counter() - started) * 1000) if "started" in locals() else 0,
                "target_count": len(missing_plans),
                "error": format_ai_error(exc),
                "payload_summary": _diagnostic_payload_summary(missing_plans),
            },
        )
        return {}, {"state": "warn", "message": f"IA indisponível nesta análise: {format_ai_error(exc)}"}


def _ai_source_matches_for_job(
    job_dir: Path,
    analysis: AnalysisResult,
    allow_ai: bool = True,
) -> tuple[dict[str, dict], dict[str, str]]:
    if not ai_configured(PROJECT_ROOT):
        return {}, {"state": "disabled", "message": "IA indisponivel: configure OPENAI_API_KEY no .env."}
    planned_targets = {plan.target_id for plan in analysis.plans}
    unmatched = [
        target
        for target in analysis.targets
        if target.object_type in {"chart", "table"} and target.shape_name not in planned_targets
    ]
    if not unmatched:
        return {}, {"state": "ok", "message": "IA nao precisou criar novos matches de datasource."}

    cache_path = job_dir / "ai_source_matches.json"
    payload: dict[str, dict] = {}
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    missing_targets = [target for target in unmatched if target.shape_name not in payload]
    if not missing_targets:
        return payload, {"state": "cached", "message": f"Matches IA carregados do cache ({len(payload)} sugestao/oes)."}
    if not allow_ai:
        return payload, {
            "state": "cached" if payload else "warn",
            "message": f"Download usou cache de IA; {len(missing_targets)} target(s) sem match cached nao foram enviados para IA.",
        }
    try:
        sent_at = _now_iso()
        started = time.perf_counter()
        payload_summary = _match_payload_summary(missing_targets, analysis.sources)
        suggestions = suggest_source_matches_with_ai(
            missing_targets,
            analysis.sources,
            existing_plan_ids=planned_targets,
            root=PROJECT_ROOT,
        )
        duration_ms = round((time.perf_counter() - started) * 1000)
        for target in missing_targets:
            payload.setdefault(
                target.shape_name,
                {
                    "datasource": "",
                    "confidence": 0,
                    "reason": "IA nao encontrou match confiavel para este target.",
                    "status": "no_match",
                },
            )
        payload.update(
            {
                item.target: {
                    "datasource": item.datasource,
                    "confidence": item.confidence,
                    "reason": item.reason,
                    "status": "matched",
                }
            }
            for item in suggestions
        )
        cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        _append_ai_log(
            job_dir,
            {
                "operation": "source_match",
                "status": "ok",
                "sent_at": sent_at,
                "returned_at": _now_iso(),
                "duration_ms": duration_ms,
                "target_count": len(missing_targets),
                "returned_count": len(suggestions),
                "payload_summary": payload_summary,
            },
        )
        if suggestions:
            return payload, {"state": "ok", "message": f"IA criou {len(suggestions)} match(es) novo(s) de datasource."}
        if payload:
            return payload, {"state": "cached", "message": f"Matches IA carregados do cache ({len(payload)} sugestao/oes)."}
        return {}, {"state": "warn", "message": "IA nao encontrou novos matches confiaveis de datasource."}
    except Exception as exc:
        _append_ai_log(
            job_dir,
            {
                "operation": "source_match",
                "status": "error",
                "sent_at": sent_at if "sent_at" in locals() else _now_iso(),
                "returned_at": _now_iso(),
                "duration_ms": round((time.perf_counter() - started) * 1000) if "started" in locals() else 0,
                "target_count": len(missing_targets),
                "error": format_ai_error(exc),
                "payload_summary": _match_payload_summary(missing_targets, analysis.sources),
            },
        )
        return {}, {"state": "warn", "message": f"IA indisponivel para match de datasource: {format_ai_error(exc)}"}


def _combine_ai_status(*statuses: dict[str, str]) -> dict[str, str]:
    clean = [status for status in statuses if status]
    if not clean:
        return {"state": "disabled", "message": "IA desativada para esta analise."}
    state = "ok"
    if any(status.get("state") == "warn" for status in clean):
        state = "warn"
    elif all(status.get("state") == "disabled" for status in clean):
        state = "disabled"
    elif any(status.get("state") == "cached" for status in clean):
        state = "cached"
    messages = []
    seen = set()
    for status in clean:
        message = status.get("message", "")
        if message and message not in seen:
            messages.append(message)
            seen.add(message)
    return {"state": state, "message": " ".join(messages)}


def _append_ai_log(job_dir: Path, event: dict) -> None:
    log_dir = job_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    event = {"created_at": _now_iso(), **event}
    with (log_dir / "ai_usage.jsonl").open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(event, ensure_ascii=False) + "\n")


def _read_ai_logs(job_dir: Path, limit: int = 12) -> list[dict]:
    log_path = job_dir / "logs" / "ai_usage.jsonl"
    if not log_path.exists():
        return []
    entries = []
    for line in log_path.read_text(encoding="utf-8").splitlines()[-limit:]:
        try:
            entries.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return list(reversed(entries))


def _diagnostic_payload_summary(plans: list) -> dict:
    return {
        "targets": [
            {
                "target": plan.target_id,
                "slide": plan.target.slide_number,
                "object_type": plan.object_type,
                "datasource": plan.datasource.file_name,
                "orientation_ppt": plan.orientation_ppt,
                "orientation_xlsx": plan.datasource.orientation,
                "categories": plan.categories[:12],
                "series": plan.series[:12],
                "value_shape": [len(plan.values), max((len(row) for row in plan.values), default=0)],
                "nearby_text": _short_text(plan.target.nearby_text, 240),
            }
            for plan in plans
        ]
    }


def _match_payload_summary(targets: list, sources: list) -> dict:
    return {
        "targets": [
            {
                "target": target.shape_name,
                "slide": target.slide_number,
                "object_type": target.object_type,
                "nearby_text": _short_text(target.nearby_text, 240),
                "categories": target.expected_categories[:12],
                "series": target.expected_series[:12],
                "candidates": [
                    {
                        "datasource": candidate.source.file_name,
                        "local_score": round(candidate.score, 4),
                        "reason": candidate.reason,
                        "xlsx_orientation": candidate.source.orientation,
                        "xlsx_categories": candidate.source.categories[:8],
                        "xlsx_series": candidate.source.series[:8],
                    }
                    for candidate in source_match_candidates(target, sources, limit=4)
                ],
            }
            for target in targets
        ]
    }


def _short_text(value, limit: int) -> str:
    text = "" if value is None else str(value)
    return text if len(text) <= limit else f"{text[:limit]}..."


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _render_cache_path(job_dir: Path) -> Path:
    return job_dir / "render_cache.json"


def _save_render_cache(job_dir: Path, context: dict) -> None:
    cached = dict(context)
    cached["notice"] = ""
    cached["error"] = ""
    cached["ai_log_entries"] = []
    _render_cache_path(job_dir).write_text(json.dumps(cached, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_render_cache(job_dir: Path) -> dict:
    cache_path = _render_cache_path(job_dir)
    if not cache_path.exists():
        return {}
    return json.loads(cache_path.read_text(encoding="utf-8"))


def _clear_render_cache(job_dir: Path) -> None:
    cache_path = _render_cache_path(job_dir)
    if cache_path.exists():
        cache_path.unlink()


def _clear_ai_cache(job_dir: Path, target_id: str | None = None) -> None:
    for cache_name in ("ai_diagnostics.json", "ai_source_matches.json"):
        cache_path = job_dir / cache_name
        if not cache_path.exists():
            continue
        if target_id is None:
            cache_path.unlink()
            continue
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
        if target_id in payload:
            payload.pop(target_id, None)
            cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _ppt_contract_for_target(target) -> dict:
    if target.object_type == "chart":
        if target.expected_orientation == "series_rows_categories_columns":
            return {
                "orientation": target.expected_orientation,
                "headers": ["", *target.expected_categories],
                "rows": [
                    [target.expected_series[index] if index < len(target.expected_series) else "", *row]
                    for index, row in enumerate(target.expected_values[:8])
                ],
            }
        return {
            "orientation": target.expected_orientation,
            "headers": ["", *target.expected_series],
            "rows": [
                [target.expected_categories[index] if index < len(target.expected_categories) else "", *row]
                for index, row in enumerate(target.expected_values[:8])
            ],
        }
    if target.object_type == "table":
        return {"orientation": "table_cells", "headers": [], "rows": target.table_cells[:8]}
    return {"orientation": target.object_type, "headers": [], "rows": []}


def _source_detected_for_plan(plan) -> dict:
    return {
        "orientation": plan.datasource.orientation,
        "headers": plan.datasource.preview_rows[0] if plan.datasource.preview_rows else [],
        "rows": plan.datasource.preview_rows[1:9] if plan.datasource.preview_rows else [],
    }


def _resolve_project(project_ref: str, squad: str, project_name: str):
    ensure_store()
    if project_ref:
        ref_squad, slug = project_ref.split("|", 1)
        project = load_project(ref_squad, slug)
        if project is None:
            raise ValueError("Projeto selecionado nao foi encontrado.")
        return project
    if not project_name.strip():
        raise ValueError("Selecione um projeto existente ou informe o nome de um novo projeto.")
    return create_project(_normalize_squad_form(squad), project_name.strip())


def _save_project_checkpoint(job_dir: Path, status: str = "in_progress") -> None:
    metadata = _load_job_metadata(job_dir)
    project_meta = metadata.get("project", {})
    project = load_project(project_meta.get("squad", ""), project_meta.get("slug", ""))
    if project is None:
        return

    checkpoint = {
        "schema_version": 1,
        "status": status,
        "job_id": metadata.get("job_id"),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "metadata": metadata,
        "manual_overrides": {
            target_id: {"filename": filename, "range": cell_range}
            for target_id, (filename, _data, cell_range) in _manual_sources_for_job(job_dir).items()
        },
        "caches": {
            cache_name: (job_dir / cache_name).exists()
            for cache_name in ("ai_source_matches.json", "ai_diagnostics.json", "render_cache.json")
        },
        "logs": {
            "ai_usage": (job_dir / "logs" / "ai_usage.jsonl").exists(),
        },
    }
    save_project_bytes(project, ["checkpoint"], "input.pptx", (job_dir / "input.pptx").read_bytes())
    save_project_bytes(project, ["checkpoint"], "datasources.zip", (job_dir / "datasources.zip").read_bytes())
    mapping_path = job_dir / "mapping.xlsx"
    if mapping_path.exists():
        save_project_bytes(project, ["checkpoint"], "mapping.xlsx", mapping_path.read_bytes())

    for target_id, (filename, data, _cell_range) in _manual_sources_for_job(job_dir).items():
        save_project_bytes(project, ["checkpoint", "overrides", target_id], filename, data)
    for cache_name in ("ai_source_matches.json", "ai_diagnostics.json", "render_cache.json"):
        cache_path = job_dir / cache_name
        if cache_path.exists():
            save_project_bytes(project, ["checkpoint"], cache_name, cache_path.read_bytes())
    ai_log_path = job_dir / "logs" / "ai_usage.jsonl"
    if ai_log_path.exists():
        save_project_bytes(project, ["checkpoint", "logs"], "ai_usage.jsonl", ai_log_path.read_bytes())
    save_project_json(project, ["checkpoint"], "checkpoint.json", checkpoint)


def _restore_project_checkpoint(project) -> str:
    try:
        checkpoint = load_project_json(project, ["checkpoint"], "checkpoint.json")
    except FileNotFoundError as exc:
        raise ValueError("Este projeto ainda nao tem preview salvo. Crie uma analise com PPTX e datasources primeiro.") from exc

    metadata = checkpoint.get("metadata") or {}
    job_id = str(checkpoint.get("job_id") or metadata.get("job_id") or "")
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        job_id = uuid.uuid4().hex
        metadata["job_id"] = job_id
    job_dir = _job_dir(job_id, create=True)
    (job_dir / "input.pptx").write_bytes(load_project_bytes(project, ["checkpoint"], "input.pptx"))
    (job_dir / "datasources.zip").write_bytes(load_project_bytes(project, ["checkpoint"], "datasources.zip"))
    mapping_path = job_dir / "mapping.xlsx"
    if (metadata.get("files") or {}).get("mapping"):
        try:
            mapping_path.write_bytes(load_project_bytes(project, ["checkpoint"], "mapping.xlsx"))
        except FileNotFoundError:
            if mapping_path.exists():
                mapping_path.unlink()
    elif mapping_path.exists():
        mapping_path.unlink()

    metadata.setdefault("project", {"squad": project.squad, "slug": project.slug, "name": project.name})
    metadata["job_id"] = job_id
    metadata["use_ai"] = ai_configured(PROJECT_ROOT) or bool(metadata.get("use_ai"))
    _save_job_metadata(job_dir, metadata)

    overrides_root = job_dir / "overrides"
    if overrides_root.exists():
        shutil.rmtree(overrides_root)
    for target_id, override in (checkpoint.get("manual_overrides") or {}).items():
        _validate_target_id(str(target_id))
        filename = safe_filename(str(override.get("filename") or "override.xlsx"))
        target_dir = overrides_root / str(target_id)
        target_dir.mkdir(parents=True, exist_ok=True)
        data = load_project_bytes(project, ["checkpoint", "overrides", str(target_id)], filename)
        (target_dir / filename).write_bytes(data)
        cell_range = str(override.get("range") or "").strip()
        if cell_range:
            (target_dir / "range.txt").write_text(cell_range, encoding="utf-8")

    cache_manifest = checkpoint.get("caches") or {}
    for cache_name in ("ai_source_matches.json", "ai_diagnostics.json", "render_cache.json"):
        cache_path = job_dir / cache_name
        if cache_manifest.get(cache_name):
            try:
                cache_path.write_bytes(load_project_bytes(project, ["checkpoint"], cache_name))
            except FileNotFoundError:
                if cache_path.exists():
                    cache_path.unlink()
        elif cache_path.exists():
            cache_path.unlink()
    log_dir = job_dir / "logs"
    ai_log_path = log_dir / "ai_usage.jsonl"
    if (checkpoint.get("logs") or {}).get("ai_usage"):
        try:
            log_dir.mkdir(parents=True, exist_ok=True)
            ai_log_path.write_bytes(load_project_bytes(project, ["checkpoint", "logs"], "ai_usage.jsonl"))
        except FileNotFoundError:
            if ai_log_path.exists():
                ai_log_path.unlink()
    elif ai_log_path.exists():
        ai_log_path.unlink()
    return job_id


def _save_project_run(job_dir: Path, output: bytes, analysis: AnalysisResult, file_name: str) -> None:
    metadata = _load_job_metadata(job_dir)
    project_meta = metadata.get("project", {})
    project = load_project(project_meta.get("squad", ""), project_meta.get("slug", ""))
    if project is None:
        return
    run = create_run(
        project,
        {
            "job_id": metadata.get("job_id"),
            "targets_found": analysis.target_count,
            "plans_generated": len(analysis.plans),
            "manual_overrides": sorted(_manual_source_names(job_dir)),
            "selected_slides": _selected_slides_for_job(job_dir),
        },
    )
    save_project_bytes(project, ["runs", run.run_id, "inputs"], metadata["files"]["pptx"], (job_dir / "input.pptx").read_bytes())
    save_project_bytes(
        project,
        ["runs", run.run_id, "inputs"],
        metadata["files"]["datasources"],
        (job_dir / "datasources.zip").read_bytes(),
    )
    mapping_path = job_dir / "mapping.xlsx"
    if mapping_path.exists():
        save_project_bytes(
            project,
            ["runs", run.run_id, "inputs"],
            metadata["files"].get("mapping") or "mapping.xlsx",
            mapping_path.read_bytes(),
        )
    manual_ranges = _manual_source_ranges(job_dir)
    for target_id, (filename, data, _cell_range) in _manual_sources_for_job(job_dir).items():
        save_project_bytes(project, ["runs", run.run_id, "overrides", target_id], filename, data)
    if manual_ranges:
        save_project_json(project, ["runs", run.run_id, "overrides"], "manual_ranges.json", manual_ranges)
    output_location = save_project_bytes(project, ["runs", run.run_id, "outputs"], file_name, output)
    save_project_json(
        project,
        ["runs", run.run_id, "reports"],
        "execution_report.json",
        {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "output": output_location,
            "selected_slides": _selected_slides_for_job(job_dir),
            "targets": [
                {
                    "target": plan.target_id,
                    "object_type": plan.object_type,
                    "datasource": plan.datasource.file_name,
                    "action": plan.action,
                    "confidence": plan.confidence,
                    "reason": plan.reason,
                }
                for plan in analysis.plans
            ],
        },
    )


def _manual_sources_for_job(job_dir: Path) -> dict[str, tuple[str, bytes, str]]:
    overrides_root = job_dir / "overrides"
    if not overrides_root.exists():
        return {}
    output: dict[str, tuple[str, bytes, str]] = {}
    for target_dir in overrides_root.iterdir():
        if not target_dir.is_dir():
            continue
        files = sorted(target_dir.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
        if files:
            range_path = target_dir / "range.txt"
            cell_range = range_path.read_text(encoding="utf-8").strip() if range_path.exists() else ""
            output[target_dir.name] = (files[0].name, files[0].read_bytes(), cell_range)
    return output


def _manual_source_names(job_dir: Path) -> dict[str, str]:
    return {target_id: filename for target_id, (filename, _data, _range) in _manual_sources_for_job(job_dir).items()}


def _manual_source_ranges(job_dir: Path) -> dict[str, str]:
    return {
        target_id: cell_range
        for target_id, (_filename, _data, cell_range) in _manual_sources_for_job(job_dir).items()
        if cell_range
    }


def _selected_slides_for_job(job_dir: Path) -> list[int]:
    metadata = _load_job_metadata(job_dir)
    slides = metadata.get("slides") or {}
    if isinstance(slides, dict):
        return [int(item) for item in slides.get("numbers") or []]
    return []


def _slide_selection_label(selected_slides: list[int]) -> str:
    if not selected_slides:
        return "Todos os slides"
    return ", ".join(str(slide) for slide in selected_slides)


def _job_dir(job_id: str, create: bool = False) -> Path:
    if not re.fullmatch(r"[a-f0-9]{32}", job_id):
        raise HTTPException(status_code=404, detail="Job nao encontrado.")
    path = RUNTIME_ROOT / job_id
    if create:
        path.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Job nao encontrado.")
    return path


def _validate_target_id(target_id: str) -> None:
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,80}", target_id):
        raise ValueError("Target invalido.")


def _validate_cell_range(cell_range: str) -> None:
    text = (cell_range or "").strip().replace("$", "")
    if not text:
        return
    ref = text.split("!", 1)[1] if "!" in text else text
    if not re.fullmatch(r"[A-Za-z]{1,4}\d{1,7}(:[A-Za-z]{1,4}\d{1,7})?", ref):
        raise ValueError("Range invalido. Use algo como D5:G12 ou Planilha1!D5:G12.")


def _validate_upload(upload: UploadFile, extension: str, message: str) -> None:
    filename = upload.filename or ""
    if not filename.lower().endswith(extension):
        raise ValueError(message)


def _save_job_metadata(job_dir: Path, payload: dict) -> None:
    (job_dir / "metadata.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_job_metadata(job_dir: Path) -> dict:
    return json.loads((job_dir / "metadata.json").read_text(encoding="utf-8"))


def _projects_by_squad() -> dict[str, list]:
    return {squad: list_projects(squad) for squad in SQUADS}


def _project_cards_by_squad() -> dict[str, list[dict]]:
    output: dict[str, list[dict]] = {}
    for squad in SQUADS:
        cards = []
        for project in list_projects(squad):
            checkpoint = _checkpoint_summary(project)
            cards.append(
                {
                    "project": project,
                    "has_checkpoint": bool(checkpoint),
                    "checkpoint_status": checkpoint.get("status", ""),
                    "checkpoint_updated_at": checkpoint.get("updated_at", ""),
                    "selected_slides": _checkpoint_slide_label(checkpoint),
                    "preview_url": f"/projects/{project.squad}/{project.slug}/preview",
                }
            )
        output[squad] = cards
    return output


def _checkpoint_summary(project) -> dict:
    try:
        return load_project_json(project, ["checkpoint"], "checkpoint.json")
    except FileNotFoundError:
        return {}


def _checkpoint_slide_label(checkpoint: dict) -> str:
    slides = ((checkpoint.get("metadata") or {}).get("slides") or {}).get("numbers") or []
    if not slides:
        return "Todos os slides"
    return ", ".join(str(slide) for slide in slides)


def _squad_labels() -> list[dict[str, str]]:
    return [{"value": squad, "label": squad.title()} for squad in SQUADS]


def _normalize_squad_form(value: str) -> str:
    normalized = value.strip().lower()
    if normalized not in SQUADS:
        raise ValueError("Squad invalido.")
    return normalized


def _error_response(request: Request, message: str, status_code: int = 400) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "error.html",
        {
            "message": message,
            "squads": _squad_labels(),
            "projects_by_squad": _projects_by_squad(),
            "project_cards_by_squad": _project_cards_by_squad(),
            "ai_available": ai_configured(PROJECT_ROOT),
        },
        status_code=status_code,
    )
