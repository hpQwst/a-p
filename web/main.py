from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import json
import os
import re
import uuid

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from ppt_automator import generate_updated_pptx
from ppt_automator.ai import ai_configured, format_ai_error
from ppt_automator.ai_transform import suggest_transform_diagnostics
from ppt_automator.project_store import (
    SQUADS,
    create_project,
    create_run,
    ensure_store,
    list_projects,
    load_project,
    safe_filename,
    save_project_bytes,
    save_project_json,
)
from worker.processor import AnalysisResult, analyze_files


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
) -> HTMLResponse:
    try:
        project = _resolve_project(project_ref, squad, project_name)
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
            "use_ai": bool(use_ai),
        },
    )
    return _render_preview(request, job_id)


@app.post("/jobs/{job_id}/targets/{target_id}/override", response_class=HTMLResponse)
async def override_target_datasource(
    request: Request,
    job_id: str,
    target_id: str,
    datasource: UploadFile = File(...),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        _validate_target_id(target_id)
        _validate_upload(datasource, ".xlsx", "Envie um XLSX para substituir o datasource deste target.")
        data = await datasource.read()
        target_dir = job_dir / "overrides" / target_id
        target_dir.mkdir(parents=True, exist_ok=True)
        for existing in target_dir.glob("*.xlsx"):
            existing.unlink()
        filename = safe_filename(datasource.filename or f"{target_id}.xlsx")
        (target_dir / filename).write_bytes(data)
        _clear_ai_cache(job_dir)
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc))
    return _render_preview(request, job_id, notice=f"Datasource {filename} aplicado ao target {target_id}.")


@app.get("/jobs/{job_id}/download")
async def download(job_id: str) -> Response:
    job_dir = _job_dir(job_id)
    pptx_path = job_dir / "input.pptx"
    datasource_path = job_dir / "datasources.zip"
    if not pptx_path.exists() or not datasource_path.exists():
        raise HTTPException(status_code=404, detail="Job nao encontrado.")

    manual_sources = _manual_sources_for_job(job_dir)
    pptx_bytes = pptx_path.read_bytes()
    datasource_bytes = datasource_path.read_bytes()
    analysis = analyze_files(pptx_bytes, datasource_bytes, manual_sources=manual_sources)
    output = generate_updated_pptx(pptx_bytes, analysis.plans)
    file_name = f"ppt_automatizado_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
    _save_project_run(job_dir, output, analysis, file_name)
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
) -> HTMLResponse:
    job_dir = _job_dir(job_id)
    metadata = _load_job_metadata(job_dir)
    try:
        analysis = analyze_files(
            (job_dir / "input.pptx").read_bytes(),
            (job_dir / "datasources.zip").read_bytes(),
            manual_sources=_manual_sources_for_job(job_dir),
        )
    except Exception as exc:
        return _error_response(request, f"Nao consegui analisar os arquivos: {exc}", status_code=400)

    ai_diagnostics, ai_status = _ai_diagnostics_for_job(job_dir, analysis)
    cards_by_slide = _cards_by_slide(analysis, _manual_source_names(job_dir), ai_diagnostics)
    return templates.TemplateResponse(
        request,
        "preview.html",
        {
            "job_id": job_id,
            "metadata": metadata,
            "squad": metadata["project"]["squad"].title(),
            "project_name": metadata["project"]["name"],
            "target_count": analysis.target_count,
            "source_count": analysis.source_count,
            "mapped_count": len(analysis.plans),
            "cards_by_slide": dict(sorted(cards_by_slide.items())),
            "ai_status": ai_status,
            "notice": notice,
            "error": error,
        },
    )


def _cards_by_slide(
    analysis: AnalysisResult,
    manual_names: dict[str, str],
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
                "ppt_contract": _ppt_contract_for_target(target),
                "source_detected": _source_detected_for_plan(plan) if plan else None,
                "ai": ai_diagnostics.get(target.shape_name),
            }
        )
    return cards


def _ai_diagnostics_for_job(job_dir: Path, analysis: AnalysisResult) -> tuple[dict[str, dict], dict[str, str]]:
    metadata = _load_job_metadata(job_dir)
    if not metadata.get("use_ai"):
        return {}, {"state": "disabled", "message": "IA desativada para esta análise."}
    cache_path = job_dir / "ai_diagnostics.json"
    if cache_path.exists():
        return json.loads(cache_path.read_text(encoding="utf-8")), {
            "state": "cached",
            "message": "Diagnóstico IA carregado do cache.",
        }
    try:
        diagnostics = suggest_transform_diagnostics(analysis.plans, root=PROJECT_ROOT)
        payload = {
            item.target: {
                "status": item.status,
                "confidence": round(item.confidence * 100, 1),
                "action": item.action,
                "reason": item.reason,
                "row_mapping": item.row_mapping,
                "column_mapping": item.column_mapping,
            }
            for item in diagnostics
        }
        cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return payload, {"state": "ok", "message": "Diagnóstico IA concluído."}
    except Exception as exc:
        return {}, {"state": "warn", "message": f"IA indisponível nesta análise: {format_ai_error(exc)}"}


def _clear_ai_cache(job_dir: Path) -> None:
    cache_path = job_dir / "ai_diagnostics.json"
    if cache_path.exists():
        cache_path.unlink()


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
    for target_id, (filename, data) in _manual_sources_for_job(job_dir).items():
        save_project_bytes(project, ["runs", run.run_id, "overrides", target_id], filename, data)
    output_location = save_project_bytes(project, ["runs", run.run_id, "outputs"], file_name, output)
    save_project_json(
        project,
        ["runs", run.run_id, "reports"],
        "execution_report.json",
        {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "output": output_location,
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


def _manual_sources_for_job(job_dir: Path) -> dict[str, tuple[str, bytes]]:
    overrides_root = job_dir / "overrides"
    if not overrides_root.exists():
        return {}
    output: dict[str, tuple[str, bytes]] = {}
    for target_dir in overrides_root.iterdir():
        if not target_dir.is_dir():
            continue
        files = sorted(target_dir.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
        if files:
            output[target_dir.name] = (files[0].name, files[0].read_bytes())
    return output


def _manual_source_names(job_dir: Path) -> dict[str, str]:
    return {target_id: filename for target_id, (filename, _data) in _manual_sources_for_job(job_dir).items()}


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
            "ai_available": ai_configured(PROJECT_ROOT),
        },
        status_code=status_code,
    )
