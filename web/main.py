from __future__ import annotations

from collections import defaultdict, OrderedDict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, replace
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import quote
import hashlib
import hmac
import json
import secrets
import os
import re
import shutil
import threading
import time
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import uuid

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from ppt_automator import generate_updated_pptx
from ppt_automator.archive_safety import (
    validate_datasource_zip_bytes,
    validate_pptx_bytes,
    validate_xlsx_bytes,
)
from ppt_automator.ai import ai_configured, format_ai_error
from ppt_automator.ai_debug import (
    log_debug_event,
    reset_ai_debug_log_path,
    reset_ai_usage_log_path,
    set_ai_debug_log_path,
    set_ai_usage_log_path,
)
from ppt_automator.embedded_workbook_writer import EmbeddedWorkbookWriterUnavailable
from ppt_automator.ppt_chart_writer import ChartSheetUnresolvedError, resolved_series_number_formats
from ppt_automator.ai_mapper import suggest_source_matches_with_ai
from ppt_automator.ai_slide_matrix_builder import SlideMatrixBuildInput, build_slide_matrices_with_ai
from ppt_automator.ai_transform import suggest_transform_diagnostics
from ppt_automator.edit_data_validator import validate_typed_edit_data
from ppt_automator.typed_matrix import normalize_typed_edit_data
from ppt_automator.slide_datasources import collect_datasource_entries, entries_for_slide
from ppt_automator.source_manifest import xlsx_source_manifest
from ppt_automator.table_normalizer import source_match_candidates
from ppt_automator.target_labeler import target_aliases, visual_label
from ppt_automator.learned_mapping import (
    mapping_entry_learning_fields,
    resolve_learned_matches,
)
from ppt_automator.xlsx_plaintext_dump import dump_xlsx_workbook, dump_xlsx_zip_entries
from ppt_automator.project_store import (
    SQUADS,
    create_project,
    create_run,
    ensure_store,
    ensure_user,
    load_project_bytes,
    load_project_json,
    load_user,
    list_mapping_templates,
    list_projects,
    list_users,
    load_mapping_template,
    load_project,
    normalize_email,
    record_admin_event,
    safe_filename,
    save_mapping_template,
    save_project_bytes,
    save_project_json,
    update_user,
)
from worker.processor import (
    AnalysisResult,
    analyze_files,
    apply_ai_source_matches_to_analysis,
    apply_ai_recommendations_to_analysis,
    apply_saved_source_matches_to_analysis,
    apply_typed_outputs_to_analysis,
    parse_slide_selection,
)
from web import audit, auth, entra


APP_ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = APP_ROOT.parent
RUNTIME_ROOT = Path(os.getenv("AUTO_PPT_RUNTIME_ROOT", PROJECT_ROOT / "workspace_data" / "web_jobs")).resolve()
RENDER_CACHE_VERSION = 7
PREVIEW_EXECUTOR = ThreadPoolExecutor(max_workers=max(int(os.getenv("AUTO_PPT_PREVIEW_WORKERS", "2") or "2"), 1))
PREVIEW_RUNNING: set[str] = set()
PREVIEW_RUNNING_LOCK = threading.Lock()
GENERATION_RUNNING: set[str] = set()
GENERATION_RUNNING_LOCK = threading.Lock()

# Cache em processo do resultado caro de analyze_files() (discovery do PPT + parse/
# recalculo de formula dos XLSX + matching deterministico). input.pptx/datasources.zip
# nunca sao reescritos depois da criacao do job, entao a unica coisa que pode invalidar
# esse resultado e mudanca de escopo de slides ou de overrides manuais por target -
# ambos cobertos pela assinatura em _analyze_files_signature. Limitado a poucos jobs
# simultaneos em memoria (nao precisa sobreviver a reinicio do processo).
ANALYZE_FILES_CACHE: "OrderedDict[str, tuple[tuple, AnalysisResult]]" = OrderedDict()
ANALYZE_FILES_CACHE_LOCK = threading.Lock()
ANALYZE_FILES_CACHE_MAX_JOBS = 8
PPT_XML_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

app = FastAPI(title="QWST Auto PPT")
app.mount("/static", StaticFiles(directory=APP_ROOT / "static"), name="static")
templates = Jinja2Templates(directory=APP_ROOT / "templates")
templates.env.filters["format_bytes"] = lambda value: _format_bytes(value)


def _asset_version(name: str) -> int:
    """Cache-busting token for a static asset (its mtime), so browsers fetch
    fresh CSS/JS after each deploy or edit instead of serving a stale copy."""
    try:
        return int((APP_ROOT / "static" / name).stat().st_mtime)
    except OSError:
        return 0


templates.env.globals["asset_version"] = _asset_version


@app.middleware("http")
async def require_team_password(request: Request, call_next):
    """Autentica e aplica isolamento real antes de qualquer rota do produto."""
    request.state.user = None
    path = request.url.path
    if auth.auth_enabled() and not auth.path_is_public(path):
        if not auth.request_is_authenticated(request.cookies):
            if _request_wants_json(request):
                return JSONResponse({"error": "Sessao expirada. Entre de novo."}, status_code=401)
            destination = path
            if request.url.query:
                destination = f"{destination}?{request.url.query}"
            return RedirectResponse(f"/login?next={quote(destination, safe='')}", status_code=303)

        email = auth.current_user(request.cookies)
        # Sessao sem identidade so existia no fallback por senha compartilhada.
        # Mantemos compatibilidade local, mas producao desliga esse modo.
        if email:
            user = ensure_user(email)
            request.state.user = user
            if not user.active:
                return templates.TemplateResponse(
                    request,
                    "account_blocked.html",
                    {"user": user},
                    status_code=403,
                )
            if path.startswith("/admin") and not user.is_admin:
                return _access_denied(request, "Apenas administradores podem abrir esta tela.")
            if not user.is_admin and not user.squad and path != "/choose-squad":
                return RedirectResponse("/choose-squad", status_code=303)
            if not user.is_admin and user.squad:
                path_squad = _squad_from_path(path)
                if path_squad and path_squad != user.squad:
                    return _access_denied(request, "Este projeto pertence a outro squad.")
                job_squad = _job_squad_from_runtime_path(path)
                if job_squad is not None and not job_squad:
                    return _access_denied(request, "Nao foi possivel validar o squad deste job.")
                if job_squad is not None and job_squad != user.squad:
                    return _access_denied(request, "Este job pertence a outro squad.")
    return await call_next(request)


def _request_wants_json(request: Request) -> bool:
    return request.headers.get("accept", "").startswith("application/json")


def _access_denied(request: Request, message: str) -> Response:
    if _request_wants_json(request):
        return JSONResponse({"error": message}, status_code=403)
    return templates.TemplateResponse(
        request,
        "error.html",
        {"message": message, "current_user": getattr(request.state, "user", None)},
        status_code=403,
    )


def _squad_from_path(path: str) -> str:
    match = re.match(r"^/projects/(squad[1-5])(?:/|$)", path)
    return match.group(1) if match else ""


def _job_squad_from_runtime_path(path: str) -> str | None:
    match = re.match(r"^/jobs/([a-f0-9]{32})(?:/|$)", path)
    if not match:
        return None
    job_dir = RUNTIME_ROOT / match.group(1)
    if not job_dir.exists():
        return None
    metadata_path = job_dir / "metadata.json"
    if not metadata_path.exists():
        return ""
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return ""
    if not isinstance(metadata, dict) or not isinstance(metadata.get("project"), dict):
        return ""
    squad = str(metadata["project"].get("squad") or "")
    return squad if squad in SQUADS else ""


def _login_page(request: Request, destination: str, error: str = "", status_code: int = 200) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "login.html",
        {
            "next_url": destination,
            "error": error,
            "entra_enabled": auth.entra_enabled(),
            "password_enabled": auth.team_password_enabled(),
            "config_problems": auth.config_problems(),
        },
        status_code=status_code,
    )


def _start_session(request: Request, destination: str, subject: str = "") -> Response:
    response = RedirectResponse(destination, status_code=303)
    response.set_cookie(
        auth.SESSION_COOKIE,
        auth.issue_session_token(subject),
        max_age=auth.SESSION_TTL_SECONDS,
        httponly=True,
        samesite="lax",
        secure=_request_is_https(request),
    )
    response.delete_cookie(auth.HANDSHAKE_COOKIE)
    return response


def _request_is_https(request: Request) -> bool:
    """No App Runner o TLS termina antes do container. O uvicorn roda com
    --proxy-headers, entao request.url.scheme ja reflete o X-Forwarded-Proto;
    o cabecalho e conferido tambem por seguranca."""
    forwarded = request.headers.get("x-forwarded-proto", "").split(",")[0].strip().lower()
    return request.url.scheme == "https" or forwarded == "https"


@app.get("/login", response_class=HTMLResponse)
async def login_form(request: Request, next: str = "/", erro: str = "") -> HTMLResponse:
    if not auth.auth_enabled() or auth.request_is_authenticated(request.cookies):
        return RedirectResponse(_safe_next(next), status_code=303)
    return _login_page(request, _safe_next(next), error=erro)


@app.post("/login", response_class=HTMLResponse)
async def login_submit(request: Request, password: str = Form(""), next: str = Form("/")) -> Response:
    destination = _safe_next(next)
    if not auth.auth_enabled():
        return RedirectResponse(destination, status_code=303)
    if not auth.team_password_enabled():
        return _login_page(request, destination, error="A senha compartilhada foi desativada. Entre com a Microsoft.", status_code=403)
    if not auth.password_matches(password):
        return _login_page(request, destination, error="Senha incorreta. Tente de novo.", status_code=401)
    return _start_session(request, destination)


@app.get("/auth/login")
async def entra_login(request: Request, next: str = "/") -> Response:
    """Inicia o login Microsoft. State e nonce vao num cookie assinado."""
    destination = _safe_next(next)
    if not auth.entra_enabled():
        return RedirectResponse(f"/login?next={quote(destination, safe='')}", status_code=303)
    state = secrets.token_urlsafe(24)
    nonce = secrets.token_urlsafe(24)
    try:
        url = entra.authorization_url(state, nonce)
    except entra.EntraError as exc:
        return _login_page(request, destination, error=str(exc), status_code=500)
    response = RedirectResponse(url, status_code=303)
    response.set_cookie(
        auth.HANDSHAKE_COOKIE,
        auth.issue_handshake_token(state, nonce, destination),
        max_age=auth.HANDSHAKE_TTL_SECONDS,
        httponly=True,
        samesite="lax",
        secure=_request_is_https(request),
    )
    return response


@app.get("/auth/callback")
async def entra_callback(
    request: Request,
    code: str = "",
    state: str = "",
    error: str = "",
    error_description: str = "",
) -> Response:
    if not auth.entra_enabled():
        return RedirectResponse("/login", status_code=303)

    handshake = auth.read_handshake_token(request.cookies.get(auth.HANDSHAKE_COOKIE, ""))
    if handshake is None:
        return _login_page(request, "/", error="O login demorou demais ou foi aberto em outra aba. Tente de novo.", status_code=400)
    expected_state, nonce, destination = handshake

    if error:
        return _login_page(request, destination, error=f"A Microsoft recusou o login: {error_description or error}", status_code=401)
    # Compara o state em tempo constante: e a defesa contra CSRF no retorno.
    if not code or not state or not hmac.compare_digest(state, expected_state):
        return _login_page(request, destination, error="Retorno de login invalido. Tente de novo.", status_code=400)

    try:
        email = await run_in_threadpool(entra.exchange_code, code, nonce)
    except entra.EntraError as exc:
        return _login_page(request, destination, error=str(exc), status_code=401)

    user = ensure_user(email)
    if not user.is_admin and not user.squad:
        destination = "/choose-squad"
    return _start_session(request, destination, subject=email)


@app.get("/auth/logout")
async def entra_logout() -> Response:
    return _clear_session()


@app.post("/logout")
async def logout() -> Response:
    return _clear_session()


def _clear_session() -> Response:
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(auth.SESSION_COOKIE)
    response.delete_cookie(auth.HANDSHAKE_COOKIE)
    return response


def _actor(request: Request) -> str:
    return audit.actor_from(request.cookies, auth)


def _request_user(request: Request):
    return getattr(request.state, "user", None)


def _safe_next(value: str) -> str:
    """So aceita caminho interno, para ninguem usar ?next= para redirecionar
    a vitima a um site externo depois do login."""
    candidate = (value or "/").strip()
    if not candidate.startswith("/") or candidate.startswith("//"):
        return "/"
    return candidate


@app.get("/choose-squad", response_class=HTMLResponse)
async def choose_squad_form(request: Request) -> Response:
    user = _request_user(request)
    if user is None:
        return RedirectResponse("/login", status_code=303)
    if user.is_admin or user.squad:
        return RedirectResponse("/", status_code=303)
    return templates.TemplateResponse(
        request,
        "choose_squad.html",
        {"user": user, "squads": _squad_labels(), "error": ""},
    )


@app.post("/choose-squad", response_class=HTMLResponse)
async def choose_squad_submit(request: Request, squad: str = Form("")) -> Response:
    user = _request_user(request)
    if user is None:
        return RedirectResponse("/login", status_code=303)
    if user.is_admin or user.squad:
        return RedirectResponse("/", status_code=303)
    try:
        selected = _normalize_squad_form(squad)
        updated = update_user(user.email, squad=selected)
        record_admin_event(
            updated.email,
            "selecionou_squad_no_primeiro_acesso",
            updated.email,
            {"squad": updated.squad},
        )
    except ValueError as exc:
        return templates.TemplateResponse(
            request,
            "choose_squad.html",
            {"user": user, "squads": _squad_labels(), "error": str(exc)},
            status_code=400,
        )
    return RedirectResponse("/", status_code=303)


@app.middleware("http")
async def production_guardrails(request: Request, call_next):
    request_id = request.headers.get("x-request-id", "").strip() or uuid.uuid4().hex
    content_length = request.headers.get("content-length", "").strip()
    if content_length.isdigit() and int(content_length) > _max_request_bytes():
        return JSONResponse(
            {"error": "A requisicao excede o limite de upload configurado.", "request_id": request_id},
            status_code=413,
            headers={"X-Request-ID": request_id},
        )
    response = await call_next(request)
    response.headers["X-Request-ID"] = request_id
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "same-origin"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    return response


@app.get("/", response_class=HTMLResponse)
async def index(request: Request, squad: str = "") -> HTMLResponse:
    ensure_store()
    user = _request_user(request)
    if user is not None and user.is_admin:
        try:
            selected_squad = _normalize_squad_form(squad or SQUADS[0])
        except ValueError:
            selected_squad = SQUADS[0]
        visible_squads = [selected_squad]
    elif user is not None and user.squad:
        selected_squad = user.squad
        visible_squads = [user.squad]
    else:
        selected_squad = _normalize_squad_form(squad or SQUADS[0])
        visible_squads = list(SQUADS)
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "squads": _squad_labels(visible_squads),
            "all_squads": _squad_labels(),
            "selected_squad": selected_squad,
            "current_user": user,
            "is_admin": bool(user and user.is_admin),
            "can_choose_squad": bool(user is None or user.is_admin),
            "projects_by_squad": _projects_by_squad(visible_squads),
            "project_cards_by_squad": _project_cards_by_squad(visible_squads),
            "resume_cards": _resume_cards(visible_squads),
            "mapping_templates_by_squad": _mapping_templates_by_squad(visible_squads),
            "ai_available": ai_configured(PROJECT_ROOT),
            "large_deck_slide_threshold": _large_deck_slide_threshold(),
        },
    )


@app.get("/admin/users", response_class=HTMLResponse)
async def admin_users(request: Request, notice: str = "", error: str = "") -> HTMLResponse:
    user = _request_user(request)
    if user is None or not user.is_admin:
        return _access_denied(request, "Apenas administradores podem abrir esta tela.")
    ensure_store()
    return templates.TemplateResponse(
        request,
        "admin_users.html",
        {
            "current_user": user,
            "users": list_users(),
            "squads": _squad_labels(),
            "notice": notice,
            "error": error,
        },
    )


@app.post("/admin/users/update", response_class=HTMLResponse)
async def admin_update_user(
    request: Request,
    email: str = Form(""),
    squad: str = Form(""),
    role: str = Form("user"),
    active: str = Form(""),
) -> Response:
    actor = _request_user(request)
    if actor is None or not actor.is_admin:
        return _access_denied(request, "Apenas administradores podem alterar usuarios.")
    try:
        before = ensure_user(email)
        after = update_user(
            email,
            squad=squad,
            role=role,
            active=bool(active),
        )
        record_admin_event(
            actor.email,
            "alterou_usuario",
            after.email,
            {
                "antes": {"squad": before.squad, "role": before.role, "active": before.active},
                "depois": {"squad": after.squad, "role": after.role, "active": after.active},
            },
        )
    except ValueError as exc:
        return RedirectResponse(
            f"/admin/users?error={quote(str(exc), safe='')}",
            status_code=303,
        )
    return RedirectResponse(
        f"/admin/users?notice={quote(f'Usuario {after.email} atualizado.', safe='')}",
        status_code=303,
    )


@app.post("/ppt-summary")
async def ppt_summary(pptx: UploadFile = File(...)) -> JSONResponse:
    try:
        _validate_upload(pptx, ".pptx", "Envie um arquivo PPTX.")
        pptx_bytes = await _read_upload_limited(pptx, "PPTX")
        validate_pptx_bytes(pptx_bytes)
        summary = _inspect_ppt_upload(pptx_bytes)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)
    threshold = _large_deck_slide_threshold()
    return JSONResponse(
        {
            **summary,
            "large_slide_threshold": threshold,
            "requires_confirmation": summary["slide_count"] > threshold,
        }
    )


@app.post("/preview", response_class=HTMLResponse)
async def preview(
    request: Request,
    project_ref: str = Form(""),
    squad: str = Form("squad1"),
    project_name: str = Form(""),
    pptx: UploadFile = File(...),
    datasources: list[UploadFile] = File(...),
    mapping: UploadFile | None = File(None),
    mapping_template_ref: str = Form(""),
    use_ai: str = Form(""),
    slides_to_update: str = Form(""),
    confirm_large_deck: str = Form(""),
) -> HTMLResponse:
    try:
        squad = _authorized_form_squad(request, squad, project_ref)
        project = _resolve_project(project_ref, squad, project_name)
        mapping_template = _resolve_mapping_template(project, mapping_template_ref)
        selected_slides = parse_slide_selection(slides_to_update)
        _validate_upload(pptx, ".pptx", "Envie um arquivo PPTX.")
        if mapping and mapping.filename:
            _validate_upload(mapping, ".xlsx", "A planilha de mapeamento precisa ser XLSX.")
        pptx_bytes = await _read_upload_limited(pptx, "PPTX")
        datasource_payloads: list[tuple[str, bytes]] = []
        for upload in datasources:
            if not upload or not (upload.filename or "").strip():
                continue
            datasource_payloads.append(
                ((upload.filename or ""), await _read_upload_limited(upload, "Planilha"))
            )
        datasource_bytes, datasources_display = _coalesce_datasources_to_zip(datasource_payloads)
        mapping_bytes = await _read_upload_limited(mapping, "XLSX de mapeamento") if mapping and mapping.filename else b""
        validate_pptx_bytes(pptx_bytes)
        validate_datasource_zip_bytes(datasource_bytes)
        if mapping_bytes:
            validate_xlsx_bytes(mapping_bytes)
        ppt_summary = _inspect_ppt_upload(pptx_bytes)
        _validate_slide_scope(ppt_summary, selected_slides)
        requires_large_confirmation = _large_scope_requires_confirmation(ppt_summary, selected_slides)
        if requires_large_confirmation and not bool(confirm_large_deck):
            raise ValueError(_large_deck_confirmation_message(ppt_summary, selected_slides))
    except Exception as exc:
        return _error_response(request, str(exc), status_code=400)

    job_id = uuid.uuid4().hex
    job_dir = _job_dir(job_id, create=True)
    (job_dir / "input.pptx").write_bytes(pptx_bytes)
    (job_dir / "datasources.zip").write_bytes(datasource_bytes)
    if mapping_bytes:
        (job_dir / "mapping.xlsx").write_bytes(mapping_bytes)
    _reset_debug_log(job_dir)
    _log_job_debug_event(
        job_dir,
        "preview_request",
        {
            "job_id": job_id,
            "project": {"squad": project.squad, "slug": project.slug, "name": project.name},
            "files": {
                "pptx": {"name": pptx.filename or "modelo.pptx", "bytes": len(pptx_bytes)},
                "datasources": {"name": datasources_display, "bytes": len(datasource_bytes)},
                "mapping": {
                    "name": mapping.filename if mapping and mapping.filename else "",
                    "bytes": len(mapping_bytes),
                },
            },
            "slides_to_update": selected_slides,
            "use_ai_requested": bool(use_ai),
            "use_ai_effective": bool(use_ai) and not bool(mapping_template),
            "auto_source_review": _auto_source_match_review_enabled() and not bool(mapping_template),
            "auto_source_review_confidence_floor": _ai_review_confidence_floor(),
            "mapping_template": _mapping_template_metadata(mapping_template),
            "ppt_summary": ppt_summary,
            "combined_upload_bytes": len(pptx_bytes) + sum(len(data) for _name, data in datasource_payloads) + len(mapping_bytes),
        },
    )

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
                "datasources": datasources_display,
                "mapping": mapping.filename if mapping and mapping.filename else "",
            },
            "slides": {
                "raw": slides_to_update.strip(),
                "numbers": selected_slides,
            },
            "ppt_summary": ppt_summary,
            "combined_upload_bytes": len(pptx_bytes) + sum(len(data) for _name, data in datasource_payloads) + len(mapping_bytes),
            "large_deck_slide_threshold": _large_deck_slide_threshold(),
            "large_deck_confirmed": bool(confirm_large_deck),
            "mapping_template": _mapping_template_metadata(mapping_template),
            "use_ai": bool(use_ai) and not bool(mapping_template),
            "auto_source_review": _auto_source_match_review_enabled() and not bool(mapping_template),
            "ignore_mapping_candidates": False,
        },
    )
    _save_project_checkpoint(job_dir, status="in_progress", include_inputs=True, reason="preview_criado")
    _init_preview_processing_state(job_dir)
    _start_preview_processing(job_dir)
    return _render_preview(request, job_id, notice="Preview iniciado. Os slides serao preenchidos conforme o processamento terminar.")


@app.get("/projects/{squad}/{slug}/preview", response_class=HTMLResponse)
async def resume_project_preview(request: Request, squad: str, slug: str) -> HTMLResponse:
    try:
        project = load_project(squad, slug)
        if project is None:
            raise ValueError("Projeto nao encontrado.")
        job_id = _restore_project_checkpoint(project)
        job_dir = _job_dir(job_id)
        if not (job_dir / "render_cache.json").exists():
            _init_preview_processing_state(job_dir)
            _start_preview_processing(job_dir)
    except Exception as exc:
        return _error_response(request, str(exc), status_code=400)
    return _render_preview(
        request,
        job_id,
        notice="Checkpoint do projeto carregado.",
        prefer_cache=True,
        allow_ai=False,
    )


@app.get("/jobs/{job_id}/preview", response_class=HTMLResponse)
async def job_preview(request: Request, job_id: str, slide: int | None = None) -> HTMLResponse:
    return _render_preview(request, job_id, selected_preview_slide=slide, allow_ai=False)


@app.get("/jobs/{job_id}/processing-status")
async def job_processing_status(job_id: str) -> JSONResponse:
    job_dir = _job_dir(job_id)
    state = _load_preview_processing_state(job_dir)
    if not state:
        payload = {
            "job_id": job_id,
            "status": "complete",
            "active": False,
            "slides": {},
            "preview_url": f"/jobs/{job_id}/preview",
            "debug_log_url": f"/jobs/{job_id}/debug-log",
        }
    else:
        payload = {
            **state,
            "job_id": job_id,
            "active": _preview_processing_is_active(state),
            "preview_url": f"/jobs/{job_id}/preview",
            "debug_log_url": f"/jobs/{job_id}/debug-log",
        }
    _log_job_debug_event(
        job_dir,
        "processing_status_response",
        {
            "status": payload.get("status"),
            "active": payload.get("active"),
            "current_stage": payload.get("current_stage"),
            "message": payload.get("message"),
            "totals": payload.get("totals"),
            "slides": payload.get("slides"),
            "preview_url": payload.get("preview_url"),
        },
    )
    return JSONResponse(payload)


@app.get("/jobs/{job_id}/debug-log")
async def job_debug_log(job_id: str) -> FileResponse:
    job_dir = _job_dir(job_id)
    path = _debug_log_path(job_dir)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Log de debug ainda nao foi criado.")
    _log_job_debug_event(
        job_dir,
        "debug_log_downloaded",
        {"current_processing_state": _load_preview_processing_state(job_dir)},
    )
    return FileResponse(path, media_type="text/plain; charset=utf-8", filename="log.txt")


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
        metadata["use_ai"] = False
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
    notice = f"Slides adicionados ao escopo: {', '.join(str(slide) for slide in added_slides)}."
    return _render_preview(request, job_id, notice=notice)


@app.post("/jobs/{job_id}/mapping-template", response_class=HTMLResponse)
async def apply_job_mapping_template(
    request: Request,
    job_id: str,
    mapping_template_ref: str = Form(""),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        metadata = _load_job_metadata(job_dir)
        project_meta = metadata.get("project") or {}
        project = load_project(project_meta.get("squad", ""), project_meta.get("slug", ""))
        if project is None:
            raise ValueError("Projeto nao encontrado.")
        mapping_template = _resolve_mapping_template(project, mapping_template_ref)
        if not mapping_template:
            raise ValueError("Selecione um mapeamento salvo.")
        metadata["mapping_template"] = _mapping_template_metadata(mapping_template)
        metadata["use_ai"] = False
        metadata["ignore_mapping_candidates"] = False
        _save_job_metadata(job_dir, metadata)
        _clear_ai_cache(job_dir)
        _clear_render_cache(job_dir)
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc))
    return _render_preview(
        request,
        job_id,
        notice=f"Mapeamento '{mapping_template.get('name')}' aplicado neste projeto.",
        allow_ai=False,
    )


@app.post("/jobs/{job_id}/review-ai", response_class=HTMLResponse)
async def review_job_with_ai(request: Request, job_id: str) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        metadata = _load_job_metadata(job_dir)
        if not ai_configured(PROJECT_ROOT):
            raise ValueError("IA indisponivel: configure OPENAI_API_KEY no .env.")
        metadata["ignore_mapping_candidates"] = True
        metadata["use_ai"] = False
        _save_job_metadata(job_dir, metadata)
        _clear_render_cache(job_dir)
        selected_slide = _preview_slide_from_request(request)
        slide_scope = [selected_slide] if selected_slide else None
        notice = await run_in_threadpool(_run_automatic_slide_ai_review, job_dir, slide_numbers=slide_scope, force=True)
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc))
    return _render_preview(
        request,
        job_id,
        notice=notice or "IA automatica executada por slide.",
        allow_ai=False,
    )


@app.post("/jobs/{job_id}/targets/{target_id}/review-ai", response_class=HTMLResponse)
async def review_target_with_ai(
    request: Request,
    job_id: str,
    target_id: str,
    manual_context: str = Form(""),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        if not ai_configured(PROJECT_ROOT):
            raise ValueError("IA indisponivel: configure OPENAI_API_KEY no .env.")
        analysis, _mapping_status, _mapping_candidates, _pause = await run_in_threadpool(
            _analysis_for_job, job_dir, apply_slide_outputs=False
        )
        canonical_target_id = _canonical_target_id(analysis.targets, target_id)
        _validate_target_id(canonical_target_id)
        metadata = _load_job_metadata(job_dir)
        metadata["use_ai"] = False
        metadata["ignore_mapping_candidates"] = True
        _save_job_metadata(job_dir, metadata)
        _clear_render_cache(job_dir)
        notice = await run_in_threadpool(
            _run_target_ai_review, job_dir, canonical_target_id, manual_context=manual_context
        )
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc), allow_ai=False)
    return _render_preview(request, job_id, notice=notice, allow_ai=False)


@app.post("/jobs/{job_id}/slides/{slide_number}/review-ai", response_class=HTMLResponse)
async def review_slide_with_ai(
    request: Request,
    job_id: str,
    slide_number: int,
    manual_context: str = Form(""),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        if not ai_configured(PROJECT_ROOT):
            raise ValueError("IA indisponivel: configure OPENAI_API_KEY no .env.")
        notice = await run_in_threadpool(
            _run_slide_ai_review, job_dir, slide_number, manual_context=manual_context
        )
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc), allow_ai=False)
    return _render_preview(request, job_id, notice=notice, allow_ai=False)


@app.post("/jobs/{job_id}/targets/{target_id}/approve", response_class=HTMLResponse)
async def approve_target(request: Request, job_id: str, target_id: str) -> HTMLResponse:
    return _set_target_state_response(request, job_id, target_id, approved=True, skipped=False)


@app.post("/jobs/{job_id}/targets/{target_id}/skip", response_class=HTMLResponse)
async def skip_target(request: Request, job_id: str, target_id: str) -> HTMLResponse:
    return _set_target_state_response(request, job_id, target_id, approved=False, skipped=True)


@app.post("/jobs/{job_id}/slides/{slide_number}/approve", response_class=HTMLResponse)
async def approve_slide(request: Request, job_id: str, slide_number: int) -> HTMLResponse:
    return _set_slide_state_response(request, job_id, slide_number, approved=True, skipped=False)


@app.post("/jobs/{job_id}/slides/{slide_number}/skip", response_class=HTMLResponse)
async def skip_slide(request: Request, job_id: str, slide_number: int) -> HTMLResponse:
    return _set_slide_state_response(request, job_id, slide_number, approved=False, skipped=True)


@app.post("/jobs/{job_id}/targets/{target_id}/override", response_class=HTMLResponse)
async def override_target_datasource(
    request: Request,
    job_id: str,
    target_id: str,
    datasource: UploadFile | None = File(None),
    existing_source: str = Form(""),
    cell_range: str = Form(""),
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        # O formulario do preview sempre envia o ID canonico. Reanalisar o deck
        # inteiro aqui só para validar esse ID duplicava o trabalho: a nova
        # análise já acontece uma vez, depois de salvar o override.
        _validate_target_id(target_id)
        _validate_cell_range(cell_range)
        if existing_source.strip():
            filename, data = _read_existing_datasource(job_dir, existing_source.strip())
        else:
            if datasource is None or not (datasource.filename or "").strip():
                raise ValueError("Escolha uma planilha existente ou envie um novo XLSX.")
            _validate_upload(datasource, ".xlsx", "Envie um XLSX para substituir a planilha deste gráfico.")
            filename = datasource.filename or f"{target_id}.xlsx"
            data = await _read_upload_limited(datasource, "XLSX")
        validate_xlsx_bytes(data)
        target_dir = job_dir / "overrides" / target_id
        target_dir.mkdir(parents=True, exist_ok=True)
        for existing in target_dir.glob("*.xlsx"):
            existing.unlink()
        filename = safe_filename(filename or f"{target_id}.xlsx")
        (target_dir / filename).write_bytes(data)
        range_path = target_dir / "range.txt"
        if cell_range.strip():
            range_path.write_text(cell_range.strip(), encoding="utf-8")
        elif range_path.exists():
            range_path.unlink()
        metadata = _load_job_metadata(job_dir)
        metadata["use_ai"] = False
        metadata["ignore_mapping_candidates"] = True
        _save_job_metadata(job_dir, metadata)
        _clear_ai_cache(job_dir, target_id=target_id)
        _clear_render_cache(job_dir)
        audit.record(
            job_dir,
            _actor(request),
            "trocou_planilha_do_grafico",
            {
                "target": target_id,
                "planilha": filename,
                "intervalo": cell_range.strip(),
                "origem": "planilha ja enviada" if existing_source.strip() else "upload novo",
            },
        )
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc))
    range_notice = f" com range {cell_range.strip()}" if cell_range.strip() else ""
    return _render_preview(
        request,
        job_id,
        notice=(
            f"Datasource {filename}{range_notice} aplicado ao target {target_id}. "
            "A correção determinística foi salva; a revisão por IA continua disponível no modo avançado."
        ),
        allow_ai=False,
    )


@app.post("/jobs/{job_id}/targets/{target_id}/chart-formats", response_class=HTMLResponse)
async def update_chart_formats(request: Request, job_id: str, target_id: str) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        _validate_target_id(target_id)
        form = await request.form()
        allowed = {"auto", "percent", "number"}
        indexes = sorted(
            {
                int(match.group(1))
                for key in form.keys()
                if (match := re.fullmatch(r"format_(\d+)", str(key))) is not None
            }
        )
        if not indexes or len(indexes) > 256:
            raise ValueError("Series do grafico nao encontradas.")
        labels = {
            index: str(form.get(f"series_label_{index}") or "").strip()[:200]
            for index in indexes
        }
        label_counts = {
            label: sum(1 for candidate in labels.values() if candidate == label)
            for label in labels.values()
            if label
        }
        overrides: dict[str, str] = {}
        for index in indexes:
            mode = str(form.get(f"format_{index}") or "auto").strip().lower()
            if mode not in allowed:
                raise ValueError("Formato invalido.")
            if mode != "auto":
                label = labels[index]
                key = label if label and label_counts.get(label) == 1 else f"__index_{index}"
                overrides[key] = mode

        metadata = _load_job_metadata(job_dir)
        all_overrides = dict(metadata.get("chart_format_overrides") or {})
        if overrides:
            all_overrides[target_id] = overrides
        else:
            all_overrides.pop(target_id, None)
        metadata["chart_format_overrides"] = all_overrides
        _save_job_metadata(job_dir, metadata)
        _clear_render_cache(job_dir)
        _discard_stale_generated_output(job_dir)
        audit.record(
            job_dir,
            _actor(request),
            "ajustou_formato_do_grafico",
            {"target": target_id, "series": overrides},
        )
        _save_project_checkpoint(job_dir, status="in_progress", reason="formato_grafico")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc), allow_ai=False)
    return _render_preview(
        request,
        job_id,
        notice="Formato das series salvo. Automático continua usando primeiro o contrato do PowerPoint.",
        allow_ai=False,
    )


@app.get("/jobs/{job_id}/download")
async def download(request: Request, job_id: str) -> Response:
    job_dir = _job_dir(job_id)
    audit.remember_actor(job_dir, _actor(request))
    if not _generated_is_current(job_dir):
        _discard_stale_generated_output(job_dir)
        await run_in_threadpool(_generate_job_output, job_dir)
    return FileResponse(
        _generated_ppt_path(job_dir),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=_generated_filename(job_dir),
    )


@app.post("/jobs/{job_id}/generate")
async def start_generation(request: Request, job_id: str) -> JSONResponse:
    job_dir = _job_dir(job_id)
    audit.remember_actor(job_dir, _actor(request))
    input_signature = _generation_input_signature(job_dir)
    if _generated_is_current(job_dir, input_signature=input_signature):
        return JSONResponse({"status": "complete", "download_url": f"/jobs/{job_id}/download"})
    _discard_stale_generated_output(job_dir)
    state = _load_generation_state(job_dir)
    if state.get("active") and state.get("input_signature") == input_signature:
        return JSONResponse({"status": state.get("status", "running"), "status_url": f"/jobs/{job_id}/generation-status"})

    _init_generation_state(job_dir, input_signature=input_signature)
    with GENERATION_RUNNING_LOCK:
        if job_id not in GENERATION_RUNNING:
            GENERATION_RUNNING.add(job_id)
            PREVIEW_EXECUTOR.submit(_generate_job_worker, job_dir)
    return JSONResponse({"status": "queued", "status_url": f"/jobs/{job_id}/generation-status"}, status_code=202)


@app.post("/jobs/{job_id}/save")
async def save_job_checkpoint(request: Request, job_id: str) -> JSONResponse:
    job_dir = _job_dir(job_id)
    actor = _actor(request)
    await run_in_threadpool(
        _save_project_checkpoint,
        job_dir,
        "in_progress",
        False,
        "salvamento_manual",
    )
    audit.record(job_dir, actor, "salvou_checkpoint_manual")
    checkpoint = _load_job_checkpoint_summary(job_dir)
    return JSONResponse(
        {
            "status": "saved",
            "saved_at": checkpoint.get("updated_at") or _now_iso(),
            "message": "Trabalho salvo. Voce pode voltar depois pela pagina inicial.",
        }
    )


@app.get("/jobs/{job_id}/generation-status")
async def generation_status(job_id: str) -> JSONResponse:
    job_dir = _job_dir(job_id)
    state = _load_generation_state(job_dir)
    return JSONResponse(
        {
            **state,
            "job_id": job_id,
            "download_url": f"/jobs/{job_id}/download" if _generated_is_current(job_dir) else "",
        }
    )


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/health/live")
async def health_live() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/health/ready")
async def health_ready() -> JSONResponse:
    try:
        RUNTIME_ROOT.mkdir(parents=True, exist_ok=True)
        probe = RUNTIME_ROOT / ".readiness"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError:
        return JSONResponse({"status": "not_ready"}, status_code=503)
    return JSONResponse({"status": "ok"})


def _preview_processing_path(job_dir: Path) -> Path:
    return job_dir / "preview_processing.json"


def _debug_log_path(job_dir: Path) -> Path:
    return job_dir / "log.txt"


def _generation_processing_path(job_dir: Path) -> Path:
    return job_dir / "generation_processing.json"


def _generated_ppt_path(job_dir: Path) -> Path:
    return job_dir / "generated.pptx"


def _generated_metadata_path(job_dir: Path) -> Path:
    return job_dir / "generated.json"


def _generation_input_signature(job_dir: Path, *, include_mapping_template: bool = True) -> str:
    """Assina somente o estado que pode alterar os bytes do PPT final."""
    digest = hashlib.sha256(
        b"auto-ppt-generation-v2-mapping-" + (b"1" if include_mapping_template else b"0")
    )

    def add_bytes(label: str, data: bytes) -> None:
        digest.update(label.encode("utf-8"))
        digest.update(b"\0")
        digest.update(data)
        digest.update(b"\0")

    # Entradas grandes sao imutaveis dentro de um job. Assinar stat evita reler
    # decks de centenas de MB a cada consulta de status/download.
    for name in ("input.pptx", "datasources.zip"):
        path = job_dir / name
        if path.exists() and path.is_file():
            stat = path.stat()
            add_bytes(f"{name}-stat", f"{stat.st_size}:{stat.st_mtime_ns}".encode("ascii"))

    for name in ("ai_source_matches.json", "ai_diagnostics.json"):
        path = job_dir / name
        if path.exists() and path.is_file():
            add_bytes(name, path.read_bytes())

    metadata_path = job_dir / "metadata.json"
    if metadata_path.exists():
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            metadata_keys = ["slides", "apply_slide_ai_outputs", "chart_format_overrides"]
            if include_mapping_template:
                metadata_keys.append("mapping_template")
            relevant_metadata = {key: metadata.get(key) for key in metadata_keys}
            add_bytes(
                "metadata-relevant",
                json.dumps(relevant_metadata, ensure_ascii=False, sort_keys=True).encode("utf-8"),
            )
        except (OSError, json.JSONDecodeError):
            add_bytes("metadata.json", metadata_path.read_bytes())

    slide_state_path = job_dir / "slide_ai_state.json"
    if slide_state_path.exists():
        try:
            slide_state = json.loads(slide_state_path.read_text(encoding="utf-8"))
            target_outputs = {
                slide: (state or {}).get("target_outputs") or {}
                for slide, state in (slide_state.get("slides") or {}).items()
            }
            add_bytes(
                "slide-ai-target-outputs",
                json.dumps(target_outputs, ensure_ascii=False, sort_keys=True).encode("utf-8"),
            )
        except (OSError, json.JSONDecodeError):
            add_bytes("slide_ai_state.json", slide_state_path.read_bytes())

    overrides_dir = job_dir / "overrides"
    if overrides_dir.exists():
        for path in sorted(item for item in overrides_dir.rglob("*") if item.is_file()):
            add_bytes(path.relative_to(job_dir).as_posix(), path.read_bytes())
    return digest.hexdigest()


def _generated_is_current(job_dir: Path, input_signature: str | None = None) -> bool:
    generated = _generated_ppt_path(job_dir)
    metadata_path = _generated_metadata_path(job_dir)
    if not generated.exists() or not metadata_path.exists():
        return False
    try:
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    expected = input_signature or _generation_input_signature(job_dir)
    return bool(expected) and metadata.get("input_signature") == expected


def _discard_stale_generated_output(job_dir: Path) -> None:
    if _generated_is_current(job_dir):
        return
    for path in (_generated_ppt_path(job_dir), _generated_metadata_path(job_dir)):
        if path.exists() and path.is_file():
            path.unlink()


OUTPUT_NAME_SEPARATOR = "__"


def _output_file_name(job_dir: Path) -> str:
    """Mantem o nome do PPT enviado e acrescenta so um sufixo de data/hora.

    Os analistas reaproveitam o mesmo nome de deck todo mes, mudando uma
    variavel; trocar o nome por 'ppt_automatizado' obrigava a renomear tudo de
    volta. O separador '__' e visivel e nao aparece em nome de arquivo normal,
    entao da para apagar do sufixo em diante sem pensar."""
    try:
        metadata = _load_job_metadata(job_dir)
        original = str((metadata.get("files") or {}).get("pptx") or "")
    except Exception:
        original = ""
    stem = Path(safe_filename(original or "ppt_atualizado.pptx")).stem or "ppt_atualizado"
    # Nao empilha sufixo quando o arquivo enviado ja saiu daqui antes.
    stem = stem.split(OUTPUT_NAME_SEPARATOR)[0].strip() or "ppt_atualizado"
    return f"{stem}{OUTPUT_NAME_SEPARATOR}{datetime.now().strftime('%Y-%m-%d_%H%M')}.pptx"


def _generated_filename(job_dir: Path) -> str:
    try:
        return str(json.loads(_generated_metadata_path(job_dir).read_text(encoding="utf-8")).get("file_name") or "ppt_automatizado.pptx")
    except Exception:
        return "ppt_automatizado.pptx"


def _load_generation_state(job_dir: Path) -> dict:
    path = _generation_processing_path(job_dir)
    if not path.exists():
        return {"status": "idle", "active": False, "message": ""}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"status": "error", "active": False, "message": "Estado de geracao invalido."}


def _save_generation_state(job_dir: Path, state: dict) -> None:
    state = {**state, "updated_at": _now_iso()}
    path = _generation_processing_path(job_dir)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def _init_generation_state(job_dir: Path, input_signature: str = "") -> None:
    _save_generation_state(
        job_dir,
        {
            "status": "queued",
            "active": True,
            "message": "Aguardando geracao do PPT.",
            "created_at": _now_iso(),
            "input_signature": input_signature or _generation_input_signature(job_dir),
            "progress": {"completed": 0, "total": 0, "percent": 0, "phase": "queued"},
        },
    )


def _generate_job_worker(job_dir: Path) -> None:
    try:
        _generate_job_output(job_dir)
    except Exception as exc:
        state = _load_generation_state(job_dir)
        state.update({"status": "error", "active": False, "message": str(exc)})
        _save_generation_state(job_dir, state)
        _log_job_debug_event(job_dir, "generation_worker_error", {"error": repr(exc)})
    finally:
        with GENERATION_RUNNING_LOCK:
            GENERATION_RUNNING.discard(job_dir.name)


def _generate_job_output(job_dir: Path) -> None:
    pptx_path = job_dir / "input.pptx"
    datasource_path = job_dir / "datasources.zip"
    if not pptx_path.exists() or not datasource_path.exists():
        raise FileNotFoundError("Job nao encontrado.")
    input_signature = _generation_input_signature(job_dir)
    state = _load_generation_state(job_dir)
    state.update({"status": "running", "active": True, "message": "Gerando PPT atualizado."})
    state["input_signature"] = input_signature
    _save_generation_state(job_dir, state)
    manual_sources = _manual_sources_for_job(job_dir)
    selected_slides = _selected_slides_for_job(job_dir)
    analysis = _cached_analyze_files(job_dir, manual_sources, selected_slides)
    analysis, _mapping_status = _apply_mapping_template_to_analysis(job_dir, analysis, skip_targets=set(manual_sources))
    ai_matches, _ai_match_status = _ai_source_matches_for_job(job_dir, analysis, allow_ai=False)
    analysis = apply_ai_source_matches_to_analysis(analysis, ai_matches)
    ai_diagnostics, _ai_status = _ai_diagnostics_for_job(job_dir, analysis, allow_ai=False)
    analysis = apply_ai_recommendations_to_analysis(analysis, ai_diagnostics)
    analysis = apply_typed_outputs_to_analysis(analysis, _slide_ai_target_outputs(job_dir))
    analysis = _apply_chart_format_overrides(analysis, _load_job_metadata(job_dir))
    try:
        output = generate_updated_pptx(
            pptx_path.read_bytes(),
            analysis.plans,
            targets=analysis.targets,
            progress_callback=lambda payload: _record_generation_progress(job_dir, payload),
        )
    except (EmbeddedWorkbookWriterUnavailable, ChartSheetUnresolvedError) as exc:
        raise RuntimeError(str(exc)) from exc
    file_name = _output_file_name(job_dir)
    published_signature = _publish_generated_output(
        job_dir,
        output,
        analysis,
        file_name,
        input_signature,
    )
    state = _load_generation_state(job_dir)
    state.update(
        {
            "status": "complete",
            "active": False,
            "message": "PPT pronto para download.",
            "input_signature": published_signature,
            "progress": {**(state.get("progress") or {}), "percent": 100, "phase": "complete"},
        }
    )
    _save_generation_state(job_dir, state)


def _publish_generated_output(
    job_dir: Path,
    output: bytes,
    analysis: AnalysisResult,
    file_name: str,
    source_signature: str,
) -> str:
    if _generation_input_signature(job_dir) != source_signature:
        raise RuntimeError("O projeto foi alterado durante a geracao. Gere novamente para incluir as mudancas.")

    # Salvar o run treina/seleciona o mapping template usado na proxima rodada.
    # Essa mutacao interna nao invalida o PPT que acabou de ser gerado: o template
    # foi derivado dos mesmos planos. Alteracoes de conteudo feitas em paralelo
    # (slides, overrides ou saidas IA) continuam invalidando a publicacao.
    content_signature = _generation_input_signature(job_dir, include_mapping_template=False)
    _save_project_run(job_dir, output, analysis, file_name)
    _save_project_checkpoint(job_dir, status="completed")
    if _generation_input_signature(job_dir, include_mapping_template=False) != content_signature:
        raise RuntimeError("O projeto foi alterado durante a geracao. Gere novamente para incluir as mudancas.")

    published_signature = _generation_input_signature(job_dir)
    generated_path = _generated_ppt_path(job_dir)
    generated_metadata_path = _generated_metadata_path(job_dir)
    generated_tmp = generated_path.with_name(f"{generated_path.name}.tmp")
    metadata_tmp = generated_metadata_path.with_name(f"{generated_metadata_path.name}.tmp")
    generated_tmp.write_bytes(output)
    metadata_tmp.write_text(
        json.dumps({"file_name": file_name, "input_signature": published_signature}, ensure_ascii=False),
        encoding="utf-8",
    )
    generated_tmp.replace(generated_path)
    metadata_tmp.replace(generated_metadata_path)
    return published_signature


def _record_generation_progress(job_dir: Path, payload: dict) -> None:
    completed = max(int(payload.get("completed") or 0), 0)
    total = max(int(payload.get("total") or 0), 0)
    phase = str(payload.get("phase") or "targets")
    percent = _object_progress_percent(completed, total, phase)
    state = _load_generation_state(job_dir)
    state.update(
        {
            "status": "running",
            "active": True,
            "message": str(payload.get("message") or state.get("message") or "Gerando PPT atualizado."),
            "progress": {
                "completed": completed,
                "total": total,
                "percent": min(max(percent, 0), 100),
                "phase": phase,
                "slide": payload.get("slide"),
                "target_id": payload.get("target_id"),
            },
        }
    )
    _save_generation_state(job_dir, state)


def _reset_debug_log(job_dir: Path) -> None:
    _debug_log_path(job_dir).write_text("", encoding="utf-8")


def _log_job_debug_event(job_dir: Path, event: str, payload: dict | None = None) -> None:
    token = set_ai_debug_log_path(_debug_log_path(job_dir))
    try:
        log_debug_event(event, payload or {})
    finally:
        reset_ai_debug_log_path(token)


def _call_with_job_debug(job_dir: Path, callback, *args, **kwargs):
    debug_token = set_ai_debug_log_path(_debug_log_path(job_dir))
    usage_token = set_ai_usage_log_path(job_dir / "logs" / "ai_usage.jsonl")
    try:
        return callback(*args, **kwargs)
    finally:
        reset_ai_usage_log_path(usage_token)
        reset_ai_debug_log_path(debug_token)


def _init_preview_processing_state(job_dir: Path) -> None:
    try:
        metadata = _load_job_metadata(job_dir)
    except Exception:
        metadata = {}
    slides = _preview_scope_slides(metadata)
    now = _now_iso()
    state = {
        "schema_version": 1,
        "status": "queued",
        "active": True,
        "created_at": now,
        "updated_at": now,
        "message": "Aguardando processamento do preview.",
        "current_stage": "queued",
        "slides": {
            str(slide): {
                "slide": slide,
                "status": "queued",
                "stage": "queued",
                "message": "Na fila.",
                "target_count": 0,
                "mapped_count": 0,
                "unmapped_count": 0,
            }
            for slide in slides
        },
        "totals": {
            "slides": len(slides),
            "targets": 0,
            "mapped": 0,
            "unmapped": 0,
        },
        "progress": {
            "completed": 0,
            "total": 0,
            "percent": 0,
            "phase": "queued",
            "message": "Preparando objetos do preview.",
        },
    }
    _save_preview_processing_state(job_dir, state)
    _log_job_debug_event(
        job_dir,
        "preview_state_initialized",
        {"slides": slides, "slide_count": len(slides)},
    )


def _start_preview_processing(job_dir: Path) -> None:
    job_id = job_dir.name
    with PREVIEW_RUNNING_LOCK:
        if job_id in PREVIEW_RUNNING:
            _log_job_debug_event(job_dir, "preview_worker_already_running", {"job_id": job_id})
            return
        PREVIEW_RUNNING.add(job_id)
    _log_job_debug_event(job_dir, "preview_worker_submitted", {"job_id": job_id})
    PREVIEW_EXECUTOR.submit(_preview_processing_worker, job_dir)


def _preview_processing_worker(job_dir: Path) -> None:
    job_id = job_dir.name
    debug_token = set_ai_debug_log_path(_debug_log_path(job_dir))
    worker_started = time.perf_counter()
    try:
        log_debug_event("preview_worker_start", {"job_id": job_id})
        _update_preview_processing_state(
            job_dir,
            status="running",
            current_stage="analysis",
            message="Lendo PPT, XLSX do escopo e contratos do Editar dados.",
            slide_status="running",
            slide_message="Analisando objetos e datasources deste slide.",
        )
        analysis_started = time.perf_counter()
        log_debug_event("analysis_start", {"apply_cached_source_matches": True, "apply_slide_outputs": False})
        analysis, mapping_status, mapping_candidates, pause_for_mapping = _analysis_for_job(
            job_dir,
            apply_cached_source_matches=True,
            apply_cached_diagnostics=False,
            apply_slide_outputs=False,
            progress_callback=lambda payload: _record_preview_progress(job_dir, payload),
        )
        log_debug_event(
            "analysis_done",
            {
                "elapsed_ms": round((time.perf_counter() - analysis_started) * 1000),
                "target_count": analysis.target_count,
                "source_count": analysis.source_count,
                "plan_count": len(analysis.plans),
                "warning_count": len(analysis.warnings),
                "warnings": analysis.warnings,
                "mapping_status": mapping_status,
                "mapping_candidate_count": len(mapping_candidates),
                "pause_for_mapping": pause_for_mapping,
            },
        )
        ai_match_status, ai_diagnostic_status = _ai_waiting_status(pause_for_mapping)
        _mark_processing_analysis_done(job_dir, analysis)

        metadata = _load_job_metadata(job_dir)
        source_match_ai_enabled = _source_match_ai_enabled(metadata)
        if source_match_ai_enabled and ai_configured(PROJECT_ROOT):
            ai_started = time.perf_counter()
            log_debug_event(
                "preview_ai_match_start",
                {
                    "operation": "source_match",
                    "use_ai": bool(metadata.get("use_ai")),
                    "auto_source_review": bool(metadata.get("auto_source_review")),
                    "confidence_floor": _ai_review_confidence_floor(),
                },
            )
            _update_preview_processing_state(
                job_dir,
                status="running",
                current_stage="ai_match",
                message="IA enxuta escolhendo datasources pendentes e sugerindo receitas estruturais.",
                slide_status="running",
                slide_message="Revisando matches pendentes ou abaixo do corte de confianca com IA.",
                only_unfinished=False,
            )
            ai_matches, ai_match_status = _ai_source_matches_for_job(job_dir, analysis, allow_ai=True)
            analysis = apply_ai_source_matches_to_analysis(analysis, ai_matches)
            log_debug_event(
                "preview_ai_match_done",
                {
                    "elapsed_ms": round((time.perf_counter() - ai_started) * 1000),
                    "status": ai_match_status,
                    "match_count": len(ai_matches),
                    "plan_count_after_ai": len(analysis.plans),
                },
            )
            _mark_processing_analysis_done(job_dir, analysis)
        else:
            log_debug_event(
                "preview_ai_match_skipped",
                {
                    "use_ai": bool(metadata.get("use_ai")),
                    "auto_source_review": bool(metadata.get("auto_source_review")),
                    "ai_configured": ai_configured(PROJECT_ROOT),
                },
            )

        _clear_render_cache(job_dir)
        _update_preview_processing_state(
            job_dir,
            status="complete",
            current_stage="complete",
            message="Preview pronto.",
            slide_status="done",
            slide_message="Cards prontos.",
            only_unfinished=False,
        )
        try:
            _save_completed_preview_cache(
                job_dir,
                analysis,
                mapping_status,
                mapping_candidates,
                _combine_ai_status(ai_match_status, ai_diagnostic_status),
            )
        except Exception as cache_exc:
            # Cache is only a speed-up for the first completed render.
            log_debug_event("render_cache_error", {"error": repr(cache_exc)})
            pass
        _save_project_checkpoint(job_dir, status="in_progress")
        log_debug_event(
            "preview_worker_complete",
            {
                "elapsed_ms": round((time.perf_counter() - worker_started) * 1000),
                "target_count": analysis.target_count,
                "source_count": analysis.source_count,
                "plan_count": len(analysis.plans),
            },
        )
    except Exception as exc:
        log_debug_event(
            "preview_worker_error",
            {
                "elapsed_ms": round((time.perf_counter() - worker_started) * 1000),
                "error": repr(exc),
            },
        )
        _update_preview_processing_state(
            job_dir,
            status="error",
            current_stage="error",
            message=f"Falha ao preparar preview: {exc}",
            slide_status="error",
            slide_message=str(exc),
            only_unfinished=False,
        )
    finally:
        with PREVIEW_RUNNING_LOCK:
            PREVIEW_RUNNING.discard(job_id)
        log_debug_event(
            "preview_worker_finished",
            {"elapsed_ms": round((time.perf_counter() - worker_started) * 1000)},
        )
        reset_ai_debug_log_path(debug_token)


def _load_preview_processing_state(job_dir: Path) -> dict:
    path = _preview_processing_path(job_dir)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def _save_preview_processing_state(job_dir: Path, state: dict) -> None:
    state["updated_at"] = _now_iso()
    path = _preview_processing_path(job_dir)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def _record_preview_progress(job_dir: Path, payload: dict) -> None:
    state = _load_preview_processing_state(job_dir) or {}
    completed = max(int(payload.get("completed") or 0), 0)
    total = max(int(payload.get("total") or 0), 0)
    phase = str(payload.get("phase") or "analysis")
    percent = _object_progress_percent(completed, total, phase)
    state["progress"] = {
        "completed": completed,
        "total": total,
        "percent": min(max(percent, 0), 100),
        "phase": phase,
        "slide": payload.get("slide"),
        "target_id": payload.get("target_id"),
        "message": str(payload.get("message") or "Analisando objetos do preview."),
    }
    state["message"] = state["progress"]["message"]
    _save_preview_processing_state(job_dir, state)


def _object_progress_percent(completed: int, total: int, phase: str) -> int:
    """100% significa job pronto, tanto no preview quanto na geracao.

    A unidade final representa o trabalho posterior aos objetos (revisao e
    persistencia no preview; empacotamento na geracao).
    """
    if phase == "complete":
        return 100
    return min(max(round(100 * max(completed, 0) / max(max(total, 0) + 1, 1)), 0), 100)


def _update_preview_processing_state(
    job_dir: Path,
    status: str,
    current_stage: str,
    message: str,
    slide_status: str | None = None,
    slide_message: str = "",
    only_unfinished: bool = True,
) -> None:
    state = _load_preview_processing_state(job_dir) or {}
    state["status"] = status
    state["active"] = status in {"queued", "running"}
    state["current_stage"] = current_stage
    state["message"] = message
    if state.get("progress"):
        state["progress"]["message"] = message
        if status == "complete":
            state["progress"]["phase"] = "complete"
            state["progress"]["percent"] = 100
    log_debug_event(
        "processing_state_update",
        {
            "status": status,
            "current_stage": current_stage,
            "message": message,
            "slide_status": slide_status,
            "slide_message": slide_message,
            "only_unfinished": only_unfinished,
        },
    )
    if slide_status:
        for slide_state in (state.get("slides") or {}).values():
            if only_unfinished and slide_state.get("status") == "done":
                continue
            slide_state["status"] = slide_status
            slide_state["stage"] = current_stage
            slide_state["message"] = slide_message or message
    _save_preview_processing_state(job_dir, state)


def _mark_processing_analysis_done(job_dir: Path, analysis: AnalysisResult) -> None:
    state = _load_preview_processing_state(job_dir) or {}
    slides_state = state.setdefault("slides", {})
    plans_by_slide: dict[int, int] = defaultdict(int)
    targets_by_slide: dict[int, int] = defaultdict(int)
    for target in analysis.targets:
        if target.object_type in {"chart", "table"}:
            targets_by_slide[target.slide_number] += 1
    for plan in analysis.plans:
        plans_by_slide[plan.target.slide_number] += 1
    for slide, target_count in targets_by_slide.items():
        mapped_count = plans_by_slide.get(slide, 0)
        slide_state = slides_state.setdefault(str(slide), {"slide": slide})
        slide_state.update(
            {
                "slide": slide,
                "status": "running",
                "stage": "analysis_done",
                "message": "Mapeamento deterministico pronto; aguardando proxima etapa.",
                "target_count": target_count,
                "mapped_count": mapped_count,
                "unmapped_count": max(target_count - mapped_count, 0),
            }
        )
    state["totals"] = {
        "slides": len(slides_state),
        "targets": sum(item.get("target_count", 0) for item in slides_state.values()),
        "mapped": sum(item.get("mapped_count", 0) for item in slides_state.values()),
        "unmapped": sum(item.get("unmapped_count", 0) for item in slides_state.values()),
    }
    state["status"] = "running"
    state["active"] = True
    state["current_stage"] = "analysis"
    state["message"] = "Analise deterministica concluida."
    _save_preview_processing_state(job_dir, state)


def _preview_processing_is_active(state: dict) -> bool:
    return bool(state) and str(state.get("status") or "") in {"queued", "running"}


def _preview_scope_slides(metadata: dict) -> list[int]:
    selected = ((metadata.get("slides") or {}).get("numbers") or []) if isinstance(metadata.get("slides"), dict) else []
    if selected:
        return sorted({int(slide) for slide in selected if int(slide) > 0})
    slide_count = int(((metadata.get("ppt_summary") or {}).get("slide_count") or 0))
    return list(range(1, slide_count + 1)) if slide_count else []


def _processing_preview_context(
    request: Request,
    job_id: str,
    metadata: dict,
    state: dict,
    notice: str = "",
    error: str = "",
    selected_preview_slide: int | None = None,
) -> dict:
    all_slides = _processing_slide_items(state)
    all_slide_numbers = [item["slide"] for item in all_slides]
    full_render_limit = _preview_full_render_slide_limit()
    preview_is_windowed = len(all_slide_numbers) > full_render_limit
    preview_selected_slide = _preview_selected_slide(
        all_slide_numbers,
        selected_preview_slide if selected_preview_slide is not None else _preview_slide_from_request(request),
    )
    if preview_is_windowed and preview_selected_slide is not None:
        visible_slides = [item for item in all_slides if item["slide"] == preview_selected_slide]
    else:
        visible_slides = all_slides
    totals = state.get("totals") or {}
    review_summary = {
        "slides": int(totals.get("slides") or len(all_slides)),
        "targets": int(totals.get("targets") or 0),
        "mapped": int(totals.get("mapped") or 0),
        "unmapped": int(totals.get("unmapped") or 0),
    }
    if _source_match_ai_enabled(metadata) and ai_configured(PROJECT_ROOT):
        ai_status = {"state": "running", "message": state.get("message") or "Processando preview em background."}
    elif _source_match_ai_enabled(metadata):
        ai_status = {"state": "disabled", "message": "IA indisponivel: configure OPENAI_API_KEY no .env."}
    else:
        ai_status = {"state": "disabled", "message": "IA nao foi acionada neste preview."}
    mapping_status = {"state": "running", "message": "Mapeamento sera exibido conforme os slides terminarem."}
    return {
        "job_id": job_id,
        "metadata": metadata,
        "squad": metadata["project"]["squad"].title(),
        "project_name": metadata["project"]["name"],
        "target_count": review_summary["targets"],
        "source_count": 0,
        "mapped_count": review_summary["mapped"],
        "cards_by_slide": {},
        "processing_slides": visible_slides,
        "slide_summaries": _processing_slide_summaries(all_slides),
        "review_summary": review_summary,
        "process_steps": _processing_steps_from_state(state),
        "slide_datasources": {},
        "slide_ai_state": {},
        "preview_is_windowed": preview_is_windowed,
        "preview_selected_slide": preview_selected_slide,
        "preview_full_render_limit": full_render_limit,
        "preview_total_slides": len(all_slide_numbers),
        "mapping_status": mapping_status,
        "mapping_candidates": [],
        "ai_status": ai_status,
        "ai_available": ai_configured(PROJECT_ROOT),
        "ai_enabled": bool(metadata.get("use_ai")),
        "async_generation": _async_generation_enabled(),
        "analysis_warnings": [],
        "slide_selection_label": _slide_selection_label(_preview_scope_slides(metadata)),
        "notice": notice,
        "error": error,
        "ai_log_entries": _read_ai_logs(_job_dir(job_id)),
        "preview_processing": {**state, "active": True},
    }


def _processing_slide_items(state: dict) -> list[dict]:
    output = []
    for key, item in (state.get("slides") or {}).items():
        slide = int(item.get("slide") or key)
        output.append(
            {
                "slide": slide,
                "status": item.get("status") or "queued",
                "stage": item.get("stage") or "",
                "message": item.get("message") or "",
                "target_count": int(item.get("target_count") or 0),
                "mapped_count": int(item.get("mapped_count") or 0),
                "unmapped_count": int(item.get("unmapped_count") or 0),
            }
        )
    return sorted(output, key=lambda item: item["slide"])


def _processing_slide_summaries(slides: list[dict]) -> list[dict]:
    return [
        {
            "slide": item["slide"],
            "target_count": item["target_count"],
            "mapped_count": item["mapped_count"],
            "unmapped_count": item["unmapped_count"],
        }
        for item in slides
    ]


def _processing_steps_from_state(state: dict) -> list[dict]:
    current = str(state.get("current_stage") or "queued")
    order = [
        ("queued", "Fila", "Job recebido"),
        ("analysis", "Analise", "PPT e XLSX do escopo"),
        ("ai_match", "IA enxuta", "Datasource e receita estrutural"),
        ("complete", "Preview", "Cards prontos"),
    ]
    current_index = next((index for index, item in enumerate(order) if item[0] == current), 0)
    if str(state.get("status")) == "complete":
        current_index = len(order) - 1
    if str(state.get("status")) == "error":
        return [
            {"state": "done", "label": "Fila", "detail": "Job recebido"},
            {"state": "error", "label": "Erro", "detail": state.get("message") or "Falha no processamento"},
        ]
    steps = []
    for index, (_key, label, detail) in enumerate(order):
        if index < current_index:
            state_name = "done"
        elif index == current_index:
            state_name = "running"
        else:
            state_name = "pending"
        steps.append({"state": state_name, "label": label, "detail": detail})
    return steps


def _save_completed_preview_cache(
    job_dir: Path,
    analysis: AnalysisResult,
    mapping_status: dict,
    mapping_candidates: list[dict],
    ai_status: dict,
) -> None:
    try:
        metadata = _load_job_metadata(job_dir)
    except Exception:
        metadata = {}
    context = _preview_context_from_analysis(
        job_dir,
        job_dir.name,
        metadata,
        analysis,
        mapping_status,
        mapping_candidates,
        ai_status,
        _load_ai_diagnostics(job_dir),
        processing_state=_load_preview_processing_state(job_dir),
    )
    _save_render_cache(job_dir, context)


def _completed_preview_cache_usable(
    request: Request,
    cached: dict,
    selected_preview_slide: int | None,
) -> bool:
    if not cached:
        return False
    if not cached.get("preview_is_windowed"):
        return True
    requested_slide = selected_preview_slide if selected_preview_slide is not None else _preview_slide_from_request(request)
    if requested_slide is None:
        return True
    try:
        return int(cached.get("preview_selected_slide") or 0) == int(requested_slide)
    except (TypeError, ValueError):
        return False


def _preview_context_from_analysis(
    job_dir: Path,
    job_id: str,
    metadata: dict,
    analysis: AnalysisResult,
    mapping_status: dict,
    mapping_candidates: list[dict],
    ai_status: dict,
    ai_diagnostics: dict[str, dict],
    notice: str = "",
    error: str = "",
    request: Request | None = None,
    selected_preview_slide: int | None = None,
    processing_state: dict | None = None,
) -> dict:
    selected_slides = _selected_slides_for_job(job_dir)
    slide_ai_state = _load_slide_ai_state(job_dir)
    all_cards_by_slide = _cards_by_slide(
        analysis,
        _manual_source_names(job_dir),
        _manual_source_ranges(job_dir),
        ai_diagnostics,
        slide_ai_state,
        apply_slide_ai_outputs=_apply_slide_ai_outputs_enabled(job_dir),
    )
    all_slide_numbers = sorted(all_cards_by_slide)
    full_render_limit = _preview_full_render_slide_limit()
    preview_is_windowed = len(all_slide_numbers) > full_render_limit
    requested_slide = selected_preview_slide
    if requested_slide is None and request is not None:
        requested_slide = _preview_slide_from_request(request)
    preview_selected_slide = _preview_selected_slide(all_slide_numbers, requested_slide)
    if preview_is_windowed and preview_selected_slide is not None:
        visible_slide_numbers = [preview_selected_slide]
    else:
        visible_slide_numbers = all_slide_numbers
    cards_by_slide = {slide: all_cards_by_slide[slide] for slide in visible_slide_numbers}
    review_summary = _review_summary(all_cards_by_slide)
    return {
        "job_id": job_id,
        "metadata": metadata,
        "squad": metadata["project"]["squad"].title(),
        "project_name": metadata["project"]["name"],
        "target_count": analysis.target_count,
        "source_count": analysis.source_count,
        "mapped_count": len(analysis.plans),
        "cards_by_slide": dict(sorted(cards_by_slide.items())),
        "slide_summaries": _slide_summaries(all_cards_by_slide),
        "review_summary": review_summary,
        "process_steps": _process_steps(analysis, review_summary, mapping_status, ai_status, slide_ai_state),
        "slide_datasources": _slide_datasource_summary(job_dir, visible_slide_numbers),
        "slide_ai_state": slide_ai_state,
        "preview_is_windowed": preview_is_windowed,
        "preview_selected_slide": preview_selected_slide,
        "preview_full_render_limit": full_render_limit,
        "preview_total_slides": len(all_slide_numbers),
        "mapping_status": mapping_status,
        "mapping_candidates": mapping_candidates,
        "ai_status": ai_status,
        "ai_available": ai_configured(PROJECT_ROOT),
        "ai_enabled": False,
        "async_generation": _async_generation_enabled(),
        "analysis_warnings": analysis.warnings,
        "slide_selection_label": _slide_selection_label(selected_slides),
        "notice": notice,
        "error": error,
        "ai_log_entries": _read_ai_logs(job_dir),
        "preview_processing": {**(processing_state or {}), "active": False},
        "processing_slides": [],
    }


def _render_preview(
    request: Request,
    job_id: str,
    notice: str = "",
    error: str = "",
    prefer_cache: bool = False,
    allow_ai: bool = True,
    allow_ai_source_matches: bool | None = None,
    ai_diagnostic_target_ids: set[str] | None = None,
    selected_preview_slide: int | None = None,
) -> HTMLResponse:
    job_dir = _job_dir(job_id)
    if prefer_cache:
        cached = _load_render_cache(job_dir)
        if cached:
            cached["notice"] = notice
            cached["error"] = error
            cached["ai_log_entries"] = _read_ai_logs(job_dir)
            cached["current_user"] = _request_user(request)
            cached["checkpoint_summary"] = _load_job_checkpoint_summary(job_dir)
            return templates.TemplateResponse(request, "preview.html", cached)
    metadata = _load_job_metadata(job_dir)
    processing_state = _load_preview_processing_state(job_dir)
    if _preview_processing_is_active(processing_state):
        _log_job_debug_event(
            job_dir,
            "render_preview_processing",
            {
                "status": processing_state.get("status"),
                "current_stage": processing_state.get("current_stage"),
            },
        )
        context = _processing_preview_context(
                request,
                job_id,
                metadata,
                processing_state,
                notice=notice,
                error=error,
                selected_preview_slide=selected_preview_slide,
            )
        context["current_user"] = _request_user(request)
        context["checkpoint_summary"] = _load_job_checkpoint_summary(job_dir)
        return templates.TemplateResponse(request, "preview.html", context)
    cached = _load_render_cache(job_dir)
    if ai_diagnostic_target_ids is None and _completed_preview_cache_usable(request, cached, selected_preview_slide):
        _log_job_debug_event(
            job_dir,
            "render_preview_cache_hit",
            {
                "preview_is_windowed": cached.get("preview_is_windowed"),
                "preview_selected_slide": cached.get("preview_selected_slide"),
            },
        )
        cached["notice"] = notice
        cached["error"] = error
        cached["ai_log_entries"] = _read_ai_logs(job_dir)
        cached["current_user"] = _request_user(request)
        cached["checkpoint_summary"] = _load_job_checkpoint_summary(job_dir)
        return templates.TemplateResponse(request, "preview.html", cached)
    try:
        render_started = time.perf_counter()
        _log_job_debug_event(job_dir, "render_preview_recompute_start", {})
        analysis, mapping_status, mapping_candidates, pause_for_mapping = _analysis_for_job(job_dir)
    except Exception as exc:
        return _error_response(request, f"Nao consegui analisar os arquivos: {exc}", status_code=400)

    target_ai_requested = ai_diagnostic_target_ids is not None
    effective_allow_ai = allow_ai and target_ai_requested and ai_configured(PROJECT_ROOT) and not pause_for_mapping
    if effective_allow_ai:
        analysis, _mapping_status, _mapping_candidates, _pause_for_mapping = _analysis_for_job(
            job_dir,
            apply_cached_diagnostics=False,
        )
        source_match_ai = effective_allow_ai if allow_ai_source_matches is None else allow_ai_source_matches
        ai_matches, ai_match_status = _ai_source_matches_for_job(
            job_dir,
            analysis,
            allow_ai=source_match_ai,
            target_ids=ai_diagnostic_target_ids,
        )
        analysis = apply_ai_source_matches_to_analysis(analysis, ai_matches)
        ai_diagnostics, ai_diagnostic_status = _ai_diagnostics_for_job(
            job_dir,
            analysis,
            allow_ai=effective_allow_ai,
            target_ids=ai_diagnostic_target_ids,
        )
        analysis = apply_ai_recommendations_to_analysis(analysis, ai_diagnostics)
    else:
        ai_diagnostics = _load_ai_diagnostics(job_dir)
        ai_match_status, ai_diagnostic_status = _ai_waiting_status(pause_for_mapping)

    ai_status = _combine_ai_status(ai_match_status, ai_diagnostic_status)
    context = _preview_context_from_analysis(
        job_dir,
        job_id,
        metadata,
        analysis,
        mapping_status,
        mapping_candidates,
        ai_status,
        ai_diagnostics,
        notice=notice,
        error=error,
        request=request,
        selected_preview_slide=selected_preview_slide,
        processing_state=processing_state,
    )
    if not context["preview_is_windowed"]:
        _save_render_cache(job_dir, context)
    context["current_user"] = _request_user(request)
    context["checkpoint_summary"] = _load_job_checkpoint_summary(job_dir)
    _log_job_debug_event(
        job_dir,
        "render_preview_recompute_done",
        {
            "elapsed_ms": round((time.perf_counter() - render_started) * 1000),
            "target_count": analysis.target_count,
            "source_count": analysis.source_count,
            "plan_count": len(analysis.plans),
            "preview_is_windowed": context["preview_is_windowed"],
            "preview_selected_slide": context["preview_selected_slide"],
        },
    )
    _save_project_checkpoint(job_dir, status="in_progress")
    return templates.TemplateResponse(request, "preview.html", context)


def _analyze_files_signature(job_dir: Path, selected_slides: list[int]) -> tuple:
    pptx_stat = (job_dir / "input.pptx").stat()
    zip_stat = (job_dir / "datasources.zip").stat()
    overrides_signature: list[tuple] = []
    overrides_root = job_dir / "overrides"
    if overrides_root.exists():
        for target_dir in sorted(overrides_root.iterdir(), key=lambda path: path.name):
            if not target_dir.is_dir():
                continue
            files = sorted(target_dir.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
            if not files:
                continue
            chosen_stat = files[0].stat()
            range_path = target_dir / "range.txt"
            cell_range = range_path.read_text(encoding="utf-8").strip() if range_path.exists() else ""
            overrides_signature.append(
                (target_dir.name, files[0].name, chosen_stat.st_mtime_ns, chosen_stat.st_size, cell_range)
            )
    return (
        pptx_stat.st_mtime_ns,
        pptx_stat.st_size,
        zip_stat.st_mtime_ns,
        zip_stat.st_size,
        tuple(sorted(selected_slides)),
        tuple(overrides_signature),
    )


def _cached_analyze_files(
    job_dir: Path,
    manual_sources: dict,
    selected_slides: list[int],
    progress_callback=None,
) -> AnalysisResult:
    signature = _analyze_files_signature(job_dir, selected_slides)
    job_key = job_dir.name
    with ANALYZE_FILES_CACHE_LOCK:
        cached = ANALYZE_FILES_CACHE.get(job_key)
        if cached is not None and cached[0] == signature:
            ANALYZE_FILES_CACHE.move_to_end(job_key)
            log_debug_event("analyze_files_cache_hit", {"job_id": job_key})
            if progress_callback:
                progress_callback(
                    {
                        "phase": "matching",
                        "completed": cached[1].target_count,
                        "total": cached[1].target_count,
                        "message": f"{cached[1].target_count} objeto(s) recuperado(s) do cache.",
                    }
                )
            return cached[1]
    analysis = analyze_files(
        (job_dir / "input.pptx").read_bytes(),
        (job_dir / "datasources.zip").read_bytes(),
        manual_sources=manual_sources,
        slide_numbers=selected_slides,
        progress_callback=progress_callback,
    )
    with ANALYZE_FILES_CACHE_LOCK:
        ANALYZE_FILES_CACHE[job_key] = (signature, analysis)
        ANALYZE_FILES_CACHE.move_to_end(job_key)
        while len(ANALYZE_FILES_CACHE) > ANALYZE_FILES_CACHE_MAX_JOBS:
            ANALYZE_FILES_CACHE.popitem(last=False)
    return analysis


def _analysis_for_job(
    job_dir: Path,
    apply_cached_source_matches: bool = True,
    apply_cached_diagnostics: bool = True,
    apply_slide_outputs: bool = True,
    progress_callback=None,
) -> tuple[AnalysisResult, dict, list[dict], bool]:
    metadata = _load_job_metadata(job_dir)
    selected_slides = _selected_slides_for_job(job_dir)
    analysis = _cached_analyze_files(
        job_dir,
        _manual_sources_for_job(job_dir),
        selected_slides,
        progress_callback=progress_callback,
    )
    mapping_candidates = []
    selected_mapping_template = _selected_mapping_template(metadata)
    if selected_mapping_template:
        analysis, mapping_status = _apply_mapping_template_to_analysis(
            job_dir,
            analysis,
            skip_targets=set(_manual_source_names(job_dir)),
        )
    else:
        mapping_candidates = _mapping_template_candidates(metadata["project"]["squad"], analysis.targets)
        mapping_status = _mapping_status_without_selection(mapping_candidates)
    pause_for_mapping = bool(mapping_candidates) and not bool(metadata.get("ignore_mapping_candidates"))

    if apply_cached_source_matches:
        analysis = apply_ai_source_matches_to_analysis(analysis, _load_ai_source_matches(job_dir))
    if apply_cached_diagnostics:
        analysis = apply_ai_recommendations_to_analysis(analysis, _load_ai_diagnostics(job_dir))
    if apply_slide_outputs:
        analysis = apply_typed_outputs_to_analysis(analysis, _slide_ai_target_outputs(job_dir))
    analysis = _apply_chart_format_overrides(analysis, metadata)
    return analysis, mapping_status, mapping_candidates, pause_for_mapping


def _run_target_ai_review(job_dir: Path, target_id: str, manual_context: str = "") -> str:
    analysis, _mapping_status, _mapping_candidates, _pause_for_mapping = _analysis_for_job(
        job_dir,
        apply_cached_source_matches=True,
        apply_cached_diagnostics=False,
        apply_slide_outputs=False,
    )
    targets_by_id = {target.target_id: target for target in analysis.targets}
    target = targets_by_id.get(target_id)
    if target is None or target.object_type not in {"chart", "table"}:
        raise ValueError("Target nao encontrado no escopo atual.")
    # Revisao de UM target (upload manual ou botao "Revisar este target"): leve e
    # sem rasterizacao. O cliente ja escolheu o XLSX; aqui a IA so precisa do
    # Editar dados do target + a estrutura do XLSX + o titulo para decidir
    # orientacao/colunas.
    return _run_slide_ai_review(
        job_dir,
        target.slide_number,
        target_ids={target_id},
        manual_context=manual_context,
    )


def _run_automatic_slide_ai_review(
    job_dir: Path,
    slide_numbers: list[int] | set[int] | None = None,
    force: bool = False,
) -> str:
    if not force and not _auto_slide_ai_enabled():
        return ""
    if not ai_configured(PROJECT_ROOT):
        return ""

    metadata = _load_job_metadata(job_dir)
    if _selected_mapping_template(metadata) and not force:
        return ""

    if not _selected_mapping_template(metadata):
        metadata["ignore_mapping_candidates"] = True
        metadata["use_ai"] = False
        _save_job_metadata(job_dir, metadata)

    analysis, _mapping_status, _mapping_candidates, _pause_for_mapping = _analysis_for_job(
        job_dir,
        apply_cached_source_matches=True,
        apply_cached_diagnostics=False,
        apply_slide_outputs=False,
    )
    requested_slides = {int(slide) for slide in (slide_numbers or []) if int(slide) > 0}
    slides = sorted(
        {
            target.slide_number
            for target in analysis.targets
            if target.object_type in {"chart", "table"}
            and (not requested_slides or target.slide_number in requested_slides)
        }
    )
    if not slides:
        return ""

    state = _load_slide_ai_state(job_dir)
    ran = 0
    skipped = 0
    deterministic_skipped = 0
    deferred = 0
    errors: list[str] = []
    max_slides = _slide_ai_max_slides_per_run()
    target_ids_by_slide = {
        slide: {
            target.target_id
            for target in analysis.targets
            if target.slide_number == slide and target.object_type in {"chart", "table"}
        }
        for slide in slides
    }
    review_flags = {
        slide: _slide_needs_ai_review(analysis, slide, target_ids_by_slide[slide]) for slide in slides
    }
    # Em decks grandes, gasta o orcamento de chamadas primeiro nos slides que
    # realmente precisam (target sem plano, confianca baixa, aviso), deixando os
    # demais para uma proxima rodada em vez de varrer 100+ slides em sequencia.
    ordered_slides = [slide for slide in slides if review_flags[slide][0]] + [
        slide for slide in slides if not review_flags[slide][0]
    ]
    for slide in ordered_slides:
        slide_target_ids = target_ids_by_slide[slide]
        existing_slide_state = ((state.get("slides") or {}).get(str(slide)) or {})
        signature = _slide_ai_signature(
            job_dir,
            analysis,
            slide,
            slide_target_ids,
            str(existing_slide_state.get("manual_context") or ""),
        )
        if _slide_ai_outputs_complete(state, slide, slide_target_ids, signature):
            skipped += 1
            continue
        should_review, review_reason = review_flags[slide]
        if not force and not should_review:
            _mark_slide_ai_skipped(state, slide, signature, review_reason)
            _save_slide_ai_state(job_dir, state)
            deterministic_skipped += 1
            _append_ai_log(
                job_dir,
                {
                    "operation": "auto_slide_review",
                    "status": "skipped",
                    "sent_at": _now_iso(),
                    "returned_at": _now_iso(),
                    "duration_ms": 0,
                    "slide": slide,
                    "target_count": len(slide_target_ids),
                    "reason": review_reason,
                },
            )
            continue
        if ran >= max_slides:
            deferred += 1
            continue
        try:
            sent_at = _now_iso()
            started = time.perf_counter()
            _run_slide_ai_review(job_dir, slide)
            state = _load_slide_ai_state(job_dir)
            slide_state = state.setdefault("slides", {}).setdefault(str(slide), {})
            slide_state["auto_reviewed_at"] = _now_iso()
            _save_slide_ai_state(job_dir, state)
            _append_ai_log(
                job_dir,
                {
                    "operation": "auto_slide_review",
                    "status": "ok",
                    "sent_at": sent_at,
                    "returned_at": _now_iso(),
                    "duration_ms": round((time.perf_counter() - started) * 1000),
                    "slide": slide,
                    "target_count": len(slide_target_ids),
                },
            )
            ran += 1
        except Exception as exc:
            errors.append(f"slide {slide}: {format_ai_error(exc)}")
            _append_ai_log(
                job_dir,
                {
                    "operation": "auto_slide_review",
                    "status": "error",
                    "sent_at": sent_at if "sent_at" in locals() else _now_iso(),
                    "returned_at": _now_iso(),
                    "duration_ms": round((time.perf_counter() - started) * 1000) if "started" in locals() else 0,
                    "slide": slide,
                    "target_count": len(slide_target_ids),
                    "error": format_ai_error(exc),
                },
            )

    if not ran and not errors:
        if skipped:
            return "IA automatica por slide ja estava atualizada."
        if deterministic_skipped:
            return f"IA automatica pulou {deterministic_skipped} slide(s): mapeamento deterministico suficiente."
        return ""
    notice = f"IA automatica revisou {ran} slide(s)."
    if skipped:
        notice += f" {skipped} slide(s) ja tinham matriz IA valida."
    if deterministic_skipped:
        notice += f" {deterministic_skipped} slide(s) ficaram no fluxo deterministico por alta confianca."
    if deferred:
        notice += (
            f" {deferred} slide(s) ficaram para a proxima rodada (limite de {max_slides} slides com IA por execucao);"
            " rode Revisar com IA novamente para continuar."
        )
    if errors:
        notice += f" {len(errors)} slide(s) falharam: {'; '.join(errors[:3])}."
    _save_project_checkpoint(job_dir, status="in_progress")
    return notice


def _auto_slide_ai_enabled() -> bool:
    value = os.getenv("AUTO_PPT_AUTO_SLIDE_AI", "0").strip().lower()
    return value not in {"0", "false", "no", "off"}


def _slide_ai_max_slides_per_run() -> int:
    # A revisao por slide dispara as 2 chamadas mais pesadas do sistema
    # (understanding estrutural + matrix builder). Este teto impede que um clique
    # em Revisar com IA num deck de 100+ slides gere centenas de chamadas numa
    # unica execucao; os slides prioritarios rodam primeiro e o restante fica
    # para a proxima rodada.
    return max(_env_int("AUTO_PPT_SLIDE_AI_MAX_SLIDES_PER_RUN", 1), 1)


def _slide_ai_outputs_complete(state: dict, slide_number: int, target_ids: set[str], signature: str = "") -> bool:
    if not target_ids:
        return False
    slide_state = ((state.get("slides") or {}).get(str(slide_number)) or {})
    if signature and slide_state.get("input_signature") != signature:
        return False
    outputs = slide_state.get("target_outputs") or {}
    for target_id in target_ids:
        output = outputs.get(target_id)
        if not output or output.get("validation_errors"):
            return False
    return True


def _slide_needs_ai_review(analysis: AnalysisResult, slide_number: int, target_ids: set[str]) -> tuple[bool, str]:
    mode = os.getenv("AUTO_PPT_AUTO_SLIDE_AI_MODE", "strict").strip().lower()
    if mode in {"all", "always"}:
        return True, "modo configurado para revisar todos os slides com IA"
    if mode in {"off", "none", "disabled"}:
        return False, "modo configurado para nao acionar IA automatica"

    plans_by_id = {plan.target_id: plan for plan in analysis.plans}
    slide_targets = [target for target in analysis.targets if target.target_id in target_ids]
    missing = [target.target_id for target in slide_targets if target.target_id not in plans_by_id]
    if missing:
        return True, f"{len(missing)} target(s) sem datasource automatico"

    confidence_floor = _auto_slide_ai_confidence_floor()
    low_confidence = [
        plan.target_id
        for plan in analysis.plans
        if plan.target_id in target_ids and float(plan.confidence or 0) < confidence_floor
    ]
    if low_confidence:
        return True, f"{len(low_confidence)} target(s) abaixo de {confidence_floor:.0%} de confianca"

    warned = [
        plan.target_id
        for plan in analysis.plans
        if plan.target_id in target_ids and (plan.warnings or "atencao" in _norm_text(plan.reason))
    ]
    if warned:
        return True, f"{len(warned)} target(s) com aviso ou candidato parecido"

    if mode in {"duplicates", "aggressive"} and _slide_has_duplicate_contracts(analysis, target_ids):
        return True, "slide tem targets com contrato PPT repetido; IA ajuda a diferenciar filtros/titulos"

    return False, "todos os targets tem match deterministico de alta confianca"


def _slide_has_duplicate_contracts(analysis: AnalysisResult, target_ids: set[str]) -> bool:
    signatures: dict[str, int] = {}
    for target in analysis.targets:
        if target.target_id not in target_ids or target.object_type not in {"chart", "table"}:
            continue
        signature = json.dumps(
            {
                "type": target.object_type,
                "orientation": target.expected_orientation,
                "categories": [_norm_text(item) for item in target.expected_categories],
                "series": [_norm_text(item) for item in target.expected_series],
                "table_shape": [len(row) for row in target.table_cells[:8]],
            },
            ensure_ascii=False,
            sort_keys=True,
        )
        signatures[signature] = signatures.get(signature, 0) + 1
    return any(count > 1 for count in signatures.values())


def _mark_slide_ai_skipped(state: dict, slide_number: int, signature: str, reason: str) -> None:
    slide_state = state.setdefault("slides", {}).setdefault(str(slide_number), {})
    slide_state["input_signature"] = signature
    slide_state["auto_review_skipped_at"] = _now_iso()
    slide_state["auto_review_skip_reason"] = reason


def _slide_ai_signature(
    job_dir: Path,
    analysis: AnalysisResult,
    slide_number: int,
    target_ids: set[str] | None = None,
    manual_context: str = "",
) -> str:
    selected_target_ids = set(target_ids or [])
    targets = [
        target
        for target in analysis.targets
        if target.slide_number == slide_number
        and target.object_type in {"chart", "table"}
        and (not selected_target_ids or target.target_id in selected_target_ids)
    ]
    selected_entries, _warnings = _datasource_entries_for_ai_review(job_dir, slide_number, targets, analysis.sources)
    manifests = _source_manifests_for_entries(analysis.sources, selected_entries)
    datasource_hashes = _datasource_entry_hashes(job_dir / "datasources.zip", selected_entries)
    payload = {
        "version": "slide-ai-v5-text-only",
        "slide": slide_number,
        "manual_context": manual_context.strip(),
        "payload_profile": _ai_payload_profile(),
        "targets": [_target_ai_payload(target) for target in targets],
        "xlsx_manifests": manifests,
        "xlsx_hashes": datasource_hashes,
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _datasource_entry_hashes(datasources_zip: Path, entries: list) -> dict[str, str]:
    output: dict[str, str] = {}
    zip_entries = [entry for entry in entries if getattr(entry, "manual_data", None) is None]
    for entry in entries:
        manual_data = getattr(entry, "manual_data", None)
        if manual_data is not None:
            output[entry.zip_path] = hashlib.sha256(manual_data).hexdigest()
    if zip_entries:
        with ZipFile(datasources_zip) as zf:
            for entry in zip_entries:
                try:
                    output[entry.zip_path] = hashlib.sha256(zf.read(entry.zip_path)).hexdigest()
                except KeyError:
                    output[entry.zip_path] = "missing"
    return output


def _xlsx_prompt_texts(xlsx_dumps: list) -> list[str]:
    mode = _xlsx_dump_mode()
    max_cells = max(_env_int("AUTO_PPT_AI_XLSX_MAX_CELLS_PER_SHEET", 800), 1)
    return [dump.as_prompt_text(max_cells_per_sheet=max_cells, mode=mode) for dump in xlsx_dumps]


def _xlsx_dump_mode() -> str:
    mode = os.getenv("AUTO_PPT_AI_XLSX_DUMP_MODE", "compact").strip().lower()
    if mode in {"full", "verbose", "debug", "raw"}:
        return "verbose"
    return "compact"


def _ai_payload_profile() -> dict:
    return {
        "xlsx_dump_mode": _xlsx_dump_mode(),
        "xlsx_max_cells_per_sheet": max(_env_int("AUTO_PPT_AI_XLSX_MAX_CELLS_PER_SHEET", 800), 1),
    }


def _ai_input_stats(xlsx_prompt_texts: list[str], xlsx_manifests: list[dict], target_payloads: list[dict]) -> dict:
    dump_bytes = sum(len(text.encode("utf-8")) for text in xlsx_prompt_texts)
    manifest_bytes = _json_size_bytes(xlsx_manifests)
    target_bytes = _json_size_bytes(target_payloads)
    return {
        **_ai_payload_profile(),
        "request_count": 1,
        "xlsx_dump_count": len(xlsx_prompt_texts),
        "xlsx_dump_bytes_utf8": dump_bytes,
        "xlsx_manifest_count": len(xlsx_manifests),
        "xlsx_manifest_bytes_utf8": manifest_bytes,
        "target_count": len(target_payloads),
        "target_payload_bytes_utf8": target_bytes,
        "content_bytes_utf8": dump_bytes + manifest_bytes + target_bytes,
        "output_contract": "typed_cells_compact",
    }


def _json_size_bytes(value) -> int:
    return len(json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))


def _format_bytes(value) -> str:
    try:
        size = float(value or 0)
    except (TypeError, ValueError):
        size = 0.0
    units = ["B", "KB", "MB", "GB"]
    unit_index = 0
    while size >= 1024 and unit_index < len(units) - 1:
        size /= 1024
        unit_index += 1
    if unit_index == 0:
        return f"{int(size)} {units[unit_index]}"
    return f"{size:.1f} {units[unit_index]}"


@dataclass(frozen=True)
class _ManualDatasourceEntry:
    zip_path: str
    file_name: str
    slide_number: int | None
    is_general: bool
    target_id: str
    manual_data: bytes
    cell_range: str = ""


def _run_slide_ai_review(
    job_dir: Path,
    slide_number: int,
    target_ids: set[str] | None = None,
    manual_context: str = "",
) -> str:
    analysis, _mapping_status, _mapping_candidates, _pause_for_mapping = _analysis_for_job(
        job_dir,
        apply_cached_source_matches=True,
        apply_cached_diagnostics=False,
        apply_slide_outputs=False,
    )
    slide_targets = [
        target
        for target in analysis.targets
        if target.slide_number == slide_number and target.object_type in {"chart", "table"}
    ]
    if target_ids:
        slide_targets = [target for target in slide_targets if target.target_id in target_ids]
    if not slide_targets:
        raise ValueError(f"Nenhum target atualizavel encontrado no slide {slide_number}.")

    state = _load_slide_ai_state(job_dir)
    slide_key = str(slide_number)
    slide_state = state.setdefault("slides", {}).setdefault(slide_key, {})
    previous_context = str(slide_state.get("manual_context") or "").strip()
    incoming_context = manual_context.strip()
    if incoming_context:
        slide_state["manual_context"] = incoming_context
    combined_context = incoming_context or previous_context
    selected_entries, warnings = _datasource_entries_for_ai_review(
        job_dir,
        slide_number,
        slide_targets,
        analysis.sources,
    )
    xlsx_dumps = _dump_datasource_entries_for_ai_review(job_dir, selected_entries)
    xlsx_prompt_texts = _xlsx_prompt_texts(xlsx_dumps)
    xlsx_manifests = _source_manifests_for_entries(analysis.sources, selected_entries)
    target_payloads = [_target_ai_payload(target) for target in slide_targets]
    signature = _slide_ai_signature(job_dir, analysis, slide_number, target_ids, combined_context)
    slide_state["ai_input_stats"] = _ai_input_stats(xlsx_prompt_texts, xlsx_manifests, target_payloads)

    # Uma unica chamada decide a fonte e monta a matriz. Antes, o mesmo dump e
    # contrato eram enviados primeiro ao "understanding" e novamente ao builder.
    understanding = {
        "mode": "single_call",
        "slide_number": slide_number,
        "slide_text": _structured_titles(slide_targets),
        "source_files": [manifest.get("file_name") for manifest in xlsx_manifests],
        "manual_context": combined_context,
    }
    slide_state["understanding"] = understanding
    slide_state["warnings"] = warnings
    slide_state["input_signature"] = signature

    matrix_result = _call_with_job_debug(
        job_dir,
        build_slide_matrices_with_ai,
        SlideMatrixBuildInput(
            slide_number=slide_number,
            slide_understanding=understanding,
            targets=target_payloads,
            xlsx_manifests=xlsx_manifests,
            xlsx_dumps=xlsx_prompt_texts,
            target_ids=[target.target_id for target in slide_targets],
            manual_context=combined_context,
        ),
        root=PROJECT_ROOT,
    )
    valid_target_ids = {target.target_id for target in slide_targets}
    valid_sources = {entry.zip_path for entry in selected_entries}
    valid_source_basenames = {Path(entry.zip_path).name for entry in selected_entries}
    target_outputs = slide_state.setdefault("target_outputs", {})
    targets_by_id = {target.target_id: target for target in slide_targets}
    validation_messages: list[str] = []
    for output in matrix_result.get("target_outputs") or []:
        target_id = str(output.get("target_id") or "")
        if target_id not in valid_target_ids:
            continue
        target = targets_by_id.get(target_id)
        object_type = target.object_type if target is not None else str(output.get("object_type") or "chart")
        output["final_edit_data"] = normalize_typed_edit_data(output.get("final_edit_data") or {})
        errors = validate_typed_edit_data(output["final_edit_data"], object_type=object_type, target=target)
        source_file = str(output.get("source_file") or "")
        if not source_file:
            errors.append("A IA nao informou o XLSX de origem (source_file) para este target.")
        elif source_file not in valid_sources and Path(source_file).name not in valid_source_basenames:
            errors.append(f"source_file '{source_file}' nao esta entre os XLSX do slide.")
        output["validation_errors"] = errors
        if errors:
            validation_messages.extend(errors)
        target_outputs[target_id] = output
    slide_state["questions_for_user"] = matrix_result.get("questions_for_user") or understanding.get("questions_for_user") or []
    _save_slide_ai_state(job_dir, state)
    _clear_render_cache(job_dir)
    count = len(matrix_result.get("target_outputs") or [])
    suffix = f" Avisos: {' '.join(validation_messages[:3])}" if validation_messages else ""
    return f"IA revisou slide {slide_number} e gerou {count} matriz(es) tipada(s).{suffix}"


def _load_slide_ai_state(job_dir: Path) -> dict:
    path = job_dir / "slide_ai_state.json"
    if not path.exists():
        return {"slides": {}}
    return json.loads(path.read_text(encoding="utf-8"))


def _save_slide_ai_state(job_dir: Path, state: dict) -> None:
    (job_dir / "slide_ai_state.json").write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _slide_ai_target_outputs(job_dir: Path) -> dict[str, dict]:
    if not _apply_slide_ai_outputs_enabled(job_dir):
        return {}
    state = _load_slide_ai_state(job_dir)
    output: dict[str, dict] = {}
    for slide_state in (state.get("slides") or {}).values():
        for target_id, target_output in (slide_state.get("target_outputs") or {}).items():
            if not target_output.get("validation_errors"):
                output[target_id] = target_output
    return output


def _apply_slide_ai_outputs_enabled(job_dir: Path) -> bool:
    metadata = _load_job_metadata(job_dir)
    if bool(metadata.get("apply_slide_ai_outputs")):
        return True
    return _env_bool("AUTO_PPT_APPLY_SLIDE_AI_OUTPUTS", False)


def _datasource_entries_for_slide(job_dir: Path, slide_number: int):
    entries = collect_datasource_entries(job_dir / "datasources.zip")
    return entries_for_slide(entries, slide_number)


def _datasource_entries_for_ai_review(job_dir: Path, slide_number: int, targets: list, sources: list) -> tuple[list, list[str]]:
    selected_entries, warnings = _datasource_entries_for_slide(job_dir, slide_number)
    manual_entries = _manual_datasource_entries_for_targets(job_dir, targets)
    if manual_entries and len(manual_entries) == len(targets):
        return manual_entries, []
    if manual_entries:
        selected_by_path = {entry.zip_path: entry for entry in selected_entries}
        for entry in manual_entries:
            selected_by_path[entry.zip_path] = entry
        selected_entries = list(selected_by_path.values())
    return _relevant_datasource_entries_for_targets(targets, selected_entries, sources), warnings


def _manual_datasource_entries_for_targets(job_dir: Path, targets: list) -> list[_ManualDatasourceEntry]:
    manual_sources = _manual_sources_for_job(job_dir)
    entries: list[_ManualDatasourceEntry] = []
    for target in targets:
        payload = manual_sources.get(target.target_id)
        if not payload:
            continue
        filename, data, cell_range = payload
        zip_path = f"upload_manual/{target.target_id}_{filename}"
        entries.append(
            _ManualDatasourceEntry(
                zip_path=zip_path,
                file_name=zip_path,
                slide_number=target.slide_number,
                is_general=False,
                target_id=target.target_id,
                manual_data=data,
                cell_range=cell_range,
            )
        )
    return entries


def _dump_datasource_entries_for_ai_review(job_dir: Path, entries: list) -> list:
    output = []
    zip_entries = [entry for entry in entries if getattr(entry, "manual_data", None) is None]
    zip_dumps_by_path = {}
    if zip_entries:
        zip_dumps_by_path = {
            dump.file_name: dump for dump in dump_xlsx_zip_entries(job_dir / "datasources.zip", zip_entries)
        }
    for entry in entries:
        manual_data = getattr(entry, "manual_data", None)
        if manual_data is not None:
            output.append(dump_xlsx_workbook(manual_data, file_name=entry.zip_path))
            continue
        dump = zip_dumps_by_path.get(entry.zip_path)
        if dump is not None:
            output.append(dump)
    return output


def _relevant_datasource_entries_for_targets(
    targets: list,
    entries: list,
    sources: list,
    limit_per_target: int = 4,
) -> list:
    """Restringe os XLSX enviados para IA de nivel de slide aos candidatos com sinal
    local para os targets revisados, em vez de mandar todos os XLSX do slide. Reduz o
    payload (dump + manifesto) sem perder cobertura: se a pontuacao local nao encontrar
    nenhum candidato (score 0 para todos), mantem o conjunto original do slide."""
    entries_by_path = {entry.zip_path: entry for entry in entries}
    sources_in_scope = [source for source in sources if source.file_name in entries_by_path]
    if not sources_in_scope:
        return entries
    relevant_paths: set[str] = set()
    for target in targets:
        candidates = source_match_candidates(target, sources_in_scope, limit=limit_per_target)
        relevant_paths.update(candidate.source.file_name for candidate in candidates if candidate.score > 0)
    if not relevant_paths:
        return entries
    restricted = [entries_by_path[path] for path in relevant_paths if path in entries_by_path]
    return restricted or entries


def _source_manifests_for_entries(sources: list, entries: list) -> list[dict]:
    entries_by_path = {entry.zip_path: entry for entry in entries}
    output: list[dict] = []
    for source in sources:
        entry = entries_by_path.get(source.file_name)
        if entry is None:
            continue
        manifest = xlsx_source_manifest(source)
        semantic = manifest.get("semantic_context") or {}
        profile = manifest.get("structural_profile") or {}
        output.append(
            {
                "file_name": manifest.get("file_name"),
                "sheet_name": manifest.get("sheet_name"),
                "used_range": manifest.get("used_range"),
                "requested_range": str(getattr(entry, "cell_range", "") or ""),
                "orientation": manifest.get("orientation"),
                "semantic_context": {
                    key: semantic.get(key)
                    for key in ("table_title", "row_group_label", "context_text", "graph_id", "ppt_tag", "variable")
                    if semantic.get(key)
                },
                "structural_profile": {
                    "name": profile.get("name"),
                    "recipe_hint": profile.get("recipe_hint") or {},
                },
                "categories": manifest.get("categories") or [],
                "series": manifest.get("series") or [],
            }
        )
    return output


def _slide_datasource_summary(job_dir: Path, slide_numbers: list[int]) -> dict[int, dict]:
    try:
        entries = collect_datasource_entries(job_dir / "datasources.zip")
    except Exception as exc:
        return {slide: {"files": [], "warnings": [str(exc)]} for slide in slide_numbers}
    output = {}
    for slide in slide_numbers:
        selected, warnings = entries_for_slide(entries, slide)
        output[slide] = {"files": [entry.zip_path for entry in selected], "warnings": warnings}
    return output


def _chart_shape_for_target(target) -> dict | None:
    """Tipo + cores reais do chart (do XML já extraído) para a prévia fiel no
    navegador. None quando não é chart ou o tipo não é suportado (só tabela)."""
    if getattr(target, "object_type", "") != "chart":
        return None
    kind = getattr(target, "chart_kind", "") or ""
    if not kind:
        return None
    colors = list(getattr(target, "chart_series_colors", []) or [])
    return {"kind": kind, "colors": colors}


def _target_ai_payload(target) -> dict:
    return {
        "target_id": target.target_id,
        "visual_label": visual_label(target.target_id),
        "shape_name": target.shape_name,
        "shape_id": target.shape_id,
        "object_type": target.object_type,
        "slide_number": target.slide_number,
        "title": getattr(target, "title", "") or "",
        "position": {
            "left_in": target.left_in,
            "top_in": target.top_in,
            "width_in": target.width_in,
            "height_in": target.height_in,
        },
        "nearby_text": target.nearby_text,
        "ppt_contract": _ppt_contract_for_target(target),
    }


def _structured_titles(targets: list) -> str:
    """Contexto compacto e ORDENADO por objeto (titulo -> id), em vez do dump do
    slide inteiro sem ordem. Ex.: 'Status de Inatividade (%) [S003_T003_CHART]'."""
    parts = []
    for target in targets:
        title = (getattr(target, "title", "") or "").strip() or (target.nearby_text or "").split(" | ")[0].strip()
        if title:
            parts.append(f"{title} [{target.target_id}]")
    return " ; ".join(parts)


def _canonical_target_id(targets: list, target_id_or_alias: str) -> str:
    text = str(target_id_or_alias or "").strip()
    for target in targets:
        if text in target_aliases(target):
            return target.target_id
    return text


def _set_target_state_response(
    request: Request,
    job_id: str,
    target_id: str,
    approved: bool,
    skipped: bool,
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        analysis, _mapping_status, _mapping_candidates, _pause = _analysis_for_job(job_dir, apply_slide_outputs=False)
        canonical_id = _canonical_target_id(analysis.targets, target_id)
        target = next((item for item in analysis.targets if item.target_id == canonical_id), None)
        if target is None:
            raise ValueError("Target nao encontrado.")
        state = _load_slide_ai_state(job_dir)
        slide_state = state.setdefault("slides", {}).setdefault(str(target.slide_number), {})
        approvals = slide_state.setdefault("target_approvals", {})
        actor = _actor(request)
        approvals[canonical_id] = {
            "approved": approved,
            "skipped": skipped,
            "updated_at": _now_iso(),
            "by": actor,
        }
        _save_slide_ai_state(job_dir, state)
        _clear_render_cache(job_dir)
        audit.record(
            job_dir,
            actor,
            "aprovou_grafico" if approved else "pulou_grafico",
            {"target": canonical_id, "slide": target.slide_number},
        )
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc), allow_ai=False)
    action = "aprovado" if approved else "pulado"
    return _render_preview(request, job_id, notice=f"Target {canonical_id} {action}.", allow_ai=False)


def _set_slide_state_response(
    request: Request,
    job_id: str,
    slide_number: int,
    approved: bool,
    skipped: bool,
) -> HTMLResponse:
    try:
        job_dir = _job_dir(job_id)
        state = _load_slide_ai_state(job_dir)
        slide_state = state.setdefault("slides", {}).setdefault(str(slide_number), {})
        actor = _actor(request)
        slide_state["approval"] = {
            "approved": approved,
            "skipped": skipped,
            "updated_at": _now_iso(),
            "by": actor,
        }
        _save_slide_ai_state(job_dir, state)
        _clear_render_cache(job_dir)
        audit.record(
            job_dir,
            actor,
            "aprovou_slide" if approved else "pulou_slide",
            {"slide": slide_number},
        )
        _save_project_checkpoint(job_dir, status="in_progress")
    except Exception as exc:
        return _render_preview(request, job_id, error=str(exc), allow_ai=False)
    action = "aprovado" if approved else "pulado"
    return _render_preview(request, job_id, notice=f"Slide {slide_number} {action}.", allow_ai=False)


def _review_summary(cards_by_slide: dict[int, list[dict]]) -> dict:
    cards = [item for items in cards_by_slide.values() for item in items]
    return {
        "slides": len(cards_by_slide),
        "targets": len(cards),
        "mapped": sum(1 for item in cards if item["has_plan"]),
        "unmapped": sum(1 for item in cards if not item["has_plan"]),
        "manual": sum(1 for item in cards if item["manual_file"]),
        "ai": sum(1 for item in cards if item["ai"]),
    }


def _process_steps(
    analysis: AnalysisResult,
    review_summary: dict,
    mapping_status: dict,
    ai_status: dict,
    slide_ai_state: dict,
) -> list[dict]:
    slide_summary = _slide_ai_progress_summary(slide_ai_state)
    unmapped = int(review_summary.get("unmapped") or 0)
    questions = int(slide_summary.get("questions") or 0)
    ai_done = int(slide_summary.get("with_outputs") or 0)
    ai_skipped = int(slide_summary.get("skipped") or 0)
    return [
        {
            "label": "Arquivos",
            "state": "done",
            "detail": "PPTX e ZIP recebidos.",
        },
        {
            "label": "Escopo",
            "state": "done",
            "detail": f"{review_summary.get('slides', 0)} slide(s), {review_summary.get('targets', 0)} target(s).",
        },
        {
            "label": "XLSX",
            "state": "done",
            "detail": f"{analysis.source_count} fonte(s) lidas com contexto semantico.",
        },
        {
            "label": "Mapeamento",
            "state": "warn" if unmapped else "done",
            "detail": f"{review_summary.get('mapped', 0)} mapeado(s), {unmapped} pendente(s).",
        },
        {
            "label": "IA seletiva",
            "state": _process_ai_state(ai_status, ai_done, ai_skipped),
            "detail": f"{ai_done} slide(s) com matriz IA, {ai_skipped} em modo deterministico.",
        },
        {
            "label": "Revisao",
            "state": "warn" if questions or unmapped else "active",
            "detail": "Revise perguntas/pendencias antes do download." if questions or unmapped else "Pronto para aprovacao e download.",
        },
    ]


def _slide_ai_progress_summary(slide_ai_state: dict) -> dict[str, int]:
    slides = (slide_ai_state or {}).get("slides") or {}
    with_outputs = 0
    skipped = 0
    questions = 0
    for slide_state in slides.values():
        if slide_state.get("target_outputs"):
            with_outputs += 1
        if slide_state.get("auto_review_skipped_at"):
            skipped += 1
        questions += len(slide_state.get("questions_for_user") or [])
    return {"with_outputs": with_outputs, "skipped": skipped, "questions": questions}


def _process_ai_state(ai_status: dict, ai_done: int, ai_skipped: int) -> str:
    state = str((ai_status or {}).get("state") or "")
    if state == "warn":
        return "warn"
    if ai_done or ai_skipped:
        return "done"
    if state in {"disabled", "cached"}:
        return "waiting"
    return "active"


def _slide_summaries(cards_by_slide: dict[int, list[dict]]) -> list[dict]:
    summaries = []
    for slide, items in sorted(cards_by_slide.items()):
        summaries.append(
            {
                "slide": slide,
                "target_count": len(items),
                "mapped_count": sum(1 for item in items if item["has_plan"]),
                "unmapped_count": sum(1 for item in items if not item["has_plan"]),
                "manual_count": sum(1 for item in items if item["manual_file"]),
                "ai_count": sum(1 for item in items if item["ai"]),
            }
        )
    return summaries


def _preview_slide_from_request(request: Request) -> int | None:
    raw = request.query_params.get("slide")
    if not raw:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def _preview_selected_slide(slides: list[int], requested_slide: int | None) -> int | None:
    if not slides:
        return None
    if requested_slide in slides:
        return requested_slide
    return slides[0]


def _preview_full_render_slide_limit() -> int:
    try:
        return max(int(os.getenv("AUTO_PPT_PREVIEW_FULL_RENDER_SLIDES", "12")), 1)
    except ValueError:
        return 12


def _resolve_mapping_template(project, mapping_template_ref: str) -> dict:
    text = (mapping_template_ref or "").strip()
    if not text:
        return {}
    if "|" not in text:
        raise ValueError("Mapeamento salvo invalido.")
    squad, slug = text.split("|", 1)
    squad = _normalize_squad_form(squad)
    if squad != project.squad:
        raise ValueError("Este mapeamento pertence a outro squad e nao pode ser usado neste projeto.")
    template = load_mapping_template(squad, slug)
    if not template:
        raise ValueError("Mapeamento salvo nao encontrado.")
    return template


def _mapping_template_metadata(template: dict) -> dict:
    if not template:
        return {}
    return {
        "squad": str(template.get("squad") or ""),
        "slug": str(template.get("slug") or ""),
        "name": str(template.get("name") or template.get("slug") or ""),
    }


def _selected_mapping_template(metadata: dict) -> dict:
    selected = metadata.get("mapping_template") or {}
    squad = str(selected.get("squad") or "")
    slug = str(selected.get("slug") or "")
    if not squad or not slug:
        return {}
    return load_mapping_template(squad, slug) or {}


def _apply_mapping_template_to_analysis(
    job_dir: Path,
    analysis: AnalysisResult,
    skip_targets: set[str] | None = None,
) -> tuple[AnalysisResult, dict]:
    metadata = _load_job_metadata(job_dir)
    template = _selected_mapping_template(metadata)
    if not template:
        return analysis, {"state": "disabled", "message": "Nenhum mapeamento salvo selecionado."}

    skip_targets = skip_targets or set()
    entries = template.get("entries") or {}
    updatable_targets = [target for target in analysis.targets if target.object_type in {"chart", "table"}]
    target_ids = {target.target_id for target in updatable_targets}
    alias_to_target = {
        alias: target.target_id
        for target in updatable_targets
        for alias in target_aliases(target)
    }

    # Camada 4: resolve por CONTEUDO (fingerprint do target + assinatura do
    # datasource), resistente a renome dos XLSX e a recriacao do deck. Cai para o
    # match por nome/alias automaticamente dentro do resolvedor.
    resolved = resolve_learned_matches(entries, updatable_targets, analysis.sources)
    saved_matches = {
        target_id: match
        for target_id, match in resolved.items()
        if target_id not in skip_targets
    }
    resolved_targets = set(saved_matches)
    matched_entry_ids = {
        alias_to_target.get(str(raw_id), str(raw_id)) for raw_id in entries
    } & target_ids
    missing_sources = sorted(matched_entry_ids - resolved_targets)

    mapped_analysis = apply_saved_source_matches_to_analysis(analysis, saved_matches)
    new_targets = sorted(target_ids - matched_entry_ids - resolved_targets)
    message = (
        f"Mapeamento '{template.get('name')}' aplicado: {len(saved_matches)} target(s) reconhecido(s)."
    )
    if new_targets:
        message += f" {len(new_targets)} target(s) novo(s) fora do mapeamento."
    if missing_sources:
        message += f" {len(missing_sources)} datasource(s) salvo(s) nao apareceram no ZIP atual."
    return mapped_analysis, {
        "state": "ok" if saved_matches else "warn",
        "message": message,
        "template": _mapping_template_metadata(template),
        "matched_count": len(saved_matches),
        "missing_source_count": len(missing_sources),
        "new_target_count": len(new_targets),
    }


def _mapping_template_candidates(squad: str, targets: list) -> list[dict]:
    updatable_targets = [target for target in targets if target.object_type in {"chart", "table"}]
    target_ids = {target.target_id for target in updatable_targets}
    alias_to_target = {
        alias: target.target_id
        for target in updatable_targets
        for alias in target_aliases(target)
    }
    if not target_ids:
        return []
    candidates = []
    for template_ref in list_mapping_templates(squad):
        template = load_mapping_template(squad, template_ref.slug) or {}
        entries = template.get("entries") or {}
        entry_ids = {alias_to_target.get(str(entry_id), str(entry_id)) for entry_id in entries}
        matched = sorted(target_ids & entry_ids)
        if not matched:
            continue
        candidates.append(
            {
                "ref": f"{template_ref.squad}|{template_ref.slug}",
                "name": template_ref.name,
                "slug": template_ref.slug,
                "updated_at": template_ref.updated_at,
                "entry_count": template_ref.entry_count,
                "matched_count": len(matched),
                "new_target_count": len(target_ids - entry_ids),
                "score": round(len(matched) / len(target_ids) * 100, 1),
            }
        )
    return sorted(candidates, key=lambda item: (item["score"], item["matched_count"]), reverse=True)


def _mapping_status_without_selection(candidates: list[dict]) -> dict:
    if not candidates:
        return {"state": "disabled", "message": "Nenhum mapeamento salvo do squad atual bateu com este PPT."}
    best = candidates[0]
    return {
        "state": "suggested",
        "message": (
            f"Mapeamento salvo encontrado no squad atual: {best['name']} "
            f"({best['matched_count']} target(s), {best['score']}%)."
        ),
    }


def _load_ai_diagnostics(job_dir: Path) -> dict:
    cache_path = job_dir / "ai_diagnostics.json"
    if not cache_path.exists():
        return {}
    return json.loads(cache_path.read_text(encoding="utf-8"))


def _load_ai_source_matches(job_dir: Path) -> dict:
    cache_path = job_dir / "ai_source_matches.json"
    if not cache_path.exists():
        return {}
    return json.loads(cache_path.read_text(encoding="utf-8"))


def _ai_waiting_status(pause_for_mapping: bool) -> tuple[dict[str, str], dict[str, str]]:
    if not ai_configured(PROJECT_ROOT):
        status = {"state": "disabled", "message": "IA indisponivel: configure OPENAI_API_KEY no .env."}
        return status, status
    if pause_for_mapping:
        status = {
            "state": "disabled",
            "message": "IA aguardando: selecione o mapeamento salvo ou revise um target especifico com IA.",
        }
        return status, status
    status = {"state": "disabled", "message": "IA disponivel por target. Abra um card e clique em Revisar este target com IA."}
    return status, status


def _cards_by_slide(
    analysis: AnalysisResult,
    manual_names: dict[str, str],
    manual_ranges: dict[str, str],
    ai_diagnostics: dict[str, dict],
    slide_ai_state: dict | None = None,
    apply_slide_ai_outputs: bool = False,
) -> dict[int, list[dict]]:
    preview_by_target = {item.target: item for item in analysis.preview}
    plan_by_target = {plan.target_id: plan for plan in analysis.plans}
    cards: dict[int, list[dict]] = defaultdict(list)
    for target in analysis.targets:
        slide_state = ((slide_ai_state or {}).get("slides") or {}).get(str(target.slide_number), {})
        target_output = (slide_state.get("target_outputs") or {}).get(target.target_id, {}) if apply_slide_ai_outputs else {}
        target_approval = (slide_state.get("target_approvals") or {}).get(target.target_id, {})
        item = preview_by_target.get(target.target_id)
        plan = plan_by_target.get(target.target_id)
        datasource = item.datasource if item else ""
        status = "mapped" if item else "unmapped"
        if manual_names.get(target.target_id):
            status = "manual"
        elif target_output:
            status = "ai"
        elif ai_diagnostics.get(target.target_id):
            status = "ai"
        elif item and item.reason.startswith("IA sugeriu"):
            status = "ai"
        origin_label = {"mapped": "Determinístico", "ai": "IA", "manual": "Manual"}.get(status, "")
        cards[target.slide_number].append(
            {
                "slide": target.slide_number,
                "target": target.target_id,
                "visual_label": visual_label(target.target_id),
                "shape_name": target.shape_name,
                "object_type": target.object_type,
                "nearby_text": target.nearby_text,
                "has_plan": item is not None,
                "datasource": datasource,
                "chart_shape": _chart_shape_for_target(target),
                "series_formats": _chart_series_format_controls(plan) if plan and target.object_type == "chart" else [],
                "action": item.action if item else "aguardando_datasource",
                "reason": item.reason if item else "Nenhum datasource compativel foi escolhido automaticamente.",
                "confidence": item.confidence if item else None,
                "headers": item.headers if item else [],
                "rows": item.rows if item else [],
                "manual_file": manual_names.get(target.target_id, ""),
                "manual_range": manual_ranges.get(target.target_id, ""),
                "ppt_contract": _ppt_contract_for_target(target),
                "source_detected": _source_detected_for_plan(plan) if plan else None,
                "source_context": _source_context_for_plan(plan) if plan else {},
                "ai": ai_diagnostics.get(target.target_id),
                "slide_ai": target_output,
                "approval": target_approval,
                "status": status,
                "origin_label": origin_label,
                "search_text": " ".join(
                    [
                        str(target.slide_number),
                        target.target_id,
                        target.shape_name,
                        target.object_type,
                        datasource,
                        target.nearby_text,
                        " ".join((plan.datasource.metadata or {}).values()) if plan else "",
                        item.reason if item else "",
                    ]
                ).lower(),
            }
        )
    return cards


def _apply_chart_format_overrides(analysis: AnalysisResult, metadata: dict) -> AnalysisResult:
    configured = metadata.get("chart_format_overrides") or {}
    plans = [
        replace(
            plan,
            series_format_overrides={
                str(key): str(value)
                for key, value in (configured.get(plan.target_id) or {}).items()
                if str(value) in {"percent", "number"}
            },
        )
        if plan.object_type == "chart"
        else plan
        for plan in analysis.plans
    ]
    return replace(analysis, plans=plans)


def _chart_series_format_controls(plan) -> list[dict]:
    automatic_formats = resolved_series_number_formats(
        plan.target,
        replace(plan, series_format_overrides={}),
    )
    formats = resolved_series_number_formats(plan.target, plan)
    controls = []
    for index, label in enumerate(plan.series):
        override = (
            plan.series_format_overrides.get(label)
            or plan.series_format_overrides.get(f"__index_{index}")
            or "auto"
        )
        number_format = formats[index] if index < len(formats) else "0.0"
        automatic_format = (
            automatic_formats[index]
            if index < len(automatic_formats)
            else "0.0"
        )
        controls.append(
            {
                "index": index,
                "label": label or f"Série {index + 1}",
                "mode": override,
                "effective_format": number_format,
                "effective_label": "Percentual" if "%" in number_format else "Número",
                "automatic_label": "Percentual" if "%" in automatic_format else "Número",
            }
        )
    return controls


def _ai_diagnostics_for_job(
    job_dir: Path,
    analysis: AnalysisResult,
    allow_ai: bool = True,
    target_ids: set[str] | None = None,
) -> tuple[dict[str, dict], dict[str, str]]:
    if not ai_configured(PROJECT_ROOT):
        return {}, {"state": "disabled", "message": "IA indisponivel: configure OPENAI_API_KEY no .env."}
    if not analysis.plans:
        return {}, {"state": "warn", "message": "IA nao tem planos para diagnosticar depois do mapeamento."}
    cache_path = job_dir / "ai_diagnostics.json"
    payload: dict[str, dict] = {}
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    eligible_plans = analysis.plans
    if target_ids is not None:
        eligible_plans = [plan for plan in analysis.plans if plan.target_id in target_ids]
        if not eligible_plans:
            return payload, {"state": "warn", "message": "Target alterado nao tem plano para diagnostico IA."}
    missing_plans = [plan for plan in eligible_plans if plan.target_id not in payload]
    if not missing_plans:
        return payload, {"state": "cached", "message": "Diagnostico IA carregado do cache."}
    if not allow_ai:
        return payload, {
            "state": "cached" if payload else "warn",
            "message": f"Preview usou apenas dados salvos; {len(missing_plans)} target(s) sem diagnostico cached nao foram enviados para IA.",
        }
    batch_limit = _ai_diagnostic_batch_limit()
    plans_to_send = missing_plans[:batch_limit]
    remaining_count = max(len(missing_plans) - len(plans_to_send), 0)
    try:
        sent_at = _now_iso()
        started = time.perf_counter()
        payload_summary = _diagnostic_payload_summary(plans_to_send)
        diagnostics = _call_with_job_debug(job_dir, suggest_transform_diagnostics, plans_to_send, root=PROJECT_ROOT)
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
                for item in diagnostics
            }
        )
        for plan in plans_to_send:
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
                "target_count": len(plans_to_send),
                "returned_count": len(diagnostics),
                "remaining_count": remaining_count,
                "payload_summary": payload_summary,
            },
        )
        message = (
            f"Diagnostico IA concluiu {len(diagnostics)} target(s) novo(s) "
            f"neste lote de {len(plans_to_send)}."
        )
        if remaining_count:
            message += f" Restam {remaining_count} target(s); clique em Revisar com IA para continuar."
        return payload, {"state": "ok", "message": message}
    except Exception as exc:
        _append_ai_log(
            job_dir,
            {
                "operation": "transform_diagnostics",
                "status": "error",
                "sent_at": sent_at if "sent_at" in locals() else _now_iso(),
                "returned_at": _now_iso(),
                "duration_ms": round((time.perf_counter() - started) * 1000) if "started" in locals() else 0,
                "target_count": len(plans_to_send),
                "error": format_ai_error(exc),
                "remaining_count": remaining_count,
                "payload_summary": _diagnostic_payload_summary(plans_to_send),
            },
        )
        return payload, {"state": "warn", "message": f"IA indisponivel nesta analise: {format_ai_error(exc)}"}


def _ai_source_matches_for_job(
    job_dir: Path,
    analysis: AnalysisResult,
    allow_ai: bool = True,
    target_ids: set[str] | None = None,
) -> tuple[dict[str, dict], dict[str, str]]:
    if not ai_configured(PROJECT_ROOT):
        return {}, {"state": "disabled", "message": "IA indisponivel: configure OPENAI_API_KEY no .env."}

    try:
        metadata = _load_job_metadata(job_dir)
    except Exception:
        metadata = {}
    manual_target_ids = set(_manual_source_names(job_dir))
    plans_by_target = {plan.target_id: plan for plan in analysis.plans}
    planned_targets = set(plans_by_target)
    reviewable_targets = [
        target
        for target in analysis.targets
        if target.object_type in {"chart", "table"} and target.target_id not in manual_target_ids
    ]
    if target_ids is not None:
        reviewable_targets = [target for target in reviewable_targets if target.target_id in target_ids]

    unmatched = [target for target in reviewable_targets if target.target_id not in planned_targets]
    # Fix A: so vale a pena perguntar a IA sobre um target pendente se existe pelo
    # menos um datasource minimamente compativel no ZIP. Sem candidato plausivel, o
    # problema e falta de XLSX - deixamos pendente para upload manual, sem gastar IA.
    plausibility_floor = _ai_source_match_plausibility_floor()

    def _best_candidate_score(target) -> float:
        candidates = source_match_candidates(target, analysis.sources, limit=1)
        return candidates[0].score if candidates else 0.0

    best_candidate_scores = {target.target_id: _best_candidate_score(target) for target in unmatched}
    no_source_targets = [t for t in unmatched if best_candidate_scores.get(t.target_id, 0.0) < plausibility_floor]
    unmatched = [t for t in unmatched if best_candidate_scores.get(t.target_id, 0.0) >= plausibility_floor]
    confidence_floor = _ai_review_confidence_floor()
    selected_mapping_template = bool(_selected_mapping_template(metadata))
    low_confidence = []
    if not selected_mapping_template:
        low_confidence = [
            target
            for target in reviewable_targets
            if target.target_id in plans_by_target
            and float(plans_by_target[target.target_id].confidence or 0) < confidence_floor
        ]
    review_targets = _unique_targets([*unmatched, *low_confidence])
    if not review_targets:
        if no_source_targets:
            return {}, {
                "state": "warn",
                "message": (
                    f"{len(no_source_targets)} target(s) sem datasource compativel no ZIP "
                    "(envie o XLSX manualmente); IA nao foi acionada."
                ),
            }
        return {}, {
            "state": "ok",
            "message": f"IA nao precisou revisar matches de datasource abaixo de {confidence_floor:.0%}.",
        }

    cache_path = job_dir / "ai_source_matches.json"
    payload: dict[str, dict] = {}
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    missing_targets = [target for target in review_targets if target.target_id not in payload]
    if not missing_targets:
        return payload, {"state": "cached", "message": f"Matches IA carregados do cache ({len(payload)} sugestao/oes)."}
    if not allow_ai:
        return payload, {
            "state": "cached" if payload else "warn",
            "message": f"Download usou cache de IA; {len(missing_targets)} target(s) sem match cached nao foram enviados para IA.",
        }
    batch_limit = _ai_source_match_batch_limit()
    max_calls = _ai_source_match_max_calls()
    low_confidence_ids = {target.target_id for target in low_confidence}
    log_debug_event(
        "source_match_review_targets",
        {
            "confidence_floor": confidence_floor,
            "unmatched_targets": [target.target_id for target in unmatched],
            "low_confidence_targets": [
                {
                    "target": target.target_id,
                    "confidence": plans_by_target[target.target_id].confidence,
                    "datasource": plans_by_target[target.target_id].datasource.file_name,
                }
                for target in low_confidence
            ],
            "missing_target_count": len(missing_targets),
            "no_source_targets": [target.target_id for target in no_source_targets],
            "plausibility_floor": plausibility_floor,
            "batch_limit": batch_limit,
            "max_calls": max_calls,
            "selected_mapping_template": selected_mapping_template,
        },
    )

    # Cada chamada leva um lote pequeno (payload enxuto por slide), mas o loop cobre
    # TODOS os targets pendentes do deck numa unica passada de preview - em vez de
    # parar no primeiro lote e exigir que o usuario clique varias vezes em decks de
    # 100+ slides. O cache e salvo apos cada lote, entao uma falha no meio preserva
    # tudo que ja foi revisado. max_calls limita o custo no pior caso.
    pending = list(missing_targets)
    sent_count = 0
    suggestion_count = 0
    calls_made = 0
    error_message = ""
    while pending and calls_made < max_calls:
        targets_to_send = pending[:batch_limit]
        pending = pending[batch_limit:]
        calls_made += 1
        review_target_ids = {target.target_id for target in targets_to_send}
        existing_plan_ids = planned_targets - review_target_ids
        sent_at = _now_iso()
        started = time.perf_counter()
        payload_summary = _match_payload_summary(targets_to_send, analysis.sources)
        try:
            suggestions = _call_with_job_debug(
                job_dir,
                suggest_source_matches_with_ai,
                targets_to_send,
                analysis.sources,
                existing_plan_ids=existing_plan_ids,
                root=PROJECT_ROOT,
            )
        except Exception as exc:
            _append_ai_log(
                job_dir,
                {
                    "operation": "source_match",
                    "status": "error",
                    "sent_at": sent_at,
                    "returned_at": _now_iso(),
                    "duration_ms": round((time.perf_counter() - started) * 1000),
                    "target_count": len(targets_to_send),
                    "error": format_ai_error(exc),
                    "batch_index": calls_made,
                    "remaining_count": len(pending),
                    "payload_summary": payload_summary,
                },
            )
            error_message = format_ai_error(exc)
            break
        duration_ms = round((time.perf_counter() - started) * 1000)
        sent_count += len(targets_to_send)
        suggestion_count += len(suggestions)
        for target in targets_to_send:
            previous_plan = plans_by_target.get(target.target_id)
            payload.setdefault(
                target.target_id,
                {
                    "datasource": "",
                    "confidence": 0,
                    "reason": "IA nao encontrou match confiavel para este target.",
                    "status": "no_match",
                    "replace_existing": target.target_id in low_confidence_ids,
                    "review_reason": "low_confidence" if target.target_id in low_confidence_ids else "unmatched",
                    "previous_datasource": previous_plan.datasource.file_name if previous_plan else "",
                    "previous_confidence": previous_plan.confidence if previous_plan else None,
                },
            )
        payload.update(
            {
                item.target: {
                    "datasource": item.datasource,
                    "confidence": item.confidence,
                    "reason": item.reason,
                    "recipe_suggestion": item.recipe_suggestion,
                    "status": "matched",
                    "replace_existing": item.target in low_confidence_ids,
                    "review_reason": "low_confidence" if item.target in low_confidence_ids else "unmatched",
                    "previous_datasource": (
                        plans_by_target[item.target].datasource.file_name if item.target in plans_by_target else ""
                    ),
                    "previous_confidence": (
                        plans_by_target[item.target].confidence if item.target in plans_by_target else None
                    ),
                }
                for item in suggestions
            }
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
                "target_count": len(targets_to_send),
                "returned_count": len(suggestions),
                "batch_index": calls_made,
                "remaining_count": len(pending),
                "payload_summary": payload_summary,
            },
        )

    remaining_count = len(pending)
    suffix = ""
    if remaining_count:
        suffix = f" Restam {remaining_count} target(s) alem do limite de {max_calls} chamada(s); clique em Revisar com IA para continuar."
    if error_message:
        state = "warn"
        message = f"IA indisponivel para match de datasource: {error_message}"
        if sent_count:
            message = (
                f"IA revisou {sent_count} target(s) em {calls_made - 1} chamada(s) antes de falhar: {error_message}"
            )
        return payload, {"state": state, "message": message}
    if suggestion_count:
        return payload, {
            "state": "ok",
            "message": (
                f"IA revisou {sent_count} match(es) de datasource em {calls_made} chamada(s) "
                f"e retornou {suggestion_count} sugestao/oes confiaveis.{suffix}"
            ),
        }
    return payload, {
        "state": "warn",
        "message": f"IA revisou {sent_count} target(s), mas nao encontrou novos matches confiaveis de datasource.{suffix}",
    }


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
                "target": target.target_id,
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
    cached["render_cache_version"] = RENDER_CACHE_VERSION
    cached["notice"] = ""
    cached["error"] = ""
    cached["ai_log_entries"] = []
    _render_cache_path(job_dir).write_text(json.dumps(cached, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_render_cache(job_dir: Path) -> dict:
    cache_path = _render_cache_path(job_dir)
    if not cache_path.exists():
        return {}
    payload = json.loads(cache_path.read_text(encoding="utf-8"))
    if payload.get("render_cache_version") != RENDER_CACHE_VERSION:
        return {}
    return payload


def _clear_render_cache(job_dir: Path) -> None:
    cache_path = _render_cache_path(job_dir)
    if cache_path.exists():
        cache_path.unlink()
    rendered_dir = job_dir / "rendered"
    if rendered_dir.exists() and rendered_dir.is_dir():
        shutil.rmtree(rendered_dir, ignore_errors=True)


def _clear_ai_cache(
    job_dir: Path,
    target_id: str | None = None,
    cache_names: tuple[str, ...] = ("ai_diagnostics.json", "ai_source_matches.json"),
) -> None:
    for cache_name in cache_names:
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
    _clear_slide_ai_state(job_dir, target_id=target_id)


def _clear_slide_ai_state(job_dir: Path, target_id: str | None = None) -> None:
    state_path = job_dir / "slide_ai_state.json"
    if not state_path.exists():
        return
    state = json.loads(state_path.read_text(encoding="utf-8"))
    if target_id is None:
        state_path.unlink()
        return
    changed = False
    for slide_state in (state.get("slides") or {}).values():
        outputs = slide_state.get("target_outputs") or {}
        if target_id in outputs:
            outputs.pop(target_id, None)
            slide_state.pop("input_signature", None)
            changed = True
    if changed:
        state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _ppt_contract_for_target(target) -> dict:
    if target.object_type == "chart":
        if target.expected_orientation == "series_rows_categories_columns":
            return {
                "orientation": target.expected_orientation,
                "headers": ["", *target.expected_categories],
                "rows": [
                    [target.expected_series[index] if index < len(target.expected_series) else "", *row]
                    for index, row in enumerate(target.expected_values)
                ],
            }
        return {
            "orientation": target.expected_orientation,
            "headers": ["", *target.expected_series],
            "rows": [
                [target.expected_categories[index] if index < len(target.expected_categories) else "", *row]
                for index, row in enumerate(target.expected_values)
            ],
        }
    if target.object_type == "table":
        return {"orientation": "table_cells", "headers": [], "rows": target.table_cells}
    return {"orientation": target.object_type, "headers": [], "rows": []}


def _source_detected_for_plan(plan) -> dict:
    return {
        "orientation": plan.datasource.orientation,
        "context": _source_context_for_plan(plan),
        "headers": _display_row(plan.datasource.preview_rows[0]) if plan.datasource.preview_rows else [],
        "rows": [_display_row(row) for row in plan.datasource.preview_rows[1:9]] if plan.datasource.preview_rows else [],
    }


def _source_context_for_plan(plan) -> dict:
    metadata = plan.datasource.metadata or {}
    return {
        "table_title": metadata.get("table_title", ""),
        "row_group_label": metadata.get("row_group_label", ""),
        "context_text": metadata.get("context_text", ""),
    }


def _display_row(row: list) -> list:
    return ["" if value is None else value for value in row]


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


def _save_project_checkpoint(
    job_dir: Path,
    status: str = "in_progress",
    include_inputs: bool = False,
    reason: str = "autosave",
) -> None:
    metadata = _load_job_metadata(job_dir)
    project_meta = metadata.get("project", {})
    project = load_project(project_meta.get("squad", ""), project_meta.get("slug", ""))
    if project is None:
        return

    try:
        previous = load_project_json(project, ["checkpoint"], "checkpoint.json")
    except FileNotFoundError:
        previous = {}
    inputs_persisted = bool(previous.get("inputs_persisted")) or include_inputs
    manual_overrides = {}
    previous_overrides = previous.get("manual_overrides") or {}
    for target_id, (filename, data, cell_range) in _manual_sources_for_job(job_dir).items():
        digest = hashlib.sha256(data).hexdigest()
        manual_overrides[target_id] = {
            "filename": filename,
            "range": cell_range,
            "sha256": digest,
        }
        previous_override = previous_overrides.get(target_id) or {}
        if previous_override.get("filename") != filename or previous_override.get("sha256") != digest:
            save_project_bytes(project, ["checkpoint", "overrides", target_id], filename, data)

    checkpoint = {
        "schema_version": 2,
        "status": status,
        "job_id": metadata.get("job_id"),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "save_reason": reason,
        "save_count": int(previous.get("save_count") or 0) + 1,
        "inputs_persisted": inputs_persisted,
        "metadata": metadata,
        "manual_overrides": manual_overrides,
        "caches": {
            cache_name: (job_dir / cache_name).exists()
            for cache_name in ("ai_source_matches.json", "ai_diagnostics.json", "render_cache.json")
        },
        "logs": {
            "ai_usage": (job_dir / "logs" / "ai_usage.jsonl").exists(),
        },
        "slide_ai_state": (job_dir / "slide_ai_state.json").exists(),
    }
    # Inputs sao imutaveis. Sobem uma unica vez, antes do worker iniciar. Os
    # autosaves seguintes gravam apenas JSON/caches pequenos: retomada robusta
    # sem reenviar um deck de centenas de MB a cada clique.
    if include_inputs or not inputs_persisted:
        save_project_bytes(project, ["checkpoint"], "input.pptx", (job_dir / "input.pptx").read_bytes())
        save_project_bytes(project, ["checkpoint"], "datasources.zip", (job_dir / "datasources.zip").read_bytes())
        mapping_path = job_dir / "mapping.xlsx"
        if mapping_path.exists():
            save_project_bytes(project, ["checkpoint"], "mapping.xlsx", mapping_path.read_bytes())
        checkpoint["inputs_persisted"] = True

    for cache_name in ("ai_source_matches.json", "ai_diagnostics.json", "render_cache.json"):
        cache_path = job_dir / cache_name
        if cache_path.exists():
            save_project_bytes(project, ["checkpoint"], cache_name, cache_path.read_bytes())
    slide_state_path = job_dir / "slide_ai_state.json"
    if slide_state_path.exists():
        save_project_bytes(project, ["checkpoint"], "slide_ai_state.json", slide_state_path.read_bytes())
    ai_log_path = job_dir / "logs" / "ai_usage.jsonl"
    if ai_log_path.exists():
        save_project_bytes(project, ["checkpoint", "logs"], "ai_usage.jsonl", ai_log_path.read_bytes())
    save_project_json(project, ["checkpoint"], "checkpoint.json", checkpoint)


def _load_job_checkpoint_summary(job_dir: Path) -> dict:
    try:
        metadata = _load_job_metadata(job_dir)
        project_meta = metadata.get("project") or {}
        project = load_project(project_meta.get("squad", ""), project_meta.get("slug", ""))
        if project is None:
            return {}
        return load_project_json(project, ["checkpoint"], "checkpoint.json")
    except (FileNotFoundError, ValueError):
        return {}


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
    metadata["use_ai"] = False
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
    slide_state_path = job_dir / "slide_ai_state.json"
    if checkpoint.get("slide_ai_state"):
        try:
            slide_state_path.write_bytes(load_project_bytes(project, ["checkpoint"], "slide_ai_state.json"))
        except FileNotFoundError:
            if slide_state_path.exists():
                slide_state_path.unlink()
    elif slide_state_path.exists():
        slide_state_path.unlink()
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
    generated_by = audit.remembered_actor(job_dir) or audit.ANONYMOUS_ACTOR
    run = create_run(
        project,
        {
            "job_id": metadata.get("job_id"),
            "generated_by": generated_by,
            "generated_by_identified": audit.is_identified(generated_by),
            "targets_found": analysis.target_count,
            "plans_generated": len(analysis.plans),
            "manual_overrides": sorted(_manual_source_names(job_dir)),
            "selected_slides": _selected_slides_for_job(job_dir),
            "mapping_template": metadata.get("mapping_template") or {},
            "slide_ai_state": (job_dir / "slide_ai_state.json").exists(),
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
    slide_state_path = job_dir / "slide_ai_state.json"
    if slide_state_path.exists():
        save_project_bytes(project, ["runs", run.run_id, "state"], "slide_ai_state.json", slide_state_path.read_bytes())
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
    _save_or_update_mapping_template(job_dir, project, analysis)


def _save_or_update_mapping_template(job_dir: Path, project, analysis: AnalysisResult) -> None:
    metadata = _load_job_metadata(job_dir)
    selected_template = _selected_mapping_template(metadata)
    existing_entries = dict((selected_template or {}).get("entries") or {})
    manual_names = _manual_source_names(job_dir)
    now = _now_iso()
    for plan in analysis.plans:
        datasource = manual_names.get(plan.target_id) or plan.datasource.file_name
        existing_entries[plan.target_id] = {
            "target_id": plan.target_id,
            "object_type": plan.object_type,
            "slide": plan.target.slide_number,
            "shape_name": plan.target.shape_name,
            "shape_id": plan.target.shape_id,
            "target_aliases": sorted(target_aliases(plan.target)),
            "datasource": datasource,
            "datasource_basename": Path(datasource).name,
            "action": plan.action,
            "confidence": round(plan.confidence, 4),
            "reason": plan.reason,
            "updated_at": now,
            **mapping_entry_learning_fields(plan.target, plan.datasource),
        }
    if not existing_entries:
        return
    template_name = str((selected_template or {}).get("name") or f"{project.name} - mapeamento")
    template_slug = str((selected_template or {}).get("slug") or "")
    template_ref = save_mapping_template(
        project,
        template_name,
        existing_entries,
        slug=template_slug,
        metadata={
            "last_job_id": metadata.get("job_id"),
            "last_pptx": (metadata.get("files") or {}).get("pptx", ""),
            "last_datasources": (metadata.get("files") or {}).get("datasources", ""),
            "selected_slides": _selected_slides_for_job(job_dir),
            "last_trained_by": audit.remembered_actor(job_dir) or audit.ANONYMOUS_ACTOR,
            "last_trained_at": datetime.now().isoformat(timespec="seconds"),
        },
    )
    audit.record(
        job_dir,
        audit.remembered_actor(job_dir),
        "treinou_memoria_do_squad",
        {"template": template_ref.slug, "entradas": len(existing_entries)},
    )
    metadata["mapping_template"] = {
        "squad": template_ref.squad,
        "slug": template_ref.slug,
        "name": template_ref.name,
    }
    _save_job_metadata(job_dir, metadata)


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


def _inspect_ppt_upload(pptx_bytes: bytes) -> dict:
    with ZipFile(BytesIO(pptx_bytes)) as zf:
        slide_paths = sorted(
            [name for name in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)],
            key=lambda name: int(re.search(r"slide(\d+)\.xml", name).group(1)),
        )
        chart_count = 0
        table_count = 0
        for slide_path in slide_paths:
            root = ET.fromstring(zf.read(slide_path))
            for frame in root.findall(".//p:graphicFrame", PPT_XML_NS):
                if frame.find(".//c:chart", PPT_XML_NS) is not None:
                    chart_count += 1
                elif frame.find(".//a:tbl", PPT_XML_NS) is not None:
                    table_count += 1
    return {
        "slide_count": len(slide_paths),
        "chart_count": chart_count,
        "table_count": table_count,
        "target_count": chart_count + table_count,
    }


def _validate_slide_scope(ppt_summary: dict, selected_slides: list[int]) -> None:
    slide_count = int(ppt_summary.get("slide_count") or 0)
    if not selected_slides or not slide_count:
        return
    invalid = [slide for slide in selected_slides if slide > slide_count]
    if invalid:
        raise ValueError(
            f"O PPT tem {slide_count} slide(s), mas o escopo inclui slide(s) fora do arquivo: "
            f"{', '.join(str(slide) for slide in invalid[:8])}."
        )


def _large_scope_requires_confirmation(ppt_summary: dict, selected_slides: list[int]) -> bool:
    threshold = _large_deck_slide_threshold()
    selected_count = len(selected_slides)
    if selected_count:
        return selected_count > threshold
    return int(ppt_summary.get("slide_count") or 0) > threshold


def _large_deck_confirmation_message(ppt_summary: dict, selected_slides: list[int]) -> str:
    threshold = _large_deck_slide_threshold()
    slide_count = int(ppt_summary.get("slide_count") or 0)
    target_count = int(ppt_summary.get("target_count") or 0)
    scope_count = len(selected_slides) or slide_count
    return (
        f"Este escopo tem {scope_count} slide(s) e o limite de confirmacao e {threshold}. "
        f"O PPT completo tem {slide_count} slide(s) e {target_count} target(s). "
        "Selecione um intervalo menor ou confirme que deseja analisar todos estes slides."
    )


def _async_generation_enabled() -> bool:
    """Gera o PPT em segundo plano e deixa o navegador acompanhar por polling.

    Ligado por padrao: gerar dentro da requisicao funciona num deck pequeno, mas
    um deck grande passa do tempo limite do App Runner e o usuario nao recebe
    arquivo nenhum. Em segundo plano o tempo de geracao deixa de competir com o
    tempo limite da requisicao."""
    return os.getenv("AUTO_PPT_ASYNC_GENERATION", "1").strip().lower() in {"1", "true", "yes", "on"}


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


def _read_existing_datasource(job_dir: Path, entry_name: str) -> tuple[str, bytes]:
    """Pull one XLSX already uploaded (inside datasources.zip) by name, so the user
    can reassign a chart to a known planilha without uploading it again."""
    zip_path = job_dir / "datasources.zip"
    if not zip_path.exists():
        raise ValueError("Não encontrei as planilhas deste projeto.")
    target = entry_name.replace("\\", "/").strip()
    with ZipFile(zip_path) as archive:
        names = [n for n in archive.namelist() if not n.endswith("/")]
        match = next((n for n in names if n == target), None)
        if match is None:
            match = next((n for n in names if n.rsplit("/", 1)[-1] == target.rsplit("/", 1)[-1]), None)
        if match is None:
            raise ValueError(f"Planilha '{entry_name}' não está entre as enviadas.")
        return match.rsplit("/", 1)[-1], archive.read(match)


def _coalesce_datasources_to_zip(payloads: list[tuple[str, bytes]]) -> tuple[bytes, str]:
    """Accept either a single .zip or several loose .xlsx files and always return
    (zip_bytes, human_label). Lets the user drop planilhas directly, no zipping."""
    if not payloads:
        raise ValueError("Envie as planilhas .xlsx (ou um .zip com elas).")
    zips = [(name, data) for name, data in payloads if name.lower().endswith(".zip")]
    xlsx = [(name, data) for name, data in payloads if name.lower().endswith(".xlsx")]
    if zips and not xlsx:
        if len(zips) > 1:
            raise ValueError("Envie um único .zip, ou várias planilhas .xlsx soltas.")
        return zips[0][1], zips[0][0] or "datasources.zip"
    if not xlsx:
        raise ValueError("Formato não suportado. Envie planilhas .xlsx ou um .zip com elas.")

    # Recusa nomes repetidos em vez de renomear escondido: o mapeamento salvo
    # procura a planilha pelo nome do arquivo antes de olhar o conteudo, entao
    # dois arquivos homonimos podem alimentar o grafico errado na proxima vez.
    names = [safe_filename(name or "planilha.xlsx") or "planilha.xlsx" for name, _data in xlsx]
    repeated = sorted({name for name in names if names.count(name) > 1})
    if repeated:
        raise ValueError(
            "Estas planilhas foram enviadas com o mesmo nome: "
            + ", ".join(repeated)
            + ". Renomeie para nomes diferentes e envie de novo — o nome do arquivo é usado "
            "para lembrar qual planilha alimenta cada gráfico."
        )

    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for final_name, (_name, data) in zip(names, xlsx):
            archive.writestr(final_name, data)
    label = xlsx[0][0] if len(xlsx) == 1 else f"{len(xlsx)} planilhas"
    return buffer.getvalue(), label


def _validate_upload(upload: UploadFile, extension: str, message: str) -> None:
    filename = upload.filename or ""
    if not filename.lower().endswith(extension):
        raise ValueError(message)


async def _read_upload_limited(upload: UploadFile, label: str) -> bytes:
    limit = _max_upload_bytes()
    data = await upload.read(limit + 1)
    if len(data) > limit:
        limit_mb = limit // (1024 * 1024)
        raise ValueError(f"{label} excede o limite de {limit_mb} MB.")
    return data


def _max_upload_bytes() -> int:
    return max(_env_int("AUTO_PPT_MAX_UPLOAD_MB", 350), 1) * 1024 * 1024


def _max_request_bytes() -> int:
    return max(_env_int("AUTO_PPT_MAX_REQUEST_MB", 600), 1) * 1024 * 1024


def _save_job_metadata(job_dir: Path, payload: dict) -> None:
    (job_dir / "metadata.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_job_metadata(job_dir: Path) -> dict:
    return json.loads((job_dir / "metadata.json").read_text(encoding="utf-8"))


def _large_deck_slide_threshold() -> int:
    return max(_env_int("AUTO_PPT_LARGE_DECK_SLIDE_THRESHOLD", 10), 1)


def _ai_source_match_batch_limit() -> int:
    default_limit = _env_int("AUTO_PPT_AI_MATCH_TARGET_LIMIT", 10)
    return max(_env_int("AUTO_PPT_AI_SOURCE_MATCH_BATCH_TARGETS", default_limit), 1)


def _ai_source_match_max_calls() -> int:
    # Teto de chamadas de IA por passada de revisao de matches. Com o batch default
    # de 10 targets, o default de 12 cobre ate 120 targets pendentes de uma vez -
    # suficiente para decks de 100+ slides sem deixar o custo do pior caso aberto.
    return max(_env_int("AUTO_PPT_AI_MATCH_MAX_CALLS", 12), 1)


# Politica de confianca do sistema, em 3 niveis independentes (valores default
# inalterados nesta consolidacao; isto so agrupa e documenta o que ja existia
# espalhado pelo arquivo, para ficar claro onde ajustar cada um):
#
# 1. table_normalizer.LOCAL_MATCH_THRESHOLD_STRONG_ID / _DEFAULT: decide se existe
#    ALGUM plano automatico para o target. Abaixo disso o target fica "sem match"
#    e so pode ser resolvido por IA ou override manual.
# 2. _ai_review_confidence_floor() (este arquivo): decide se um plano ja aceito
#    pelo matching local ainda assim recebe uma segunda opiniao enxuta de IA
#    (troca de datasource + receita estrutural) antes do preview ficar pronto.
# 3. _auto_slide_ai_confidence_floor() (este arquivo): decide se a revisao mais
#    pesada por slide (understanding + matriz tipada) roda automaticamente quando
#    AUTO_PPT_AUTO_SLIDE_AI esta ligado.
def _ai_review_confidence_floor() -> float:
    return min(max(_env_float("AUTO_PPT_AI_REVIEW_CONFIDENCE_FLOOR", 0.80), 0.0), 1.0)


def _auto_slide_ai_confidence_floor() -> float:
    return min(max(_env_float("AUTO_PPT_AUTO_SLIDE_AI_CONFIDENCE_FLOOR", 0.82), 0.0), 1.0)


def _ai_source_match_plausibility_floor() -> float:
    # Abaixo deste score bruto do melhor candidato, entende-se que NAO existe um
    # datasource compativel no ZIP para o target (o problema e falta de XLSX, nao
    # de match). Nesses casos nao gastamos uma chamada de IA - o alvo fica pendente
    # para upload manual. Evita, por exemplo, mandar tabelas "Base:" sem fonte.
    return min(max(_env_float("AUTO_PPT_AI_MATCH_PLAUSIBILITY_FLOOR", 0.30), 0.0), 1.0)


def _auto_source_match_review_enabled() -> bool:
    return _env_bool("AUTO_PPT_AI_AUTO_SOURCE_REVIEW", False)


def _source_match_ai_enabled(metadata: dict) -> bool:
    if bool(metadata.get("use_ai")):
        return True
    if metadata.get("mapping_template"):
        return False
    return bool(metadata.get("auto_source_review", _auto_source_match_review_enabled()))


def _ai_diagnostic_batch_limit() -> int:
    return max(_env_int("AUTO_PPT_AI_DIAGNOSTIC_BATCH_TARGETS", 10), 1)


def _env_int(name: str, default: int) -> int:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _env_float(name: str, default: float) -> float:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return float(value)
    except ValueError:
        return default


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on", "sim"}


def _norm_text(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _unique_targets(targets: list) -> list:
    seen = set()
    output = []
    for target in targets:
        if target.target_id in seen:
            continue
        seen.add(target.target_id)
        output.append(target)
    return output


def _projects_by_squad(squads: list[str] | None = None) -> dict[str, list]:
    return {squad: list_projects(squad) for squad in (squads or SQUADS)}


def _mapping_templates_by_squad(squads: list[str] | None = None) -> dict[str, list]:
    return {squad: list_mapping_templates(squad) for squad in (squads or SQUADS)}


def _project_cards_by_squad(squads: list[str] | None = None) -> dict[str, list[dict]]:
    output: dict[str, list[dict]] = {}
    for squad in (squads or SQUADS):
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


def _resume_cards(squads: list[str] | None = None) -> list[dict]:
    """Lista plana de projetos com progresso salvo ('continue de onde parou'),
    do mais recente para o mais antigo - independente de squad."""
    cards: list[dict] = []
    for squad in (squads or SQUADS):
        for project in list_projects(squad):
            checkpoint = _checkpoint_summary(project)
            if not checkpoint:
                continue
            status = str(checkpoint.get("status") or "")
            cards.append(
                {
                    "squad": project.squad,
                    "name": project.name,
                    "slug": project.slug,
                    "status": status,
                    "status_label": _checkpoint_status_label(status),
                    "status_kind": "done" if status.lower() in {"complete", "concluido", "concluído", "done"} else "progress",
                    "slides": _checkpoint_slide_label(checkpoint),
                    "updated_at": str(checkpoint.get("updated_at") or ""),
                    "updated_label": _pretty_datetime(checkpoint.get("updated_at")),
                    "preview_url": f"/projects/{project.squad}/{project.slug}/preview",
                }
            )
    cards.sort(key=lambda card: card["updated_at"], reverse=True)
    return cards


def _checkpoint_status_label(status: str) -> str:
    mapping = {
        "complete": "Concluído",
        "done": "Concluído",
        "in_progress": "Em andamento",
        "processing": "Processando",
        "error": "Com erro",
    }
    return mapping.get(status.lower(), status.replace("_", " ").title() if status else "Em andamento")


def _pretty_datetime(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "").split(".")[0])
    except ValueError:
        return text
    return parsed.strftime("%d/%m/%Y %H:%M")


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


def _squad_labels(squads: list[str] | None = None) -> list[dict[str, str]]:
    return [{"value": squad, "label": squad.title()} for squad in (squads or SQUADS)]


def _normalize_squad_form(value: str) -> str:
    normalized = value.strip().lower()
    if normalized not in SQUADS:
        raise ValueError("Squad invalido.")
    return normalized


def _authorized_form_squad(request: Request, squad: str, project_ref: str = "") -> str:
    user = _request_user(request)
    requested = _normalize_squad_form(squad)
    if project_ref:
        try:
            ref_squad, _slug = project_ref.split("|", 1)
            ref_squad = _normalize_squad_form(ref_squad)
        except ValueError as exc:
            raise ValueError("Projeto selecionado invalido.") from exc
        if ref_squad != requested:
            raise ValueError("O projeto selecionado nao pertence ao squad informado.")
    if user is None or user.is_admin:
        return requested
    if not user.squad or requested != user.squad:
        raise ValueError("Voce so pode criar ou abrir projetos do seu squad.")
    if project_ref and ref_squad != user.squad:
        raise ValueError("Voce so pode abrir projetos do seu squad.")
    return user.squad


def _error_response(request: Request, message: str, status_code: int = 400) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "error.html",
        {
            "message": message,
            "error_title": "Nao consegui concluir esta acao",
            "current_user": _request_user(request),
        },
        status_code=status_code,
    )
