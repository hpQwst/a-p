# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A tool that updates PowerPoint decks from spreadsheet data. Users upload a `.pptx` template and a `.zip` of `.xlsx` datasources; the app discovers updatable targets (charts, PowerPoint tables) in the deck, matches each one to the right datasource (deterministically, with optional AI assistance), normalizes/transposes the data, and writes a new `.pptx` that still opens correctly in PowerPoint (including "Edit Data" on charts).

The product is organized around **Squads** (`squad1`–`squad5`) → **Projetos** → **Modelos de mapeamento** (reusable per-squad mapping memory) → **Execuções** (each generation run, saved with inputs/output/report). See `ppt_automator/project_store.py`.

## Running the app

```powershell
pip install -r requirements.txt
uvicorn web.main:app --host 0.0.0.0 --port 8501
```

This is the current, actively developed entry point (FastAPI + Jinja2 templates in `web/`). Upload the model `.pptx` and the datasources `.zip` through the UI; the flow is `Projeto` → `Arquivos` → `Preview` → `Download`.

`app.py` (Streamlit) is a **legacy** entry point that targets the older `ppt_automator/core.py` API (numeric-named charts only, no tables/text targets). Don't extend it — new work goes through `web/main.py` + the `ppt_automator` "core novo" modules.

### AI (optional)

The app works without AI; when `OPENAI_API_KEY` is set (via `.env`, based on `.env.example`) it's used by default during preview. Validate the OpenAI connection before starting the server:

```powershell
.\.venv\Scripts\python.exe scripts\check_openai.py
```

Key env vars (see `.env.example` for the full set and defaults): `OPENAI_API_KEY`, `OPENAI_MODEL`, `AUTO_PPT_AI_AUTO_SOURCE_REVIEW`, `AUTO_PPT_AI_REVIEW_CONFIDENCE_FLOOR`, `AUTO_PPT_AUTO_SLIDE_AI`, `AUTO_PPT_APPLY_SLIDE_AI_OUTPUTS`, `AUTO_PPT_STORAGE_BACKEND` (`local` or `s3`).

## Tests

Tests use `unittest`, run per-module (no pytest in `requirements.txt`):

```powershell
python -m unittest tests.test_mb_update_targets
python -m unittest discover tests
python scripts/smoke_test.py
```

Several regression tests (`test_mb_update_targets.py`, `test_hugo_matching.py`, `test_andre_web_flow.py`) depend on fixture decks that live **outside** the repo (e.g. `C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\mb`), configurable via env vars like `AUTO_PPT_MB_TEST_DIR`, `AUTO_PPT_HUGO_TEST_DIR`, `AUTO_PPT_ANDRE_TEST_DIR`. They `unittest.skipUnless(...)` themselves when the fixtures are absent — don't be surprised by skipped tests locally, and don't try to fabricate the missing fixtures.

## Architecture

### The core pipeline (`ppt_automator/`)

The "core novo" works with a generic `PptTarget` (not just charts) — a slide can have several updatable targets: real PowerPoint charts, PowerPoint tables, and eventually text boxes/shapes. Pipeline stages, roughly in data-flow order:

1. `ppt_discovery.py` — discovers `PptTarget`s in the PPTX (`chart`, `table`, `text`, `shape`).
2. `xlsx_parser.py` — parses XLSX datasources without assuming a fixed legacy layout.
3. `slide_datasources.py` — scopes which datasource files are relevant to which slide.
4. `table_normalizer.py` — builds the `TransformPlan` (align vs. transpose) per target.
5. `ai_mapper.py` / `ai_transform.py` / `ai_slide_understanding.py` / `ai_slide_matrix_builder.py` — optional AI review layers (see below).
6. `ppt_chart_writer.py` / `embedded_workbook_writer.py` — writes chart XML + the chart's embedded Excel workbook.
7. `ppt_table_writer.py` — writes PowerPoint table cells directly in DrawingML XML.
8. `preview_model.py` — builds the UI-friendly preview model.
9. `engine.py` — orchestrates analysis, preview, and final `.pptx` generation (`analyze_update_package`, `generate_updated_pptx`).

`ppt_automator/__init__.py` re-exports both the new engine API (`analyze_update_package`, `generate_updated_pptx`, `discover_ppt_targets`, `PptTarget`, `TransformPlan`, `parse_datasource_zip`, ...) and the legacy `core.py` API used by `app.py` (`build_chart_jobs`, `generate_pptx`, `ChartTarget`, ...). When adding functionality, extend the new-engine modules, not `core.py`.

### Matching datasources to targets

Datasource XLSX files don't need to be named after the PPT chart's numeric shape name (e.g. `7792738590`). When the name matches, that's used as a shortcut; otherwise the system compares columns, rows, table question, mapping variable/opening, and optional metadata embedded in the XLSX to suggest the most likely datasource (`table_normalizer.py` source-match scoring, plus optional embedded hint rows like `PPT_TAG`, `graph_id`, `var_analise`, `abertura`).

Once a project has a successful download, the system creates/updates a **mapping template** for that project's squad. On the next update, if targets and datasource names still line up, the template's de-para is applied before AI runs; new targets show up in the preview to be added to the same template.

### Preserving "Edit Data" (no python-pptx chart.replace_data / no full openpyxl workbook save)

To keep a chart's "Edit Data" working and to run without Microsoft Office, the writer path is **serverless OOXML surgery**: open the `.pptx`/`.xlsx` as ZIP/OPC packages, patch only the necessary parts (`openxml_zip.py`, `embedded_workbook_writer.py`), and update the chart's visual XML cache — never `python-pptx`'s chart replace, never a full `openpyxl.save()` of the embedded workbook. This is what makes generation work in headless Linux containers (App Runner) without Office/COM. Don't reintroduce those shortcuts even if they look simpler — they were deliberately avoided (see `docs/pptx_serverless_openxml_poc.md` and the README section "Gráficos editáveis e Excel embutido").

Formula evaluation for datasource XLSX uses only the internal, AST-restricted evaluator (`core.py: _SimpleFormulaEvaluator`) covering `SUM`/`SOMA`, `SUMPRODUCT`/`SOMARPRODUTO`, `AVERAGE`/`MEDIA`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`/`SE`, `SUMIF`/`SOMASE`, `COUNTIF`/`CONT.SE`. Unsupported formulas fail by default; `AUTO_PPT_FORMULA_FALLBACK=cached` is an explicit opt-in to use a value already cached in the XLSX. Original files are never mutated in place. XLSX files with pivot-cache metadata are read from an in-memory sanitized copy because recent `openpyxl` versions can reject legacy pivot records; the source file remains untouched.

Preview and per-slide AI review use extracted OpenXML contracts, titles and structured XLSX text only. Never add Office, COM, LibreOffice, PDF conversion or slide-image rendering back to the runtime.

### AI layers (all optional, degrade gracefully without a key)

- **Enxuta (lean) source-match review**: runs automatically for unmatched targets or deterministic matches below `AUTO_PPT_AI_REVIEW_CONFIDENCE_FLOOR` (default 0.80). Chooses the datasource per target and may suggest a small structural `recipe_suggestion` (`keep`, `transpose`, `drop_and_keep`, `drop_and_transpose`); the matrix itself is still built and validated by deterministic code, never written directly by the AI.
- **Per-slide AI** (`AUTO_PPT_AUTO_SLIDE_AI=1`): does not run automatically on the initial preview (keeps it fast, avoids the AI restructuring targets the deterministic normalizer already mapped well). Enable explicitly to investigate a hard deck; `AUTO_PPT_APPLY_SLIDE_AI_OUTPUTS=1` is required to actually apply AI-generated matrices to the final PPT.
- Cost/latency controls: `AUTO_PPT_AI_XLSX_DUMP_MODE` (`compact` default, `verbose` for deep investigation).

### Web layer (`web/main.py`)

FastAPI app. Preview runs as a **background job**: `POST /preview` creates a job dir under `workspace_data/web_jobs/<job_id>/` (or `AUTO_PPT_RUNTIME_ROOT`), writes uploaded files, persists the immutable inputs once in the project checkpoint, and submits work to a `ThreadPoolExecutor` (`_preview_processing_worker`). `GET /jobs/{id}/processing-status` is polled by the frontend while `preview_processing.json` tracks real object progress; generation does the same in `generation_processing.json`. A render cache (`_save_render_cache`/`_load_render_cache`) avoids recomputation once a preview is complete. Most mutating endpoints (`override`, `review-ai`, `mapping-template`, `slides`) clear this cache and autosave only small state/caches. `POST /jobs/{id}/save` is the explicit save button. After a restart, the project checkpoint restores inputs and restarts unfinished analysis. All job-scoped debug events are appended to `log.txt` in the job dir via `ppt_automator/ai_debug.py`.

`worker/processor.py` is the thin orchestration layer between `web/main.py` and `ppt_automator/engine.py`: `AnalysisResult` (targets + sources + plans + preview + warnings), `analyze_files`, and helpers to layer AI source-matches, AI diagnostics, and manual per-target datasource overrides on top of the deterministic analysis.

### Storage (`ppt_automator/project_store.py`)

Single storage abstraction switched by `AUTO_PPT_STORAGE_BACKEND` (`local` default, or `s3` with `AUTO_PPT_S3_BUCKET`/`AUTO_PPT_S3_PREFIX`). In dev, everything lives under `workspace_data/` (git-ignored). Layout: `squads/<squad>/projects/<slug>/{checkpoint,runs,memory}`, `squads/<squad>/mapping_templates/<slug>/template.json`, hashed user profiles under `users/profiles/`, and immutable admin events under `admin_audit/`. Manual mapping corrections are appended to `memory/corrections.json` per project for audit/future learning.

## Auth (`web/auth.py`, `web/entra.py`)

Production uses **Microsoft Entra (OIDC)** only. Authorization Code flow via MSAL is single-tenant: `entra.exchange_code` rejects any `tid` that isn't ours. State and nonce ride in a signed cookie (`qwst_oidc`), not process memory.

On first login, `project_store.ensure_user` creates a profile. Bootstrap emails from `AUTO_PPT_BOOTSTRAP_ADMINS` become admins; everyone else must choose one of `squad1`–`squad5` exactly once. Middleware enforces the squad on project routes, job routes and listings. Common users never change their own squad. Admins use `/admin/users` to change squad, activate/deactivate and promote/revoke admins; every change is recorded.

The old team password is disabled in App Runner with `AUTO_PPT_TEAM_PASSWORD_ENABLED=0`; it remains available only for isolated local development.

The session cookie never carries the password or any Microsoft token — only an expiry, the signed-in email, and an HMAC over both. The signing key comes from `AUTO_PPT_SESSION_SECRET` (falling back to the team password), derived rather than stored so every instance validates the same cookie.

`/static/*`, `/health*`, `/login`, `/logout` and `/auth/*` stay public. **No Entra and no password = app fully open** — fine for local dev, never in production.

Gotcha: MSAL adds `openid`/`profile`/`offline_access` itself and raises `ValueError` if they're passed in, so `auth.SCOPES` only lists `email`.

## Deploy

**AWS App Runner** (`us-east-1`), single container that does everything in-process — no separate workers, no queue. Shared state lives in S3 (`AUTO_PPT_STORAGE_BACKEND=s3`); authorization filters that state so common users only see their assigned squad while admins may select any squad.

Code lives in **Azure DevOps** (`qwst-auto-ppt`); AWS never connects to the repo. `deploy.ps1` packs the current commit with `git archive`, uploads the zip to S3 and CodeBuild builds from that — CodeBuild cannot read Azure Repos natively, and this avoids mirroring or cross-cloud credentials. No local Docker needed, no console setup.

```powershell
.\infra\aws\deploy.ps1 -EntraTenantId "..." -EntraClientId "..." -EntraClientSecret "..." -EntraRedirectUri "https://.../auth/callback"
```

Same command publishes a new version. Infra is `infra/aws/apprunner.yaml` + `infra/aws/deploy.ps1`; the full runbook is in `DEPLOYMENT.md`.

`infra/aws/configure-cost-controls.ps1` maintains the `squad4e5-auto-ppt-monthly` USD 20 budget. Cost allocation tag activation may require the payer account. Real-deck benchmarks are documented in `DEPLOYMENT.md`; observed peak stayed below 640 MB, so the service remains at 2 GB.

Two constraints that are deliberate, not accidental:

- **Only touch AWS resources named `squad4*`/`squad5*`** — everything else in the account belongs to other people. Resources for this project are `squad4e5-auto-ppt`.
- **App Runner max instances = 1.** `project_store.py` writes JSON objects; each write is atomic, but two instances writing the same object could lose an update. Raising `MaxSize` requires ETag-conditional writes first.

The old ECS Fargate + ALB + on-demand-worker architecture was removed (never deployed, complexity the usage volume didn't justify). It's in git history.

## graphify

This project has a knowledge graph at graphify-out/ with god nodes, community structure, and cross-file relationships.

Rules:
- For codebase questions, first run `graphify query "<question>"` when graphify-out/graph.json exists. Use `graphify path "<A>" "<B>"` for relationships and `graphify explain "<concept>"` for focused concepts. These return a scoped subgraph, usually much smaller than GRAPH_REPORT.md or raw grep output.
- If graphify-out/wiki/index.md exists, use it for broad navigation instead of raw source browsing.
- Read graphify-out/GRAPH_REPORT.md only for broad architecture review or when query/path/explain do not surface enough context.
- After modifying code, run `graphify update .` to keep the graph current (AST-only, no API cost).
