# Graph Report - automatizador-ppt  (2026-07-08)

## Corpus Check
- 66 files · ~57,261 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 1017 nodes · 3195 edges · 40 communities (30 shown, 10 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 60 edges (avg confidence: 0.58)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `f515ad1b`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- Legacy Chart Job Core
- Embedded Workbook Writer
- AI Client & Debug Logging
- Streamlit Legacy App
- Slide Datasources & Rendering
- PPT Target Discovery
- PPTX Surgery Test Script
- Slide AI Review Pipeline
- XLSX Parsing & Metadata
- Edit Data Validation
- Source Matching & Normalization
- Preview Web Endpoints
- Analysis Worker Orchestration
- Chart XML Writer
- AI Mapping Payloads
- Upload & Job Status API
- Job Debug & Cache State
- Frontend Preview UI
- render_slide_with_target_labels
- Job Analysis Caching
- Preview Context & Flags
- PptTarget
- ai.py
- _run_automatic_slide_ai_review
- typed_matrix.py
- slide_renderer.py
- Fargate Deploy Script
- CodeBuild Docker Image
- Web App Package Init
- Worker Package Boundary
- ai_transform.py
- xlsx_source_manifest
- RuntimeError
- InputFile
- Path
- _error_response
- test_ai_scope.py
- Element
- ZipFile

## God Nodes (most connected - your core abstractions)
1. `PptTarget` - 96 edges
2. `TransformPlan` - 47 edges
3. `_render_preview()` - 33 edges
4. `ParsedXlsxTable` - 28 edges
5. `preview()` - 24 edges
6. `_analysis_for_job()` - 24 edges
7. `discover_ppt_targets()` - 23 edges
8. `_run_slide_ai_review()` - 23 edges
9. `analyze_update_package()` - 23 edges
10. `_preview_context_from_analysis()` - 21 edges

## Surprising Connections (you probably didn't know these)
- `AiScopeTests` --uses--> `AiSourceMatchSuggestion`  [INFERRED]
  tests/test_ai_scope.py → ppt_automator/ai_mapper.py
- `_ai_source_matches_for_job()` --indirect_call--> `suggest_source_matches_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_mapper.py
- `AiMapperTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_ai_mapper.py → ppt_automator/ppt_discovery.py
- `AiScopeTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_ai_scope.py → ppt_automator/ppt_discovery.py
- `LearnedMappingTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_learned_mapping.py → ppt_automator/ppt_discovery.py

## Import Cycles
- None detected.

## Communities (40 total, 10 thin omitted)

### Community 0 - "Legacy Chart Job Core"
Cohesion: 0.06
Nodes (88): uploaded_source_table(), FormulaMode, _average(), _best_target_label(), _best_text_score(), build_auto_chart_jobs(), build_chart_job(), build_chart_jobs() (+80 more)

### Community 1 - "Embedded Workbook Writer"
Cohesion: 0.08
Nodes (52): assert_valid_typed_edit_data(), _chart_series_capacity_errors(), EditDataValidationError, Any, ValueError, O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp, validate_typed_edit_data(), _collect_namespace() (+44 more)

### Community 2 - "AI Client & Debug Logging"
Cohesion: 0.07
Nodes (64): BaseException, Exception, ai_configured(), AiMatchSuggestion, AiUnavailableError, build_openai_client(), ai_debug_log(), _append_jsonl() (+56 more)

### Community 3 - "Streamlit Legacy App"
Cohesion: 0.11
Nodes (52): build_review_jobs(), file_bytes(), generated_output_for(), item_match_score(), list_match_pct(), missing_items(), norm_text(), Path (+44 more)

### Community 4 - "Slide Datasources & Rendering"
Cohesion: 0.08
Nodes (44): Image, InputFile, read_bytes(), collect_datasource_entries(), entries_for_slide(), parse_slide_number_from_path(), _parse_slide_suffix(), _parse_slide_token() (+36 more)

### Community 5 - "PPT Target Discovery"
Cohesion: 0.17
Nodes (36): Element, _cache_values(), _chart_orientation(), _chart_target(), _chart_value_format(), _chart_workbook_lookup_enabled(), discover_ppt_targets(), _formula_range_values() (+28 more)

### Community 6 - "PPTX Surgery Test Script"
Cohesion: 0.15
Nodes (37): Namespace, _compress_zip_payload(), _excel_col(), _first_table_path(), _join_xlsx_path(), main(), _matrix_strings(), _number_or_original() (+29 more)

### Community 7 - "Slide AI Review Pipeline"
Cohesion: 0.12
Nodes (31): _apply_slide_ai_outputs_enabled(), _auto_source_match_review_enabled(), _cards_by_slide(), _completed_preview_cache_usable(), _diagnostic_payload_summary(), _display_row(), _load_slide_ai_state(), _match_payload_summary() (+23 more)

### Community 8 - "XLSX Parsing & Metadata"
Cohesion: 0.11
Nodes (47): date, datetime, Any, _recipe_hint(), _signature(), _structural_profile(), _structure_name(), _take() (+39 more)

### Community 9 - "Edit Data Validation"
Cohesion: 0.17
Nodes (22): _chart_value_text(), ChartSheetUnresolvedError, _disable_chart_auto_update(), effective_value_format(), _excel_col(), Any, Element, ZipFile (+14 more)

### Community 10 - "Source Matching & Normalization"
Cohesion: 0.16
Nodes (34): _aligned_value(), _assign_slide_plans(), _best_axis_alignment(), _best_match(), _best_text_score(), _calibrated_confidence(), _context_variants(), _coverage_score() (+26 more)

### Community 11 - "Preview Web Endpoints"
Cohesion: 0.30
Nodes (24): HTMLResponse, Request, apply_job_mapping_template(), approve_slide(), approve_target(), _canonical_target_id(), _clear_render_cache(), _job_dir() (+16 more)

### Community 12 - "Analysis Worker Orchestration"
Cohesion: 0.12
Nodes (41): AiSourceMatchSuggestion, ManualSourceMap, ManualSourcePayload, preview_update_package(), build_preview(), _display_format(), _display_value(), _format_pt_number() (+33 more)

### Community 13 - "Chart XML Writer"
Cohesion: 0.14
Nodes (15): Nome real da primeira aba do workbook embutido, para resolver target.sheet_name, resolve_default_sheet_name(), analyze_update_package(), _build_slide_aware_plans(), _datasource_names_for_slides(), generate_updated_pptx(), Any, InputFile (+7 more)

### Community 14 - "AI Mapping Payloads"
Cohesion: 0.44
Nodes (8): _find_graphic_frame(), _format_value(), _matches_categories(), _norm(), Any, Element, _set_cell_text(), _write_row()

### Community 15 - "Upload & Job Status API"
Cohesion: 0.12
Nodes (26): JSONResponse, UploadFile, _checkpoint_slide_label(), _checkpoint_status_label(), _checkpoint_summary(), _error_response(), index(), _inspect_ppt_upload() (+18 more)

### Community 16 - "Job Debug & Cache State"
Cohesion: 0.12
Nodes (31): Path, _ai_waiting_status(), _call_with_job_debug(), _clear_ai_cache(), _clear_slide_ai_state(), _combine_ai_status(), _datasource_entries_for_ai_review(), _datasource_entries_for_slide() (+23 more)

### Community 17 - "Frontend Preview UI"
Cohesion: 0.15
Nodes (17): activeSquad(), applyPreviewFilters(), deckScopeSize(), formatBytes(), inspectPpt(), parseSlideScope(), previewFilterState(), progressStepsForMessage() (+9 more)

### Community 18 - "render_slide_with_target_labels"
Cohesion: 0.11
Nodes (31): mapping_entry_learning_fields(), Any, ParsedXlsxTable, Camada 4 - aprendizado do mapeamento (match "perfeito na 2a vez").  A cada downl, Campos de aprendizado a persistir na entrada do template de mapeamento., Traduz as entradas do template salvo em matches para os targets/sources     ATUA, _reason_for(), resolve_learned_matches() (+23 more)

### Community 19 - "Job Analysis Caching"
Cohesion: 0.14
Nodes (29): FileResponse, Response, _ai_review_confidence_floor(), _ai_source_match_plausibility_floor(), _ai_source_matches_for_job(), _analysis_for_job(), _analyze_files_signature(), _apply_mapping_template_to_analysis() (+21 more)

### Community 20 - "Preview Context & Flags"
Cohesion: 0.11
Nodes (24): _ai_input_stats(), _ai_payload_profile(), _builder_payload_from_understanding(), _datasource_entry_hashes(), _dump_datasource_entries_for_ai_review(), _env_bool(), _json_size_bytes(), _NoRender (+16 more)

### Community 21 - "PptTarget"
Cohesion: 0.16
Nodes (30): _ai_match_target_payload(), AiSourceMatchSuggestion, build_ai_mapping_payload(), _env_float(), _env_int(), _non_negative_int(), plans_to_ai_payload(), Any (+22 more)

### Community 22 - "ai.py"
Cohesion: 0.19
Nodes (13): PptTarget, _match_target(), Element, rename_targets_in_slide_xml(), _normalize_table(), normalize_to_target(), PptTargetRenamerTests, _target() (+5 more)

### Community 23 - "_run_automatic_slide_ai_review"
Cohesion: 0.23
Nodes (12): _ai_diagnostic_batch_limit(), _ai_diagnostics_for_job(), _ai_source_match_batch_limit(), _ai_source_match_max_calls(), _append_ai_log(), _auto_slide_ai_enabled(), _env_int(), _mark_slide_ai_skipped() (+4 more)

### Community 24 - "typed_matrix.py"
Cohesion: 0.27
Nodes (8): build_transform_plans(), _hungarian(), Atribuicao 1:1 de custo minimo (Kuhn-Munkres, O(n^3)). Retorna, para cada     li, Match deterministico por SLIDE resolvido como atribuicao global 1:1.      Em vez, ParsedXlsxTable, SlideAssignmentTests, _source(), _target()

### Community 26 - "Fargate Deploy Script"
Cohesion: 0.43
Nodes (5): Assert-AwsOk(), Ensure-Role(), Get-RoleArn(), Put-InlinePolicy(), Save-JsonFile()

### Community 31 - "ai_transform.py"
Cohesion: 0.45
Nodes (10): _central_header(), _compress_payload(), _datetime(), _filename_bytes(), _local_header(), Any, ValueError, _raw_local_record_ends() (+2 more)

### Community 32 - "xlsx_source_manifest"
Cohesion: 0.48
Nodes (5): _key_value_workbook(), _paragraph_child_names(), _slide_with_table(), _table_values(), TableKeyValueUpdateTests

### Community 37 - "test_ai_scope.py"
Cohesion: 0.28
Nodes (8): AiScopeTests, _analysis(), _plan(), AnalysisResult, ParsedXlsxTable, TransformPlan, _source(), _target()

## Knowledge Gaps
- **1 isolated node(s):** `CodeBuild Buildspec (ECR image build)`
  These have ≤1 connection - possible missing edges or undocumented components.
- **10 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `PptTarget` connect `ai.py` to `Legacy Chart Job Core`, `xlsx_source_manifest`, `Streamlit Legacy App`, `Slide Datasources & Rendering`, `PPT Target Discovery`, `PPTX Surgery Test Script`, `test_ai_scope.py`, `Edit Data Validation`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `render_slide_with_target_labels`, `PptTarget`, `typed_matrix.py`?**
  _High betweenness centrality (0.129) - this node is a cross-community bridge._
- **Why does `TransformPlan` connect `Chart XML Writer` to `Legacy Chart Job Core`, `AI Client & Debug Logging`, `Streamlit Legacy App`, `PPTX Surgery Test Script`, `Edit Data Validation`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `AI Mapping Payloads`, `ai.py`, `typed_matrix.py`?**
  _High betweenness centrality (0.033) - this node is a cross-community bridge._
- **Why does `ParsedXlsxTable` connect `XLSX Parsing & Metadata` to `Legacy Chart Job Core`, `Streamlit Legacy App`, `Chart XML Writer`, `PptTarget`, `ai.py`?**
  _High betweenness centrality (0.017) - this node is a cross-community bridge._
- **Are the 16 inferred relationships involving `PptTarget` (e.g. with `AiSourceMatchSuggestion` and `ChartSheetUnresolvedError`) actually correct?**
  _`PptTarget` has 16 INFERRED edges - model-reasoned connections that need verification._
- **Are the 12 inferred relationships involving `Path` (e.g. with `read_bytes()` and `_filename_context_score()`) actually correct?**
  _`Path` has 12 INFERRED edges - model-reasoned connections that need verification._
- **Are the 6 inferred relationships involving `TransformPlan` (e.g. with `AiTransformDiagnostic` and `ChartSheetUnresolvedError`) actually correct?**
  _`TransformPlan` has 6 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Formato numerico que o template ja define para os valores (o 'contrato').      P`, `Titulo do objeto: a caixa de texto imediatamente ACIMA (ou colada ao topo)     d`, `Render vazio para revisoes leves (sem imagem): pula a renderizacao cara do     s` to the rest of the system?**
  _33 weakly-connected nodes found - possible documentation gaps or missing edges._