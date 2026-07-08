# Graph Report - automatizador-ppt  (2026-07-08)

## Corpus Check
- 66 files · ~56,284 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 994 nodes · 3180 edges · 38 communities (32 shown, 6 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 55 edges (avg confidence: 0.53)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `65d2ee66`
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

## God Nodes (most connected - your core abstractions)
1. `PptTarget` - 96 edges
2. `TransformPlan` - 55 edges
3. `ParsedXlsxTable` - 42 edges
4. `_render_preview()` - 33 edges
5. `preview()` - 24 edges
6. `_analysis_for_job()` - 24 edges
7. `discover_ppt_targets()` - 23 edges
8. `analyze_update_package()` - 23 edges
9. `_preview_context_from_analysis()` - 21 edges
10. `normalize_to_target()` - 21 edges

## Surprising Connections (you probably didn't know these)
- `AiScopeTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_ai_scope.py → ppt_automator/ppt_discovery.py
- `LearnedMappingTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_learned_mapping.py → ppt_automator/ppt_discovery.py
- `MappingTemplateTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_mapping_templates.py → ppt_automator/ppt_discovery.py
- `PptDiscoveryTargetKeyTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_ppt_discovery_target_keys.py → ppt_automator/ppt_discovery.py
- `PptTargetRenamerTests` --uses--> `PptTarget`  [INFERRED]
  tests/test_ppt_target_renamer.py → ppt_automator/ppt_discovery.py

## Import Cycles
- None detected.

## Communities (38 total, 6 thin omitted)

### Community 0 - "Legacy Chart Job Core"
Cohesion: 0.06
Nodes (92): uploaded_source_table(), FormulaMode, _average(), _best_target_label(), _best_text_score(), build_auto_chart_jobs(), build_chart_job(), build_chart_jobs() (+84 more)

### Community 1 - "Embedded Workbook Writer"
Cohesion: 0.11
Nodes (42): _collect_namespace(), EmbeddedWorkbookWriterUnavailable, _excel_col(), _first_table_path(), _join_xlsx_path(), _matrix_cell_is_text(), _matrix_cell_number(), _matrix_cell_text() (+34 more)

### Community 2 - "AI Client & Debug Logging"
Cohesion: 0.15
Nodes (28): ai_debug_log(), _append_jsonl(), current_ai_debug_log_path(), _json_size(), log_ai_request(), log_ai_response(), log_debug_event(), Any (+20 more)

### Community 3 - "Streamlit Legacy App"
Cohesion: 0.13
Nodes (49): build_review_jobs(), file_bytes(), generated_output_for(), item_match_score(), list_match_pct(), missing_items(), norm_text(), Path (+41 more)

### Community 4 - "Slide Datasources & Rendering"
Cohesion: 0.16
Nodes (20): _cell_sort_key(), _cell_type(), _compact_cell(), _display_value(), dump_xlsx_workbook(), dump_xlsx_zip_entries(), _merged_lookup(), _node_text() (+12 more)

### Community 5 - "PPT Target Discovery"
Cohesion: 0.13
Nodes (39): _cache_values(), _chart_orientation(), _chart_target(), _chart_value_format(), _chart_workbook_lookup_enabled(), discover_ppt_targets(), _formula_range_values(), _formula_sheet_name() (+31 more)

### Community 6 - "PPTX Surgery Test Script"
Cohesion: 0.15
Nodes (37): Namespace, _compress_zip_payload(), _excel_col(), _first_table_path(), _join_xlsx_path(), main(), _matrix_strings(), _number_or_original() (+29 more)

### Community 7 - "Slide AI Review Pipeline"
Cohesion: 0.11
Nodes (35): _ai_input_stats(), _ai_payload_profile(), _ai_review_confidence_floor(), _auto_slide_ai_confidence_floor(), _auto_source_match_review_enabled(), _cards_by_slide(), _completed_preview_cache_usable(), _display_row() (+27 more)

### Community 8 - "XLSX Parsing & Metadata"
Cohesion: 0.16
Nodes (35): date, datetime, _date_period_label(), _extract_context_metadata(), _extract_metadata(), _find_header_row(), _find_label_col(), _first_row_group_label() (+27 more)

### Community 9 - "Edit Data Validation"
Cohesion: 0.16
Nodes (23): chart_replacements(), _chart_value_text(), ChartSheetUnresolvedError, _disable_chart_auto_update(), effective_value_format(), _excel_col(), Any, Element (+15 more)

### Community 10 - "Source Matching & Normalization"
Cohesion: 0.09
Nodes (48): _aligned_value(), _assign_slide_plans(), _best_axis_alignment(), _best_match(), _best_text_score(), build_transform_plans(), _calibrated_confidence(), _context_variants() (+40 more)

### Community 11 - "Preview Web Endpoints"
Cohesion: 0.30
Nodes (25): HTMLResponse, Request, apply_job_mapping_template(), approve_slide(), approve_target(), _canonical_target_id(), _clear_render_cache(), _job_dir() (+17 more)

### Community 12 - "Analysis Worker Orchestration"
Cohesion: 0.11
Nodes (44): AiSourceMatchSuggestion, ManualSourceMap, ManualSourcePayload, build_preview(), _display_format(), _display_value(), _format_pt_number(), _matrix_for_preview() (+36 more)

### Community 13 - "Chart XML Writer"
Cohesion: 0.08
Nodes (23): InputFile, analyze_update_package(), _build_slide_aware_plans(), _datasource_names_for_slides(), generate_updated_pptx(), preview_update_package(), Any, InputFile (+15 more)

### Community 14 - "AI Mapping Payloads"
Cohesion: 0.24
Nodes (14): _find_graphic_frame(), _format_value(), _matches_categories(), _norm(), Any, Element, _set_cell_text(), update_table_slide_xml() (+6 more)

### Community 15 - "Upload & Job Status API"
Cohesion: 0.18
Nodes (18): JSONResponse, UploadFile, _init_preview_processing_state(), _inspect_ppt_upload(), job_processing_status(), _large_deck_confirmation_message(), _large_deck_slide_threshold(), _large_scope_requires_confirmation() (+10 more)

### Community 16 - "Job Debug & Cache State"
Cohesion: 0.21
Nodes (13): FileResponse, _ai_waiting_status(), _combine_ai_status(), job_debug_log(), _load_ai_diagnostics(), _load_preview_processing_state(), _mark_processing_analysis_done(), _preview_processing_path() (+5 more)

### Community 17 - "Frontend Preview UI"
Cohesion: 0.15
Nodes (17): activeSquad(), applyPreviewFilters(), deckScopeSize(), formatBytes(), inspectPpt(), parseSlideScope(), previewFilterState(), progressStepsForMessage() (+9 more)

### Community 18 - "render_slide_with_target_labels"
Cohesion: 0.11
Nodes (31): mapping_entry_learning_fields(), Any, ParsedXlsxTable, Camada 4 - aprendizado do mapeamento (match "perfeito na 2a vez").  A cada downl, Campos de aprendizado a persistir na entrada do template de mapeamento., Traduz as entradas do template salvo em matches para os targets/sources     ATUA, _reason_for(), resolve_learned_matches() (+23 more)

### Community 19 - "Job Analysis Caching"
Cohesion: 0.18
Nodes (21): AnalysisResult, Response, _analysis_for_job(), _analyze_files_signature(), _apply_mapping_template_to_analysis(), _cached_analyze_files(), download(), _load_ai_source_matches() (+13 more)

### Community 20 - "Preview Context & Flags"
Cohesion: 0.14
Nodes (26): Path, _apply_slide_ai_outputs_enabled(), _builder_payload_from_understanding(), _call_with_job_debug(), _clear_ai_cache(), _clear_slide_ai_state(), _datasource_entries_for_slide(), _datasource_entry_hashes() (+18 more)

### Community 21 - "PptTarget"
Cohesion: 0.17
Nodes (30): _ai_match_target_payload(), AiSourceMatchSuggestion, build_ai_mapping_payload(), _env_float(), _env_int(), _non_negative_int(), plans_to_ai_payload(), Any (+22 more)

### Community 22 - "ai.py"
Cohesion: 0.16
Nodes (23): BaseException, Exception, ai_configured(), AiMatchSuggestion, AiUnavailableError, build_openai_client(), _env_float(), _env_int() (+15 more)

### Community 23 - "_run_automatic_slide_ai_review"
Cohesion: 0.18
Nodes (17): _ai_diagnostic_batch_limit(), _ai_diagnostics_for_job(), _ai_source_match_batch_limit(), _ai_source_match_max_calls(), _ai_source_matches_for_job(), _append_ai_log(), _auto_slide_ai_enabled(), _diagnostic_payload_summary() (+9 more)

### Community 24 - "typed_matrix.py"
Cohesion: 0.18
Nodes (22): assert_valid_typed_edit_data(), _chart_series_capacity_errors(), EditDataValidationError, Any, ValueError, O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp, validate_typed_edit_data(), cell_to_excel_value() (+14 more)

### Community 25 - "slide_renderer.py"
Cohesion: 0.29
Nodes (13): Image, _cached_deck_pdf(), _export_slide_with_libreoffice(), _export_slide_with_powerpoint(), _fallback_slide_canvas(), _font(), InputFile, Path (+5 more)

### Community 26 - "Fargate Deploy Script"
Cohesion: 0.43
Nodes (5): Assert-AwsOk(), Ensure-Role(), Get-RoleArn(), Put-InlinePolicy(), Save-JsonFile()

### Community 31 - "ai_transform.py"
Cohesion: 0.29
Nodes (11): AiTransformDiagnostic, _edit_data_matrix(), _mapping_to_dict(), _plan_axes(), _plan_payload(), Any, Path, _recommended_edit_data() (+3 more)

### Community 32 - "xlsx_source_manifest"
Cohesion: 0.36
Nodes (9): Any, _recipe_hint(), _signature(), _structural_profile(), _structure_name(), _take(), _take_rows(), xlsx_source_manifest() (+1 more)

### Community 36 - "_error_response"
Cohesion: 0.24
Nodes (12): _checkpoint_slide_label(), _checkpoint_status_label(), _checkpoint_summary(), _error_response(), index(), _mapping_templates_by_squad(), _pretty_datetime(), _project_cards_by_squad() (+4 more)

### Community 37 - "test_ai_scope.py"
Cohesion: 0.38
Nodes (5): AiScopeTests, _analysis(), _plan(), _source(), _target()

## Knowledge Gaps
- **1 isolated node(s):** `CodeBuild Buildspec (ECR image build)`
  These have ≤1 connection - possible missing edges or undocumented components.
- **6 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `PptTarget` connect `PptTarget` to `Legacy Chart Job Core`, `PPT Target Discovery`, `PPTX Surgery Test Script`, `test_ai_scope.py`, `Edit Data Validation`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `render_slide_with_target_labels`, `slide_renderer.py`?**
  _High betweenness centrality (0.120) - this node is a cross-community bridge._
- **Why does `TransformPlan` connect `Analysis Worker Orchestration` to `Legacy Chart Job Core`, `test_ai_scope.py`, `PPTX Surgery Test Script`, `Edit Data Validation`, `Source Matching & Normalization`, `Chart XML Writer`, `AI Mapping Payloads`, `PptTarget`, `ai_transform.py`?**
  _High betweenness centrality (0.052) - this node is a cross-community bridge._
- **Why does `analyze_update_package()` connect `Chart XML Writer` to `Legacy Chart Job Core`, `PPT Target Discovery`, `Analysis Worker Orchestration`, `PptTarget`, `ai.py`, `ai_transform.py`?**
  _High betweenness centrality (0.030) - this node is a cross-community bridge._
- **Are the 16 inferred relationships involving `PptTarget` (e.g. with `AiSourceMatchSuggestion` and `ChartSheetUnresolvedError`) actually correct?**
  _`PptTarget` has 16 INFERRED edges - model-reasoned connections that need verification._
- **Are the 4 inferred relationships involving `Path` (e.g. with `read_bytes()` and `_filename_context_score()`) actually correct?**
  _`Path` has 4 INFERRED edges - model-reasoned connections that need verification._
- **Are the 8 inferred relationships involving `TransformPlan` (e.g. with `AiSourceMatchSuggestion` and `AiTransformDiagnostic`) actually correct?**
  _`TransformPlan` has 8 INFERRED edges - model-reasoned connections that need verification._
- **Are the 6 inferred relationships involving `ParsedXlsxTable` (e.g. with `AiSourceMatchSuggestion` and `AiMapperTests`) actually correct?**
  _`ParsedXlsxTable` has 6 INFERRED edges - model-reasoned connections that need verification._