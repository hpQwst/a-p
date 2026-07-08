# Graph Report - .  (2026-07-08)

## Corpus Check
- cluster-only mode — file stats not available

## Summary
- 930 nodes · 3191 edges · 30 communities (27 shown, 3 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 62 edges (avg confidence: 0.55)
- Token cost: 27,332 input · 474 output

## Graph Freshness
- Built from commit: `b0881dab`
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
- Job Analysis Caching
- Preview Context & Flags
- Table Writer & Tests
- AI Call Batching Limits
- AI Scope Tests
- Preview Model Building
- Source Manifest & Plans
- Fargate Deploy Script
- CodeBuild Docker Image
- Web App Package Init
- Worker Package Boundary

## God Nodes (most connected - your core abstractions)
1. `PptTarget` - 82 edges
2. `ParsedXlsxTable` - 64 edges
3. `TransformPlan` - 56 edges
4. `_render_preview()` - 36 edges
5. `AnalysisResult` - 35 edges
6. `_analysis_for_job()` - 27 edges
7. `analyze_update_package()` - 26 edges
8. `_run_slide_ai_review()` - 26 edges
9. `preview()` - 25 edges
10. `log_debug_event()` - 24 edges

## Surprising Connections (you probably didn't know these)
- `_ai_source_matches_for_job()` --indirect_call--> `suggest_source_matches_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_mapper.py
- `_run_slide_ai_review()` --indirect_call--> `build_slide_matrices_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_slide_matrix_builder.py
- `_run_slide_ai_review()` --indirect_call--> `suggest_slide_understanding()`  [INFERRED]
  web/main.py → ppt_automator/ai_slide_understanding.py
- `_ai_diagnostics_for_job()` --indirect_call--> `suggest_transform_diagnostics()`  [INFERRED]
  web/main.py → ppt_automator/ai_transform.py
- `download()` --indirect_call--> `EmbeddedWorkbookWriterUnavailable`  [INFERRED]
  web/main.py → ppt_automator/embedded_workbook_writer.py

## Import Cycles
- None detected.

## Hyperedges (group relationships)
- **Jinja2 UI templates rendered by the FastAPI web layer** — web_templates_index, web_templates_preview, web_templates_error, web_main, web_static_style [EXTRACTED 0.90]
- **Upload → background preview job → per-slide review → download flow** — web_templates_index, web_main, web_templates_preview, web_templates_preview_workflow [INFERRED 0.80]

## Communities (30 total, 3 thin omitted)

### Community 0 - "Legacy Chart Job Core"
Cohesion: 0.06
Nodes (89): uploaded_source_table(), FormulaMode, _average(), _best_target_label(), _best_text_score(), build_auto_chart_jobs(), build_chart_job(), build_chart_jobs() (+81 more)

### Community 1 - "Embedded Workbook Writer"
Cohesion: 0.06
Nodes (55): _collect_namespace(), EmbeddedWorkbookWriterUnavailable, _excel_col(), _first_table_path(), _join_xlsx_path(), _matrix_cell_is_text(), _matrix_cell_number(), _matrix_cell_text() (+47 more)

### Community 2 - "AI Client & Debug Logging"
Cohesion: 0.07
Nodes (59): BaseException, Exception, AiMatchSuggestion, AiUnavailableError, build_openai_client(), ai_debug_log(), _append_jsonl(), current_ai_debug_log_path() (+51 more)

### Community 3 - "Streamlit Legacy App"
Cohesion: 0.12
Nodes (51): build_review_jobs(), file_bytes(), generated_output_for(), item_match_score(), list_match_pct(), missing_items(), norm_text(), Path (+43 more)

### Community 4 - "Slide Datasources & Rendering"
Cohesion: 0.08
Nodes (43): Image, InputFile, read_bytes(), collect_datasource_entries(), parse_slide_number_from_path(), _parse_slide_suffix(), _parse_slide_token(), InputFile (+35 more)

### Community 5 - "PPT Target Discovery"
Cohesion: 0.10
Nodes (43): _cache_values(), _chart_orientation(), _chart_target(), _chart_workbook_lookup_enabled(), discover_ppt_targets(), _formula_range_values(), _formula_sheet_name(), _formula_single_value() (+35 more)

### Community 6 - "PPTX Surgery Test Script"
Cohesion: 0.15
Nodes (37): Namespace, _compress_zip_payload(), _excel_col(), _first_table_path(), _join_xlsx_path(), main(), _matrix_strings(), _number_or_original() (+29 more)

### Community 7 - "Slide AI Review Pipeline"
Cohesion: 0.09
Nodes (39): SlideUnderstandingInput, entries_for_slide(), _ai_input_stats(), _ai_payload_profile(), _ai_review_confidence_floor(), _auto_slide_ai_confidence_floor(), _builder_payload_from_understanding(), _cards_by_slide() (+31 more)

### Community 8 - "XLSX Parsing & Metadata"
Cohesion: 0.16
Nodes (35): date, datetime, _date_period_label(), _extract_context_metadata(), _extract_metadata(), _find_header_row(), _find_label_col(), _first_row_group_label() (+27 more)

### Community 9 - "Edit Data Validation"
Cohesion: 0.18
Nodes (22): assert_valid_typed_edit_data(), _chart_series_capacity_errors(), EditDataValidationError, Any, ValueError, O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp, validate_typed_edit_data(), cell_to_excel_value() (+14 more)

### Community 10 - "Source Matching & Normalization"
Cohesion: 0.18
Nodes (28): _aligned_value(), _best_axis_alignment(), _best_match(), _best_source_for_target(), _best_text_score(), _context_variants(), _coverage_score(), _domain_text_score() (+20 more)

### Community 11 - "Preview Web Endpoints"
Cohesion: 0.26
Nodes (28): HTMLResponse, ai_configured(), Request, _ai_waiting_status(), apply_job_mapping_template(), approve_slide(), approve_target(), _canonical_target_id() (+20 more)

### Community 12 - "Analysis Worker Orchestration"
Cohesion: 0.18
Nodes (26): ManualSourceMap, ManualSourcePayload, SavedSourceMatchMap, _analysis_warnings(), analyze_files(), apply_ai_recommendations(), apply_ai_recommendations_to_analysis(), _apply_manual_sources() (+18 more)

### Community 13 - "Chart XML Writer"
Cohesion: 0.19
Nodes (22): chart_replacements(), _chart_value_text(), ChartSheetUnresolvedError, _disable_chart_auto_update(), _excel_col(), _is_percent_text(), _numeric_value(), _plan_number_format() (+14 more)

### Community 14 - "AI Mapping Payloads"
Cohesion: 0.19
Nodes (27): _ai_match_target_payload(), build_ai_mapping_payload(), _env_float(), _env_int(), _non_negative_int(), plans_to_ai_payload(), Any, _recipe_schema() (+19 more)

### Community 15 - "Upload & Job Status API"
Cohesion: 0.12
Nodes (24): JSONResponse, UploadFile, _checkpoint_slide_label(), _checkpoint_summary(), _error_response(), index(), _inspect_ppt_upload(), job_processing_status() (+16 more)

### Community 16 - "Job Debug & Cache State"
Cohesion: 0.18
Nodes (24): reset_ai_debug_log_path(), set_ai_debug_log_path(), _call_with_job_debug(), _clear_ai_cache(), _clear_slide_ai_state(), _combine_ai_status(), _debug_log_path(), _init_preview_processing_state() (+16 more)

### Community 17 - "Frontend Preview UI"
Cohesion: 0.13
Nodes (21): Squad Mapping Template (Modelo de Mapeamento), activeSquad(), applyPreviewFilters(), deckScopeSize(), inspectPpt(), parseSlideScope(), previewFilterState(), progressStepsForMessage() (+13 more)

### Community 19 - "Job Analysis Caching"
Cohesion: 0.18
Nodes (19): FileResponse, target_aliases(), Response, _analysis_for_job(), _analyze_files_signature(), _apply_mapping_template_to_analysis(), _cached_analyze_files(), download() (+11 more)

### Community 20 - "Preview Context & Flags"
Cohesion: 0.15
Nodes (18): _apply_slide_ai_outputs_enabled(), _auto_source_match_review_enabled(), _env_bool(), _load_slide_ai_state(), _preview_context_from_analysis(), _preview_full_render_slide_limit(), _preview_selected_slide(), _preview_slide_from_request() (+10 more)

### Community 21 - "Table Writer & Tests"
Cohesion: 0.24
Nodes (14): _find_graphic_frame(), _format_value(), _matches_categories(), _norm(), Any, Element, _set_cell_text(), update_table_slide_xml() (+6 more)

### Community 22 - "AI Call Batching Limits"
Cohesion: 0.18
Nodes (17): _ai_diagnostic_batch_limit(), _ai_diagnostics_for_job(), _ai_source_match_batch_limit(), _ai_source_match_max_calls(), _ai_source_matches_for_job(), _append_ai_log(), _auto_slide_ai_enabled(), _diagnostic_payload_summary() (+9 more)

### Community 23 - "AI Scope Tests"
Cohesion: 0.30
Nodes (10): AiSourceMatchSuggestion, AiTransformDiagnostic, AiScopeTests, _analysis(), _plan(), _source(), _target(), AnalysisResult (+2 more)

### Community 24 - "Preview Model Building"
Cohesion: 0.24
Nodes (12): build_preview(), _display_value(), _format_pt_percent(), _matrix_for_preview(), _plan_is_percentage(), PreviewTarget, Any, _to_number() (+4 more)

### Community 25 - "Source Manifest & Plans"
Cohesion: 0.22
Nodes (15): _build_slide_aware_plans(), Any, _recipe_hint(), _signature(), _structural_profile(), _structure_name(), _take(), _take_rows() (+7 more)

### Community 26 - "Fargate Deploy Script"
Cohesion: 0.43
Nodes (5): Assert-AwsOk(), Ensure-Role(), Get-RoleArn(), Put-InlinePolicy(), Save-JsonFile()

## Knowledge Gaps
- **2 isolated node(s):** `CodeBuild Buildspec (ECR image build)`, `Preview Background Job Workflow`
  These have ≤1 connection - possible missing edges or undocumented components.
- **3 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `PptTarget` connect `AI Mapping Payloads` to `Legacy Chart Job Core`, `Embedded Workbook Writer`, `Slide Datasources & Rendering`, `PPT Target Discovery`, `PPTX Surgery Test Script`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `Chart XML Writer`, `Job Analysis Caching`, `Table Writer & Tests`, `AI Scope Tests`, `Preview Model Building`, `Source Manifest & Plans`?**
  _High betweenness centrality (0.077) - this node is a cross-community bridge._
- **Why does `TransformPlan` connect `Preview Model Building` to `Legacy Chart Job Core`, `Embedded Workbook Writer`, `AI Client & Debug Logging`, `PPTX Surgery Test Script`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `Table Writer & Tests`, `AI Scope Tests`, `Source Manifest & Plans`?**
  _High betweenness centrality (0.046) - this node is a cross-community bridge._
- **Why does `generate_updated_pptx()` connect `Embedded Workbook Writer` to `Legacy Chart Job Core`, `Slide Datasources & Rendering`, `PPT Target Discovery`, `Slide AI Review Pipeline`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `Job Analysis Caching`, `Table Writer & Tests`, `Preview Model Building`?**
  _High betweenness centrality (0.043) - this node is a cross-community bridge._
- **Are the 13 inferred relationships involving `PptTarget` (e.g. with `AiSourceMatchSuggestion` and `ChartSheetUnresolvedError`) actually correct?**
  _`PptTarget` has 13 INFERRED edges - model-reasoned connections that need verification._
- **Are the 9 inferred relationships involving `ParsedXlsxTable` (e.g. with `AiSourceMatchSuggestion` and `SourceMatchCandidate`) actually correct?**
  _`ParsedXlsxTable` has 9 INFERRED edges - model-reasoned connections that need verification._
- **Are the 9 inferred relationships involving `TransformPlan` (e.g. with `AiSourceMatchSuggestion` and `AiTransformDiagnostic`) actually correct?**
  _`TransformPlan` has 9 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Utilities for mapping datasource workbooks into mapped PowerPoint charts.`, `Permite modelo mais barato por operacao (ex.: OPENAI_MODEL_SOURCE_MATCH para o`, `Raised when a workbook contains formulas but no calculator can resolve them.` to the rest of the system?**
  _16 weakly-connected nodes found - possible documentation gaps or missing edges._