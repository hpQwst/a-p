# Graph Report - automatizador-ppt  (2026-07-08)

## Corpus Check
- 66 files · ~56,039 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 989 nodes · 3208 edges · 30 communities (24 shown, 6 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 56 edges (avg confidence: 0.53)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `1bab42c1`
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
- typed_matrix.py
- Fargate Deploy Script
- CodeBuild Docker Image
- Web App Package Init
- Worker Package Boundary
- RuntimeError
- InputFile
- Path

## God Nodes (most connected - your core abstractions)
1. `PptTarget` - 96 edges
2. `TransformPlan` - 55 edges
3. `ParsedXlsxTable` - 42 edges
4. `_render_preview()` - 35 edges
5. `AnalysisResult` - 33 edges
6. `_analysis_for_job()` - 27 edges
7. `preview()` - 25 edges
8. `discover_ppt_targets()` - 23 edges
9. `analyze_update_package()` - 23 edges
10. `normalize_to_target()` - 21 edges

## Surprising Connections (you probably didn't know these)
- `download()` --indirect_call--> `ChartSheetUnresolvedError`  [INFERRED]
  web/main.py → ppt_automator/ppt_chart_writer.py
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

## Communities (30 total, 6 thin omitted)

### Community 0 - "Legacy Chart Job Core"
Cohesion: 0.06
Nodes (90): uploaded_source_table(), FormulaMode, _average(), _best_target_label(), _best_text_score(), build_auto_chart_jobs(), build_chart_job(), build_chart_jobs() (+82 more)

### Community 1 - "Embedded Workbook Writer"
Cohesion: 0.11
Nodes (42): _collect_namespace(), EmbeddedWorkbookWriterUnavailable, _excel_col(), _first_table_path(), _join_xlsx_path(), _matrix_cell_is_text(), _matrix_cell_number(), _matrix_cell_text() (+34 more)

### Community 2 - "AI Client & Debug Logging"
Cohesion: 0.07
Nodes (61): BaseException, Exception, ai_configured(), AiMatchSuggestion, AiUnavailableError, build_openai_client(), ai_debug_log(), _append_jsonl() (+53 more)

### Community 3 - "Streamlit Legacy App"
Cohesion: 0.11
Nodes (52): build_review_jobs(), file_bytes(), generated_output_for(), item_match_score(), list_match_pct(), missing_items(), norm_text(), Path (+44 more)

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
Cohesion: 0.08
Nodes (49): visual_label(), _ai_diagnostic_batch_limit(), _ai_input_stats(), _ai_payload_profile(), _ai_review_confidence_floor(), _ai_source_match_batch_limit(), _ai_source_match_max_calls(), _apply_slide_ai_outputs_enabled() (+41 more)

### Community 8 - "XLSX Parsing & Metadata"
Cohesion: 0.15
Nodes (37): date, datetime, _date_period_label(), _extract_context_metadata(), _extract_metadata(), _find_header_row(), _find_label_col(), _first_row_group_label() (+29 more)

### Community 9 - "Edit Data Validation"
Cohesion: 0.16
Nodes (23): chart_replacements(), _chart_value_text(), ChartSheetUnresolvedError, _disable_chart_auto_update(), effective_value_format(), _excel_col(), Any, Element (+15 more)

### Community 10 - "Source Matching & Normalization"
Cohesion: 0.05
Nodes (94): _ai_match_target_payload(), AiSourceMatchSuggestion, build_ai_mapping_payload(), _env_float(), _env_int(), _non_negative_int(), plans_to_ai_payload(), Any (+86 more)

### Community 11 - "Preview Web Endpoints"
Cohesion: 0.30
Nodes (24): HTMLResponse, Request, apply_job_mapping_template(), approve_slide(), approve_target(), _canonical_target_id(), _clear_render_cache(), _job_dir() (+16 more)

### Community 12 - "Analysis Worker Orchestration"
Cohesion: 0.13
Nodes (38): AiSourceMatchSuggestion, ManualSourceMap, ManualSourcePayload, build_preview(), _display_format(), _display_value(), _format_pt_number(), _matrix_for_preview() (+30 more)

### Community 13 - "Chart XML Writer"
Cohesion: 0.07
Nodes (35): Image, InputFile, analyze_update_package(), _build_slide_aware_plans(), _datasource_names_for_slides(), generate_updated_pptx(), preview_update_package(), Any (+27 more)

### Community 14 - "AI Mapping Payloads"
Cohesion: 0.24
Nodes (14): _find_graphic_frame(), _format_value(), _matches_categories(), _norm(), Any, Element, _set_cell_text(), update_table_slide_xml() (+6 more)

### Community 15 - "Upload & Job Status API"
Cohesion: 0.12
Nodes (25): JSONResponse, UploadFile, _checkpoint_slide_label(), _checkpoint_summary(), _error_response(), index(), _inspect_ppt_upload(), job_processing_status() (+17 more)

### Community 16 - "Job Debug & Cache State"
Cohesion: 0.16
Nodes (25): Path, _ai_waiting_status(), _analyze_files_signature(), _call_with_job_debug(), _clear_ai_cache(), _clear_slide_ai_state(), _combine_ai_status(), _debug_log_path() (+17 more)

### Community 17 - "Frontend Preview UI"
Cohesion: 0.15
Nodes (17): activeSquad(), applyPreviewFilters(), deckScopeSize(), formatBytes(), inspectPpt(), parseSlideScope(), previewFilterState(), progressStepsForMessage() (+9 more)

### Community 18 - "render_slide_with_target_labels"
Cohesion: 0.11
Nodes (30): mapping_entry_learning_fields(), Any, ParsedXlsxTable, Camada 4 - aprendizado do mapeamento (match "perfeito na 2a vez").  A cada downl, Campos de aprendizado a persistir na entrada do template de mapeamento., Traduz as entradas do template salvo em matches para os targets/sources     ATUA, _reason_for(), resolve_learned_matches() (+22 more)

### Community 19 - "Job Analysis Caching"
Cohesion: 0.20
Nodes (23): FileResponse, target_aliases(), Response, _ai_source_matches_for_job(), _analysis_for_job(), _apply_mapping_template_to_analysis(), _cached_analyze_files(), download() (+15 more)

### Community 20 - "Preview Context & Flags"
Cohesion: 0.13
Nodes (21): _ai_diagnostics_for_job(), _append_ai_log(), _auto_slide_ai_enabled(), _builder_payload_from_understanding(), _datasource_entries_for_slide(), _datasource_entry_hashes(), _load_slide_ai_state(), _mark_slide_ai_skipped() (+13 more)

### Community 24 - "typed_matrix.py"
Cohesion: 0.18
Nodes (22): assert_valid_typed_edit_data(), _chart_series_capacity_errors(), EditDataValidationError, Any, ValueError, O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp, validate_typed_edit_data(), cell_to_excel_value() (+14 more)

### Community 26 - "Fargate Deploy Script"
Cohesion: 0.43
Nodes (5): Assert-AwsOk(), Ensure-Role(), Get-RoleArn(), Put-InlinePolicy(), Save-JsonFile()

## Knowledge Gaps
- **1 isolated node(s):** `CodeBuild Buildspec (ECR image build)`
  These have ≤1 connection - possible missing edges or undocumented components.
- **6 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `PptTarget` connect `Source Matching & Normalization` to `Legacy Chart Job Core`, `Streamlit Legacy App`, `PPT Target Discovery`, `PPTX Surgery Test Script`, `Edit Data Validation`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `render_slide_with_target_labels`, `Job Analysis Caching`?**
  _High betweenness centrality (0.116) - this node is a cross-community bridge._
- **Why does `TransformPlan` connect `Source Matching & Normalization` to `Legacy Chart Job Core`, `AI Client & Debug Logging`, `Streamlit Legacy App`, `PPTX Surgery Test Script`, `Edit Data Validation`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `Job Analysis Caching`?**
  _High betweenness centrality (0.052) - this node is a cross-community bridge._
- **Why does `analyze_update_package()` connect `Chart XML Writer` to `Legacy Chart Job Core`, `AI Client & Debug Logging`, `PPT Target Discovery`, `XLSX Parsing & Metadata`, `Source Matching & Normalization`?**
  _High betweenness centrality (0.030) - this node is a cross-community bridge._
- **Are the 16 inferred relationships involving `PptTarget` (e.g. with `AiSourceMatchSuggestion` and `ChartSheetUnresolvedError`) actually correct?**
  _`PptTarget` has 16 INFERRED edges - model-reasoned connections that need verification._
- **Are the 4 inferred relationships involving `Path` (e.g. with `read_bytes()` and `_filename_context_score()`) actually correct?**
  _`Path` has 4 INFERRED edges - model-reasoned connections that need verification._
- **Are the 8 inferred relationships involving `TransformPlan` (e.g. with `AiSourceMatchSuggestion` and `AiTransformDiagnostic`) actually correct?**
  _`TransformPlan` has 8 INFERRED edges - model-reasoned connections that need verification._
- **Are the 6 inferred relationships involving `ParsedXlsxTable` (e.g. with `AiSourceMatchSuggestion` and `AiMapperTests`) actually correct?**
  _`ParsedXlsxTable` has 6 INFERRED edges - model-reasoned connections that need verification._