# Graph Report - automatizador-ppt  (2026-07-08)

## Corpus Check
- 61 files · ~53,292 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 942 nodes · 2980 edges · 32 communities (27 shown, 5 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 69 edges (avg confidence: 0.64)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `40dc88b7`
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
- Table Writer & Tests
- AI Call Batching Limits
- AI Scope Tests
- Preview Model Building
- Source Manifest & Plans
- Fargate Deploy Script
- CodeBuild Docker Image
- Web App Package Init
- Worker Package Boundary
- Decimal

## God Nodes (most connected - your core abstractions)
1. `TransformPlan` - 54 edges
2. `PptTarget` - 40 edges
3. `_render_preview()` - 36 edges
4. `AnalysisResult` - 32 edges
5. `ParsedXlsxTable` - 28 edges
6. `_analysis_for_job()` - 27 edges
7. `preview()` - 26 edges
8. `_run_slide_ai_review()` - 26 edges
9. `analyze_update_package()` - 23 edges
10. `_preview_context_from_analysis()` - 22 edges

## Surprising Connections (you probably didn't know these)
- `AnalysisResult` --uses--> `AiSourceMatchSuggestion`  [INFERRED]
  worker/processor.py → ppt_automator/ai_mapper.py
- `_ai_source_matches_for_job()` --indirect_call--> `suggest_source_matches_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_mapper.py
- `_run_slide_ai_review()` --indirect_call--> `build_slide_matrices_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_slide_matrix_builder.py
- `_run_slide_ai_review()` --indirect_call--> `suggest_slide_understanding()`  [INFERRED]
  web/main.py → ppt_automator/ai_slide_understanding.py
- `AiScopeTests` --uses--> `AiTransformDiagnostic`  [INFERRED]
  tests/test_ai_scope.py → ppt_automator/ai_transform.py

## Import Cycles
- None detected.

## Communities (32 total, 5 thin omitted)

### Community 0 - "Legacy Chart Job Core"
Cohesion: 0.06
Nodes (90): uploaded_source_table(), FormulaMode, _average(), _best_target_label(), _best_text_score(), build_auto_chart_jobs(), build_chart_job(), build_chart_jobs() (+82 more)

### Community 1 - "Embedded Workbook Writer"
Cohesion: 0.14
Nodes (32): Decimal, _collect_namespace(), EmbeddedWorkbookWriterUnavailable, _excel_col(), _first_table_path(), _join_xlsx_path(), _matrix_cell_is_text(), _matrix_cell_number() (+24 more)

### Community 2 - "AI Client & Debug Logging"
Cohesion: 0.08
Nodes (47): BaseException, ChartJob, Exception, AiMatchSuggestion, AiUnavailableError, build_openai_client(), _env_float(), _env_int() (+39 more)

### Community 3 - "Streamlit Legacy App"
Cohesion: 0.11
Nodes (52): build_review_jobs(), file_bytes(), generated_output_for(), item_match_score(), list_match_pct(), missing_items(), norm_text(), Path (+44 more)

### Community 4 - "Slide Datasources & Rendering"
Cohesion: 0.10
Nodes (31): InputFile, read_bytes(), collect_datasource_entries(), entries_for_slide(), parse_slide_number_from_path(), _parse_slide_suffix(), _parse_slide_token(), InputFile (+23 more)

### Community 5 - "PPT Target Discovery"
Cohesion: 0.10
Nodes (45): _cache_values(), _chart_orientation(), _chart_target(), _chart_workbook_lookup_enabled(), discover_ppt_targets(), _formula_range_values(), _formula_sheet_name(), _formula_single_value() (+37 more)

### Community 6 - "PPTX Surgery Test Script"
Cohesion: 0.15
Nodes (37): Namespace, _compress_zip_payload(), _excel_col(), _first_table_path(), _join_xlsx_path(), main(), _matrix_strings(), _number_or_original() (+29 more)

### Community 7 - "Slide AI Review Pipeline"
Cohesion: 0.09
Nodes (42): _ai_diagnostic_batch_limit(), _ai_input_stats(), _ai_payload_profile(), _ai_source_match_batch_limit(), _ai_source_match_max_calls(), _apply_slide_ai_outputs_enabled(), _auto_slide_ai_confidence_floor(), _auto_source_match_review_enabled() (+34 more)

### Community 8 - "XLSX Parsing & Metadata"
Cohesion: 0.15
Nodes (37): date, datetime, _date_period_label(), _extract_context_metadata(), _extract_metadata(), _find_header_row(), _find_label_col(), _first_row_group_label() (+29 more)

### Community 9 - "Edit Data Validation"
Cohesion: 0.17
Nodes (21): assert_valid_typed_edit_data(), _chart_series_capacity_errors(), EditDataValidationError, Any, O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp, validate_typed_edit_data(), cell_to_excel_value(), cell_to_python_value() (+13 more)

### Community 10 - "Source Matching & Normalization"
Cohesion: 0.11
Nodes (47): _find_graphic_frame(), _format_value(), _matches_categories(), _norm(), Any, Element, _set_cell_text(), update_table_slide_xml() (+39 more)

### Community 11 - "Preview Web Endpoints"
Cohesion: 0.22
Nodes (23): HTMLResponse, Request, approve_slide(), approve_target(), _completed_preview_cache_usable(), _error_response(), index(), _job_dir() (+15 more)

### Community 12 - "Analysis Worker Orchestration"
Cohesion: 0.12
Nodes (35): ManualSourceMap, ManualSourcePayload, AiSourceMatchSuggestion, SavedSourceMatchMap, AiScopeTests, _analysis(), _plan(), _source() (+27 more)

### Community 13 - "Chart XML Writer"
Cohesion: 0.06
Nodes (47): analyze_update_package(), _build_slide_aware_plans(), _datasource_names_for_slides(), generate_updated_pptx(), preview_update_package(), Any, InputFile, ParsedXlsxTable (+39 more)

### Community 14 - "AI Mapping Payloads"
Cohesion: 0.24
Nodes (27): _ai_match_target_payload(), build_ai_mapping_payload(), _env_float(), _env_int(), _non_negative_int(), plans_to_ai_payload(), Any, ParsedXlsxTable (+19 more)

### Community 15 - "Upload & Job Status API"
Cohesion: 0.19
Nodes (21): UploadFile, ValueError, _canonical_target_id(), _inspect_ppt_upload(), _large_deck_confirmation_message(), _large_deck_slide_threshold(), _large_scope_requires_confirmation(), _normalize_squad_form() (+13 more)

### Community 16 - "Job Debug & Cache State"
Cohesion: 0.16
Nodes (22): FileResponse, JSONResponse, _clear_ai_cache(), _clear_slide_ai_state(), _init_preview_processing_state(), job_debug_log(), job_processing_status(), _load_ai_diagnostics() (+14 more)

### Community 17 - "Frontend Preview UI"
Cohesion: 0.16
Nodes (17): Squad Mapping Template (Modelo de Mapeamento), activeSquad(), applyPreviewFilters(), deckScopeSize(), inspectPpt(), parseSlideScope(), previewFilterState(), progressStepsForMessage() (+9 more)

### Community 18 - "render_slide_with_target_labels"
Cohesion: 0.27
Nodes (13): Image, _cached_deck_pdf(), _export_slide_with_libreoffice(), _export_slide_with_powerpoint(), _fallback_slide_canvas(), _font(), InputFile, Path (+5 more)

### Community 19 - "Job Analysis Caching"
Cohesion: 0.14
Nodes (34): Response, _analysis_for_job(), _analyze_files_signature(), apply_job_mapping_template(), _apply_mapping_template_to_analysis(), _auto_slide_ai_enabled(), _cached_analyze_files(), _clear_render_cache() (+26 more)

### Community 20 - "Preview Context & Flags"
Cohesion: 0.21
Nodes (12): _builder_payload_from_understanding(), _datasource_entries_for_slide(), _datasource_entry_hashes(), O understanding acabou de classificar quais XLSX do slide tem partes uteis     (, Restringe os XLSX enviados para IA de nivel de slide aos candidatos com sinal, _relevant_datasource_entries_for_targets(), _render_slide_for_job(), _run_slide_ai_review() (+4 more)

### Community 21 - "Table Writer & Tests"
Cohesion: 0.45
Nodes (10): _central_header(), _compress_payload(), _datetime(), _filename_bytes(), _local_header(), Any, ValueError, _raw_local_record_ends() (+2 more)

### Community 22 - "AI Call Batching Limits"
Cohesion: 0.20
Nodes (16): ai_configured(), _ai_diagnostics_for_job(), _ai_review_confidence_floor(), _ai_source_matches_for_job(), _ai_waiting_status(), _append_ai_log(), _call_with_job_debug(), _combine_ai_status() (+8 more)

### Community 23 - "AI Scope Tests"
Cohesion: 0.53
Nodes (4): env_int(), image_data_url(), _optimized_image_bytes(), Path

### Community 25 - "Source Manifest & Plans"
Cohesion: 0.13
Nodes (23): ai_debug_log(), _append_jsonl(), current_ai_debug_log_path(), _json_size(), log_ai_request(), log_ai_response(), log_debug_event(), Any (+15 more)

### Community 26 - "Fargate Deploy Script"
Cohesion: 0.43
Nodes (5): Assert-AwsOk(), Ensure-Role(), Get-RoleArn(), Put-InlinePolicy(), Save-JsonFile()

## Knowledge Gaps
- **3 isolated node(s):** `CodeBuild Buildspec (ECR image build)`, `Error Page Template`, `Squad Mapping Template (Modelo de Mapeamento)`
  These have ≤1 connection - possible missing edges or undocumented components.
- **5 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `TransformPlan` connect `Chart XML Writer` to `Legacy Chart Job Core`, `AI Client & Debug Logging`, `Streamlit Legacy App`, `PPTX Surgery Test Script`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `AI Mapping Payloads`, `Job Analysis Caching`?**
  _High betweenness centrality (0.076) - this node is a cross-community bridge._
- **Why does `PptTarget` connect `PPT Target Discovery` to `Legacy Chart Job Core`, `Streamlit Legacy App`, `PPTX Surgery Test Script`, `Source Matching & Normalization`, `Analysis Worker Orchestration`, `Source Manifest & Plans`?**
  _High betweenness centrality (0.039) - this node is a cross-community bridge._
- **Why does `Index / Upload Page Template` connect `Frontend Preview UI` to `Slide AI Review Pipeline`?**
  _High betweenness centrality (0.037) - this node is a cross-community bridge._
- **Are the 7 inferred relationships involving `TransformPlan` (e.g. with `AiSourceMatchSuggestion` and `AiTransformDiagnostic`) actually correct?**
  _`TransformPlan` has 7 INFERRED edges - model-reasoned connections that need verification._
- **Are the 7 inferred relationships involving `PptTarget` (e.g. with `AiMapperTests` and `AiScopeTests`) actually correct?**
  _`PptTarget` has 7 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Permite modelo mais barato por operacao (ex.: OPENAI_MODEL_SOURCE_MATCH para o`, `O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp`, `Nome real da primeira aba do workbook embutido, para resolver target.sheet_name` to the rest of the system?**
  _17 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Legacy Chart Job Core` be split into smaller, more focused modules?**
  _Cohesion score 0.06325247079964061 - nodes in this community are weakly interconnected._