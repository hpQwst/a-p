# Graph Report - automatizador-ppt  (2026-07-08)

## Corpus Check
- 65 files · ~55,564 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 978 nodes · 3334 edges · 27 communities (24 shown, 3 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 67 edges (avg confidence: 0.55)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `e3006afe`
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
- AI Call Batching Limits
- Fargate Deploy Script
- CodeBuild Docker Image
- Web App Package Init
- Worker Package Boundary

## God Nodes (most connected - your core abstractions)
1. `PptTarget` - 96 edges
2. `ParsedXlsxTable` - 82 edges
3. `TransformPlan` - 56 edges
4. `_render_preview()` - 36 edges
5. `AnalysisResult` - 35 edges
6. `_analysis_for_job()` - 27 edges
7. `analyze_update_package()` - 26 edges
8. `_run_slide_ai_review()` - 26 edges
9. `preview()` - 25 edges
10. `log_debug_event()` - 24 edges

## Surprising Connections (you probably didn't know these)
- `uploaded_source_table()` --calls--> `read_source_table_from_workbook()`  [EXTRACTED]
  app.py → ppt_automator/core.py
- `_ai_source_matches_for_job()` --indirect_call--> `suggest_source_matches_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_mapper.py
- `_run_slide_ai_review()` --indirect_call--> `build_slide_matrices_with_ai()`  [INFERRED]
  web/main.py → ppt_automator/ai_slide_matrix_builder.py
- `_run_slide_ai_review()` --indirect_call--> `suggest_slide_understanding()`  [INFERRED]
  web/main.py → ppt_automator/ai_slide_understanding.py
- `_ai_diagnostics_for_job()` --indirect_call--> `suggest_transform_diagnostics()`  [INFERRED]
  web/main.py → ppt_automator/ai_transform.py

## Import Cycles
- None detected.

## Communities (27 total, 3 thin omitted)

### Community 0 - "Legacy Chart Job Core"
Cohesion: 0.06
Nodes (89): FormulaMode, _average(), _best_target_label(), _best_text_score(), build_auto_chart_jobs(), build_chart_job(), build_chart_jobs(), _build_chart_matrix() (+81 more)

### Community 1 - "Embedded Workbook Writer"
Cohesion: 0.08
Nodes (52): assert_valid_typed_edit_data(), _chart_series_capacity_errors(), EditDataValidationError, Any, ValueError, O writer (ppt_chart_writer.py) so atualiza os <c:ser> que ja existem no     temp, validate_typed_edit_data(), _collect_namespace() (+44 more)

### Community 2 - "AI Client & Debug Logging"
Cohesion: 0.07
Nodes (59): BaseException, Exception, AiMatchSuggestion, AiUnavailableError, build_openai_client(), ai_debug_log(), _append_jsonl(), current_ai_debug_log_path() (+51 more)

### Community 3 - "Streamlit Legacy App"
Cohesion: 0.12
Nodes (52): build_review_jobs(), file_bytes(), generated_output_for(), item_match_score(), list_match_pct(), missing_items(), norm_text(), Path (+44 more)

### Community 4 - "Slide Datasources & Rendering"
Cohesion: 0.16
Nodes (20): _cell_sort_key(), _cell_type(), _compact_cell(), _display_value(), dump_xlsx_workbook(), dump_xlsx_zip_entries(), _merged_lookup(), _node_text() (+12 more)

### Community 5 - "PPT Target Discovery"
Cohesion: 0.06
Nodes (68): mapping_entry_learning_fields(), Any, Camada 4 - aprendizado do mapeamento (match "perfeito na 2a vez").  A cada downl, Campos de aprendizado a persistir na entrada do template de mapeamento., Traduz as entradas do template salvo em matches para os targets/sources     ATUA, _reason_for(), resolve_learned_matches(), _resolve_source() (+60 more)

### Community 6 - "PPTX Surgery Test Script"
Cohesion: 0.15
Nodes (37): Namespace, _compress_zip_payload(), _excel_col(), _first_table_path(), _join_xlsx_path(), main(), _matrix_strings(), _number_or_original() (+29 more)

### Community 7 - "Slide AI Review Pipeline"
Cohesion: 0.11
Nodes (35): _ai_review_confidence_floor(), _apply_slide_ai_outputs_enabled(), _auto_slide_ai_confidence_floor(), _auto_source_match_review_enabled(), _cards_by_slide(), _completed_preview_cache_usable(), _display_row(), _env_bool() (+27 more)

### Community 8 - "XLSX Parsing & Metadata"
Cohesion: 0.16
Nodes (35): date, datetime, _date_period_label(), _extract_context_metadata(), _extract_metadata(), _find_header_row(), _find_label_col(), _first_row_group_label() (+27 more)

### Community 9 - "Edit Data Validation"
Cohesion: 0.16
Nodes (23): chart_replacements(), _chart_value_text(), ChartSheetUnresolvedError, _disable_chart_auto_update(), effective_value_format(), _excel_col(), Any, Element (+15 more)

### Community 10 - "Source Matching & Normalization"
Cohesion: 0.06
Nodes (83): _ai_match_target_payload(), build_ai_mapping_payload(), _env_float(), _env_int(), _non_negative_int(), plans_to_ai_payload(), Any, _recipe_schema() (+75 more)

### Community 11 - "Preview Web Endpoints"
Cohesion: 0.23
Nodes (30): HTMLResponse, ai_configured(), Request, _ai_waiting_status(), apply_job_mapping_template(), approve_slide(), approve_target(), _canonical_target_id() (+22 more)

### Community 12 - "Analysis Worker Orchestration"
Cohesion: 0.10
Nodes (49): ManualSourceMap, ManualSourcePayload, AiSourceMatchSuggestion, AiTransformDiagnostic, build_preview(), _display_format(), _display_value(), _format_pt_number() (+41 more)

### Community 13 - "Chart XML Writer"
Cohesion: 0.06
Nodes (38): Nome real da primeira aba do workbook embutido, para resolver target.sheet_name, resolve_default_sheet_name(), analyze_update_package(), _build_slide_aware_plans(), _datasource_names_for_slides(), generate_updated_pptx(), preview_update_package(), Any (+30 more)

### Community 14 - "AI Mapping Payloads"
Cohesion: 0.24
Nodes (14): _find_graphic_frame(), _format_value(), _matches_categories(), _norm(), Any, Element, _set_cell_text(), update_table_slide_xml() (+6 more)

### Community 15 - "Upload & Job Status API"
Cohesion: 0.16
Nodes (20): UploadFile, _checkpoint_slide_label(), _checkpoint_summary(), _error_response(), index(), _inspect_ppt_upload(), _large_deck_confirmation_message(), _large_deck_slide_threshold() (+12 more)

### Community 16 - "Job Debug & Cache State"
Cohesion: 0.16
Nodes (20): FileResponse, JSONResponse, reset_ai_debug_log_path(), set_ai_debug_log_path(), _call_with_job_debug(), _combine_ai_status(), _debug_log_path(), _init_preview_processing_state() (+12 more)

### Community 17 - "Frontend Preview UI"
Cohesion: 0.16
Nodes (17): Squad Mapping Template (Modelo de Mapeamento), activeSquad(), applyPreviewFilters(), deckScopeSize(), inspectPpt(), parseSlideScope(), previewFilterState(), progressStepsForMessage() (+9 more)

### Community 18 - "render_slide_with_target_labels"
Cohesion: 0.29
Nodes (13): Image, _cached_deck_pdf(), _export_slide_with_libreoffice(), _export_slide_with_powerpoint(), _fallback_slide_canvas(), _font(), InputFile, Path (+5 more)

### Community 19 - "Job Analysis Caching"
Cohesion: 0.18
Nodes (25): Response, _analysis_for_job(), _analyze_files_signature(), _apply_mapping_template_to_analysis(), _cached_analyze_files(), _clear_ai_cache(), _clear_slide_ai_state(), download() (+17 more)

### Community 20 - "Preview Context & Flags"
Cohesion: 0.15
Nodes (17): _ai_input_stats(), _ai_payload_profile(), _builder_payload_from_understanding(), _datasource_entries_for_slide(), _datasource_entry_hashes(), _json_size_bytes(), O understanding acabou de classificar quais XLSX do slide tem partes uteis     (, Restringe os XLSX enviados para IA de nivel de slide aos candidatos com sinal (+9 more)

### Community 22 - "AI Call Batching Limits"
Cohesion: 0.18
Nodes (17): _ai_diagnostic_batch_limit(), _ai_diagnostics_for_job(), _ai_source_match_batch_limit(), _ai_source_match_max_calls(), _ai_source_matches_for_job(), _append_ai_log(), _auto_slide_ai_enabled(), _diagnostic_payload_summary() (+9 more)

### Community 26 - "Fargate Deploy Script"
Cohesion: 0.43
Nodes (5): Assert-AwsOk(), Ensure-Role(), Get-RoleArn(), Put-InlinePolicy(), Save-JsonFile()

## Knowledge Gaps
- **3 isolated node(s):** `CodeBuild Buildspec (ECR image build)`, `Error Page Template`, `Squad Mapping Template (Modelo de Mapeamento)`
  These have ≤1 connection - possible missing edges or undocumented components.
- **3 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `PptTarget` connect `Source Matching & Normalization` to `Legacy Chart Job Core`, `PPT Target Discovery`, `PPTX Surgery Test Script`, `Edit Data Validation`, `Analysis Worker Orchestration`, `Chart XML Writer`, `AI Mapping Payloads`, `render_slide_with_target_labels`?**
  _High betweenness centrality (0.093) - this node is a cross-community bridge._
- **Why does `ParsedXlsxTable` connect `Source Matching & Normalization` to `Legacy Chart Job Core`, `PPT Target Discovery`, `XLSX Parsing & Metadata`, `Analysis Worker Orchestration`, `Chart XML Writer`?**
  _High betweenness centrality (0.050) - this node is a cross-community bridge._
- **Why does `TransformPlan` connect `Analysis Worker Orchestration` to `Legacy Chart Job Core`, `AI Client & Debug Logging`, `PPTX Surgery Test Script`, `Edit Data Validation`, `Source Matching & Normalization`, `Chart XML Writer`, `AI Mapping Payloads`?**
  _High betweenness centrality (0.045) - this node is a cross-community bridge._
- **Are the 16 inferred relationships involving `PptTarget` (e.g. with `AiSourceMatchSuggestion` and `ChartSheetUnresolvedError`) actually correct?**
  _`PptTarget` has 16 INFERRED edges - model-reasoned connections that need verification._
- **Are the 12 inferred relationships involving `ParsedXlsxTable` (e.g. with `AiSourceMatchSuggestion` and `SourceMatchCandidate`) actually correct?**
  _`ParsedXlsxTable` has 12 INFERRED edges - model-reasoned connections that need verification._
- **Are the 9 inferred relationships involving `TransformPlan` (e.g. with `AiSourceMatchSuggestion` and `AiTransformDiagnostic`) actually correct?**
  _`TransformPlan` has 9 INFERRED edges - model-reasoned connections that need verification._
- **What connects `Utilities for mapping datasource workbooks into mapped PowerPoint charts.`, `Permite modelo mais barato por operacao (ex.: OPENAI_MODEL_SOURCE_MATCH para o`, `Raised when a workbook contains formulas but no calculator can resolve them.` to the rest of the system?**
  _31 weakly-connected nodes found - possible documentation gaps or missing edges._