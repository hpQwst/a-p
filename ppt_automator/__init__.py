"""Utilities for mapping datasource workbooks into mapped PowerPoint charts."""

from .core import (
    ChartJob,
    ChartTarget,
    MappingRow,
    SourceMatch,
    SourceTable,
    build_auto_chart_jobs,
    build_chart_job,
    build_chart_jobs,
    generate_pptx,
    load_datasource_tables,
    load_datasources,
    load_mapping,
    load_ppt_targets,
    read_source_table_from_workbook,
    suggest_source_matches,
)
from .engine import analyze_update_package, generate_updated_pptx, preview_update_package
from .ppt_discovery import PptTarget, discover_ppt_targets
from .table_normalizer import TransformPlan
from .xlsx_parser import ParsedXlsxTable, parse_datasource_zip, parse_xlsx_table

__all__ = [
    "ChartJob",
    "ChartTarget",
    "MappingRow",
    "SourceMatch",
    "SourceTable",
    "build_auto_chart_jobs",
    "build_chart_job",
    "build_chart_jobs",
    "generate_pptx",
    "load_datasource_tables",
    "load_datasources",
    "load_mapping",
    "load_ppt_targets",
    "read_source_table_from_workbook",
    "suggest_source_matches",
    "PptTarget",
    "ParsedXlsxTable",
    "TransformPlan",
    "analyze_update_package",
    "discover_ppt_targets",
    "generate_updated_pptx",
    "parse_datasource_zip",
    "parse_xlsx_table",
    "preview_update_package",
]
