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
]
