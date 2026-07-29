from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
import csv
from io import StringIO
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Iterable

REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
if str(REPOSITORY_ROOT) not in sys.path:
    sys.path.insert(0, str(REPOSITORY_ROOT))

from ppt_automator import analyze_update_package, generate_updated_pptx
from ppt_automator.regression_validation import (
    validate_generated_pptx,
    validate_rendered_slides,
)


DEFAULT_FIXTURES = Path(
    os.getenv(
        "AUTO_PPT_REAL_FIXTURE_ROOT",
        r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos",
    )
)
DEFAULT_OUTPUT = Path("workspace_data") / "real_deck_regression"


@dataclass(frozen=True)
class RegressionCase:
    name: str
    pptx: Path
    datasources: Path
    large: bool = False


def fixture_cases(root: Path) -> dict[str, RegressionCase]:
    andre_ppt = "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
    return {
        "andre": RegressionCase(
            "andre",
            root / "andre" / andre_ppt,
            root / "andre" / "datasources.zip",
        ),
        "andre-enxuto": RegressionCase(
            "andre-enxuto",
            root / "andre" / "Natura_2Q26_RelacionalCB_modelo_mapeado-enxuto.pptx",
            root / "andre" / "datasources.zip",
        ),
        "hugo": RegressionCase(
            "hugo",
            root / "hugo" / andre_ppt,
            root / "hugo" / "datasources.zip",
        ),
        "mb": RegressionCase(
            "mb",
            root / "mb" / "MBTESTE_formula.pptx",
            root / "mb" / "datasources.zip",
        ),
        "mb2": RegressionCase(
            "mb2",
            root / "mb2" / "C Experiência 1Q26_TRIMESTRAL_v01.pptx",
            root / "mb2" / "datasources.zip",
            large=True,
        ),
    }


def run_case(
    case: RegressionCase,
    output_root: Path,
    *,
    render: bool,
) -> dict[str, object]:
    missing = [path for path in (case.pptx, case.datasources) if not path.exists()]
    if missing:
        raise FileNotFoundError(", ".join(str(path) for path in missing))

    case_output = output_root / case.name
    case_output.mkdir(parents=True, exist_ok=True)
    started = time.perf_counter()
    targets, sources, plans = analyze_update_package(case.pptx, case.datasources)
    analysis_seconds = time.perf_counter() - started

    generation_started = time.perf_counter()
    generated = generate_updated_pptx(case.pptx, plans, targets=targets)
    generation_seconds = time.perf_counter() - generation_started
    generated_path = case_output / f"{case.pptx.stem}-generated.pptx"
    generated_path.write_bytes(generated)

    structure = validate_generated_pptx(
        case.pptx,
        generated,
        plans,
    )
    render_report: dict[str, object] | None = None
    if render:
        original_render = case_output / "render-original"
        generated_render = case_output / "render-generated"
        _render_with_powerpoint(case.pptx, original_render)
        _render_with_powerpoint(generated_path, generated_render)
        render_report = asdict(
            validate_rendered_slides(
                case.pptx,
                original_render,
                generated_render,
                plans,
            )
        )

    return {
        "case": case.name,
        "pptx": str(case.pptx),
        "datasources": str(case.datasources),
        "large": case.large,
        "targets": len(targets),
        "sources": len(sources),
        "plans": len(plans),
        "slides_updated": sorted({plan.target.slide_number for plan in plans}),
        "warnings": sum(len(plan.warnings) for plan in plans),
        "analysis_seconds": round(analysis_seconds, 3),
        "generation_seconds": round(generation_seconds, 3),
        "total_seconds": round(time.perf_counter() - started, 3),
        "generated_pptx": str(generated_path.resolve()),
        "structure": asdict(structure),
        "render": render_report,
    }


def _render_with_powerpoint(pptx: Path, output_dir: Path) -> dict[str, object]:
    script = Path(__file__).with_name("render_pptx_with_powerpoint.ps1")
    command = [
        "powershell.exe",
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-InputPptx",
        str(pptx),
        "-OutputDir",
        str(output_dir),
    ]
    powerpoint_before = _powerpoint_process_ids()
    try:
        completed = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            timeout=300,
        )
    except subprocess.TimeoutExpired:
        # O COM pode deixar POWERPNT.EXE vivo quando o processo chamador e
        # encerrado. Matamos somente PIDs que nasceram durante este render; uma
        # instancia que ja pertencia ao usuario nunca entra nesta lista.
        for process_id in sorted(_powerpoint_process_ids() - powerpoint_before):
            subprocess.run(
                ["taskkill.exe", "/PID", str(process_id), "/T", "/F"],
                check=False,
                capture_output=True,
                text=True,
            )
        raise
    payload_lines = [line for line in completed.stdout.splitlines() if line.strip()]
    if not payload_lines:
        raise RuntimeError(f"PowerPoint nao retornou relatorio ao renderizar {pptx}.")
    return json.loads(payload_lines[-1])


def _powerpoint_process_ids() -> set[int]:
    if os.name != "nt":
        return set()
    completed = subprocess.run(
        [
            "tasklist.exe",
            "/FI",
            "IMAGENAME eq POWERPNT.EXE",
            "/FO",
            "CSV",
            "/NH",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    output: set[int] = set()
    for row in csv.reader(StringIO(completed.stdout)):
        if len(row) < 2 or row[0].casefold() != "powerpnt.exe":
            continue
        try:
            output.add(int(row[1]))
        except ValueError:
            continue
    return output


def _selected_cases(
    all_cases: dict[str, RegressionCase],
    names: Iterable[str],
    include_large: bool,
) -> list[RegressionCase]:
    requested = list(names)
    if requested:
        unknown = sorted(set(requested) - set(all_cases))
        if unknown:
            raise ValueError(f"Casos desconhecidos: {', '.join(unknown)}")
        return [all_cases[name] for name in requested]
    return [
        case
        for case in all_cases.values()
        if include_large or not case.large
    ]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Regressao real: analisa, gera, valida e opcionalmente renderiza PPTs.",
    )
    parser.add_argument("--fixtures-root", type=Path, default=DEFAULT_FIXTURES)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--case", action="append", default=[])
    parser.add_argument("--include-large", action="store_true")
    parser.add_argument("--render", action="store_true")
    args = parser.parse_args()

    cases = _selected_cases(
        fixture_cases(args.fixtures_root),
        args.case,
        args.include_large,
    )
    args.output_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, object]] = []
    failures: list[dict[str, str]] = []
    for case in cases:
        print(f"[regression] {case.name}: iniciando", flush=True)
        try:
            result = run_case(case, args.output_dir, render=args.render)
            results.append(result)
            print(
                f"[regression] {case.name}: OK "
                f"({result['plans']} planos, {result['total_seconds']}s)",
                flush=True,
            )
        except Exception as exc:
            failures.append({"case": case.name, "error": f"{type(exc).__name__}: {exc}"})
            print(f"[regression] {case.name}: FALHOU - {exc}", file=sys.stderr, flush=True)

    report = {
        "fixtures_root": str(args.fixtures_root.resolve()),
        "output_dir": str(args.output_dir.resolve()),
        "rendered_with_powerpoint": bool(args.render),
        "results": results,
        "failures": failures,
    }
    report_path = args.output_dir / "report.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(report, ensure_ascii=False))
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
