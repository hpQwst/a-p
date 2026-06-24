from pathlib import Path
from io import BytesIO
import sys
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from ppt_automator import (
    build_auto_chart_jobs,
    build_chart_job,
    build_chart_jobs,
    generate_pptx,
    read_source_table_from_workbook,
)

PPT = ROOT / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
MAPPING = ROOT / "Natura_2Q26_RelacionalCB_modelo_mapeamento.xlsx"
DATASOURCES = ROOT / "datasources.zip"
OUTPUT_DIR = ROOT / "outputs"


def main() -> None:
    jobs = build_chart_jobs(PPT, MAPPING, DATASOURCES)
    ok_jobs = [job for job in jobs if job.ok]
    print(f"jobs={len(jobs)} ok={len(ok_jobs)} pending={len(jobs) - len(ok_jobs)}")
    for job in jobs:
        slide = job.target.slide_number if job.target else None
        print(f"{job.graph_id} status={job.status} slide={slide} message={job.message}")
    context_count = sum(1 for job in jobs if job.target and job.target.nearby_text)
    print(f"context_targets={context_count}")
    if context_count == 0:
        raise SystemExit("Nao consegui extrair contexto textual dos slides.")

    OUTPUT_DIR.mkdir(exist_ok=True)
    output_path = OUTPUT_DIR / "Natura_2Q26_RelacionalCB_automatizado_smoke.pptx"
    output_path.write_bytes(generate_pptx(PPT, ok_jobs))
    print(f"output={output_path}")

    renamed_zip = _renamed_datasources_zip(DATASOURCES)
    renamed_jobs = build_chart_jobs(PPT, MAPPING, renamed_zip)
    renamed_ok_jobs = [job for job in renamed_jobs if job.ok]
    print(
        f"renamed_sources jobs={len(renamed_jobs)} ok={len(renamed_ok_jobs)} "
        f"pending={len(renamed_jobs) - len(renamed_ok_jobs)}"
    )
    if len(renamed_ok_jobs) != len(renamed_jobs):
        raise SystemExit("Auto-match falhou com datasources renomeados.")

    sample = ok_jobs[0]
    manual_job = build_chart_job(sample.mapping, sample.source, sample.target)
    print(f"manual_review_job={manual_job.graph_id} status={manual_job.status}")

    manual_source = _manual_source_from_zip(DATASOURCES, sample.source.file_name)
    uploaded_job = build_chart_job(sample.mapping, manual_source, sample.target)
    print(f"manual_upload_job={uploaded_job.graph_id} status={uploaded_job.status}")

    auto_jobs = build_auto_chart_jobs(PPT, renamed_zip)
    print(f"auto_mode jobs={len(auto_jobs)} ok={sum(1 for job in auto_jobs if job.ok)}")
    if len(auto_jobs) < 8:
        raise SystemExit("Modo automatico encontrou poucos matches.")


def _renamed_datasources_zip(path: Path) -> bytes:
    output = BytesIO()
    with ZipFile(path) as zin, ZipFile(output, "w", ZIP_DEFLATED) as zout:
        counter = 1
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.lower().endswith(".xlsx"):
                info.filename = f"datasources/tabela_sem_id_{counter:02d}.xlsx"
                counter += 1
            zout.writestr(info, data)
    return output.getvalue()


def _manual_source_from_zip(path: Path, source_name: str):
    with ZipFile(path) as zf:
        data = zf.read(source_name)
    return read_source_table_from_workbook(
        data,
        f"upload_manual/{Path(source_name).name}",
        formula_mode="auto",
        source_key=f"manual:{Path(source_name).name}",
    )


if __name__ == "__main__":
    main()
