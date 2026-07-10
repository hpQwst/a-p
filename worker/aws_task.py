from __future__ import annotations

import argparse
from pathlib import Path
import traceback

from worker import aws_runtime


def main() -> int:
    parser = argparse.ArgumentParser(description="Executa um job Auto PPT no Fargate.")
    parser.add_argument("--job-id", required=True)
    parser.add_argument("--operation", required=True, choices=["preview", "generate"])
    args = parser.parse_args()

    from web.main import _generate_job_output, _preview_processing_worker, _save_generation_state, _init_generation_state

    runtime_root = Path(__import__("os").environ.get("AUTO_PPT_RUNTIME_ROOT", "/tmp/auto-ppt-jobs"))
    job_dir = runtime_root / args.job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    aws_runtime.hydrate_job(args.job_id, job_dir)
    try:
        if args.operation == "preview":
            _preview_processing_worker(job_dir)
        else:
            _init_generation_state(job_dir)
            _generate_job_output(job_dir)
        aws_runtime.upload_job(job_dir)
        return 0
    except Exception as exc:
        if args.operation == "generate":
            _save_generation_state(job_dir, {"status": "error", "active": False, "message": str(exc)})
        (job_dir / "worker_error.txt").write_text(traceback.format_exc(), encoding="utf-8")
        aws_runtime.upload_job(job_dir)
        raise


if __name__ == "__main__":
    raise SystemExit(main())
