from __future__ import annotations

import hashlib
from pathlib import Path, PurePosixPath
from typing import Iterable
import os


class AwsExecutionError(RuntimeError):
    pass


def execution_backend() -> str:
    return os.getenv("AUTO_PPT_EXECUTION_BACKEND", "local").strip().lower() or "local"


def uses_fargate_workers() -> bool:
    return execution_backend() == "fargate"


def job_bucket() -> str:
    return os.getenv("AUTO_PPT_S3_BUCKET", "").strip()


def job_prefix(job_id: str, *parts: str) -> str:
    base = os.getenv("AUTO_PPT_S3_PREFIX", "auto-ppt").strip().strip("/") or "auto-ppt"
    cleaned = [base, "jobs", job_id]
    cleaned.extend(part.strip("/") for part in parts if part)
    return "/".join(cleaned)


def upload_job(job_dir: Path) -> None:
    bucket = _required_bucket()
    client = _s3_client()
    for file_path in _job_files(job_dir):
        relative = file_path.relative_to(job_dir).as_posix()
        client.upload_file(str(file_path), bucket, job_prefix(job_dir.name, relative))


def hydrate_job(job_id: str, job_dir: Path) -> None:
    bucket = _required_bucket()
    client = _s3_client()
    prefix = job_prefix(job_id) + "/"
    token = None
    found = False
    while True:
        request = {"Bucket": bucket, "Prefix": prefix}
        if token:
            request["ContinuationToken"] = token
        response = client.list_objects_v2(**request)
        for item in response.get("Contents", []):
            key = str(item.get("Key") or "")
            relative = key.removeprefix(prefix)
            if not relative:
                continue
            destination = _safe_destination(job_dir, relative)
            destination.parent.mkdir(parents=True, exist_ok=True)
            client.download_file(bucket, key, str(destination))
            found = True
        if not response.get("IsTruncated"):
            break
        token = response.get("NextContinuationToken")
    if not found:
        raise FileNotFoundError(f"Job {job_id} nao encontrado no S3.")


def upload_job_file(job_dir: Path, filename: str) -> None:
    path = _safe_destination(job_dir, filename)
    if path.exists():
        _s3_client().upload_file(str(path), _required_bucket(), job_prefix(job_dir.name, filename))


def refresh_job_file(job_id: str, job_dir: Path, filename: str) -> bool:
    destination = _safe_destination(job_dir, filename)
    try:
        _s3_client().download_file(_required_bucket(), job_prefix(job_id, filename), str(destination))
        return True
    except Exception as exc:
        if _not_found(exc):
            return False
        raise


def launch_fargate_job(job_id: str, operation: str, launch_token: str) -> str:
    cluster = os.getenv("AUTO_PPT_ECS_CLUSTER", "").strip()
    task_definition = os.getenv("AUTO_PPT_ECS_WORKER_TASK_DEFINITION", "").strip()
    subnets = _csv_env("AUTO_PPT_ECS_SUBNETS")
    security_groups = _csv_env("AUTO_PPT_ECS_SECURITY_GROUPS")
    if not cluster or not task_definition or not subnets or not security_groups:
        raise AwsExecutionError(
            "Configuracao Fargate incompleta: AUTO_PPT_ECS_CLUSTER, AUTO_PPT_ECS_WORKER_TASK_DEFINITION, "
            "AUTO_PPT_ECS_SUBNETS e AUTO_PPT_ECS_SECURITY_GROUPS sao obrigatorias."
        )
    response = _ecs_client().run_task(
        cluster=cluster,
        taskDefinition=task_definition,
        launchType="FARGATE",
        platformVersion="LATEST",
        clientToken=_client_token(job_id, operation, launch_token),
        networkConfiguration={
            "awsvpcConfiguration": {
                "subnets": subnets,
                "securityGroups": security_groups,
                "assignPublicIp": "ENABLED",
            }
        },
        overrides={
            "containerOverrides": [
                {
                    "name": "worker",
                    "command": ["python", "-m", "worker.aws_task", "--job-id", job_id, "--operation", operation],
                }
            ]
        },
        startedBy="auto-ppt-web",
        tags=[{"key": "JobId", "value": job_id}, {"key": "Operation", "value": operation}],
    )
    failures = response.get("failures") or []
    if failures:
        raise AwsExecutionError(f"ECS nao iniciou o job: {failures[0]}")
    tasks = response.get("tasks") or []
    if not tasks:
        raise AwsExecutionError("ECS nao retornou uma task para o job.")
    return str(tasks[0].get("taskArn") or "")


def _client_token(job_id: str, operation: str, launch_token: str) -> str:
    digest = hashlib.sha256(f"{job_id}:{operation}:{launch_token}".encode("utf-8")).hexdigest()
    return digest[:36]


def _job_files(job_dir: Path) -> Iterable[Path]:
    for path in job_dir.rglob("*"):
        if path.is_file() and not path.name.startswith("."):
            yield path


def _safe_destination(root: Path, relative: str) -> Path:
    pure = PurePosixPath(relative.replace("\\", "/"))
    if pure.is_absolute() or ".." in pure.parts:
        raise AwsExecutionError("Chave de job insegura no S3.")
    return root.joinpath(*pure.parts)


def _required_bucket() -> str:
    bucket = job_bucket()
    if not bucket:
        raise AwsExecutionError("AUTO_PPT_S3_BUCKET e obrigatorio para jobs Fargate.")
    return bucket


def _csv_env(name: str) -> list[str]:
    return [item.strip() for item in os.getenv(name, "").split(",") if item.strip()]


def _s3_client():
    import boto3

    return boto3.client("s3")


def _ecs_client():
    import boto3

    return boto3.client("ecs")


def _not_found(exc: Exception) -> bool:
    response = getattr(exc, "response", {}) or {}
    code = str((response.get("Error") or {}).get("Code") or "")
    return code in {"404", "NoSuchKey", "NoSuchBucket", "NotFound"}
