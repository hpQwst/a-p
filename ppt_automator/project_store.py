from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
import json
import os
import re
import uuid


SQUADS = [f"squad{i}" for i in range(1, 6)]
DEFAULT_DATA_ROOT = "workspace_data"


@dataclass(frozen=True)
class ProjectRef:
    project_id: str
    squad: str
    slug: str
    name: str
    description: str
    created_at: str
    updated_at: str
    backend: str


@dataclass(frozen=True)
class MappingTemplateRef:
    template_id: str
    squad: str
    slug: str
    name: str
    description: str
    created_at: str
    updated_at: str
    entry_count: int
    source_project_slug: str
    source_project_name: str
    backend: str


@dataclass(frozen=True)
class RunRef:
    run_id: str
    created_at: str


def storage_backend() -> str:
    return os.getenv("AUTO_PPT_STORAGE_BACKEND", "local").strip().lower() or "local"


def storage_prefix() -> str:
    return os.getenv("AUTO_PPT_S3_PREFIX", "auto-ppt").strip().strip("/")


def data_root() -> Path:
    return Path(os.getenv("AUTO_PPT_DATA_ROOT", DEFAULT_DATA_ROOT)).resolve()


def slugify(value: str) -> str:
    slug = value.strip().lower()
    slug = re.sub(r"[^a-z0-9]+", "-", slug)
    slug = re.sub(r"-+", "-", slug).strip("-")
    return slug or "projeto"


def safe_filename(value: str) -> str:
    name = Path(value or "arquivo").name
    name = re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip(" .")
    return name or "arquivo"


def normalize_squad(value: str) -> str:
    squad = value.strip().lower()
    if squad not in SQUADS:
        raise ValueError(f"Squad invalido: {value}")
    return squad


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def ensure_store() -> str:
    backend = storage_backend()
    if backend == "local":
        root = data_root()
        for squad in SQUADS:
            (root / "squads" / squad / "projects").mkdir(parents=True, exist_ok=True)
            (root / "squads" / squad / "mapping_templates").mkdir(parents=True, exist_ok=True)
        return str(root)
    if backend == "s3":
        bucket = _s3_bucket()
        if not bucket:
            raise ValueError("AUTO_PPT_S3_BUCKET precisa estar configurado quando AUTO_PPT_STORAGE_BACKEND=s3.")
        return f"s3://{bucket}/{storage_prefix()}"
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def list_projects(squad: str) -> list[ProjectRef]:
    squad = normalize_squad(squad)
    backend = storage_backend()
    if backend == "local":
        projects_root = data_root() / "squads" / squad / "projects"
        if not projects_root.exists():
            return []
        projects = []
        for project_dir in sorted(projects_root.iterdir()):
            meta_path = project_dir / "project.json"
            if meta_path.exists():
                projects.append(_project_from_meta(_read_json_file(meta_path), backend))
        return sorted(projects, key=lambda item: item.updated_at, reverse=True)
    if backend == "s3":
        client = _s3_client()
        bucket = _s3_bucket(required=True)
        prefix = _s3_project_prefix(squad, "")
        projects = []
        token = None
        while True:
            kwargs = {"Bucket": bucket, "Prefix": prefix}
            if token:
                kwargs["ContinuationToken"] = token
            response = client.list_objects_v2(**kwargs)
            for item in response.get("Contents", []):
                key = item.get("Key", "")
                if key.endswith("/project.json"):
                    projects.append(_project_from_meta(_read_json_s3(key), backend))
            if not response.get("IsTruncated"):
                break
            token = response.get("NextContinuationToken")
        return sorted(projects, key=lambda item: item.updated_at, reverse=True)
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def load_project(squad: str, slug: str) -> ProjectRef | None:
    squad = normalize_squad(squad)
    slug = slugify(slug)
    backend = storage_backend()
    if backend == "local":
        meta_path = data_root() / "squads" / squad / "projects" / slug / "project.json"
        if not meta_path.exists():
            return None
        return _project_from_meta(_read_json_file(meta_path), backend)
    if backend == "s3":
        key = _s3_project_prefix(squad, slug, "project.json")
        try:
            return _project_from_meta(_read_json_s3(key), backend)
        except FileNotFoundError:
            return None
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def create_project(squad: str, name: str, description: str = "") -> ProjectRef:
    squad = normalize_squad(squad)
    backend = storage_backend()
    base_slug = slugify(name)
    slug = base_slug
    counter = 2
    while load_project(squad, slug):
        slug = f"{base_slug}-{counter}"
        counter += 1
    now = utc_now()
    meta = {
        "id": str(uuid.uuid4()),
        "squad": squad,
        "slug": slug,
        "name": name.strip(),
        "description": description.strip(),
        "created_at": now,
        "updated_at": now,
        "backend": backend,
        "schema_version": 1,
    }
    if backend == "local":
        project_root = data_root() / "squads" / squad / "projects" / slug
        for folder in ["templates", "runs", "memory", "memory/manual_sources"]:
            (project_root / folder).mkdir(parents=True, exist_ok=True)
        _write_json_file(project_root / "project.json", meta)
    elif backend == "s3":
        _write_json_s3(_s3_project_prefix(squad, slug, "project.json"), meta)
    else:
        raise ValueError(f"Backend de storage nao suportado: {backend}")
    return _project_from_meta(meta, backend)


def list_mapping_templates(squad: str) -> list[MappingTemplateRef]:
    squad = normalize_squad(squad)
    backend = storage_backend()
    if backend == "local":
        templates_root = data_root() / "squads" / squad / "mapping_templates"
        if not templates_root.exists():
            return []
        templates = []
        for template_dir in sorted(templates_root.iterdir()):
            meta_path = template_dir / "template.json"
            if meta_path.exists():
                templates.append(_mapping_template_from_payload(_read_json_file(meta_path), backend))
        return sorted(templates, key=lambda item: item.updated_at, reverse=True)
    if backend == "s3":
        client = _s3_client()
        bucket = _s3_bucket(required=True)
        prefix = _s3_mapping_template_prefix(squad, "")
        templates = []
        token = None
        while True:
            kwargs = {"Bucket": bucket, "Prefix": prefix}
            if token:
                kwargs["ContinuationToken"] = token
            response = client.list_objects_v2(**kwargs)
            for item in response.get("Contents", []):
                key = item.get("Key", "")
                if key.endswith("/template.json"):
                    templates.append(_mapping_template_from_payload(_read_json_s3(key), backend))
            if not response.get("IsTruncated"):
                break
            token = response.get("NextContinuationToken")
        return sorted(templates, key=lambda item: item.updated_at, reverse=True)
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def load_mapping_template(squad: str, slug: str) -> dict | None:
    squad = normalize_squad(squad)
    slug = slugify(slug)
    backend = storage_backend()
    if backend == "local":
        path = data_root() / "squads" / squad / "mapping_templates" / slug / "template.json"
        if not path.exists():
            return None
        return _read_json_file(path)
    if backend == "s3":
        key = _s3_mapping_template_prefix(squad, slug, "template.json")
        try:
            return _read_json_s3(key)
        except FileNotFoundError:
            return None
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def save_mapping_template(
    project: ProjectRef,
    name: str,
    entries: dict,
    slug: str = "",
    description: str = "",
    metadata: dict | None = None,
) -> MappingTemplateRef:
    squad = normalize_squad(project.squad)
    backend = storage_backend()
    clean_entries = _normalize_mapping_entries(entries)
    template_slug = slugify(slug or name or project.name)
    existing = load_mapping_template(squad, template_slug)
    if not slug:
        base_slug = template_slug
        counter = 2
        while existing:
            template_slug = f"{base_slug}-{counter}"
            existing = load_mapping_template(squad, template_slug)
            counter += 1
    now = utc_now()
    origin_project = (existing or {}).get("origin_project") or {
        "squad": project.squad,
        "slug": project.slug,
        "name": project.name,
    }
    payload = {
        "schema_version": 1,
        "id": str((existing or {}).get("id") or uuid.uuid4()),
        "squad": squad,
        "slug": template_slug,
        "name": (name or (existing or {}).get("name") or project.name).strip(),
        "description": (description or (existing or {}).get("description") or "").strip(),
        "created_at": str((existing or {}).get("created_at") or now),
        "updated_at": now,
        "origin_project": origin_project,
        "last_project": {
            "squad": project.squad,
            "slug": project.slug,
            "name": project.name,
        },
        "entry_count": len(clean_entries),
        "entries": clean_entries,
        "metadata": {
            **((existing or {}).get("metadata") or {}),
            **(metadata or {}),
        },
    }
    if backend == "local":
        _write_json_file(data_root() / "squads" / squad / "mapping_templates" / template_slug / "template.json", payload)
    elif backend == "s3":
        _write_json_s3(_s3_mapping_template_prefix(squad, template_slug, "template.json"), payload)
    else:
        raise ValueError(f"Backend de storage nao suportado: {backend}")
    return _mapping_template_from_payload(payload, backend)


def create_run(project: ProjectRef, metadata: dict | None = None) -> RunRef:
    created_at = utc_now()
    stamp = datetime.now(UTC).strftime("%Y%m%d_%H%M%S")
    run = RunRef(run_id=f"{stamp}_{uuid.uuid4().hex[:8]}", created_at=created_at)
    payload = {
        "run_id": run.run_id,
        "created_at": created_at,
        "project": {
            "squad": project.squad,
            "slug": project.slug,
            "name": project.name,
        },
        "metadata": metadata or {},
    }
    save_project_json(project, ["runs", run.run_id], "run.json", payload)
    _touch_project(project)
    return run


def save_project_bytes(project: ProjectRef, parts: list[str], filename: str, data: bytes) -> str:
    filename = safe_filename(filename)
    backend = storage_backend()
    if backend == "local":
        path = _local_project_path(project, parts, filename)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return str(path)
    if backend == "s3":
        key = _s3_project_prefix(project.squad, project.slug, *parts, filename)
        _s3_client().put_object(Bucket=_s3_bucket(required=True), Key=key, Body=data)
        return f"s3://{_s3_bucket(required=True)}/{key}"
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def save_project_json(project: ProjectRef, parts: list[str], filename: str, payload: dict | list) -> str:
    data = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    return save_project_bytes(project, parts, filename, data)


def load_project_bytes(project: ProjectRef, parts: list[str], filename: str) -> bytes:
    filename = safe_filename(filename)
    backend = storage_backend()
    if backend == "local":
        path = _local_project_path(project, parts, filename)
        if not path.exists():
            raise FileNotFoundError(str(path))
        return path.read_bytes()
    if backend == "s3":
        key = _s3_project_prefix(project.squad, project.slug, *parts, filename)
        client = _s3_client()
        try:
            response = client.get_object(Bucket=_s3_bucket(required=True), Key=key)
        except client.exceptions.NoSuchKey as exc:
            raise FileNotFoundError(key) from exc
        return response["Body"].read()
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def load_project_json(project: ProjectRef, parts: list[str], filename: str):
    return json.loads(load_project_bytes(project, parts, filename).decode("utf-8"))


def append_memory_correction(project: ProjectRef, correction: dict) -> str:
    corrections = load_memory_corrections(project)
    corrections.append(correction)
    return save_project_json(project, ["memory"], "corrections.json", corrections)


def load_memory_corrections(project: ProjectRef) -> list[dict]:
    backend = storage_backend()
    if backend == "local":
        path = _local_project_path(project, ["memory"], "corrections.json")
        if not path.exists():
            return []
        return _read_json_file(path)
    if backend == "s3":
        key = _s3_project_prefix(project.squad, project.slug, "memory", "corrections.json")
        try:
            return _read_json_s3(key)
        except FileNotFoundError:
            return []
    raise ValueError(f"Backend de storage nao suportado: {backend}")


def _project_from_meta(meta: dict, backend: str) -> ProjectRef:
    return ProjectRef(
        project_id=str(meta.get("id") or ""),
        squad=str(meta.get("squad") or ""),
        slug=str(meta.get("slug") or ""),
        name=str(meta.get("name") or meta.get("slug") or ""),
        description=str(meta.get("description") or ""),
        created_at=str(meta.get("created_at") or ""),
        updated_at=str(meta.get("updated_at") or ""),
        backend=backend,
    )


def _mapping_template_from_payload(payload: dict, backend: str) -> MappingTemplateRef:
    origin_project = payload.get("origin_project") or payload.get("last_project") or {}
    return MappingTemplateRef(
        template_id=str(payload.get("id") or ""),
        squad=str(payload.get("squad") or ""),
        slug=str(payload.get("slug") or ""),
        name=str(payload.get("name") or payload.get("slug") or ""),
        description=str(payload.get("description") or ""),
        created_at=str(payload.get("created_at") or ""),
        updated_at=str(payload.get("updated_at") or ""),
        entry_count=int(payload.get("entry_count") or len(payload.get("entries") or {})),
        source_project_slug=str(origin_project.get("slug") or ""),
        source_project_name=str(origin_project.get("name") or ""),
        backend=backend,
    )


def _normalize_mapping_entries(entries: dict) -> dict:
    output = {}
    for target_id, entry in (entries or {}).items():
        clean_target_id = str(target_id or "").strip()
        if not clean_target_id:
            continue
        if isinstance(entry, dict):
            clean_entry = {
                str(key): value
                for key, value in entry.items()
                if isinstance(value, (str, int, float, bool, list, dict)) or value is None
            }
        else:
            clean_entry = {"datasource": str(entry or "").strip()}
        clean_entry["target_id"] = clean_target_id
        output[clean_target_id] = clean_entry
    return output


def _touch_project(project: ProjectRef) -> None:
    meta = {
        "squad": project.squad,
        "id": project.project_id,
        "slug": project.slug,
        "name": project.name,
        "description": project.description,
        "created_at": project.created_at,
        "updated_at": utc_now(),
        "backend": storage_backend(),
        "schema_version": 1,
    }
    if storage_backend() == "local":
        _write_json_file(data_root() / "squads" / project.squad / "projects" / project.slug / "project.json", meta)
    else:
        _write_json_s3(_s3_project_prefix(project.squad, project.slug, "project.json"), meta)


def _local_project_path(project: ProjectRef, parts: list[str], filename: str) -> Path:
    return data_root() / "squads" / project.squad / "projects" / project.slug / Path(*parts) / filename


def _read_json_file(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json_file(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _s3_bucket(required: bool = False) -> str:
    bucket = os.getenv("AUTO_PPT_S3_BUCKET", "").strip()
    if required and not bucket:
        raise ValueError("AUTO_PPT_S3_BUCKET nao configurado.")
    return bucket


def _s3_project_prefix(squad: str, slug: str, *parts: str) -> str:
    segments = [storage_prefix(), "squads", squad, "projects"]
    if slug:
        segments.append(slug)
    segments.extend(part.strip("/") for part in parts if part)
    return "/".join(segment for segment in segments if segment)


def _s3_mapping_template_prefix(squad: str, slug: str, *parts: str) -> str:
    segments = [storage_prefix(), "squads", squad, "mapping_templates"]
    if slug:
        segments.append(slug)
    segments.extend(part.strip("/") for part in parts if part)
    return "/".join(segment for segment in segments if segment)


def _s3_client():
    import boto3

    return boto3.client("s3")


def _read_json_s3(key: str):
    client = _s3_client()
    try:
        response = client.get_object(Bucket=_s3_bucket(required=True), Key=key)
    except client.exceptions.NoSuchKey as exc:
        raise FileNotFoundError(key) from exc
    return json.loads(response["Body"].read().decode("utf-8"))


def _write_json_s3(key: str, payload) -> None:
    _s3_client().put_object(
        Bucket=_s3_bucket(required=True),
        Key=key,
        Body=json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        ContentType="application/json; charset=utf-8",
    )
