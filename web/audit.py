"""Quem fez o que.

A identidade sai da sessao: com login Microsoft e o e-mail da pessoa; com a
senha compartilhada nao existe pessoa, e o registro diz isso em vez de fingir
que sabe. Deixar explicito importa, porque um registro que parece nominal mas
nao e seria pior que nenhum.

As acoes ficam em `memory/corrections.json` do projeto, que ja existia para esse
proposito e ate agora so era escrito pelo Streamlit legado.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any
import json

from ppt_automator.project_store import append_memory_correction, load_project

SHARED_PASSWORD_ACTOR = "senha compartilhada (sem identificacao)"
ANONYMOUS_ACTOR = "acesso aberto (sem identificacao)"


def actor_from(cookies: dict, auth_module) -> str:
    """Nome de quem esta agindo, para gravar junto da acao."""
    if not auth_module.auth_enabled():
        return ANONYMOUS_ACTOR
    email = auth_module.current_user(cookies)
    return email or SHARED_PASSWORD_ACTOR


def is_identified(actor: str) -> bool:
    return bool(actor) and actor not in {SHARED_PASSWORD_ACTOR, ANONYMOUS_ACTOR}


def record(job_dir: Path, actor: str, action: str, details: dict[str, Any] | None = None) -> None:
    """Guarda a acao no projeto dono do job. Falha aqui nunca derruba a
    operacao do usuario: registro e importante, mas nao mais que o trabalho."""
    entry = {
        "at": datetime.now().isoformat(timespec="seconds"),
        "actor": actor or ANONYMOUS_ACTOR,
        "identified": is_identified(actor),
        "action": action,
        **(details or {}),
    }
    try:
        metadata = json.loads((job_dir / "metadata.json").read_text(encoding="utf-8"))
        project_meta = metadata.get("project") or {}
        entry.setdefault("job_id", metadata.get("job_id"))
        project = load_project(str(project_meta.get("squad") or ""), str(project_meta.get("slug") or ""))
        if project is None:
            return
        append_memory_correction(project, entry)
    except Exception:
        return


def remember_actor(job_dir: Path, actor: str) -> None:
    """Guarda quem mexeu por ultimo no job, para o trabalho em segundo plano
    (geracao do PPT) saber a quem atribuir."""
    try:
        path = job_dir / "metadata.json"
        metadata = json.loads(path.read_text(encoding="utf-8"))
        metadata["last_actor"] = actor
        path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        return


def remembered_actor(job_dir: Path) -> str:
    try:
        metadata = json.loads((job_dir / "metadata.json").read_text(encoding="utf-8"))
        return str(metadata.get("last_actor") or "")
    except Exception:
        return ""
