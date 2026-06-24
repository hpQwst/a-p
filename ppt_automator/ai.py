from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
from typing import Any, Iterable, Sequence

from .core import ChartJob, SourceTable


@dataclass(frozen=True)
class AiMatchSuggestion:
    graph_id: str
    datasource: str
    confidence: float
    reason: str


class AiUnavailableError(RuntimeError):
    pass


def build_openai_client(root: Path | str | None = None) -> tuple[Any, str]:
    load_local_env(root)
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise AiUnavailableError("OPENAI_API_KEY nao configurada.")

    try:
        from openai import OpenAI
    except Exception as exc:
        raise AiUnavailableError("Pacote openai nao instalado. Rode pip install -r requirements.txt.") from exc

    model = os.getenv("OPENAI_MODEL", "gpt-5.5")
    base_url = os.getenv("OPENAI_BASE_URL")
    timeout = _env_float("OPENAI_TIMEOUT_SECONDS", 120.0)
    max_retries = _env_int("OPENAI_MAX_RETRIES", 1)
    client_kwargs: dict[str, Any] = {
        "api_key": api_key,
        "timeout": timeout,
        "max_retries": max_retries,
    }
    if base_url:
        client_kwargs["base_url"] = base_url
    return OpenAI(**client_kwargs), model


def format_ai_error(exc: Exception) -> str:
    chain = _exception_chain(exc)
    chain_text = " | ".join(str(item).strip() for item in chain if str(item).strip())
    if "[WinError 10013]" in chain_text:
        return (
            "Conexao com a OpenAI bloqueada pelo Windows, firewall, proxy corporativo ou sandbox "
            "([WinError 10013]). Rode o app pelo PowerShell normal ou libere o python.exe na rede."
        )
    if "CERTIFICATE_VERIFY_FAILED" in chain_text:
        return "Falha de certificado SSL ao conectar na OpenAI. Verifique proxy corporativo/inspecao SSL."
    if "invalid_api_key" in chain_text.lower() or "incorrect api key" in chain_text.lower():
        return "OPENAI_API_KEY invalida. Confira a chave no .env."
    if "model" in chain_text.lower() and ("not found" in chain_text.lower() or "does not exist" in chain_text.lower()):
        return "Modelo OpenAI nao disponivel para essa chave. Confira OPENAI_MODEL no .env."
    if chain_text and str(exc).strip() == "Connection error." and len(chain) > 1:
        return f"Erro de conexao com a OpenAI: {chain_text}"
    return str(exc).strip() or type(exc).__name__


def load_local_env(root: Path | str | None = None) -> None:
    env_path = Path(root or Path.cwd()) / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip().lstrip("\ufeff")
        value = value.strip().strip('"').strip("'")
        current = os.environ.get(key)
        if key and value and (current is None or not current.strip()):
            os.environ[key] = value


def ai_configured(root: Path | str | None = None) -> bool:
    load_local_env(root)
    return bool(os.getenv("OPENAI_API_KEY"))


def suggest_datasource_with_ai(
    job: ChartJob,
    sources: Iterable[SourceTable],
    root: Path | str | None = None,
) -> AiMatchSuggestion:
    client, model = build_openai_client(root)

    prompt_payload = {
        "task": "Escolha qual datasource XLSX deve alimentar o grafico do PowerPoint.",
        "graph": {
            "graph_id": job.graph_id,
            "slide": job.target.slide_number if job.target else None,
            "nearby_text": job.target.nearby_text if job.target else "",
            "slide_text": job.target.slide_text if job.target else "",
            "headers": job.target.headers if job.target else [],
            "rows": job.target.rows if job.target else [],
            "mapping_variable": job.mapping.var_analise if job.mapping else "",
            "mapping_break": job.mapping.abertura if job.mapping else "",
            "current_reason": job.match_reason,
            "current_candidates": job.match_candidates,
        },
        "datasources": [
            {
                "file_name": source.file_name,
                "question": source.question,
                "headers": source.headers,
                "rows": source.rows[:20],
                "metadata": source.metadata,
            }
            for source in sources
        ],
        "rules": [
            "Prefira compatibilidade de linhas e colunas antes de nomes de arquivo.",
            "Use o texto proximo do grafico para entender a variavel.",
            "Se estiver incerto, escolha o melhor candidato e use confidence menor.",
            "Responda somente com JSON valido no schema pedido.",
        ],
    }

    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "graph_id": {"type": "string"},
            "datasource": {"type": "string"},
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "reason": {"type": "string"},
        },
        "required": ["graph_id", "datasource", "confidence", "reason"],
    }

    response = client.responses.create(
        model=model,
        store=False,
        input=[
            {
                "role": "system",
                "content": (
                    "Voce e um analista de automacao de PowerPoint. "
                    "Seu trabalho e escolher o XLSX correto para um grafico, sem alterar dados."
                ),
            },
            {"role": "user", "content": json.dumps(prompt_payload, ensure_ascii=False)},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "ppt_datasource_match",
                "schema": schema,
                "strict": True,
            }
        },
    )
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    data = json.loads(text)
    return AiMatchSuggestion(
        graph_id=str(data.get("graph_id") or job.graph_id),
        datasource=str(data.get("datasource") or ""),
        confidence=float(data.get("confidence") or 0),
        reason=str(data.get("reason") or ""),
    )


def suggest_datasources_with_ai(
    jobs: Sequence[ChartJob],
    sources: Iterable[SourceTable],
    root: Path | str | None = None,
) -> list[AiMatchSuggestion]:
    client, model = build_openai_client(root)

    sources_payload = [
        {
            "file_name": source.file_name,
            "question": source.question,
            "headers": source.headers,
            "rows": source.rows[:24],
            "metadata": source.metadata,
        }
        for source in sources
    ]
    graphs_payload = [
        {
            "graph_id": job.graph_id,
            "slide": job.target.slide_number if job.target else None,
            "nearby_text": job.target.nearby_text if job.target else "",
            "slide_text": job.target.slide_text if job.target else "",
            "headers": job.target.headers if job.target else [],
            "rows": (job.target.rows if job.target else [])[:24],
            "mapping_variable": job.mapping.var_analise if job.mapping else "",
            "mapping_break": job.mapping.abertura if job.mapping else "",
            "deterministic_datasource": job.source.file_name if job.source else "",
            "deterministic_score": job.match_score,
            "deterministic_reason": job.match_reason,
            "top_candidates": job.match_candidates[:5],
        }
        for job in jobs
        if job.target is not None
    ]

    prompt_payload = {
        "task": "Revise o mapeamento de datasources XLSX para graficos do PowerPoint.",
        "graphs": graphs_payload,
        "datasources": sources_payload,
        "rules": [
            "Escolha exatamente um datasource para cada graph_id.",
            "Prefira compatibilidade semantica entre texto do slide/pergunta do XLSX, mas confira linhas e colunas.",
            "Se a heuristica parece correta, mantenha deterministic_datasource.",
            "Se a heuristica parece errada, escolha outro datasource e explique.",
            "Use confidence alto somente quando texto, linhas e colunas contam a mesma historia.",
            "Nao invente nomes: datasource deve ser exatamente um file_name da lista.",
        ],
    }

    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "suggestions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "graph_id": {"type": "string"},
                        "datasource": {"type": "string"},
                        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                        "reason": {"type": "string"},
                    },
                    "required": ["graph_id", "datasource", "confidence", "reason"],
                },
            }
        },
        "required": ["suggestions"],
    }

    response = client.responses.create(
        model=model,
        store=False,
        input=[
            {
                "role": "system",
                "content": (
                    "Voce e um revisor de qualidade de automacao de PowerPoint. "
                    "Seu objetivo e reduzir erro de mapeamento, nao gerar conteudo novo."
                ),
            },
            {"role": "user", "content": json.dumps(prompt_payload, ensure_ascii=False)},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "ppt_datasource_batch_match",
                "schema": schema,
                "strict": True,
            }
        },
    )
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    data = json.loads(text)
    suggestions: list[AiMatchSuggestion] = []
    valid_sources = {str(source["file_name"]) for source in sources_payload}
    requested_graphs = {str(graph["graph_id"]) for graph in graphs_payload}
    for item in data.get("suggestions", []):
        graph_id = str(item.get("graph_id") or "")
        datasource = str(item.get("datasource") or "")
        if graph_id not in requested_graphs or datasource not in valid_sources:
            continue
        suggestions.append(
            AiMatchSuggestion(
                graph_id=graph_id,
                datasource=datasource,
                confidence=float(item.get("confidence") or 0),
                reason=str(item.get("reason") or ""),
            )
        )
    return suggestions


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)


def _env_float(name: str, default: float) -> float:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return float(value)
    except ValueError:
        return default


def _env_int(name: str, default: int) -> int:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _exception_chain(exc: Exception) -> list[BaseException]:
    chain: list[BaseException] = []
    current: BaseException | None = exc
    seen: set[int] = set()
    while current is not None and id(current) not in seen:
        chain.append(current)
        seen.add(id(current))
        current = current.__cause__ or current.__context__
    return chain
