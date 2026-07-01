from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import base64
import json

from .ai import build_openai_client


@dataclass(frozen=True)
class SlideUnderstandingInput:
    slide_number: int
    slide_image_path: Path | None
    visual_map: list[dict[str, str]]
    slide_text: str
    targets: list[dict[str, Any]]
    xlsx_dumps: list[str]
    manual_context: str = ""


def suggest_slide_understanding(payload: SlideUnderstandingInput, root: Path | str | None = None) -> dict[str, Any]:
    client, model = build_openai_client(root)
    user_payload = {
        "slide_number": payload.slide_number,
        "visual_target_map": payload.visual_map,
        "slide_text": payload.slide_text,
        "targets": payload.targets,
        "xlsx_plaintext_dumps": payload.xlsx_dumps,
        "manual_context": payload.manual_context,
        "instructions": [
            "Interprete o papel visual de cada target usando a imagem do slide e os IDs visuais.",
            "Interprete os XLSX somente pelo dump textual/JSON. Nao use imagem de planilha.",
            "Nao invente valores. Nesta etapa, apenas descreva partes uteis e como extrair.",
            "Se o arquivo nao tiver dados suficientes, retorne questions_for_user.",
        ],
    }
    content: list[dict[str, Any]] = [{"type": "input_text", "text": json.dumps(user_payload, ensure_ascii=False)}]
    if payload.slide_image_path and payload.slide_image_path.exists():
        content.insert(0, {"type": "input_image", "image_url": _image_data_url(payload.slide_image_path)})
    response = client.responses.create(
        model=model,
        store=False,
        input=[
            {
                "role": "system",
                "content": (
                    "Voce entende slides de PowerPoint e planilhas complexas. "
                    "Seu trabalho e criar um entendimento por slide, mapeando alvos visuais para partes uteis dos XLSX. "
                    "Use valores raw apenas como referencia textual; nao crie matriz final nesta etapa."
                ),
            },
            {"role": "user", "content": content},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "slide_understanding",
                "schema": _schema(),
                "strict": True,
            }
        },
    )
    text = getattr(response, "output_text", "") or _response_text_fallback(response)
    return json.loads(text)


def _schema() -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "slide_number": {"type": "integer"},
            "slide_understanding": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "title": {"type": "string"},
                    "main_metric": {"type": "string"},
                    "targets": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "target_id": {"type": "string"},
                                "visual_label": {"type": "string"},
                                "semantic_name": {"type": "string"},
                                "role": {"type": "string"},
                            },
                            "required": ["target_id", "visual_label", "semantic_name", "role"],
                        },
                    },
                },
                "required": ["title", "main_metric", "targets"],
            },
            "xlsx_understanding": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "file_name": {"type": "string"},
                        "meaning": {"type": "string"},
                        "usable_parts": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "part_id": {"type": "string"},
                                    "meaning": {"type": "string"},
                                    "how_to_extract": {"type": "string"},
                                },
                                "required": ["part_id", "meaning", "how_to_extract"],
                            },
                        },
                    },
                    "required": ["file_name", "meaning", "usable_parts"],
                },
            },
            "questions_for_user": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["slide_number", "slide_understanding", "xlsx_understanding", "questions_for_user"],
    }


def _image_data_url(path: Path) -> str:
    return "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")


def _response_text_fallback(response: Any) -> str:
    try:
        return response.output[0].content[0].text
    except Exception:
        return str(response)
