from __future__ import annotations

from pathlib import Path
import os
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ppt_automator import analyze_update_package
from ppt_automator.ai import build_openai_client, format_ai_error, load_local_env
from ppt_automator.ai_transform import suggest_transform_diagnostics


def main() -> int:
    root = ROOT
    load_local_env(root)
    print(f"Repo: {root}")
    print(f"OPENAI_API_KEY: {'ok' if os.getenv('OPENAI_API_KEY') else 'ausente'}")
    print(f"OPENAI_MODEL: {os.getenv('OPENAI_MODEL', 'gpt-5.5')}")
    try:
        client, model = build_openai_client(root)
        response = client.responses.create(
            model=model,
            store=False,
            input="Responda somente: ok",
            max_output_tokens=16,
        )
    except Exception as exc:
        print(f"OpenAI: falhou - {format_ai_error(exc)}")
        return 1
    text = (getattr(response, "output_text", "") or "").strip()
    print(f"OpenAI: ok - {text[:80]}")
    if "--diagnostics-sample" in sys.argv:
        return _run_diagnostics_sample(root)
    return 0


def _run_diagnostics_sample(root: Path) -> int:
    sample_dir = Path(os.getenv("AUTO_PPT_ANDRE_TEST_DIR", r"C:\Users\HugoRocha\Documents\automatizador-ppt-arquivos\andre"))
    ppt = sample_dir / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
    datasources = sample_dir / "datasources.zip"
    if not ppt.exists() or not datasources.exists():
        print(f"Diagnostico real: pulado, arquivos nao encontrados em {sample_dir}")
        return 0
    try:
        _targets, _sources, plans = analyze_update_package(ppt, datasources)
        plan = next((item for item in plans if item.target_id == "7792738590"), plans[0])
        diagnostics = suggest_transform_diagnostics([plan], root=root)
    except Exception as exc:
        print(f"Diagnostico real: falhou - {format_ai_error(exc)}")
        return 1
    if not diagnostics:
        print("Diagnostico real: falhou - a IA nao retornou diagnostico.")
        return 1
    item = diagnostics[0]
    print(f"Diagnostico real: ok - target={item.target} status={item.status} confidence={item.confidence:.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
