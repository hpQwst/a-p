from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from pathlib import Path
import re
import unicodedata

import pandas as pd
import streamlit as st

from ppt_automator.ai import AiUnavailableError, ai_configured, suggest_datasources_with_ai
from ppt_automator import (
    build_auto_chart_jobs,
    build_chart_job,
    build_chart_jobs,
    generate_pptx,
    load_datasource_tables,
    load_mapping,
    load_ppt_targets,
    read_source_table_from_workbook,
)
from ppt_automator.project_store import (
    SQUADS,
    append_memory_correction,
    create_project,
    create_run,
    ensure_store,
    list_projects,
    load_memory_corrections,
    load_project,
    safe_filename,
    save_project_bytes,
    save_project_json,
    storage_backend,
)


ROOT = Path(__file__).parent
DEFAULT_PPT = ROOT / "Natura_2Q26_RelacionalCB_modelo_mapeado.pptx"
DEFAULT_MAPPING = ROOT / "Natura_2Q26_RelacionalCB_modelo_mapeamento.xlsx"
DEFAULT_DATASOURCES = ROOT / "datasources.zip"


st.set_page_config(page_title="Automatizador de PPT", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.4rem; }
    [data-testid="stMetricValue"] { font-size: 1.45rem; }
    div[data-testid="stDataFrame"] { border: 1px solid #e6e8ec; }
    .status-ok { color: #217346; font-weight: 700; }
    .status-warn { color: #9a5b00; font-weight: 700; }
    .flow-note {
        color: #4b5563;
        font-size: 0.95rem;
        line-height: 1.45;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Automatizador de PowerPoint")
st.markdown(
    "<div class='flow-note'>Siga as etapas em ordem: escolha o projeto, prepare os arquivos, revise os "
    "mapeamentos, confira os dados, valide as correspondências e gere o PPT final.</div>",
    unsafe_allow_html=True,
)

try:
    store_location = ensure_store()
except Exception as exc:
    st.error(f"Nao consegui preparar o armazenamento do app: {exc}")
    st.stop()

with st.sidebar:
    st.header("Etapa 0")
    st.caption("Escolha o squad e o projeto. Cada execucao fica salva sem sobrescrever historico.")
    squad_label = st.selectbox("Squad", [squad.title() for squad in SQUADS], key="squad_label")
    selected_squad = squad_label.lower()
    projects = list_projects(selected_squad)
    project_labels = {"": "Selecione ou crie um projeto"}
    project_labels.update({project.slug: project.name for project in projects})
    project_options = list(project_labels)
    selected_project_memory = st.session_state.setdefault("selected_project_by_squad", {})
    default_project_slug = selected_project_memory.get(selected_squad, "")
    if default_project_slug not in project_options:
        default_project_slug = ""
    selected_project_slug = st.selectbox(
        "Projeto",
        project_options,
        index=project_options.index(default_project_slug),
        format_func=lambda slug: project_labels.get(slug, slug),
        key=f"project_select_{selected_squad}",
    )
    if selected_project_slug:
        selected_project_memory[selected_squad] = selected_project_slug
    with st.expander("Criar novo projeto"):
        with st.form("create_project_form", clear_on_submit=True):
            new_project_name = st.text_input("Nome do projeto")
            new_project_description = st.text_area("Descricao curta", height=80)
            submitted_project = st.form_submit_button("Criar projeto")
            if submitted_project:
                if not new_project_name.strip():
                    st.warning("Informe um nome para o projeto.")
                else:
                    project = create_project(selected_squad, new_project_name, new_project_description)
                    st.session_state["selected_project_by_squad"][selected_squad] = project.slug
                    st.success("Projeto criado.")
                    st.rerun()

    selected_project = load_project(selected_squad, selected_project_slug) if selected_project_slug else None
    if selected_project:
        st.success(f"{selected_project.name} ativo")
        st.caption(f"Storage: {storage_backend()} | {store_location}")
        try:
            st.caption(f"Memoria do projeto: {len(load_memory_corrections(selected_project))} correcao(oes)")
        except Exception:
            st.caption("Memoria do projeto indisponivel no momento.")
    else:
        st.info("Crie ou selecione um projeto para continuar.")

    if selected_project:
        st.divider()
        st.header("Etapa 1")
        st.caption("Escolha o modo de trabalho e envie os arquivos de entrada.")
        workflow_mode = st.radio(
            "Modo",
            ["Com planilha de mapeamento", "Automático"],
            help="No modo automatico, o app tenta descobrir quais XLSX entram em quais graficos do PPT.",
        )
        ppt_file = st.file_uploader("PowerPoint modelo", type=["pptx"])
        mapping_file = st.file_uploader(
            "Planilha de mapeamento",
            type=["xlsx"],
            disabled=workflow_mode == "Automático",
        )
        datasource_file = st.file_uploader("ZIP com datasources", type=["zip"])
        use_examples = st.toggle("Usar arquivos desta pasta", value=True)
        st.divider()
        st.header("Opções")
        st.caption("Ajustes que afetam leitura, mapeamento e validação.")
        st.caption("Fórmulas do Excel serão sempre calculadas antes da leitura dos dados.")
        auto_match_sources = st.toggle("Encontrar datasources automaticamente", value=True)
        ai_ready = ai_configured(ROOT)
        use_ai_mapping = st.toggle(
            "Usar IA no mapeamento",
            value=ai_ready,
            disabled=not ai_ready,
            help="A IA revisa os pares grafico/datasource e preenche a escolha sugerida.",
        )
        if not ai_ready:
            st.caption("IA opcional: crie um .env com OPENAI_API_KEY para habilitar.")
        respect_update_flag = st.toggle("Respeitar coluna atualizar_grafico", value=False)
    else:
        workflow_mode = "Com planilha de mapeamento"
        ppt_file = None
        mapping_file = None
        datasource_file = None
        use_examples = False
        auto_match_sources = True
        use_ai_mapping = False
        respect_update_flag = False


def file_bytes(uploaded, fallback: Path) -> bytes | None:
    if uploaded is not None:
        return uploaded.getvalue()
    if use_examples and fallback.exists():
        return fallback.read_bytes()
    return None


def uploaded_source_table(uploaded, graph_id: str):
    return read_source_table_from_workbook(
        uploaded.getvalue(),
        f"upload_manual/{graph_id}_{uploaded.name}",
        formula_mode="auto",
        source_key=f"manual:{graph_id}:{uploaded.name}",
    )


def short_hash(data: bytes | None) -> str:
    return sha256(data or b"").hexdigest()[:16]


def render_flow_steps(steps: list[tuple[str, str, str]]) -> None:
    cols = st.columns(len(steps))
    for col, (title, detail, status) in zip(cols, steps):
        if status == "done":
            badge = "Concluida"
        elif status == "active":
            badge = "Atual"
        elif status == "blocked":
            badge = "Pendente"
        else:
            badge = "Proxima"
        col.markdown(f"**{title}**")
        col.caption(badge)
        col.write(detail)


def selected_jobs_key(selected_jobs) -> str:
    parts = [
        f"{job.graph_id}:{job.source.file_name if job.source else ''}:{job.match_score:.4f}"
        for job in selected_jobs
    ]
    return "|".join(
        [
            short_hash(ppt_bytes),
            workflow_mode,
            formula_mode,
            *sorted(parts),
        ]
    )


def generated_output_for(selected_jobs) -> bytes:
    key = selected_jobs_key(selected_jobs)
    if st.session_state.get("generated_output_key") != key:
        st.session_state["generated_output"] = generate_pptx(ppt_bytes, selected_jobs)
        st.session_state["generated_output_key"] = key
    return st.session_state["generated_output"]


def norm_text(value) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def item_match_score(target, choices) -> float:
    target_norm = norm_text(target)
    if not target_norm:
        return 0.0
    scores = []
    for choice in choices:
        choice_norm = norm_text(choice)
        if not choice_norm:
            continue
        if target_norm == choice_norm:
            scores.append(1.0)
        elif target_norm in choice_norm or choice_norm in target_norm:
            scores.append(0.9)
        else:
            target_tokens = set(target_norm.split())
            choice_tokens = set(choice_norm.split())
            scores.append(len(target_tokens & choice_tokens) / max(len(target_tokens), 1))
    return max(scores, default=0.0)


def list_match_pct(targets, choices) -> float:
    clean_targets = [value for value in targets if norm_text(value)]
    if not clean_targets:
        return 0.0
    matched = sum(1 for value in clean_targets if item_match_score(value, choices) >= 0.68)
    return round(matched / len(clean_targets) * 100, 1)


def missing_items(targets, choices, limit: int = 6) -> str:
    missing = [str(value) for value in targets if norm_text(value) and item_match_score(value, choices) < 0.68]
    return ", ".join(missing[:limit])


def filled_values_pct(job) -> float:
    values = [value for row in job.values for value in row]
    if not values:
        return 0.0
    filled = sum(1 for value in values if value not in (None, ""))
    return round(filled / len(values) * 100, 1)


def validation_status(row) -> tuple[str, str]:
    alerts = []
    if row["linhas_ok_%"] < 70:
        alerts.append("linhas pouco compatíveis")
    if row["colunas_ok_%"] < 80:
        alerts.append("colunas pouco compatíveis")
    if row["valores_preenchidos_%"] < 70:
        alerts.append("muitos valores vazios")
    if row.get("alerta_ia"):
        alerts.append(str(row["alerta_ia"]))
    if row.get("confianca_ia") is not None and not pd.isna(row.get("confianca_ia")) and row["confianca_ia"] < 72:
        alerts.append("confiança baixa da IA")
    status = "Revisar" if alerts else "OK"
    return status, "; ".join(alerts) if alerts else "Sem alertas."


def chart_job_report(job) -> dict:
    return {
        "graph_id": job.graph_id,
        "status": job.status,
        "slide": job.target.slide_number if job.target else None,
        "grafico": job.mapping.var_analise if job.mapping else "",
        "abertura": job.mapping.abertura if job.mapping else "",
        "datasource": job.source.file_name if job.source else "",
        "match_score": round(job.match_score * 100, 2) if job.match_score else 0.0,
        "match_reason": job.match_reason,
        "linhas": job.rows,
        "colunas": job.headers,
        "contexto_slide": job.target.nearby_text if job.target else "",
    }


if selected_project is None:
    render_flow_steps(
        [
            ("Etapa 0", "Escolher squad e projeto", "active"),
            ("Etapa 1", "Enviar arquivos", "blocked"),
            ("Etapa 2", "Revisar mapeamento", "blocked"),
            ("Etapa 3", "Conferir dados", "blocked"),
            ("Etapa 4", "Validar correspondências", "blocked"),
            ("Etapa 5", "Gerar PPT", "blocked"),
        ]
    )
    st.subheader("Etapa 0 - Criar ou selecionar projeto")
    st.info(
        "Escolha um dos squads na lateral e crie um projeto. Depois disso o app salva cada execucao, "
        "os arquivos usados, o PPT gerado e as correcoes aplicadas, sem sobrescrever o historico."
    )
    st.stop()


ppt_bytes = file_bytes(ppt_file, DEFAULT_PPT)
mapping_bytes = None if workflow_mode == "Automático" else file_bytes(mapping_file, DEFAULT_MAPPING)
datasource_bytes = file_bytes(datasource_file, DEFAULT_DATASOURCES)

if not ppt_bytes or not datasource_bytes or (workflow_mode != "Automático" and not mapping_bytes):
    missing = []
    if not ppt_bytes:
        missing.append("PowerPoint modelo")
    if workflow_mode != "Automático" and not mapping_bytes:
        missing.append("Planilha de mapeamento")
    if not datasource_bytes:
        missing.append("ZIP com datasources")
    render_flow_steps(
        [
            ("Etapa 0", f"{selected_project.name}", "done"),
            ("Etapa 1", "Enviar arquivos", "active"),
            ("Etapa 2", "Revisar mapeamento", "blocked"),
            ("Etapa 3", "Conferir dados", "blocked"),
            ("Etapa 4", "Validar correspondências", "blocked"),
            ("Etapa 5", "Gerar PPT", "blocked"),
        ]
    )
    st.subheader("Etapa 1 - Preparar arquivos")
    if workflow_mode == "Automático":
        st.info("No modo Automático, envie o PPT modelo e o ZIP com os XLSX. A planilha de mapeamento não é obrigatória.")
    else:
        st.info("No modo com mapeamento, envie o PPT modelo, a planilha de mapeamento e o ZIP com os XLSX.")
    if missing:
        st.warning("Falta enviar: " + ", ".join(missing) + ".")
    st.stop()

try:
    formula_mode = "auto"
    if workflow_mode == "Automático":
        jobs = build_auto_chart_jobs(
            ppt_bytes,
            datasource_bytes,
            formula_mode=formula_mode,
        )
        mappings = [job.mapping for job in jobs if job.mapping]
    else:
        jobs = build_chart_jobs(
            ppt_bytes,
            mapping_bytes,
            datasource_bytes,
            formula_mode=formula_mode,
            respect_update_flag=respect_update_flag,
            auto_match_sources=auto_match_sources,
        )
        mappings = load_mapping(mapping_bytes, formula_mode=formula_mode)
    source_tables = load_datasource_tables(datasource_bytes, formula_mode=formula_mode)
    targets = load_ppt_targets(ppt_bytes)
except Exception as exc:
    st.error(f"Não consegui ler os arquivos: {exc}")
    st.stop()

manual_source_tables = st.session_state.setdefault("manual_source_tables", {})
manual_source_choice = st.session_state.setdefault("manual_source_choice", {})
source_tables = [*source_tables, *manual_source_tables.values()]

ok_jobs = [job for job in jobs if job.ok]
warning_jobs = [job for job in jobs if not job.ok]
needs_review_jobs = [job for job in jobs if not job.ok or (job.match_score and job.match_score < 0.8)]
mapping_by_graph = {mapping.graph_id: mapping for mapping in mappings}
target_by_graph = targets
source_by_file = {source.file_name: source for source in source_tables}
source_names = [""] + [source.file_name for source in source_tables]
job_by_graph = {job.graph_id: job for job in jobs}

ai_suggestions_by_graph = {}
ai_mapping_error = ""
if use_ai_mapping and jobs:
    ai_cache_key = "|".join(
        [
            workflow_mode,
            formula_mode,
            short_hash(ppt_bytes),
            short_hash(mapping_bytes),
            short_hash(datasource_bytes),
            ",".join(sorted(job.graph_id for job in jobs)),
        ]
    )
    if st.session_state.get("ai_mapping_cache_key") != ai_cache_key:
        with st.spinner("IA revisando o mapeamento dos graficos..."):
            try:
                suggestions = suggest_datasources_with_ai(jobs, source_tables, root=ROOT)
                st.session_state["ai_mapping_suggestions"] = [
                    {
                        "graph_id": suggestion.graph_id,
                        "datasource": suggestion.datasource,
                        "confidence": suggestion.confidence,
                        "reason": suggestion.reason,
                    }
                    for suggestion in suggestions
                ]
                st.session_state["ai_mapping_error"] = ""
            except AiUnavailableError as exc:
                st.session_state["ai_mapping_suggestions"] = []
                st.session_state["ai_mapping_error"] = str(exc)
            except Exception as exc:
                st.session_state["ai_mapping_suggestions"] = []
                st.session_state["ai_mapping_error"] = f"Falha ao consultar IA: {exc}"
            st.session_state["ai_mapping_cache_key"] = ai_cache_key
    ai_mapping_error = st.session_state.get("ai_mapping_error", "")
    ai_suggestions_by_graph = {
        item["graph_id"]: item for item in st.session_state.get("ai_mapping_suggestions", [])
    }

col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Mapeamentos", len(jobs))
col2.metric("Prontos", len(ok_jobs))
col3.metric("Pendências", len(warning_jobs))
col4.metric("Slides afetados", len({job.target.slide_number for job in ok_jobs if job.target}))
col5.metric("Revisar", len(needs_review_jobs))
col6.metric("IA", len(ai_suggestions_by_graph) if use_ai_mapping else 0)

render_flow_steps(
    [
        ("Etapa 0", f"{selected_project.name}", "done"),
        ("Etapa 1", "Arquivos carregados", "done"),
        ("Etapa 2", "Conferir correspondências", "active"),
        ("Etapa 3", "Conferir dados", "next"),
        ("Etapa 4", "Validar destino", "next"),
        ("Etapa 5", "Gerar PPT", "next"),
    ]
)

st.caption(
    "Os arquivos XLSX nao precisam ter o mesmo nome do graph_id: o sistema compara colunas, linhas, "
    "pergunta da tabela e metadados opcionais para sugerir o datasource certo."
)
if ai_mapping_error:
    st.warning(ai_mapping_error)
elif use_ai_mapping and ai_suggestions_by_graph:
    st.success("IA revisou o mapeamento e preencheu as sugestoes na tabela de conferencia.")

summary_rows = []
for job in jobs:
    mapping = job.mapping
    target = job.target
    source = job.source
    ai_suggestion = ai_suggestions_by_graph.get(job.graph_id)
    ai_source = ai_suggestion["datasource"] if ai_suggestion else ""
    ai_confidence = float(ai_suggestion["confidence"]) if ai_suggestion else 0.0
    ai_reason = ai_suggestion["reason"] if ai_suggestion else ""
    suggested_source = source.file_name if source else ""
    manual_choice = manual_source_choice.get(job.graph_id, "")
    default_choice = manual_choice or ai_source or suggested_source
    if not default_choice and job.match_candidates:
        default_choice = job.match_candidates[0][0]
    ai_disagrees = bool(ai_source and suggested_source and ai_source != suggested_source)
    ai_low_confidence = bool(ai_suggestion and ai_confidence < 0.72)
    review_alert = ""
    if ai_disagrees:
        review_alert = "IA divergiu da heuristica"
    if ai_low_confidence:
        review_alert = "Baixa confianca da IA" if not review_alert else f"{review_alert}; baixa confianca"
    default_use = bool(manual_choice) or (
        job.ok and not ai_low_confidence and not (ai_disagrees and ai_confidence < 0.85)
    )
    summary_rows.append(
        {
            "usar": default_use,
            "status": job.status,
            "slide": target.slide_number if target else mapping.numero_slide if mapping else None,
            "graph_id": job.graph_id,
            "variavel": mapping.var_analise if mapping else "",
            "abertura": mapping.abertura if mapping else "",
            "datasource_sugerido": suggested_source,
            "datasource_ia": ai_source,
            "correcao_manual": manual_choice,
            "datasource_escolhido": default_choice,
            "confianca_ia": round(ai_confidence * 100, 1) if ai_suggestion else None,
            "alerta_ia": review_alert,
            "score": round(job.match_score * 100, 1) if job.match_score else None,
            "match": job.match_reason,
            "motivo_ia": ai_reason,
            "candidatos": ", ".join(
                f"{Path(name).name} ({score:.0%})" for name, score, _reason in job.match_candidates[:3]
            ),
            "contexto_slide": target.nearby_text[:240] if target else "",
            "destino": target.chart_path if target else "",
            "mensagem": job.message,
        }
    )

summary_df = pd.DataFrame(summary_rows)

tab_review, tab_preview, tab_validate, tab_generate = st.tabs(
    ["Etapa 2 - Mapeamento", "Etapa 3 - Dados", "Etapa 4 - Validação", "Etapa 5 - Gerar PPT"]
)

with tab_review:
    st.subheader("Etapa 2 - Revisar mapeamento")
    st.write(
        "Confira qual datasource será usado em cada gráfico. Ajuste o campo "
        "`Datasource escolhido` quando a sugestão não estiver correta e desmarque itens que não devem entrar no PPT."
    )
    edited = st.data_editor(
        summary_df,
        use_container_width=True,
        hide_index=True,
        disabled=[c for c in summary_df.columns if c not in {"usar", "datasource_escolhido"}],
        column_config={
            "usar": st.column_config.CheckboxColumn("Usar"),
            "status": st.column_config.TextColumn("Status"),
            "slide": st.column_config.NumberColumn("Slide", format="%d"),
            "graph_id": st.column_config.TextColumn("ID"),
            "datasource_sugerido": st.column_config.TextColumn("Datasource sugerido", width="medium"),
            "datasource_ia": st.column_config.TextColumn("Datasource IA", width="medium"),
            "correcao_manual": st.column_config.TextColumn("Correção manual", width="medium"),
            "datasource_escolhido": st.column_config.SelectboxColumn(
                "Datasource escolhido",
                options=source_names,
                width="medium",
            ),
            "confianca_ia": st.column_config.NumberColumn("Conf. IA", format="%.1f%%"),
            "alerta_ia": st.column_config.TextColumn("Alerta IA", width="medium"),
            "score": st.column_config.NumberColumn("Score", format="%.1f%%"),
            "match": st.column_config.TextColumn("Match", width="large"),
            "motivo_ia": st.column_config.TextColumn("Motivo IA", width="large"),
            "candidatos": st.column_config.TextColumn("Candidatos", width="large"),
            "contexto_slide": st.column_config.TextColumn("Contexto no slide", width="large"),
            "mensagem": st.column_config.TextColumn("Mensagem", width="large"),
        },
    )
    with st.expander("Corrigir um gráfico usando outro XLSX"):
        st.write(
            "Use isto quando a sugestão estiver errada ou quando o arquivo correto não estiver no ZIP. "
            "Escolha o gráfico, envie o XLSX correto e o app passará a usar esse arquivo apenas nesse gráfico."
        )
        fix_graph_id = st.selectbox(
            "Gráfico que será corrigido",
            [job.graph_id for job in jobs],
            format_func=lambda graph_id: next(
                (
                    f"Slide {job.target.slide_number if job.target else '?'} - {graph_id} - "
                    f"{job.mapping.var_analise if job.mapping else ''}"
                    for job in jobs
                    if job.graph_id == graph_id
                ),
                graph_id,
            ),
            key="manual_fix_graph_id",
        )
        manual_xlsx = st.file_uploader(
            "XLSX correto para este gráfico",
            type=["xlsx"],
            key="manual_fix_xlsx",
        )
        if st.button("Usar este XLSX neste gráfico", disabled=manual_xlsx is None):
            try:
                table = uploaded_source_table(manual_xlsx, fix_graph_id)
                saved_source = save_project_bytes(
                    selected_project,
                    ["memory", "manual_sources"],
                    f"{fix_graph_id}_{manual_xlsx.name}",
                    manual_xlsx.getvalue(),
                )
                append_memory_correction(
                    selected_project,
                    {
                        "created_at": datetime.now().isoformat(timespec="seconds"),
                        "graph_id": fix_graph_id,
                        "uploaded_file": manual_xlsx.name,
                        "saved_source": saved_source,
                        "detected_question": table.question,
                        "detected_headers": table.headers,
                        "detected_rows": table.rows[:20],
                    },
                )
                st.session_state["manual_source_tables"][table.file_name] = table
                st.session_state["manual_source_choice"][fix_graph_id] = table.file_name
                st.session_state.pop("generated_output_key", None)
                st.success("Correção aplicada. Recarregando o mapeamento...")
                st.rerun()
            except Exception as exc:
                st.error(f"Não consegui ler esse XLSX: {exc}")


def build_review_jobs(edited_df: pd.DataFrame):
    def confidence_value(value) -> float:
        if value is None or pd.isna(value):
            return 0.0
        return float(value) / 100

    selected = []
    invalid = []
    for _, row in edited_df.iterrows():
        if not bool(row.get("usar")):
            continue
        graph_id = str(row.get("graph_id") or "")
        source_file = str(row.get("datasource_escolhido") or "")
        mapping = mapping_by_graph.get(graph_id)
        target = target_by_graph.get(graph_id)
        source = source_by_file.get(source_file)
        if mapping is None or target is None or source is None:
            invalid.append(graph_id)
            continue
        existing = job_by_graph.get(graph_id)
        if existing and existing.ok and existing.source and existing.source.file_name == source_file:
            if str(row.get("datasource_ia") or "") == source_file:
                selected.append(
                    build_chart_job(
                        mapping,
                        source,
                        target,
                        match_score=confidence_value(row.get("confianca_ia")),
                        match_reason=f"Datasource confirmado pela IA: {row.get('motivo_ia') or ''}",
                    )
                )
            else:
                selected.append(existing)
            continue
        reason = "Datasource escolhido manualmente na revisão."
        score = 0.0
        if str(row.get("datasource_ia") or "") == source_file:
            reason = f"Datasource escolhido pela IA: {row.get('motivo_ia') or ''}"
            score = confidence_value(row.get("confianca_ia"))
        selected.append(
            build_chart_job(
                mapping,
                source,
                target,
                match_score=score,
                match_reason=reason,
            )
        )
    return selected, invalid


selected_jobs, invalid_selected_graphs = build_review_jobs(edited)
if invalid_selected_graphs:
    st.warning(
        "Alguns itens marcados para uso ainda nao tem grafico ou datasource valido: "
        + ", ".join(invalid_selected_graphs)
    )


with tab_preview:
    st.subheader("Etapa 3 - Conferir dados")
    st.write(
        "Escolha um gráfico e confira a matriz que será gravada no PowerPoint. "
        "Essa etapa ajuda a pegar troca de abertura, linha faltando ou valor estranho antes da geração."
    )
    if not selected_jobs:
        st.info("Ainda nao ha graficos prontos para pre-visualizar.")
    else:
        selected_id = st.selectbox(
            "Gráfico",
            [job.graph_id for job in selected_jobs],
            format_func=lambda graph_id: next(
                f"Slide {job.target.slide_number} - {graph_id} - {job.mapping.var_analise}"
                for job in selected_jobs
                if job.graph_id == graph_id
            ),
        )
        selected = next(job for job in selected_jobs if job.graph_id == selected_id)
        left, right = st.columns([1.1, 1])
        with left:
            st.caption("Dados que serão gravados no gráfico")
            st.dataframe(
                pd.DataFrame(selected.values, index=selected.rows, columns=selected.headers),
                use_container_width=True,
            )
        with right:
            st.caption("Destino no PowerPoint")
            st.write(
                {
                    "slide": selected.target.slide_number,
                    "contexto_slide": selected.target.nearby_text,
                    "chart": selected.target.chart_path,
                    "workbook_embutido": selected.target.embedded_workbook_path,
                    "posicao_pol": (
                        selected.target.left_in,
                        selected.target.top_in,
                        selected.target.width_in,
                        selected.target.height_in,
                    ),
                    "datasource": selected.source.file_name if selected.source else "",
                    "score_match": round(selected.match_score * 100, 1),
                }
            )
            if selected.source and selected.source.respondents:
                st.caption("Base / respondentes")
                st.dataframe(pd.DataFrame([selected.source.respondents]), use_container_width=True)

with tab_validate:
    st.subheader("Etapa 4 - Validar correspondências")
    st.write(
        "Confira o resumo final de destino: qual arquivo alimenta qual gráfico, em qual slide, "
        "e se as linhas/colunas do XLSX combinam com o que o gráfico espera."
    )
    if not selected_jobs:
        st.info("Ainda nao ha graficos prontos para validar.")
    else:
        validation_rows = []
        for job in selected_jobs:
            source = job.source
            target = job.target
            row = {
                "status": "",
                "slide": target.slide_number if target else None,
                "graph_id": job.graph_id,
                "grafico": job.mapping.var_analise if job.mapping else "",
                "arquivo_usado": source.file_name if source else "",
                "pergunta_xlsx": source.question if source else "",
                "contexto_slide": target.nearby_text if target else "",
                "linhas_ok_%": list_match_pct(target.rows if target else [], source.rows if source else []),
                "colunas_ok_%": list_match_pct(target.headers if target else [], source.headers if source else []),
                "valores_preenchidos_%": filled_values_pct(job),
                "linhas_nao_encontradas": missing_items(target.rows if target else [], source.rows if source else []),
                "colunas_nao_encontradas": missing_items(target.headers if target else [], source.headers if source else []),
                "motivo": job.match_reason,
            }
            status, alertas = validation_status(row)
            row["status"] = status
            row["alertas"] = alertas
            validation_rows.append(row)
        validation_df = pd.DataFrame(validation_rows)
        review_count = int((validation_df["status"] == "Revisar").sum()) if not validation_df.empty else 0
        if review_count:
            st.warning(f"{review_count} correspondência(s) precisam de revisão antes de gerar o PPT.")
        else:
            st.success("Todas as correspondências selecionadas passaram na validação lógica.")
        st.dataframe(validation_df, use_container_width=True, hide_index=True)

with tab_generate:
    st.subheader("Etapa 5 - Gerar PowerPoint")
    st.write(
        "Depois de revisar mapeamento, dados e validação de destino, gere o arquivo final."
    )
    st.write(f"{len(selected_jobs)} gráficos serão atualizados.")
    if st.button("Gerar PowerPoint atualizado", type="primary", disabled=not selected_jobs):
        with st.spinner("Gerando PowerPoint atualizado..."):
            output = generated_output_for(selected_jobs)
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        file_name = f"ppt_automatizado_{stamp}.pptx"
        run = create_run(
            selected_project,
            {
                "workflow_mode": workflow_mode,
                "graphs_updated": len(selected_jobs),
                "ai_mapping": bool(use_ai_mapping),
                "formula_mode": "auto",
            },
        )
        input_manifest = []
        ppt_name = safe_filename(ppt_file.name if ppt_file else DEFAULT_PPT.name)
        input_manifest.append(
            {
                "type": "ppt_template",
                "name": ppt_name,
                "saved_to": save_project_bytes(selected_project, ["runs", run.run_id, "inputs"], ppt_name, ppt_bytes),
            }
        )
        if mapping_bytes:
            mapping_name = safe_filename(mapping_file.name if mapping_file else DEFAULT_MAPPING.name)
            input_manifest.append(
                {
                    "type": "mapping",
                    "name": mapping_name,
                    "saved_to": save_project_bytes(
                        selected_project,
                        ["runs", run.run_id, "inputs"],
                        mapping_name,
                        mapping_bytes,
                    ),
                }
            )
        datasource_name = safe_filename(datasource_file.name if datasource_file else DEFAULT_DATASOURCES.name)
        input_manifest.append(
            {
                "type": "datasources_zip",
                "name": datasource_name,
                "saved_to": save_project_bytes(
                    selected_project,
                    ["runs", run.run_id, "inputs"],
                    datasource_name,
                    datasource_bytes,
                ),
            }
        )
        output_location = save_project_bytes(
            selected_project,
            ["runs", run.run_id, "outputs"],
            file_name,
            output,
        )
        report_location = save_project_json(
            selected_project,
            ["runs", run.run_id, "reports"],
            "execution_report.json",
            {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "project": {
                    "squad": selected_project.squad,
                    "slug": selected_project.slug,
                    "name": selected_project.name,
                },
                "inputs": input_manifest,
                "output": output_location,
                "jobs": [chart_job_report(job) for job in selected_jobs],
            },
        )
        st.success(f"Execução salva no projeto: {run.run_id}")
        st.caption(f"Output: {output_location}")
        st.caption(f"Relatorio: {report_location}")
        st.download_button(
            "Baixar PPT atualizado",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
