from __future__ import annotations


from typing import Any
from zipfile import ZipFile
import re
import unicodedata
import xml.etree.ElementTree as ET



from .ppt_discovery import CHART_NS, NS, PptTarget
from .table_normalizer import TransformPlan




PERCENT_FORMAT = "0%"
DECIMAL_FORMAT = "0.0"


def effective_value_format(template_format: str) -> str:
    """Preserva o contrato visual existente; General cai para decimal legivel."""
    fmt = (template_format or "").strip()
    return fmt if _meaningful_number_format(fmt) else DECIMAL_FORMAT


def resolved_series_number_formats(target: PptTarget, plan: TransformPlan) -> list[str]:
    """Resolve cada serie: override manual > PPT > XLSX > inferencia conservadora."""
    target_labels = list(target.expected_series or [])
    target_formats = list(target.series_value_formats or [])
    target_global_format = (
        target.value_format
        if not any(_meaningful_number_format(fmt) for fmt in target_formats)
        else ""
    )
    source_labels = list(plan.datasource.series or [])
    source_formats = list(plan.datasource.series_number_formats or [])
    explicit = (
        plan.number_format
        if plan.number_format and plan.number_format != "thousands_pt_br"
        else ""
    )
    output: list[str] = []
    for index, label in enumerate(plan.series):
        target_format = _format_for_series(
            label,
            index,
            target_labels,
            target_formats,
            allow_index_fallback=len(target_labels) == len(plan.series),
        )
        source_format = _format_for_series(
            label,
            index,
            source_labels,
            source_formats,
            allow_index_fallback=len(source_labels) == len(plan.series),
        )
        automatic = next(
            (
                fmt
                for fmt in (target_format, target_global_format, source_format, explicit)
                if _meaningful_number_format(fmt)
            ),
            "",
        )
        if not automatic:
            automatic = _inferred_series_format(plan, index, label)
        mode = (
            plan.series_format_overrides.get(label)
            or plan.series_format_overrides.get(_series_key(label))
            or plan.series_format_overrides.get(f"__index_{index}")
            or "auto"
        )
        output.append(_format_for_mode(automatic, mode))
    return output


class ChartSheetUnresolvedError(RuntimeError):
    pass


def chart_replacements(zf: ZipFile, target: PptTarget, plan: TransformPlan) -> dict[str, bytes]:
    if not target.chart_xml:
        return {}
    return {
        target.chart_xml: _updated_chart_xml_bytes(zf, target, plan),
    }




def _updated_chart_xml_bytes(zf: ZipFile, target: PptTarget, plan: TransformPlan) -> bytes:
    root = ET.fromstring(zf.read(target.chart_xml))
    _disable_chart_auto_update(root)
    series_elements = root.findall(".//c:ser", NS)
    if not target.sheet_name:
        raise ChartSheetUnresolvedError(
            f"Nao foi possivel identificar a aba do workbook embutido do grafico "
            f"'{target.shape_name}' (slide {target.slide_number}). Para nao gravar uma "
            f"referencia de aba incorreta no 'Editar dados', a geracao foi interrompida "
            f"para este alvo; revise o template ou aplique um XLSX manualmente."
        )
    sheet = _sheet_ref(target.sheet_name)
    series_number_formats = resolved_series_number_formats(target, plan)

    if plan.orientation_ppt == "series_rows_categories_columns":
        end_col = _excel_col(len(plan.categories) + 1)
        for index, ser in enumerate(series_elements[: len(plan.series)]):
            excel_row = index + 2
            values = plan.values[index] if index < len(plan.values) else []
            number_format = series_number_formats[index] if index < len(series_number_formats) else DECIMAL_FORMAT
            _update_series_text(ser, f"{sheet}!$A${excel_row}", plan.series[index])
            _update_series_categories(ser, f"{sheet}!$B$1:${end_col}$1", plan.categories)
            _update_series_values(ser, f"{sheet}!$B${excel_row}:${end_col}${excel_row}", values, number_format)
            _update_series_data_label_number_format(ser, number_format)
    else:
        end_row = len(plan.categories) + 1
        for index, ser in enumerate(series_elements[: len(plan.series)]):
            excel_col = _excel_col(index + 2)
            values = [row[index] if index < len(row) else None for row in plan.values]
            number_format = series_number_formats[index] if index < len(series_number_formats) else DECIMAL_FORMAT
            _update_series_text(ser, f"{sheet}!${excel_col}$1", plan.series[index])
            _update_series_categories(ser, f"{sheet}!$A$2:$A${end_row}", plan.categories)
            _update_series_values(ser, f"{sheet}!${excel_col}$2:${excel_col}${end_row}", values, number_format)
            _update_series_data_label_number_format(ser, number_format)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)




def _disable_chart_auto_update(root: ET.Element) -> None:
    for external_data in root.findall(".//c:externalData", NS):
        auto_update = external_data.find("./c:autoUpdate", NS)
        if auto_update is None:
            auto_update = ET.SubElement(external_data, f"{{{CHART_NS}}}autoUpdate")
        auto_update.attrib["val"] = "0"


def _update_series_text(ser: ET.Element, formula: str, label: str) -> None:
    tx = ser.find("./c:tx/c:strRef", NS)
    if tx is None:
        tx_parent = ser.find("./c:tx", NS)
        if tx_parent is None:
            tx_parent = ET.SubElement(ser, f"{{{CHART_NS}}}tx")
        tx = ET.SubElement(tx_parent, f"{{{CHART_NS}}}strRef")
    _set_formula(tx, formula)
    cache = tx.find("./c:strCache", NS)
    if cache is None:
        cache = ET.SubElement(tx, f"{{{CHART_NS}}}strCache")
    _set_cache_values(cache, [label], numeric=False)


def _update_series_categories(ser: ET.Element, formula: str, labels: list[str]) -> None:
    cat = ser.find("./c:cat/c:strRef", NS)
    if cat is None:
        cat_parent = ser.find("./c:cat", NS)
        if cat_parent is None:
            cat_parent = ET.SubElement(ser, f"{{{CHART_NS}}}cat")
        cat = ET.SubElement(cat_parent, f"{{{CHART_NS}}}strRef")
    _set_formula(cat, formula)
    cache = cat.find("./c:strCache", NS)
    if cache is None:
        cache = ET.SubElement(cat, f"{{{CHART_NS}}}strCache")
    _set_cache_values(cache, labels, numeric=False)


def _update_series_values(ser: ET.Element, formula: str, values: list[Any], number_format: str = "") -> None:
    val = ser.find("./c:val/c:numRef", NS)
    if val is None:
        val_parent = ser.find("./c:val", NS)
        if val_parent is None:
            val_parent = ET.SubElement(ser, f"{{{CHART_NS}}}val")
        val = ET.SubElement(val_parent, f"{{{CHART_NS}}}numRef")
    _set_formula(val, formula)
    cache = val.find("./c:numCache", NS)
    if cache is None:
        cache = ET.SubElement(val, f"{{{CHART_NS}}}numCache")
    _set_cache_values(cache, values, numeric=True, number_format=number_format)


def _set_formula(parent: ET.Element, formula: str) -> None:
    formula_el = parent.find("./c:f", NS)
    if formula_el is None:
        formula_el = ET.SubElement(parent, f"{{{CHART_NS}}}f")
    formula_el.text = formula


def _set_cache_values(cache: ET.Element, values: list[Any], numeric: bool, number_format: str = "") -> None:
    if numeric and number_format:
        _set_format_code(cache, number_format)

    for child in list(cache):
        if child.tag in {f"{{{CHART_NS}}}ptCount", f"{{{CHART_NS}}}pt"}:
            cache.remove(child)

    insert_at = 0
    for i, child in enumerate(list(cache)):
        if child.tag == f"{{{CHART_NS}}}formatCode":
            insert_at = i + 1

    cache.insert(insert_at, ET.Element(f"{{{CHART_NS}}}ptCount", {"val": str(len(values))}))
    for offset, value in enumerate(values):
        pt = ET.Element(f"{{{CHART_NS}}}pt", {"idx": str(offset)})
        v = ET.SubElement(pt, f"{{{CHART_NS}}}v")
        v.text = _chart_value_text(value, numeric=numeric, number_format=number_format)
        cache.insert(insert_at + offset + 1, pt)


def _set_format_code(cache: ET.Element, number_format: str) -> None:
    format_el = cache.find("./c:formatCode", NS)
    if format_el is None:
        format_el = ET.Element(f"{{{CHART_NS}}}formatCode")
        cache.insert(0, format_el)
    format_el.text = number_format


def _update_series_data_label_number_format(ser: ET.Element, number_format: str) -> None:
    if not number_format:
        return
    d_lbls = ser.find("./c:dLbls", NS)
    if d_lbls is None:
        d_lbls = ET.SubElement(ser, f"{{{CHART_NS}}}dLbls")
    num_fmt = d_lbls.find("./c:numFmt", NS)
    if num_fmt is None:
        num_fmt = ET.Element(f"{{{CHART_NS}}}numFmt")
        d_lbls.insert(0, num_fmt)
    num_fmt.attrib["formatCode"] = number_format
    num_fmt.attrib["sourceLinked"] = "0"


def _format_for_series(
    label: str,
    index: int,
    labels: list[str],
    formats: list[str],
    *,
    allow_index_fallback: bool,
) -> str:
    wanted = _series_key(label)
    if wanted:
        for candidate_index, candidate in enumerate(labels):
            if _series_key(candidate) == wanted and candidate_index < len(formats):
                return formats[candidate_index]
    if allow_index_fallback and index < len(formats):
        return formats[index]
    return ""


def _series_key(label: str) -> str:
    text = unicodedata.normalize("NFKD", str(label or "")).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def _meaningful_number_format(number_format: str) -> bool:
    return str(number_format or "").strip().lower() not in {"", "general", "@"}


def _format_for_mode(automatic: str, mode: str) -> str:
    normalized_mode = str(mode or "auto").strip().lower()
    if normalized_mode == "percent":
        return automatic if "%" in automatic else PERCENT_FORMAT
    if normalized_mode == "number":
        return automatic if _meaningful_number_format(automatic) and "%" not in automatic else DECIMAL_FORMAT
    return effective_value_format(automatic)


def _inferred_series_format(plan: TransformPlan, index: int, label: str) -> str:
    values = (
        plan.values[index]
        if plan.orientation_ppt == "series_rows_categories_columns" and index < len(plan.values)
        else [row[index] for row in plan.values if index < len(row)]
    )
    numeric = [_to_float(value) for value in values]
    numeric = [value for value in numeric if value is not None]
    label_key = _series_key(label)
    percent_hint = any(token in label_key for token in ("percent", "share", "taxa", "proporcao"))
    if numeric and percent_hint and all(abs(value) <= 1 for value in numeric):
        return PERCENT_FORMAT
    return DECIMAL_FORMAT


def _chart_value_text(value: Any, numeric: bool, number_format: str = "") -> str:
    if value is None:
        return "0" if numeric else ""
    if numeric:
        parsed = _to_float(value)
        if parsed is not None:
            # Verbatim: gravamos exatamente o numero do XLSX, com precisao total.
            # O formatCode cuida da exibicao (1 casa); "Editar dados" mantem tudo.
            return f"{parsed:.12g}"
    if isinstance(value, float):
        return f"{value:.12g}"
    return str(value)


def _to_float(value: Any) -> float | None:
    """Converte para numero SEM aplicar qualquer escala (nunca divide/multiplica
    por 100). Um '%' textual e apenas removido; o valor numerico e preservado."""
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").strip()
    text = re.sub(r"^,", "0,", text)
    text = re.sub(r"^-,", "-0,", text)
    text = text.replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _sheet_ref(sheet_name: str) -> str:
    if any(ch in sheet_name for ch in " -'"):
        return "'" + sheet_name.replace("'", "''") + "'"
    return sheet_name


def _excel_col(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
