from __future__ import annotations

from collections import defaultdict
from typing import Iterable
import xml.etree.ElementTree as ET

from .ppt_discovery import NS, PptTarget


def rename_targets_in_slide_xml(slide_xml: bytes, targets: Iterable[PptTarget]) -> bytes:
    target_list = list(targets)
    if not target_list:
        return slide_xml

    root = ET.fromstring(slide_xml)
    targets_by_shape_id = {str(target.shape_id): target for target in target_list if target.shape_id}
    targets_by_shape_name = defaultdict(list)
    for target in target_list:
        if target.shape_name:
            targets_by_shape_name[target.shape_name].append(target)

    changed = False
    for frame in root.findall(".//p:graphicFrame", NS):
        cnv = frame.find("./p:nvGraphicFramePr/p:cNvPr", NS)
        if cnv is None:
            continue
        object_type = "chart" if frame.find(".//c:chart", NS) is not None else "table" if frame.find(".//a:tbl", NS) is not None else ""
        target = _match_target(cnv, object_type, targets_by_shape_id, targets_by_shape_name)
        if target is not None and cnv.attrib.get("name") != target.target_id:
            cnv.attrib["name"] = target.target_id
            changed = True

    for shape in root.findall(".//p:sp", NS):
        cnv = shape.find("./p:nvSpPr/p:cNvPr", NS)
        if cnv is None:
            continue
        target = _match_target(cnv, "text", targets_by_shape_id, targets_by_shape_name)
        if target is None:
            target = _match_target(cnv, "shape", targets_by_shape_id, targets_by_shape_name)
        if target is not None and cnv.attrib.get("name") != target.target_id:
            cnv.attrib["name"] = target.target_id
            changed = True

    if not changed:
        return slide_xml
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _match_target(
    cnv: ET.Element,
    object_type: str,
    targets_by_shape_id: dict[str, PptTarget],
    targets_by_shape_name: dict[str, list[PptTarget]],
) -> PptTarget | None:
    shape_id = str(cnv.attrib.get("id") or "")
    shape_name = str(cnv.attrib.get("name") or "")
    candidates = []
    if shape_id and shape_id in targets_by_shape_id:
        candidates.append(targets_by_shape_id[shape_id])
    candidates.extend(targets_by_shape_name.get(shape_name, []))
    for target in candidates:
        if object_type and target.object_type != object_type:
            continue
        return target
    return None
