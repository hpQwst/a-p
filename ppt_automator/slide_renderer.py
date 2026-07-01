from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import BinaryIO
from zipfile import ZipFile
import re
import tempfile
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageFont

from .ppt_discovery import PML_NS, PptTarget, read_bytes
from .target_labeler import visual_label


InputFile = str | Path | bytes | bytearray | BinaryIO


@dataclass(frozen=True)
class RenderedSlide:
    slide_number: int
    image_path: Path
    visual_map: list[dict[str, str]]
    warning: str = ""


def render_slide_with_target_labels(
    pptx_file: InputFile,
    targets: list[PptTarget],
    slide_number: int,
    output_dir: Path,
    width_px: int = 1440,
) -> RenderedSlide:
    slide_targets = [target for target in targets if target.slide_number == slide_number]
    output_dir.mkdir(parents=True, exist_ok=True)
    image_path = output_dir / f"slide_{slide_number:03d}.png"
    ppt_bytes = read_bytes(pptx_file)
    slide_width_in, slide_height_in = _slide_size_inches(ppt_bytes)
    height_px = max(int(width_px * slide_height_in / slide_width_in), 1)
    warning = ""

    try:
        base_image_path = _export_slide_with_powerpoint(
            pptx_file,
            ppt_bytes,
            slide_number,
            output_dir,
            width_px,
            height_px,
        )
        image = Image.open(base_image_path).convert("RGB")
    except Exception as exc:
        warning = f"PowerPoint COM indisponivel para renderizar o slide real: {exc}"
        image = _fallback_slide_canvas(slide_number, width_px, height_px)

    draw = ImageDraw.Draw(image)
    small_font = _font(16)

    visual_map: list[dict[str, str]] = []
    for target in slide_targets:
        label = visual_label(target.target_id)
        visual_map.append({"visual_label": label, "target_id": target.target_id})
        left = int((target.left_in / slide_width_in) * width_px)
        top = int((target.top_in / slide_height_in) * height_px)
        right = int(((target.left_in + target.width_in) / slide_width_in) * width_px)
        bottom = int(((target.top_in + target.height_in) / slide_height_in) * height_px)
        right = max(right, left + 24)
        bottom = max(bottom, top + 24)
        draw.rectangle([left, top, right, bottom], outline="#0f6b52", width=4)
        label_box = [left, max(0, top - 30), left + 58, max(28, top)]
        draw.rectangle(label_box, fill="#0f6b52")
        draw.text((label_box[0] + 8, label_box[1] + 3), label, fill="white", font=small_font)
        draw.text((left + 6, bottom + 4), target.target_id, fill="#0f6b52", font=small_font)

    image.save(image_path)
    return RenderedSlide(slide_number=slide_number, image_path=image_path, visual_map=visual_map, warning=warning)


def _fallback_slide_canvas(slide_number: int, width_px: int, height_px: int) -> Image.Image:
    image = Image.new("RGB", (width_px, height_px), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width_px - 1, height_px - 1], outline="#d9e0e7", width=2)
    draw.text((24, 18), f"Slide {slide_number}", fill="#18212b", font=_font(20))
    return image


def _export_slide_with_powerpoint(
    pptx_file: InputFile,
    ppt_bytes: bytes,
    slide_number: int,
    output_dir: Path,
    width_px: int,
    height_px: int,
) -> Path:
    import pythoncom
    import win32com.client

    temp_input: Path | None = None
    powerpoint = None
    presentation = None
    export_path = output_dir / f"_slide_{slide_number:03d}_base.png"
    try:
        if isinstance(pptx_file, (str, Path)):
            ppt_path = Path(pptx_file).resolve()
        else:
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False, dir=output_dir) as temp_file:
                temp_file.write(ppt_bytes)
                temp_input = Path(temp_file.name)
            ppt_path = temp_input.resolve()

        pythoncom.CoInitialize()
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(str(ppt_path), True, False, False)
        presentation.Slides(slide_number).Export(str(export_path.resolve()), "PNG", width_px, height_px)
        if not export_path.exists():
            raise RuntimeError("PowerPoint nao gerou o PNG do slide.")
        return export_path
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        if temp_input is not None:
            try:
                temp_input.unlink()
            except OSError:
                pass


def _slide_size_inches(ppt_bytes: bytes) -> tuple[float, float]:
    with ZipFile(BytesIO(ppt_bytes)) as zf:
        if "ppt/presentation.xml" not in zf.namelist():
            return 13.333, 7.5
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
    sld_size = root.find(f".//{{{PML_NS}}}sldSz")
    if sld_size is None:
        return 13.333, 7.5
    emu_per_inch = 914400
    width = int(sld_size.attrib.get("cx", "12192000")) / emu_per_inch
    height = int(sld_size.attrib.get("cy", "6858000")) / emu_per_inch
    return max(width, 1), max(height, 1)


def _font(size: int):
    for name in ("arial.ttf", "segoeui.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()
