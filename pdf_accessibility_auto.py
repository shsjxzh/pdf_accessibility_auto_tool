#!/usr/bin/env python3
"""
Best-effort accessibility improvements for an already tagged PDF.

This tool is designed for PDFs that already contain a StructTreeRoot.
It can:
- add /Alt text to tagged Figure, Table, and Formula elements
- add /Summary text for tagged tables
- remove nested /Alt-like metadata beneath those elements
- promote missing table header cells from TD to TH
- promote likely heading paragraphs to H1/H2/H3
- set document title and language metadata
- optionally delete selected pages
- optionally keep visible running page numbers while removing overlapping annotations

It does not build a full accessibility tag tree from an untagged PDF.
"""

from __future__ import annotations

import argparse
import json
import math
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import fitz
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, BooleanObject, DictionaryObject, IndirectObject, NameObject, TextStringObject


CAPTION_RE = re.compile(r"^(Figure|Table)\s+([A-Za-z0-9.-]+):\s*(.*)")
ALT_ELIGIBLE_TAGS = {"/Figure", "/Table", "/Formula"}


@dataclass
class VisibleCaption:
    page: int
    kind: str
    label: str
    body: str
    bbox: tuple[float, float, float, float]


@dataclass
class StructElement:
    ref: IndirectObject
    tag_type: str
    page_num: int | None
    bbox: tuple[float, float, float, float] | None


def normalize(text: str) -> str:
    return " ".join(text.split())


def clean_text(text: str) -> str:
    replacements = {
        "\u00ad": "",
        "\u2013": "-",
        "\u2014": "-",
        "\u2192": "->",
        "\u2212": "-",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return normalize(text)


def clean_formula_text(text: str) -> str:
    text = clean_text(text)
    symbol_words = {
        "\u2208": " in ",
        "\u2209": " not in ",
        "\u2200": " for all ",
        "\u2203": " there exists ",
        "\u2264": " less than or equal to ",
        "\u2265": " greater than or equal to ",
        "\u2260": " not equal to ",
        "\u221d": " proportional to ",
        "\u2211": " sum ",
        "\u220f": " product ",
        "\u221a": " square root ",
        "\u2192": " maps to ",
        "\u2190": " from ",
        "\u2194": " iff ",
        "\u2225": " parallel to ",
        "\u22a4": " transpose ",
        "\u22a5": " perpendicular to ",
    }
    for src, dst in symbol_words.items():
        text = text.replace(src, dst)
    text = re.sub(r"\(\s*\d+(?:\.\d+)*\s*\)", "", text)
    text = re.sub(r"(?<!\S)[()]+(?!\S)", "", text)
    text = re.sub(r"^\s*[()]+\s*", "", text)
    text = re.sub(r"\s*[()]+\s*$", "", text)
    return normalize(text)


def temp_path_for(pdf_path: Path, suffix: str) -> Path:
    return pdf_path.with_name(f"{pdf_path.stem}.{suffix}.tmp.pdf")


def to_roman_lower(num: int) -> str:
    vals = [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
        (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
        (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ]
    out = []
    for value, symbol in vals:
        while num >= value:
            out.append(symbol)
            num -= value
    return "".join(out)


def compute_page_labels(reader: PdfReader) -> dict[int, str]:
    total = len(reader.pages)
    labels = {i: str(i + 1) for i in range(total)}
    page_labels = reader.trailer["/Root"].get("/PageLabels")
    if page_labels is None:
        return labels

    nums = page_labels.get_object().get("/Nums")
    if not isinstance(nums, ArrayObject):
        return labels

    starts = []
    for i in range(0, len(nums), 2):
        start = int(nums[i])
        spec = nums[i + 1].get_object() if isinstance(nums[i + 1], IndirectObject) else nums[i + 1]
        starts.append((start, spec))
    starts.sort(key=lambda item: item[0])

    for idx, (start, spec) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else total
        prefix = str(spec.get("/P") or "")
        style = str(spec.get("/S") or "")
        current = int(spec.get("/St", 1))
        for page_index in range(start, end):
            if style == "/r":
                label = to_roman_lower(current)
            elif style == "/D":
                label = str(current)
            else:
                label = ""
            labels[page_index] = prefix + label
            current += 1
    return labels


def find_page_number_rects(pdf_path: Path, labels_by_index: dict[int, str]) -> dict[int, list[tuple[float, float, float, float]]]:
    pdf = fitz.open(pdf_path)
    rects_by_page: dict[int, list[tuple[float, float, float, float]]] = defaultdict(list)
    for page_index in range(pdf.page_count):
        label = labels_by_index.get(page_index, "").strip()
        if not label:
            continue
        # Keep visible Roman front-matter page numbers intact.
        if re.fullmatch(r"[ivxlcdm]+", label.lower()):
            continue
        page = pdf.load_page(page_index)
        page_rect = page.rect
        for rect in page.search_for(label):
            if rect.width > 50 or rect.height > 25:
                continue
            near_top = rect.y1 <= 60
            near_bottom = rect.y0 >= page_rect.height - 60
            if near_top or near_bottom:
                rects_by_page[page_index].append((rect.x0, rect.y0, rect.x1, rect.y1))
    pdf.close()
    return rects_by_page


def delete_pages(pdf_path: Path, pages_to_delete: list[int]) -> Path:
    if not pages_to_delete:
        return pdf_path
    pdf = fitz.open(pdf_path)
    zero_based = sorted({page - 1 for page in pages_to_delete if 1 <= page <= pdf.page_count}, reverse=True)
    if not zero_based:
        pdf.close()
        return pdf_path
    for page_index in zero_based:
        pdf.delete_page(page_index)
    out_path = temp_path_for(pdf_path, "pages-deleted")
    pdf.save(out_path, garbage=4, deflate=True)
    pdf.close()
    return out_path


def cleanup_formula_glyph_artifacts(pdf_path: Path) -> tuple[Path, int]:
    pdf = fitz.open(pdf_path)
    removed = 0
    replacements = {0x2208, 0x2209}
    for page in pdf:
        rects: list[fitz.Rect] = []
        raw = page.get_text("rawdict")
        for block in raw.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    for ch in span.get("chars", []):
                        char = ch.get("c", "")
                        if len(char) == 1:
                            if ord(char) in replacements:
                                rect = fitz.Rect(ch["bbox"])
                                rect = fitz.Rect(rect.x0 - 0.2, rect.y0 - 0.2, rect.x1, rect.y1 + 0.2)
                                rects.append(rect)
                        if ch.get("c") in {"（", "）"}:
                            rect = fitz.Rect(ch["bbox"])
                            rects.append(fitz.Rect(rect.x0 - 0.5, rect.y0 - 0.5, rect.x1 + 0.5, rect.y1 + 0.5))
        for rect in rects:
            page.add_redact_annot(rect, fill=(1, 1, 1))
        if rects:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            removed += len(rects)
    if removed == 0:
        pdf.close()
        return pdf_path, 0
    out_path = temp_path_for(pdf_path, "formula-glyph-clean")
    pdf.save(out_path, garbage=4, deflate=True)
    pdf.close()
    return out_path, removed


def extract_visible_captions(pdf_path: Path) -> dict[int, list[VisibleCaption]]:
    pdf = fitz.open(pdf_path)
    by_page: dict[int, list[VisibleCaption]] = defaultdict(list)
    for page_index in range(pdf.page_count):
        page = pdf.load_page(page_index)
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0 or not block.get("lines"):
                continue
            line_texts = ["".join(span["text"] for span in line["spans"]).strip() for line in block["lines"]]
            if not line_texts:
                continue
            match = CAPTION_RE.match(line_texts[0])
            if not match:
                continue
            kind, label, first_body = match.group(1), match.group(2), match.group(3)
            body = clean_text(" ".join([first_body] + line_texts[1:]))
            by_page[page_index + 1].append(
                VisibleCaption(
                    page=page_index + 1,
                    kind=kind,
                    label=label,
                    body=body,
                    bbox=tuple(block["lines"][0]["bbox"]),
                )
            )
    pdf.close()
    return by_page


def detect_heading_candidates(pdf_path: Path) -> dict[int, list[str]]:
    pdf = fitz.open(pdf_path)
    by_page: dict[int, list[str]] = defaultdict(list)
    for page_index in range(pdf.page_count):
        page = pdf.load_page(page_index)
        blocks = []
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0 or not block.get("lines"):
                continue
            text = clean_text(" ".join("".join(span["text"] for span in line["spans"]) for line in block["lines"]))
            if not text:
                continue
            blocks.append((block["bbox"][1], text))
        blocks.sort(key=lambda item: item[0])

        for _, text in blocks[:8]:
            if re.fullmatch(r"CHAPTER\s+\d+", text) or re.fullmatch(r"APPENDIX\s+[A-Z]", text):
                by_page[page_index + 1].append("H1")
            elif re.fullmatch(r"[A-Z0-9 ,:;()'/-]{12,}", text) and any(ch.isalpha() for ch in text):
                by_page[page_index + 1].append("H1")
            elif re.match(r"^\d+\.\d+(\.\d+)*\s+\S", text) or re.match(r"^[A-Z]\.\d+(\.\d+)*\s+\S", text):
                dots = text.split()[0].count(".")
                by_page[page_index + 1].append(f"H{min(3, dots + 1)}")
    pdf.close()
    return by_page


def build_page_ref_map(reader: PdfReader) -> dict[tuple[int, int], int]:
    page_ref_map: dict[tuple[int, int], int] = {}
    for page_num, page in enumerate(reader.pages, start=1):
        ref = page.indirect_reference
        page_ref_map[(ref.idnum, ref.generation)] = page_num
    return page_ref_map


def resolve_page_num(obj: DictionaryObject | IndirectObject, page_ref_map: dict[tuple[int, int], int]) -> int | None:
    if isinstance(obj, IndirectObject):
        obj = obj.get_object()
    if not isinstance(obj, DictionaryObject):
        return None

    page_ref = obj.get("/Pg")
    if page_ref is not None:
        return page_ref_map.get((page_ref.idnum, page_ref.generation))

    kids = obj.get("/K")
    if isinstance(kids, ArrayObject):
        for kid in kids:
            page_num = resolve_page_num(kid, page_ref_map)
            if page_num is not None:
                return page_num
    elif isinstance(kids, (DictionaryObject, IndirectObject)):
        return resolve_page_num(kids, page_ref_map)
    return None


def element_bbox(obj: DictionaryObject) -> tuple[float, float, float, float] | None:
    attrs = obj.get("/A")
    if isinstance(attrs, IndirectObject):
        attrs = attrs.get_object()
    if isinstance(attrs, ArrayObject):
        for item in attrs:
            candidate = item.get_object() if isinstance(item, IndirectObject) else item
            if hasattr(candidate, "get"):
                bbox = candidate.get("/BBox")
                if isinstance(bbox, ArrayObject) and len(bbox) == 4:
                    return tuple(float(x) for x in bbox)
    elif hasattr(attrs, "get"):
        bbox = attrs.get("/BBox")
        if isinstance(bbox, ArrayObject) and len(bbox) == 4:
            return tuple(float(x) for x in bbox)
    return None


def collect_struct_elements(reader: PdfReader) -> list[StructElement]:
    page_ref_map = build_page_ref_map(reader)
    struct = reader.trailer["/Root"].get("/StructTreeRoot")
    if struct is None:
        raise ValueError("Input PDF does not contain /StructTreeRoot")
    struct = struct.get_object()

    seen: set[tuple[int, int]] = set()
    elements: list[StructElement] = []

    def walk(obj: DictionaryObject | ArrayObject | IndirectObject | None, inside_alt_branch: bool = False) -> None:
        if obj is None:
            return
        if isinstance(obj, IndirectObject):
            key = (obj.idnum, obj.generation)
            if key in seen:
                return
            seen.add(key)
            ref = obj
            obj = obj.get_object()
        else:
            ref = None

        if isinstance(obj, DictionaryObject):
            tag_type = obj.get("/S")
            current_inside = inside_alt_branch
            if tag_type in ALT_ELIGIBLE_TAGS and ref is not None and not inside_alt_branch:
                elements.append(
                    StructElement(
                        ref=ref,
                        tag_type=str(tag_type)[1:],
                        page_num=resolve_page_num(ref, page_ref_map),
                        bbox=element_bbox(obj),
                    )
                )
                current_inside = True
            elif tag_type in ALT_ELIGIBLE_TAGS:
                current_inside = True
            walk(obj.get("/K"), current_inside)
        elif isinstance(obj, ArrayObject):
            for kid in obj:
                walk(kid, inside_alt_branch)

    walk(struct.get("/K"))
    return elements


def distance_score(element: StructElement, caption: VisibleCaption) -> tuple[float, float, float]:
    if element.bbox is None:
        return (0.0, 0.0, 0.0)
    ex0, ey0, ex1, ey1 = element.bbox
    cx0, cy0, cx1, _ = caption.bbox
    vertical = cy0 - ey1
    if vertical < -24:
        vertical_penalty = 5000 + abs(vertical)
    elif vertical < 0:
        vertical_penalty = 1000 + abs(vertical)
    else:
        vertical_penalty = vertical
    horizontal = abs(((ex0 + ex1) / 2) - ((cx0 + cx1) / 2))
    top_bias = ey0
    return (vertical_penalty, horizontal, top_bias)


def clipped_text_fallback(pdf: fitz.Document, element: StructElement) -> str | None:
    if element.page_num is None or element.bbox is None:
        return None
    page = pdf.load_page(element.page_num - 1)
    rect = fitz.Rect(element.bbox)
    text = clean_text(page.get_text("text", clip=rect))
    if not text:
        return None
    if element.tag_type == "Formula":
        text = clean_formula_text(text)
        if not text:
            return None
    words = text.split()
    snippet = " ".join(words[:40])
    return f"{element.tag_type}. {snippet}"


def table_headers_from_bbox(pdf: fitz.Document, element: StructElement, caption: VisibleCaption | None = None) -> str | None:
    if element.tag_type != "Table" or element.page_num is None or element.bbox is None:
        return None
    page = pdf.load_page(element.page_num - 1)
    rect = fitz.Rect(element.bbox)
    words = page.get_text("words", clip=rect)
    if not words:
        return None

    if caption is not None:
        words = [word for word in words if word[1] >= caption.bbox[3] + 1]
        if not words:
            return None

    top_y = min(word[1] for word in words)
    row_band = max(10.0, min(18.0, (rect.y1 - rect.y0) * 0.12))
    header_words = [word for word in words if word[1] <= top_y + row_band]
    if not header_words:
        return None

    lines: dict[float, list[tuple[float, str]]] = defaultdict(list)
    for x0, y0, x1, y1, text, *_ in header_words:
        bucket = round(y0 / 3) * 3
        lines[bucket].append((x0, text))

    ordered_lines = []
    for y in sorted(lines):
        parts = [text for _, text in sorted(lines[y])]
        line = normalize(" ".join(parts))
        if line:
            ordered_lines.append(line)

    header_text = normalize(" ; ".join(ordered_lines))
    if len(header_text.split()) < 2:
        return None
    if "." in header_text or len(header_text.split()) > 16:
        return None
    return header_text or None


def generic_fallback(element: StructElement) -> str:
    if element.page_num is not None:
        if element.tag_type == "Formula":
            return f"Formula on page {element.page_num}."
        return f"Uncaptioned {element.tag_type.lower()} on page {element.page_num}."
    return f"Uncaptioned {element.tag_type.lower()}."


def alt_from_caption(element: StructElement, caption: VisibleCaption | None, table_headers: str | None) -> str | None:
    if caption is None:
        return None
    if element.tag_type == "Table":
        text = f"Table {caption.label}. {caption.body}"
        if table_headers:
            text += f" Headers: {table_headers}."
        return text
    if element.tag_type == "Formula":
        body = clean_formula_text(caption.body)
        if body:
            return f"Formula near {caption.kind} {caption.label}. {body}"
        return f"Formula near {caption.kind} {caption.label}."
    return f"{caption.kind} {caption.label}. {caption.body}"


def remove_nested_alt_text(root_ref: IndirectObject) -> None:
    seen: set[tuple[int, int]] = set()

    def walk(obj: DictionaryObject | ArrayObject | IndirectObject | None, depth: int = 0) -> None:
        if obj is None:
            return
        if isinstance(obj, IndirectObject):
            key = (obj.idnum, obj.generation)
            if key in seen:
                return
            seen.add(key)
            obj = obj.get_object()
        if isinstance(obj, DictionaryObject):
            if depth >= 1:
                for key in ["/Alt", "/ActualText", "/Summary"]:
                    if key in obj:
                        del obj[key]
            walk(obj.get("/K"), depth + 1)
        elif isinstance(obj, ArrayObject):
            for kid in obj:
                walk(kid, depth)

    walk(root_ref, 0)


def ensure_table_headers(reader: PdfReader) -> int:
    struct = reader.trailer["/Root"].get("/StructTreeRoot")
    if struct is None:
        return 0
    struct = struct.get_object()
    seen: set[tuple[int, int]] = set()
    fixed = 0

    def walk(obj: DictionaryObject | ArrayObject | IndirectObject | None) -> None:
        nonlocal fixed
        if obj is None:
            return
        if isinstance(obj, IndirectObject):
            key = (obj.idnum, obj.generation)
            if key in seen:
                return
            seen.add(key)
            obj = obj.get_object()
        if isinstance(obj, DictionaryObject):
            if obj.get("/S") == "/Table":
                kids = obj.get("/K")
                rows = []
                if isinstance(kids, ArrayObject):
                    for kid in kids:
                        kid_obj = kid.get_object() if isinstance(kid, IndirectObject) else kid
                        if isinstance(kid_obj, DictionaryObject) and kid_obj.get("/S") == "/TR":
                            rows.append(kid_obj)
                has_th = False
                for row in rows:
                    row_k = row.get("/K")
                    if isinstance(row_k, ArrayObject):
                        for cell in row_k:
                            cell_obj = cell.get_object() if isinstance(cell, IndirectObject) else cell
                            if isinstance(cell_obj, DictionaryObject) and cell_obj.get("/S") == "/TH":
                                has_th = True
                                break
                    if has_th:
                        break
                if not has_th and rows:
                    row_k = rows[0].get("/K")
                    if isinstance(row_k, ArrayObject):
                        for cell in row_k:
                            cell_obj = cell.get_object() if isinstance(cell, IndirectObject) else cell
                            if isinstance(cell_obj, DictionaryObject) and cell_obj.get("/S") == "/TD":
                                cell_obj[NameObject("/S")] = NameObject("/TH")
                                fixed += 1
            walk(obj.get("/K"))
        elif isinstance(obj, ArrayObject):
            for kid in obj:
                walk(kid)

    walk(struct.get("/K"))
    return fixed


def promote_headings(reader: PdfReader, heading_candidates: dict[int, list[str]]) -> int:
    struct = reader.trailer["/Root"].get("/StructTreeRoot")
    if struct is None:
        return 0
    struct = struct.get_object()
    seen: set[tuple[int, int]] = set()
    page_ref_map = build_page_ref_map(reader)
    p_by_page: dict[int, list[IndirectObject]] = defaultdict(list)

    def walk(obj: DictionaryObject | ArrayObject | IndirectObject | None) -> None:
        if obj is None:
            return
        if isinstance(obj, IndirectObject):
            key = (obj.idnum, obj.generation)
            if key in seen:
                return
            seen.add(key)
            ref = obj
            obj = obj.get_object()
        else:
            ref = None
        if isinstance(obj, DictionaryObject):
            if obj.get("/S") == "/P" and ref is not None:
                page_num = resolve_page_num(ref, page_ref_map)
                if page_num is not None:
                    p_by_page[page_num].append(ref)
            walk(obj.get("/K"))
        elif isinstance(obj, ArrayObject):
            for kid in obj:
                walk(kid)

    walk(struct.get("/K"))

    changed = 0
    for page_num, levels in heading_candidates.items():
        if not levels or not p_by_page.get(page_num):
            continue
        page_ps = p_by_page[page_num]
        # Skip the page-number paragraph if present and only retag the next few top-of-page paragraphs.
        start_idx = 1 if len(page_ps) > 1 else 0
        for idx, level in enumerate(levels):
            target_idx = start_idx + idx
            if target_idx >= len(page_ps):
                break
            obj = page_ps[target_idx].get_object()
            target_tag = NameObject(f"/{level}")
            if obj.get("/S") != target_tag:
                obj[NameObject("/S")] = target_tag
                changed += 1
    return changed


def assign_alt_text(reader: PdfReader, fitz_pdf: fitz.Document, visible_by_page: dict[int, list[VisibleCaption]]) -> dict[str, int]:
    elements = collect_struct_elements(reader)
    elements_by_page: dict[int | None, list[StructElement]] = defaultdict(list)
    for element in elements:
        elements_by_page[element.page_num].append(element)

    assigned = 0
    fallback_used = 0

    for page_num, page_elements in elements_by_page.items():
        captions = list(visible_by_page.get(page_num or -1, []))
        page_elements.sort(key=lambda item: (math.inf if item.bbox is None else item.bbox[1], item.tag_type))
        used_caption_indexes: set[int] = set()

        for element in page_elements:
            matched_caption = None
            if element.tag_type in {"Figure", "Table"} and captions:
                best_idx = None
                best_score = None
                for idx, caption in enumerate(captions):
                    if idx in used_caption_indexes:
                        continue
                    score = distance_score(element, caption)
                    if best_score is None or score < best_score:
                        best_score = score
                        best_idx = idx
                if best_idx is not None:
                    matched_caption = captions[best_idx]
                    used_caption_indexes.add(best_idx)
            header_text = table_headers_from_bbox(fitz_pdf, element, matched_caption)

            alt_text = alt_from_caption(element, matched_caption, header_text)
            if alt_text is None:
                alt_text = clipped_text_fallback(fitz_pdf, element)
            if alt_text is None and header_text and element.tag_type == "Table":
                alt_text = f"Table. Headers: {header_text}."
            if alt_text is None:
                alt_text = generic_fallback(element)
            if matched_caption is None:
                fallback_used += 1

            obj = element.ref.get_object()
            remove_nested_alt_text(element.ref)
            final_text = clean_formula_text(alt_text) if element.tag_type == "Formula" else clean_text(alt_text)
            obj[NameObject("/Alt")] = TextStringObject(final_text)
            if element.tag_type == "Formula":
                obj[NameObject("/ActualText")] = TextStringObject(final_text)
            if element.tag_type == "Table":
                obj[NameObject("/Summary")] = TextStringObject(clean_text(alt_text))
            assigned += 1

    return {"assigned": assigned, "fallback_used": fallback_used}


def apply_document_metadata(writer: PdfWriter, title: str, language: str) -> None:
    writer.add_metadata({"/Title": title})
    writer.root_object[NameObject("/Lang")] = TextStringObject(language)
    viewer_prefs = writer.root_object.get(NameObject("/ViewerPreferences"))
    if viewer_prefs is None:
        viewer_prefs = DictionaryObject()
    viewer_prefs[NameObject("/DisplayDocTitle")] = BooleanObject(True)
    writer.root_object[NameObject("/ViewerPreferences")] = viewer_prefs


def strip_page_number_annotations(reader: PdfReader, rects_by_page: dict[int, list[tuple[float, float, float, float]]]) -> int:
    removed = 0
    for page_index, rects in rects_by_page.items():
        if page_index >= len(reader.pages):
            continue
        page = reader.pages[page_index]
        annots = page.get("/Annots")
        if not isinstance(annots, ArrayObject):
            continue
        kept = ArrayObject()
        for annot_ref in annots:
            annot = annot_ref.get_object() if isinstance(annot_ref, IndirectObject) else annot_ref
            rect = annot.get("/Rect") if hasattr(annot, "get") else None
            remove = False
            if isinstance(rect, ArrayObject) and len(rect) == 4:
                ax0, ay0, ax1, ay1 = [float(x) for x in rect]
                annot_rect = fitz.Rect(ax0, ay0, ax1, ay1)
                for target in rects:
                    if annot_rect.intersects(fitz.Rect(target)):
                        remove = True
                        break
            if remove:
                removed += 1
            else:
                kept.append(annot_ref)
        if kept:
            page[NameObject("/Annots")] = kept
        elif "/Annots" in page:
            del page["/Annots"]
    return removed


def count_missing_alt(reader: PdfReader) -> dict[str, int]:
    struct = reader.trailer["/Root"].get("/StructTreeRoot")
    if struct is None:
        return {"Figure": 0, "Table": 0, "Formula": 0}
    struct = struct.get_object()
    seen: set[tuple[int, int]] = set()
    counts = {"Figure": 0, "Table": 0, "Formula": 0}

    def walk(obj: DictionaryObject | ArrayObject | IndirectObject | None) -> None:
        if obj is None:
            return
        if isinstance(obj, IndirectObject):
            key = (obj.idnum, obj.generation)
            if key in seen:
                return
            seen.add(key)
            obj = obj.get_object()
        if isinstance(obj, DictionaryObject):
            tag_type = obj.get("/S")
            if tag_type == "/Figure" and "/Alt" not in obj:
                counts["Figure"] += 1
            elif tag_type == "/Table" and "/Alt" not in obj:
                counts["Table"] += 1
            elif tag_type == "/Formula" and "/Alt" not in obj:
                counts["Formula"] += 1
            walk(obj.get("/K"))
        elif isinstance(obj, ArrayObject):
            for kid in obj:
                walk(kid)

    walk(struct.get("/K"))
    return counts


def main() -> int:
    parser = argparse.ArgumentParser(description="Best-effort accessibility improvements for a tagged PDF.")
    parser.add_argument("input_pdf", type=Path, help="Tagged input PDF")
    parser.add_argument("-o", "--output", type=Path, help="Output PDF path")
    parser.add_argument("--title", help="Document title metadata")
    parser.add_argument("--lang", default="en-US", help="Document language")
    parser.add_argument(
        "--delete-page",
        action="append",
        type=int,
        default=[],
        help="1-based page number to delete before processing; repeatable",
    )
    parser.add_argument(
        "--clean-running-page-numbers",
        action="store_true",
        help="Keep visible running page numbers, but remove overlapping annotations in header/footer regions",
    )
    parser.add_argument(
        "--disable-heading-promotion",
        action="store_true",
        help="Do not retag likely headings as H1/H2/H3",
    )
    parser.add_argument(
        "--disable-table-header-promotion",
        action="store_true",
        help="Do not promote first-row TD cells to TH when a table has no headers",
    )
    args = parser.parse_args()

    input_pdf = args.input_pdf.resolve()
    output_pdf = args.output.resolve() if args.output else input_pdf.with_name(f"{input_pdf.stem}.accessible.pdf")
    title = args.title if args.title else input_pdf.stem

    working_pdf = input_pdf
    temp_paths: list[Path] = []

    page_number_rects: dict[int, list[tuple[float, float, float, float]]] = {}
    if args.clean_running_page_numbers:
        source_reader = PdfReader(str(working_pdf))
        page_labels = compute_page_labels(source_reader)
        page_number_rects = find_page_number_rects(working_pdf, page_labels)

    if args.delete_page:
        deleted_pdf = delete_pages(working_pdf, args.delete_page)
        if deleted_pdf != working_pdf:
            temp_paths.append(deleted_pdf)
            working_pdf = deleted_pdf

    formula_glyph_artifacts_removed = 0
    cleaned_pdf, formula_glyph_artifacts_removed = cleanup_formula_glyph_artifacts(working_pdf)
    if cleaned_pdf != working_pdf:
        temp_paths.append(cleaned_pdf)
        working_pdf = cleaned_pdf

    visible_by_page = extract_visible_captions(working_pdf)
    heading_candidates = detect_heading_candidates(working_pdf)
    fitz_pdf = fitz.open(working_pdf)
    reader = PdfReader(str(working_pdf))

    before = count_missing_alt(reader)
    table_headers_fixed = 0
    if not args.disable_table_header_promotion:
        table_headers_fixed = ensure_table_headers(reader)
    headings_promoted = 0
    if not args.disable_heading_promotion:
        headings_promoted = promote_headings(reader, heading_candidates)
    page_number_annots_removed = 0
    if page_number_rects:
        page_number_annots_removed = strip_page_number_annotations(reader, page_number_rects)
    stats = assign_alt_text(reader, fitz_pdf, visible_by_page)
    after = count_missing_alt(reader)
    fitz_pdf.close()

    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    apply_document_metadata(writer, title, args.lang)
    with output_pdf.open("wb") as handle:
        writer.write(handle)

    for temp_path in temp_paths:
        try:
            temp_path.unlink()
        except OSError:
            pass

    print(
        json.dumps(
            {
                "input_pdf": str(input_pdf),
                "output_pdf": str(output_pdf),
                "deleted_pages": args.delete_page,
                "clean_running_page_numbers": args.clean_running_page_numbers,
                "missing_alt_before": before,
                "missing_alt_after": after,
                "table_header_cells_promoted": table_headers_fixed,
                "headings_promoted": headings_promoted,
                "page_number_annots_removed": page_number_annots_removed,
                "page_number_rects_redacted": 0,
                "formula_glyph_artifacts_removed": formula_glyph_artifacts_removed,
                "title": title,
                "lang": args.lang,
                "assigned": stats["assigned"],
                "fallback_used": stats["fallback_used"],
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
