"""
Вычисляет соответствие абзацев и блоков страницам на основе оглавления и разрывов страниц
"""
from typing import Any, Dict, List, Sequence, Tuple

from .utils import compact_upper


def build_paragraph_page_lookup(
    blocks: Sequence[Dict[str, Any]], toc_lookup: Dict[str, int]
) -> Dict[str, Dict[int, int]]:
    """Сопоставляет абзацы и блоки со страницами по оглавлению"""
    anchors: List[Tuple[int, int]] = []
    for b in blocks:
        if b["kind"] != "paragraph":
            continue
        title = compact_upper(b["text"])
        if title in toc_lookup:
            anchors.append((b["paragraph_index"], toc_lookup[title]))

    page_by_para: Dict[int, int] = {}
    page_by_block: Dict[int, int] = {}

    if not anchors:
        return {"paragraphs": page_by_para, "blocks": page_by_block}

    anchor_pos = 0
    current_page = anchors[0][1]
    next_anchor_idx = anchors[1][0] if len(anchors) > 1 else None

    for b in blocks:
        if b["kind"] == "paragraph":
            para_idx = b["paragraph_index"]

            while next_anchor_idx is not None and para_idx >= next_anchor_idx:
                anchor_pos += 1
                current_page = anchors[anchor_pos][1]
                next_anchor_idx = anchors[anchor_pos + 1][0] if anchor_pos + 1 < len(anchors) else None

            if b.get("page_break_before") and para_idx != anchors[anchor_pos][0]:
                current_page += 1

            page_by_para[para_idx] = current_page
            page_by_block[b["block_index"]] = current_page

            if b.get("has_page_break") and next_anchor_idx is not None and para_idx + 1 < next_anchor_idx:
                current_page += 1
        else:
            page_by_block[b["block_index"]] = current_page

    return {"paragraphs": page_by_para, "blocks": page_by_block}


def add_pages_to_findings(findings: List[Dict[str, Any]], page_lookup: Dict[str, Dict[int, int]]) -> None:
    """Подставляет номера страниц в findings по paragraph/block индексам"""
    for f in findings:
        loc = f.get("location", {})
        if "page" in loc:
            continue
        if "paragraph" in loc:
            page = page_lookup["paragraphs"].get(loc["paragraph"] - 1)
            if page is not None:
                loc["page"] = page
        elif "block" in loc:
            page = page_lookup["blocks"].get(loc["block"] - 1)
            if page is not None:
                loc["page"] = page
