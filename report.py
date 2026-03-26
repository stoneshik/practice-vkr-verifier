"""
Формирует итоговый отчёт: агрегирует найденные ошибки, считает статистику и подготавливает JSON
"""
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Sequence

from docx import Document

from .checks.file_checks import check_file_format
from .checks.formatting_checks import (
    check_body_formatting,
    check_heading_specific_rules,
    check_page_numbering,
    check_lists
)
from .checks.objects_checks import check_figures, check_tables, check_formulae
from .checks.special_checks import (
    check_appendices,
    check_abbreviations_list,
    check_terms_list,
    check_sources
)
from .checks.structure_checks import (
    check_required_sections,
    check_heading_numbers_and_start_pages,
    check_toc,
    check_optional_sections
)
from .constants import HEADING_NUMBER_RE
from .docx_reader import (
    validate_input_file,
    build_document_model,
    extract_raw_paragraph_texts,
    extract_raw_toc_entries,
    build_toc_page_lookup
)
from .pagination import build_paragraph_page_lookup, add_pages_to_findings
from .utils import block_text, is_heading_like


def extract_structure_overview(
    blocks: Sequence[Dict[str, Any]], page_lookup: Dict[str, Dict[int, int]]
) -> Dict[str, Any]:
    """Формирует краткий обзор структуры документа для итогового отчёта"""
    headings = []
    for b in blocks:
        if b["kind"] != "paragraph":
            continue
        text = block_text(b)
        if not text or not is_heading_like(text):
            continue
        m = HEADING_NUMBER_RE.match(text)
        headings.append({
            "paragraph": b["paragraph_index"] + 1,
            "text": text,
            "level": (m.group(1).count(".") + 1) if m else 1,
            "page_break_before": bool(b.get("page_break_before")),
            "has_page_break": b.get("has_page_break"),
            "page": page_lookup["paragraphs"].get(b["paragraph_index"]),
        })
    figures = [b["paragraph_index"] + 1 for b in blocks if b["kind"] == "paragraph" and b.get("has_drawing")]
    tables = [b["table_index"] + 1 for b in blocks if b["kind"] == "table"]
    return {"headings": headings[:200], "figure_paragraphs": figures[:200], "tables": tables[:200]}


def analyze_docx(path: Path) -> Dict[str, Any]:
    """Полностью анализирует DOCX и возвращает JSON-совместимый отчёт"""
    validate_input_file(path)
    doc = Document(str(path))
    model = build_document_model(doc)
    blocks = model["blocks"]
    sections = model["sections"]
    raw_paragraph_texts = extract_raw_paragraph_texts(path)
    raw_toc_entries = extract_raw_toc_entries(path)
    toc_lookup = build_toc_page_lookup(raw_toc_entries)
    page_lookup = build_paragraph_page_lookup(blocks, toc_lookup)

    findings: List[Dict[str, Any]] = []
    findings.extend(check_file_format(path))
    findings.extend(check_required_sections(blocks, raw_paragraph_texts))
    findings.extend(check_heading_numbers_and_start_pages(blocks, toc_lookup, page_lookup))
    findings.extend(check_toc(blocks, raw_toc_entries, page_lookup))
    findings.extend(check_body_formatting(blocks, sections))
    findings.extend(check_page_numbering(sections))
    findings.extend(check_heading_specific_rules(blocks))
    findings.extend(check_figures(blocks))
    findings.extend(check_tables(blocks))
    findings.extend(check_appendices(blocks, toc_lookup, page_lookup))
    findings.extend(check_optional_sections(blocks))
    findings.extend(check_abbreviations_list(blocks))
    findings.extend(check_terms_list(blocks))
    findings.extend(check_formulae(blocks))
    findings.extend(check_sources(blocks, toc_lookup))
    findings.extend(check_lists(blocks))
    add_pages_to_findings(findings, page_lookup)

    severity_counts = Counter(f["severity"] for f in findings)
    rule_counts = Counter(f["rule"] for f in findings)
    return {
        "status": "completed_with_findings" if any(f["severity"] in {"error", "warning"} for f in findings) else "completed",
        "file": {"name": path.name, "path": str(path), "size_bytes": path.stat().st_size, "suffix": path.suffix.lower()},
        "statistics": {**model["stats"], "finding_count": len(findings)},
        "summary": {
            "errors": severity_counts.get("error", 0),
            "warnings": severity_counts.get("warning", 0),
            "info": severity_counts.get("info", 0),
            "by_rule": dict(rule_counts),
        },
        "document": {
            "structure_overview": extract_structure_overview(blocks, page_lookup),
            "sections": sections,
            "raw_toc": raw_toc_entries,
        },
        "findings": findings,
    }
