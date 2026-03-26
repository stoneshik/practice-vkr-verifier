"""
Содержит константы: регулярные выражения, допустимые форматы, списки разделов и настройки анализа
"""
import re


# Настройки
ALLOWED_EXTENSIONS = {".docx"}
ALLOWED_MIME_TYPES = {"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}

REQUIRED_SECTIONS = [
    "СОДЕРЖАНИЕ",
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
]
OPTIONAL_SECTIONS = [
    "СПИСОК СОКРАЩЕНИЙ И УСЛОВНЫХ ОБОЗНАЧЕНИЙ",
    "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ",
    "СПИСОК ИЛЛЮСТРАТИВНОГО МАТЕРИАЛА",
]
STRUCTURAL_SECTIONS = set(REQUIRED_SECTIONS + OPTIONAL_SECTIONS)
COMMON_UNITS = [
    "мм", "см", "м", "км", "г", "кг", "мг", "л", "мл", "Вт", "кВт", "м2", "м3", "%", "°C"
]
XML_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

# Регулярные выражения
FIGURE_CAPTION_RE = re.compile(r"^Рисунок\s+(\d+(?:\.\d+)*|[А-ЯA-Z])\s*[—–-]\s+.+$", re.IGNORECASE)
TABLE_CAPTION_RE = re.compile(r"^Таблица\s+(\d+(?:\.\d+)*|[А-ЯA-Z])\s*[—–-]\s+.+$", re.IGNORECASE)
APPENDIX_RE = re.compile(r"^ПРИЛОЖЕНИЕ\s+([А-ЯA-Z])(?:\s|$)", re.IGNORECASE)
HEADING_NUMBER_RE = re.compile(r"^(\d+(?:\.\d+)*)(\.?)(?:\s+)(.+)$")
HEADING_NUMBER_BAD_TRAILING_DOT_RE = re.compile(r"^(\d+)\.\s+.+$")
TOC_LINE_RE = re.compile(r"^(.+?)(?:\.{2,}|\s{2,}|\t+)(\d+)\s*$")
LIST_MARKER_RE = re.compile(r"^(\s*)([-•*]|\d+\)|[а-яa-z]\))\s+(.+)$", re.IGNORECASE)
SOURCE_ENTRY_RE = re.compile(r"^\s*(\d+)\.\s+.+$")
FORMULA_NUMBER_RE = re.compile(r"\(\s*\d+(?:\.\d+)?\s*\)\s*$")
OMML_RE = re.compile(r"<m:oMath|<m:oMathPara")
PAGE_BREAK_RE = re.compile(r"w:type=\"page\"|w:pageBreakBefore")
PAGE_FIELD_RE = re.compile(r"<w:instrText[^>]*>PAGE</w:instrText>", re.IGNORECASE)
FIELD_PAGE_RE = re.compile(r"\bPAGE\b", re.IGNORECASE)
