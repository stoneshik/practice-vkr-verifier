"""
Формирует внутреннюю модель документа: извлекает абзацы, таблицы и их свойства для последующего анализа
"""
from typing import Any, Dict, Optional


def make_finding(
    rule: str,
    severity: str,
    message: str,
    actual: Any = None,
    expected: Any = None,
    paragraph: Optional[int] = None,
    block: Optional[int] = None,
    section: Optional[int] = None,
    page: Optional[int] = None,
    recommendation: Optional[str] = None,
    evidence: Optional[Any] = None,
) -> Dict[str, Any]:
    """Собирает объект одного найденного несоответствия"""
    location = {k: v for k, v in {
        "paragraph": paragraph,
        "block": block,
        "section": section,
        "page": page,
    }.items() if v is not None}
    finding = {
        "rule": rule,
        "severity": severity,
        "message": message,
        "location": location,
        "actual": actual,
        "expected": expected,
        "recommendation": recommendation,
    }
    if evidence is not None:
        finding["evidence"] = evidence
    return finding
