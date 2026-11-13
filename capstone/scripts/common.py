"""공통 유틸: 경로, 설정, 숫자 파싱."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config" / "기준값.json"
OUT = ROOT / "out"

부가항목 = ["견적서 비교", "내년도 부하·조업도", "현 장비 재원 소개", "생산성 향상 산출"]


def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


_NUM = re.compile(r"-?\d[\d,]*(?:\.\d+)?")


def to_number(text) -> float | None:
    """'499.5', '1,200', '5대' → 숫자. 숫자가 없으면 None."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    m = _NUM.search(str(text))
    if not m:
        return None
    return float(m.group().replace(",", ""))


def numbers_in(text: str) -> list[float]:
    return [float(x.replace(",", "")) for x in _NUM.findall(text or "")]


def norm(s: str | None) -> str:
    """비교용 정규화: 공백 제거, 소문자."""
    return re.sub(r"\s+", "", (s or "")).lower()


def fmt_num(x: float | None) -> str:
    if x is None:
        return "-"
    return f"{x:,.0f}" if float(x).is_integer() else f"{x:,.1f}"


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def read_json(path: Path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)
