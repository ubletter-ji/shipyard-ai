"""투자신청서(.pptx, 표준양식) → 구조화 JSON.

사용: python scripts/extract.py inbox/신청서.pptx [...]  → out/extract/<신청서ID>.json
"""
from __future__ import annotations

import hashlib
import re
import sys
from pathlib import Path

from pptx import Presentation

from common import OUT, 부가항목, norm, to_number, write_json

# 표준양식에 미리 들어있는 작성 안내 문구 (신청자가 지우지 않아도 '내용'으로 보지 않음)
_안내문구 = [
    "※ 투자유형에 맞는 필수 항목을 작성",
    "업체별 금액·설립연도, 선정 사유 기재",
    "현재 부하와 설치 후 수치를 숫자로",                      # 2026-11-11 양식
    "운용연수, 수량, 교체 사유",
    "금액 / JMH / 인원 감소 / 공정 단축 효과와 산출 근거",
    "내년 부하·조업도를 현재 → 설치 후 숫자로",               # 2026-11-12 양식
    "운용연수, 노후 수량, 교체 수량, 교체 사유",
    "금액 · JMH · 인력 감소 (3가지 모두) + 산출 근거",
    "※ 내용이 길면 뒤에 슬라이드를 추가하고, 슬라이드 제목을 항목명과 같게 작성",
]
_무시할_슬라이드_제목 = ["투자신청서 작성 안내"]

상세_키 = ["투자시작", "투자경과", "투자금액(백만원)", "투자유형", "현수량", "필요수량",
          "투자수량", "부 유형", "AREA", "계획 NO", "선정 업체", "견적 수"]


def file_hash(path: Path) -> str:
    """원본 지문. 점검 후 신청서가 바뀌었는지 회의자료 단계에서 확인하는 데 쓴다."""
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _cell_text(cell) -> str:
    return "\n".join(p.text for p in cell.text_frame.paragraphs).strip()


def _table_rows(shape) -> list[list[str]]:
    return [[_cell_text(c) for c in row.cells] for row in shape.table.rows]


def _shape_text(shape) -> str:
    if shape.has_text_frame:
        return shape.text_frame.text.strip()
    if shape.has_table:
        return "\n".join(" | ".join(r) for r in _table_rows(shape))
    return ""


def _is_guidance(line: str) -> bool:
    s = line.strip().lstrip("-·• ").strip()
    if not s:
        return True
    return any(norm(s) == norm(g.lstrip("-·• ")) for g in _안내문구)


def _section_header(line: str) -> str | None:
    """'- 견적서 비교 (2곳 이상) [공통 필수]' → '견적서 비교'"""
    s = line.strip().lstrip("-·•■□▶ ").strip()
    for name in 부가항목:
        if norm(s).startswith(norm(name)):
            return name
    return None


def split_부가(text: str) -> dict[str, str]:
    """부가 칸 텍스트를 항목별로 나눈다. 안내 문구만 있는 항목은 빈 문자열."""
    sections: dict[str, list[str]] = {}
    current = None
    for line in (text or "").splitlines():
        header = _section_header(line)
        if header:
            current = header
            sections.setdefault(current, [])
            # 헤더 줄 뒤에 ':' 로 바로 내용을 적은 경우
            rest = re.split(r"[:：]", line, maxsplit=1)
            if len(rest) == 2 and rest[1].strip() and "필수]" not in rest[1]:
                sections[current].append(rest[1].strip())
            continue
        if current and not _is_guidance(line):
            sections[current].append(line.strip())
    return {k: "\n".join(v).strip() for k, v in sections.items()}


선정표시 = re.compile(r"선정|채택|낙찰")


def parse_quotes(text: str, table_rows: list[list[str]] | None) -> list[dict]:
    """견적서 비교 → [{업체, 금액, 설립연도, 선정, 비고}]"""
    quotes: list[dict] = []
    if table_rows and len(table_rows) >= 2:
        header = [norm(h) for h in table_rows[0]]

        def col(*names):
            for i, h in enumerate(header):
                if any(norm(n) in h for n in names):
                    return i
            return None

        ci, ca, cy, cn = col("업체"), col("금액"), col("설립"), col("비고", "선정")
        for r in table_rows[1:]:
            if ci is None or not r[ci].strip():
                continue
            note = r[cn] if cn is not None else ""
            quotes.append({
                "업체": r[ci].strip(),
                "금액": to_number(r[ca]) if ca is not None else None,
                "설립연도": int(to_number(r[cy])) if cy is not None and to_number(r[cy]) else None,
                "선정": bool(선정표시.search(note)),
                "비고": note.strip(),
            })
        if quotes:
            return quotes
    # 텍스트 형식: "- A사: 499.5백만원, 설립 2008년 (선정)"
    pat = re.compile(r"^[-·•\s]*([^:：,]+?)\s*[:：]\s*([\d,\.]+)\s*(?:백만원)?(.*)$")
    for line in (text or "").splitlines():
        m = pat.match(line)
        if not m:
            continue
        name, amount, rest = m.groups()
        if re.match(r"(선정|채택|낙찰)", name.strip()):
            continue
        y = re.search(r"설립\s*(\d{4})", rest)
        quotes.append({
            "업체": name.strip(),
            "금액": to_number(amount),
            "설립연도": int(y.group(1)) if y else None,
            "선정": bool(선정표시.search(rest)),
            "비고": rest.strip(" ,()"),
        })
    return quotes


def extract(path: Path) -> dict:
    prs = Presentation(str(path))
    data: dict = {
        "신청서ID": path.stem,
        "파일": path.name,
        "경로": str(path.resolve()),
        "원본해시": file_hash(path),
        "투자명": "",
        "투자효과": "",
        "상세": {},
        "부가": {},
        "부가_출처": {},
        "견적": [],
        "선정사유": "",
        "슬라이드": [],
        "경고": [],
    }
    부가_raw = ""
    extra_tables: dict[str, list[list[str]]] = {}

    for no, slide in enumerate(prs.slides, start=1):
        texts = [(_shape_text(sh), sh) for sh in slide.shapes]
        full = "\n".join(t for t, _ in texts if t)
        title = ""
        if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
            title = slide.shapes.title.text_frame.text.strip()
        elif texts:
            first_tb = next((t for t, sh in texts if sh.has_text_frame and t), "")
            title = first_tb.splitlines()[0] if first_tb else ""
        if any(norm(title) == norm(x) for x in _무시할_슬라이드_제목):
            continue
        data["슬라이드"].append({"번호": no, "제목": title, "텍스트": full})

        if any(sh.shape_type == 13 for _, sh in texts):  # PICTURE
            data["경고"].append(f"슬라이드 {no}: 이미지가 있음 — 이미지 속 글자는 읽지 않음")

        for _, sh in texts:
            if not sh.has_table:
                continue
            rows = _table_rows(sh)
            flat = [c for r in rows for c in r]
            # 메인 표: 투자명 / 투자효과 / 부가
            if any(norm(c) == "투자명" for c in flat):
                for r in rows:
                    for i, c in enumerate(r):
                        if norm(c) == "투자명" and i + 1 < len(r):
                            data["투자명"] = r[i + 1].strip()
                        if norm(c) == "투자효과" and i + 1 < len(r):
                            data["투자효과"] = r[i + 1].strip()
                        if norm(c).startswith("투자상세") and "부가" in c and i + 1 < len(r):
                            부가_raw = r[i + 1]
                            data["부가_슬라이드"] = no
            # 상세 표: 헤더 행 / 값 행 반복
            if any(norm(c) == norm("투자금액(백만원)") for c in flat):
                for ri in range(0, len(rows) - 1, 2):
                    for k, v in zip(rows[ri], rows[ri + 1]):
                        if k.strip():
                            data["상세"][k.strip()] = v.strip()
                data["상세_슬라이드"] = no
            # 추가 슬라이드의 표 (예: 견적서 비교)
            sec = _section_header(title) if title else None
            if sec:
                extra_tables[sec] = rows

        # 추가 슬라이드: 제목 = 부가 항목명
        sec = _section_header(title) if title else None
        if sec and no != data.get("부가_슬라이드"):
            body = "\n".join(t for t, sh in texts if t and t.splitlines()[0] != title)
            prev = data["부가"].get(sec, "")
            data["부가"][sec] = (prev + "\n" + body).strip()
            data["부가_출처"][sec] = f"슬라이드 {no}"

    for sec, body in split_부가(부가_raw).items():
        if body:
            prev = data["부가"].get(sec, "")
            data["부가"][sec] = (body + ("\n" + prev if prev else "")).strip()
            data["부가_출처"].setdefault(sec, f"슬라이드 {data.get('부가_슬라이드', 1)} 부가 칸")

    q_text = data["부가"].get("견적서 비교", "")
    data["견적"] = parse_quotes(q_text, extra_tables.get("견적서 비교"))
    m = re.search(r"(?:선정|채택)\s*사유\s*[:：]\s*(.+)", q_text)
    data["선정사유"] = m.group(1).strip() if m else ""

    # 다른 양식에서 표준양식으로 변환한 파일이면 원본 정보를 붙인다 (out/표준화/<이름>.json)
    side = path.with_suffix(".json")
    if side.exists():
        import json
        info = json.loads(side.read_text(encoding="utf-8"))
        if "원본해시" in info:
            data["변환원본"] = {"경로": info["원본경로"], "해시": info["원본해시"], "통과": info.get("통과", False)}
            data["경고"].append(f"표준양식 변환본 — 원본 {Path(info['원본경로']).name}")

    missing = [k for k in 상세_키 if k not in data["상세"]]
    if missing:
        data["경고"].append("상세 표에서 찾지 못한 칸: " + ", ".join(missing))
    return data


def main(paths: list[str]) -> list[Path]:
    outs = []
    for p in paths:
        d = extract(Path(p))
        out = OUT / "extract" / f"{d['신청서ID']}.json"
        write_json(out, d)
        outs.append(out)
        print(f"[추출] {p} → {out.relative_to(OUT.parent)}  ({d['투자명']})")
        for w in d["경고"]:
            print(f"   ! {w}")
    return outs


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("사용: python scripts/extract.py <신청서.pptx> [...]")
    main(sys.argv[1:])
