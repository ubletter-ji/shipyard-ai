"""다른 양식의 투자신청서 → 표준양식 (양식만 바꾸고 내용은 그대로).

어떤 글이 어느 칸인지 정하는 일(매핑)은 에이전트가 하고(.claude/skills/invest-standardize/SKILL.md),
이 스크립트는 ① 원본 글자 덤프 ② 매핑대로 표준양식 채우기 ③ 내용 보존 검증 을 한다.

사용:
  python scripts/standardize.py detect [파일...]          표준양식인지 (인자 없으면 inbox 전체)
  python scripts/standardize.py dump <원본.pptx>          원본의 모든 글자를 줄 단위로 (매핑 작성용)
  python scripts/standardize.py build <매핑.json> [--out 폴더]
                                                         표준양식 변환본 + 기록(json) 생성, 검증 자동 실행
  python scripts/standardize.py verify <매핑.json>        원본 대비 추가(지어낸 글)·누락 검사만
  python scripts/standardize.py targets                   점검 대상 목록 (표준양식 원본 + 검증 통과 변환본)

매핑 JSON (값은 전부 원본 글자 그대로 — 고쳐 쓰기·요약·단위 환산 금지):
{
  "원본": "inbox/xxx.pptx",
  "투자명": "...", "투자효과": "...",
  "상세": {"투자시작": "", "투자경과": "", "투자금액(백만원)": "", "투자유형": "", "현수량": "", "필요수량": "",
          "투자수량": "", "부 유형": "", "AREA": "", "계획 NO": "", "선정 업체": "", "견적 수": ""},
  "부가": {"견적서 비교": "...", "내년도 부하·조업도": "...", "현 장비 재원 소개": "...", "생산성 향상 산출": "..."},
  "추가": [{"제목": "견적서 비교", "표": [["업체","금액(백만원)","설립연도","비고"], [...]], "본문": "..."}],
  "라벨": ["1. 투자 건명:", ...],          원본에서 '칸 이름' 역할만 하던 글 (표준양식의 칸 이름으로 대체됨)
  "기타원문": ["..."],                     어느 칸에도 맞지 않는 내용 → '원문 (양식 외 내용)' 슬라이드에 그대로
  "원본슬라이드첨부": [3]                  이미지·도표처럼 글자로 옮길 수 없는 슬라이드 → 원본 그대로 복사
}
"""
from __future__ import annotations

import argparse
import contextlib
import io
import re
import sys
from pathlib import Path

from pptx import Presentation

from common import OUT, ROOT, norm, read_json, write_json
from extract import extract, file_hash

TEMPLATE = ROOT / "표준양식_수정.pptx"
INBOX = ROOT / "inbox"
STD_OUT = OUT / "표준화"
상세_키 = ["투자시작", "투자경과", "투자금액(백만원)", "투자유형", "현수량", "필요수량",
          "투자수량", "부 유형", "AREA", "계획 NO", "선정 업체", "견적 수"]
기타_제목 = "원문 (양식 외 내용)"
표준_머리글 = {"업체", "금액(백만원)", "설립연도", "비고"}


# ---------------- 원본 읽기
def _shape_lines(shape, slide_no):
    if shape.shape_type == 6:  # GROUP
        for sh in shape.shapes:
            yield from _shape_lines(sh, slide_no)
        return
    if shape.shape_type == 13:  # PICTURE
        yield slide_no, "이미지", "[이미지 — 글자로 옮길 수 없음: 원본슬라이드첨부 대상]"
        return
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            if p.text.strip():
                yield slide_no, "글", p.text.strip()
    if getattr(shape, "has_table", False) and shape.has_table:
        for ri, row in enumerate(shape.table.rows):
            for ci, cell in enumerate(row.cells):
                for p in cell.text_frame.paragraphs:
                    if p.text.strip():
                        yield slide_no, f"표 {ri + 1}행 {ci + 1}열", p.text.strip()
    if getattr(shape, "has_chart", False) and shape.has_chart:
        yield slide_no, "도표", "[도표 — 원본슬라이드첨부 대상]"


def source_lines(path: Path) -> list[tuple[int, str, str]]:
    prs = Presentation(str(path))
    out = []
    for no, slide in enumerate(prs.slides, start=1):
        for sh in slide.shapes:
            out.extend(_shape_lines(sh, no))
    return out


def is_standard(path: Path) -> bool:
    """표준양식 구조(투자명 칸이 있는 메인 표 + 투자금액(백만원) 칸이 있는 상세 표)인지."""
    cells = set()
    for slide in Presentation(str(path)).slides:
        for sh in slide.shapes:
            if sh.has_table:
                cells |= {norm(c.text) for r in sh.table.rows for c in r.cells}
    return "투자명" in cells and norm("투자금액(백만원)") in cells


# ---------------- 검증
_ws = re.compile(r"\s+")


def _n(s: str) -> str:
    return _ws.sub(" ", s or "").strip()


def _value_lines(m: dict) -> list[str]:
    vals = [m.get("투자명", ""), m.get("투자효과", "")]
    vals += list((m.get("상세") or {}).values())
    vals += list((m.get("부가") or {}).values())
    for ex in m.get("추가") or []:
        rows = ex.get("표") or []
        body_rows = rows[1:] if rows and {_n(c) for c in rows[0]} <= 표준_머리글 else rows
        vals += [c for r in body_rows for c in r]
        vals.append(ex.get("본문", ""))
    vals += m.get("기타원문") or []
    lines = []
    for v in vals:
        lines += [_n(x) for x in str(v or "").splitlines() if _n(x)]
    return lines


def _leftover_is_content(s: str) -> bool:
    s = re.sub(r"^\s*(\d+|[①-⑳]|[가-하])[.)]\s*", "", s)          # 번호 매김 "1." "①" "가)"
    s = re.sub(r"[\s:：\-–·•▶■□○●※,./|()\[\]<>『』「」~]", "", s)
    return bool(s)


def verify(m: dict) -> dict:
    src = Path(m["원본"])
    src = src if src.is_absolute() else ROOT / src
    lines = source_lines(src)
    full = _n(" ".join(t for _, _, t in lines))
    attach = set(m.get("원본슬라이드첨부") or [])

    # 1) 추가: 매핑 값의 모든 줄이 원본에 그대로 있어야 한다
    added = [v for v in _value_lines(m) if v not in full]

    # 2) 누락: 원본의 모든 줄이 값·라벨·기타원문 중 하나로 덮여야 한다
    pieces = sorted(set(_value_lines(m)) | {_n(x) for x in m.get("라벨") or [] if _n(x)}, key=len, reverse=True)
    missing = []
    for no, kind, text in lines:
        if no in attach:
            continue
        if kind in ("이미지", "도표"):
            missing.append(f"슬라이드 {no}: {text}")
            continue
        rem = _n(text)
        for p in pieces:
            rem = rem.replace(p, " ")
        if _leftover_is_content(rem):
            missing.append(f"슬라이드 {no} {kind}: \"{text}\" (남은 글자: \"{_n(rem)}\")")

    warn = []
    amt = _n((m.get("상세") or {}).get("투자금액(백만원)", ""))
    if amt and re.search(r"억|천만|만원|(?<!백만)원", amt) and "백만원" not in amt:
        warn.append(f"투자금액 \"{amt}\" 단위가 백만원이 아님 — 환산하지 않았음 (점검 단계에서 금액 확인 필요)")
    lab_bad = [x for x in m.get("라벨") or [] if _n(x) and _n(x) not in full]
    added += [f"(라벨) {x}" for x in lab_bad]
    return {"추가": added, "누락": missing, "경고": warn, "통과": not added and not missing}


# ---------------- 표준양식 채우기
def _fill(prs, m: dict):
    from make_samples import add_extra_slide, remove_slide, set_cell
    remove_slide(prs, 1)  # 작성 안내
    slide = prs.slides[0]
    부가 = m.get("부가") or {}
    부가_text = "\n\n".join(f"- {k}\n{v.strip()}" for k, v in 부가.items() if str(v or "").strip())
    for shp in slide.shapes:
        if not shp.has_table:
            continue
        texts = [[c.text.strip() for c in r.cells] for r in shp.table.rows]
        if any(c == "투자명" for r in texts for c in r):
            for ri, r in enumerate(texts):
                for ci, c in enumerate(r):
                    if c == "투자명":
                        set_cell(shp.table.cell(ri, ci + 1), m.get("투자명", ""))
                    elif c == "투자효과":
                        set_cell(shp.table.cell(ri, ci + 1), m.get("투자효과", ""))
                    elif c.replace("\n", "").replace(" ", "") == "투자상세부가":
                        set_cell(shp.table.cell(ri, ci + 1), 부가_text)
        elif any("투자금액" in c for r in texts for c in r):
            for ri in range(0, len(texts) - 1, 2):
                for ci, key in enumerate(texts[ri]):
                    if key in 상세_키:  # 양식의 예시 값은 모두 지우고 매핑 값만 넣는다
                        set_cell(shp.table.cell(ri + 1, ci), (m.get("상세") or {}).get(key, ""))
    for ex in m.get("추가") or []:
        add_extra_slide(prs, ex["제목"], ex.get("본문", ""), ex.get("표") or None)
    if m.get("기타원문"):
        add_extra_slide(prs, 기타_제목, "\n".join(m["기타원문"]), None)


def build(mapping_path: Path, out_dir: Path = STD_OUT) -> dict:
    m = read_json(mapping_path)
    src = Path(m["원본"])
    src = src if src.is_absolute() else ROOT / src
    if is_standard(src):
        sys.exit(f"{src.name} 는 이미 표준양식입니다. 변환하지 않습니다.")
    for k in (m.get("상세") or {}):
        if k not in 상세_키:
            sys.exit(f"상세 칸 이름이 표준양식에 없음: {k}  (가능: {', '.join(상세_키)})")

    prs = Presentation(str(TEMPLATE))
    _fill(prs, m)
    if m.get("원본슬라이드첨부"):
        from meeting import copy_slide
        sp = Presentation(str(src))
        for no in m["원본슬라이드첨부"]:
            copy_slide(sp.slides[no - 1], prs)
    prs.core_properties.subject = f"표준양식 변환본 (원본: {src.name})"
    out_dir.mkdir(parents=True, exist_ok=True)
    dst = out_dir / f"{src.stem}.pptx"
    prs.save(str(dst))

    res = verify(m)
    # 왕복 확인: 변환본을 점검기와 같은 방식으로 읽었을 때 매핑 값이 그대로 나오는가
    with contextlib.redirect_stdout(io.StringIO()):
        d = extract(dst)
    rt = []
    if _n(d["투자명"]) != _n(m.get("투자명", "")):
        rt.append(f"투자명: \"{d['투자명']}\"")
    for k, v in (m.get("상세") or {}).items():
        if _n(d["상세"].get(k, "")) != _n(v):
            rt.append(f"상세 {k}: \"{d['상세'].get(k, '')}\" ≠ \"{v}\"")
    for k, v in (m.get("부가") or {}).items():
        if _n(v) and _n(v) not in _n(d["부가"].get(k, "")):
            rt.append(f"부가 {k} 가 변환본에서 다르게 읽힘")
    res["왕복불일치"] = rt
    res["통과"] = res["통과"] and not rt
    record = {"원본경로": str(src.resolve()), "원본해시": file_hash(src), "변환본": str(dst.resolve()),
              "매핑": str(Path(mapping_path).resolve()), **res}
    write_json(dst.with_suffix(".json"), record)
    return record


def targets(inbox: Path = INBOX) -> tuple[list[Path], list[tuple[Path, str]]]:
    """점검 대상: 표준양식 원본은 그대로, 다른 양식은 검증 통과한 최신 변환본. 나머지는 (원본, 사유)."""
    ok, bad = [], []
    for p in sorted(inbox.glob("*.pptx")):
        if p.name.startswith("~$"):
            continue
        if is_standard(p):
            ok.append(p)
            continue
        side = STD_OUT / f"{p.stem}.json"
        if not side.exists():
            bad.append((p, "표준양식이 아니고 변환본 없음 — invest-standardize 로 변환 필요"))
            continue
        rec = read_json(side)
        if rec.get("원본해시") != file_hash(p):
            bad.append((p, "변환 후 원본이 바뀜 — 다시 변환 필요"))
        elif not rec.get("통과"):
            bad.append((p, f"변환 검증 실패 (추가 {len(rec.get('추가', []))} / 누락 {len(rec.get('누락', []))}"
                           f" / 왕복 {len(rec.get('왕복불일치', []))})"))
        else:
            ok.append(STD_OUT / f"{p.stem}.pptx")
    return ok, bad


def _print_result(rec: dict):
    print(f"[변환] {Path(rec['변환본']).name if '변환본' in rec else ''}  → {'통과' if rec['통과'] else '실패'}")
    for k in ("추가", "누락", "왕복불일치", "경고"):
        for x in rec.get(k, []):
            print(f"   {k}: {x}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["detect", "dump", "build", "verify", "targets"])
    ap.add_argument("paths", nargs="*")
    ap.add_argument("--out", default=str(STD_OUT))
    a = ap.parse_args()
    if a.cmd == "detect":
        for p in [Path(x) for x in a.paths] or sorted(INBOX.glob("*.pptx")):
            print(f"{'표준양식' if is_standard(p) else '다른 양식'}\t{p}")
    elif a.cmd == "dump":
        for no, kind, text in source_lines(Path(a.paths[0])):
            print(f"S{no}\t{kind}\t{text}")
    elif a.cmd == "build":
        rec = build(Path(a.paths[0]), Path(a.out))
        _print_result(rec)
        sys.exit(0 if rec["통과"] else 1)
    elif a.cmd == "verify":
        rec = verify(read_json(Path(a.paths[0])))
        _print_result(rec)
        sys.exit(0 if rec["통과"] else 1)
    elif a.cmd == "targets":
        ok, bad = targets()
        for p in ok:
            print(p)
        for p, why in bad:
            print(f"[제외] {p.name}: {why}", file=sys.stderr)


if __name__ == "__main__":
    main()
