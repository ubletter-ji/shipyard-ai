"""경영계획 ↔ 투자신청서 매칭 (구현시나리오_보완.xlsx '계획 매칭 규칙' M-0~M-9).

경영계획.xlsx 열: 계획 NO | 투자명 | 투자유형 | 장비 종류 | 위치(AREA) | 수량 | 계획금액(백만원)
집행이력.xlsx 열: 계획 NO | 신청일 | 투자명 | 집행금액(백만원)
"""
from __future__ import annotations

import re
from pathlib import Path

from common import fmt_num, norm, to_number

계획외_표기 = {"", "계획外", "계획외", "없음", "-", "해당없음"}


def _unit(name: str) -> str:
    m = re.search(r"\d\s*(대|개|기|set|SET|식|EA|ea)", name or "")
    return m.group(1) if m else "대"


def load_plans(path: Path) -> list[dict]:
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h or "").strip() for h in rows[0]]

    def g(r, *names):
        for n in names:
            for i, h in enumerate(hdr):
                if norm(n) in norm(h):
                    return r[i]
        return None

    plans = []
    for r in rows[1:]:
        if not g(r, "계획 NO"):
            continue
        plans.append({
            "계획NO": str(g(r, "계획 NO")).strip(),
            "투자명": str(g(r, "투자명") or "").strip(),
            "투자유형": str(g(r, "투자유형") or "").strip(),
            "장비종류": str(g(r, "장비 종류", "장비종류") or "").strip(),
            "AREA": str(g(r, "AREA", "위치") or "").strip(),
            "수량": to_number(g(r, "수량")),
            "계획금액": to_number(g(r, "계획금액", "금액")) or 0.0,
        })
    return plans


def load_history(path: Path | None) -> list[dict]:
    if not path or not Path(path).exists():
        return []
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [norm(str(h or "")) for h in rows[0]]
    ino = next(i for i, h in enumerate(hdr) if "계획no" in h)
    iamt = next(i for i, h in enumerate(hdr) if "금액" in h)
    inm = next((i for i, h in enumerate(hdr) if "투자명" in h), None)
    return [{"계획NO": str(r[ino]).strip(), "집행금액": to_number(r[iamt]) or 0.0,
             "투자명": r[inm] if inm is not None else ""} for r in rows[1:] if r[ino]]


def _candidates(app: dict, plans: list[dict]) -> list[dict]:
    """투자명에 장비 종류가 들어간 계획. '용접기'와 '로봇용접기'처럼 겹치면 더 긴 쪽만 남긴다."""
    c = [p for p in plans if p["장비종류"] and norm(p["장비종류"]) in norm(app["투자명"])]
    return [p for p in c if not any(q is not p and norm(p["장비종류"]) in norm(q["장비종류"])
                                     and norm(p["장비종류"]) != norm(q["장비종류"]) for q in c)]


def match(apps: list[dict], plans: list[dict], history: list[dict] | None = None,
          show_unrequested: bool = True) -> dict:
    """apps: [{투자명, 투자유형, 투자금액, AREA, 계획NO, 수량}] → {rows, 미신청}"""
    history = history or []
    by_no = {norm(p["계획NO"]): p for p in plans}
    used: set[str] = set()
    rows = []

    for a in apps:
        rules, notes, check = [], [], False
        plan = None
        no = (a.get("계획NO") or "").strip()
        실행 = a.get("투자금액") or 0.0

        if norm(no) in by_no:                                   # M-0
            plan = by_no[norm(no)]
            rules.append("M-0")
            if plan["장비종류"] and norm(plan["장비종류"]) not in norm(a["투자명"]):
                notes.append(f"계획 {plan['계획NO']}({plan['투자명']})와 장비 종류 다름 — 계획 NO 확인 필요")
                check = True
        else:
            cands = _candidates(a, [p for p in plans if norm(p["계획NO"]) not in used])
            explicit_out = norm(no) in {norm(x) for x in 계획외_표기} and no != ""
            if no and not explicit_out:
                notes.append(f"계획 NO '{no}' 가 경영계획에 없음")
                check = True
            if len(cands) == 0:                                 # M-4
                rules.append("M-4")
                notes.insert(0, "계획 外")
                check = True
            elif len(cands) >= 2:                               # M-5
                rules.append("M-5")
                notes.insert(0, "후보 " + ", ".join(p["계획NO"] for p in cands))
                check = True
            elif explicit_out:
                rules.append("M-4")
                notes.insert(0, "계획 外")
                notes.append(f"유사 계획 {cands[0]['계획NO']}({cands[0]['투자명']}) 있음")
                check = True
            else:                                               # 매칭 제안
                plan = cands[0]
                rules.append("M-0(제안)")
                notes.append(f"계획 NO 미기재 → {plan['계획NO']} 제안")
                check = True

        if plan:
            used.add(norm(plan["계획NO"]))
            기집행 = sum(h["집행금액"] for h in history if norm(h["계획NO"]) == norm(plan["계획NO"]))
            diffs = []
            if plan["AREA"] and a.get("AREA") and norm(plan["AREA"]) != norm(a["AREA"]):       # M-2
                rules.append("M-2")
                diffs.append(f"계획 {plan['AREA']} → 실행 {a['AREA']} 로 변경")
            if 기집행 == 0 and plan["수량"] is not None and a.get("수량") is not None \
                    and plan["수량"] != a["수량"]:                                                # M-3
                rules.append("M-3")
                u = _unit(a["투자명"])
                diffs.append(f"계획 {fmt_num(plan['수량'])}{u} → 실행 {fmt_num(a['수량'])}{u}로 변경")
            if plan["투자유형"] and a.get("투자유형") and plan["투자유형"] != a["투자유형"]:      # M-9
                rules.append("M-9")
                diffs.append(f"유형 {plan['투자유형']} → {a['투자유형']} 변경")
                check = True
            if 기집행 > 0:                                                                       # M-7
                rules.append("M-7")
                누적 = 기집행 + 실행
                diffs.append(f"누적 {fmt_num(누적)} / 계획 {fmt_num(plan['계획금액'])} "
                             f"(잔여 {fmt_num(plan['계획금액'] - 누적)})")
            if 기집행 + 실행 > plan["계획금액"]:                                                  # M-6
                rules.append("M-6")
                diffs.append("계획 금액 초과")
            if not diffs:                                                                        # M-1
                rules.append("M-1")
            notes = diffs + notes

        계획금액 = plan["계획금액"] if plan else 0.0
        rows.append({
            "신청서ID": a.get("신청서ID", ""),
            "투자명": a["투자명"],
            "투자유형": a.get("투자유형", ""),
            "계획NO": plan["계획NO"] if plan else (no or "계획外"),
            "경영계획": 계획금액,
            "실행금액": 실행,
            "차이금액": 실행 - 계획금액,
            "비고": " / ".join(notes),
            "규칙": rules,
            "확인필요": check,
        })

    미신청 = []
    for p in plans:
        if norm(p["계획NO"]) in used:
            continue
        기집행 = sum(h["집행금액"] for h in history if norm(h["계획NO"]) == norm(p["계획NO"]))
        미신청.append({"계획NO": p["계획NO"], "투자명": p["투자명"], "투자유형": p["투자유형"],
                    "경영계획": p["계획금액"], "기집행": 기집행, "비고": "미신청", "규칙": ["M-8"]})
    return {"rows": rows, "미신청": 미신청 if show_unrequested else []}
