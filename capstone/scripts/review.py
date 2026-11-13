"""투자신청서 검토 규칙 엔진 (구현시나리오_보완.xlsx '검토 기준' R-101~R-401).

사용:
  python scripts/review.py check inbox/*.pptx    # 추출 + 규칙 판정 → out/review/<ID>.json
  python scripts/review.py report                # out/review/*.json → out/검토결과.xlsx
  python scripts/review.py pending               # 에이전트가 판단해야 할(AI대기) 행 목록
  python scripts/review.py judge <ID> <규칙ID> <판정> "<근거>" ["<보완요청문구>"] [--type 유형]
                                                 # 에이전트 판단 기록 (AI대기 행만 가능)

판정 값: 충족 / 보완 / 판단 필요 / 해당 없음
AI대기=True 인 행은 검토 에이전트가 신청서 내용을 읽고 판정·근거를 채운다.
"""
from __future__ import annotations

import datetime as dt
import re
import sys
from pathlib import Path

from common import OUT, fmt_num, load_config, norm, numbers_in, read_json, to_number, write_json
import extract as ex

충족, 보완, 판단, 해당없음 = "충족", "보완", "판단 필요", "해당 없음"
REVIEW_DIR = OUT / "review"
REPORT = OUT / "검토결과.xlsx"
확인_OK = {"o", "○", "0", "ㅇ", "확인"}

규칙명 = {
    "R-100": "표준양식",
    "R-101": "투자명", "R-102": "투자 수량", "R-103": "투자유형", "R-104": "투자금액",
    "R-105": "견적서", "R-106": "선정 사유", "R-107": "신생 업체", "R-108": "투자경과",
    "R-201": "유형 일치", "R-202": "부 유형", "R-203": "유형별 필수 항목",
    "R-301": "필요성", "R-302": "부하/조업도",
    "R-311": "운용연수", "R-312": "교체 수량", "R-313": "교체 사유", "R-314": "노후 수량",
    "R-321": "정량 효과(금액·JMH·인력)", "R-322": "효과 산출 근거",
}


def _row(rid, 판정, 근거="", 문구="", ai=False):
    return {"규칙ID": rid, "검토항목": 규칙명[rid], "판정": 판정, "근거": 근거,
            "보완요청문구": 문구 if 판정 == 보완 else "", "AI대기": ai, "검토자확인": ""}


def 신생_미선정_문구(cheaper: list[dict], sel: dict, today: dt.date) -> str:
    """최저가 신생 업체를 선정하지 않은 사유 — 신청서에 그대로 쓸 수 있는 문장."""
    names = ", ".join(f"{q['업체']}({q['설립연도']}년 설립, 업력 {today.year - q['설립연도']}년)" for q in cheaper)
    sel_age = f"업력 {today.year - sel['설립연도']}년의 " if sel.get("설립연도") else ""
    return (f"최저가를 제시한 {names}는 설립 5년이 되지 않은 신생 업체로, 장기 A/S·부품 공급의 지속성과 "
            f"재무 안정성을 확인하기 어려워 선정에서 제외하였으며, 설치 후 사후 지원까지 안정적으로 받을 수 있는 "
            f"{sel_age}{sel['업체']}를 선정함.")


def _qty_from_name(name: str) -> float | None:
    m = re.search(r"(\d[\d,]*)\s*(대|개|기|set|SET|식|EA|ea)", name or "")
    return to_number(m.group(1)) if m else None


def check(d: dict, cfg: dict, today: dt.date | None = None) -> dict:
    today = today or dt.date.today()
    s = d["상세"]
    loc = f"슬라이드 {d.get('상세_슬라이드', 1)} 상세 표"
    부가 = d.get("부가", {})
    출처 = d.get("부가_출처", {})
    rows: list[dict] = []
    conv = d.get("변환원본")
    if conv:
        rows.append(_row("R-100", 충족, f"다른 양식 → 표준양식 변환본 (원본 {Path(conv['경로']).name}, "
                                       "내용 보존 검증 통과: 추가 0 / 누락 0)"))
    else:
        rows.append(_row("R-100", 충족, "표준양식"))
    유형 = s.get("투자유형", "").strip()
    부유형 = s.get("부 유형", "").strip()
    금액 = to_number(s.get("투자금액(백만원)"))
    quotes = d.get("견적", [])
    선정명 = s.get("선정 업체", "").strip()
    sel = next((q for q in quotes if q["선정"]), None) or next(
        (q for q in quotes if 선정명 and norm(q["업체"]) == norm(선정명)), None)

    # ---- 검토 1 필수 요소
    if d.get("투자명"):
        rows.append(_row("R-101", 충족, f'투자명 칸 "{d["투자명"]}"'))
    else:
        rows.append(_row("R-101", 보완, "투자명 칸 비어 있음", "투자명 칸에 투자명을 기재해 주십시오."))

    # 수량 (2026-11-13): 현수량 / 필요수량 / 투자수량
    #   노후교체 = 투자수량 ≤ 필요수량 (교체가 필요한 만큼까지)
    #   그 외    = 투자수량 ≤ 필요수량 − 현수량 (부족분만큼만)
    qs = {k: to_number(s.get(k)) for k in ("현수량", "필요수량", "투자수량")}
    if all(v is not None for v in qs.values()):
        현, 필요, 투자 = qs["현수량"], qs["필요수량"], qs["투자수량"]
        base = f"{loc}: 현수량 {fmt_num(현)}, 필요수량 {fmt_num(필요)}, 투자수량 {fmt_num(투자)}"
        if 유형 == "노후교체":
            한도, 식 = 필요, f"노후교체: 투자수량 ≤ 필요수량 {fmt_num(필요)}"
            안내 = f"노후교체 투자는 투자수량이 필요수량({fmt_num(필요)}) 이하여야 합니다."
        else:
            한도, 식 = 필요 - 현, f"투자수량 ≤ 필요 − 현 = {fmt_num(필요 - 현)}"
            안내 = (f"투자수량은 필요수량({fmt_num(필요)}) − 현수량({fmt_num(현)}) = {fmt_num(필요 - 현)} 이하여야 합니다."
                  if 필요 > 현 else
                  f"필요수량({fmt_num(필요)})이 현수량({fmt_num(현)}) 이하라 추가로 투자할 수량이 없습니다.")
        if 투자 <= 0:
            rows.append(_row("R-102", 보완, base + " — 투자수량이 0 이하",
                             "상세 표의 투자수량을 1 이상 숫자로 기재해 주십시오."))
        elif 투자 <= 한도:
            rows.append(_row("R-102", 충족, base + f" — {식}"))
        else:
            rows.append(_row("R-102", 보완, base + f" — {식} 초과",
                             f"{안내} 투자수량({fmt_num(투자)})을 확인해 주십시오."))
    else:
        miss = [k for k, v in qs.items() if v is None]
        rows.append(_row("R-102", 보완, f"{loc}: {', '.join(miss)} 숫자 없음",
                         f"상세 표의 {', '.join(miss)} 칸을 숫자로 기재해 주십시오."))

    if 유형 in cfg["투자유형"]:
        rows.append(_row("R-103", 충족, f'{loc}: 투자유형 "{유형}"'))
    else:
        rows.append(_row("R-103", 보완, f'{loc}: 투자유형 "{유형 or "공란"}"',
                         "투자유형을 생산필수 / 노후교체 / 생산성 향상(자동화) 중 하나로 기재해 주십시오.", ai=True))

    if 금액 is None:
        rows.append(_row("R-104", 보완, f"{loc}: 투자금액 없음", "상세 표의 투자금액(백만원)을 기재해 주십시오."))
    elif not sel or sel.get("금액") is None:
        rows.append(_row("R-104", 보완, f"투자금액 {fmt_num(금액)}, 선정 업체 견적 금액 확인 불가",
                         "견적서 비교에 선정 업체와 견적 금액을 기재해 주십시오."))
    else:
        tol = cfg.get("R104_견적금액_허용차이_퍼센트")
        diff = abs(금액 - sel["금액"]) / sel["금액"] * 100 if sel["금액"] else 100
        base = f"신청 {fmt_num(금액)} vs {sel['업체']} 견적 {fmt_num(sel['금액'])} (차이 {diff:.1f}%)"
        if tol is None:
            rows.append(_row("R-104", 충족 if diff == 0 else 판단, base + " — 허용 차이 [확정 필요]"))
        elif diff <= tol:
            rows.append(_row("R-104", 충족, base))
        else:
            rows.append(_row("R-104", 보완, base, f"신청 금액({fmt_num(금액)}백만원)과 선정 업체 견적 금액"
                             f"({fmt_num(sel['금액'])}백만원)이 다릅니다. 금액을 확인해 주십시오."))

    n = len(quotes)
    stated = to_number(s.get("견적 수"))
    q_src = 출처.get("견적서 비교", "견적서 비교 없음")
    if n >= cfg["R105_최소_견적수"] and (stated is None or int(stated) == n):
        rows.append(_row("R-105", 충족, f"{q_src}: 견적 {n}곳 (" + ", ".join(q["업체"] for q in quotes) + ")"))
    elif n < cfg["R105_최소_견적수"]:
        rows.append(_row("R-105", 보완, f"{q_src}: 견적 {n}곳",
                         f"견적서를 {cfg['R105_최소_견적수']}곳 이상 받아 견적서 비교에 업체별 금액·설립연도를 기재해 주십시오."))
    else:
        rows.append(_row("R-105", 보완, f"상세 표 견적 수 {fmt_num(stated)} ≠ 견적서 비교 {n}곳",
                         "상세 표의 견적 수와 견적서 비교의 업체 수가 다릅니다. 확인해 주십시오."))

    # 견적 선정 (2026-11-13): 최저가가 원칙. 단, 업력 N년(기본 5년) 미만 신생 업체는 최저가여도 미선정.
    lim = cfg.get("R107_신생업체_기준_년")
    업력 = lambda q: today.year - q["설립연도"] if q.get("설립연도") else None
    신생 = lambda q: lim is not None and 업력(q) is not None and 업력(q) < lim
    priced = [q for q in quotes if q.get("금액") is not None]
    if not sel or not priced:
        rows.append(_row("R-106", 보완, "선정 업체 표시 없음", "견적서 비교에서 선정 업체를 표시해 주십시오."))
    else:
        low = min(priced, key=lambda q: q["금액"])
        cheaper = [q for q in priced if sel["금액"] is not None and q["금액"] < sel["금액"]]
        if sel["금액"] is not None and sel["금액"] <= low["금액"]:
            rows.append(_row("R-106", 충족, f"최저가 업체 {sel['업체']} 선정"))
        elif cheaper and all(신생(q) for q in cheaper):
            # 더 싼 곳이 모두 신생 업체 → 미선정이 기준에 맞다
            who = ", ".join(f"{q['업체']}({q['설립연도']}년 설립, 업력 {업력(q)}년)" for q in cheaper)
            if d.get("선정사유"):
                rows.append(_row("R-106", 충족, f"최저가 {who} — 업력 {lim}년 미만 신생 업체라 미선정(기준 부합). "
                                               f"선정 사유: \"{d['선정사유']}\""))
            else:
                rows.append(_row("R-106", 보완, f"최저가 {who} 미선정은 기준에 맞으나 사유 기재 없음",
                                 "견적서 비교에 '선정 사유:'를 기재해 주십시오. (예) " + 신생_미선정_문구(cheaper, sel, today)))
        elif d.get("선정사유"):
            rows.append(_row("R-106", 충족, f"최저가 {low['업체']} 아님, 선정 사유: \"{d['선정사유']}\""))
        else:
            rows.append(_row("R-106", 보완, f"최저가 {low['업체']}({fmt_num(low['금액'])}) 대신 {sel['업체']} 선정, 사유 없음",
                             f"최저가 업체({low['업체']})가 아닌 {sel['업체']}를 선정한 사유를 '선정 사유:' 로 기재해 주십시오."))

    if not sel:
        rows.append(_row("R-107", 해당없음, "선정 업체 확인 불가 (R-106 참조)"))
    elif not sel.get("설립연도"):
        rows.append(_row("R-107", 보완, f"{sel['업체']} 설립연도 없음",
                         f"견적서 비교에 {sel['업체']}의 설립연도를 기재해 주십시오."))
    elif lim is None:
        rows.append(_row("R-107", 판단, f"{sel['업체']} 설립 {sel['설립연도']} (업력 {업력(sel)}년) — 기준 [확정 필요]"))
    elif 신생(sel):
        rows.append(_row("R-107", 보완, f"선정 업체 {sel['업체']} 설립 {sel['설립연도']} (업력 {업력(sel)}년 < {lim}년): 신생 업체",
                         f"{sel['업체']}는 업력 {업력(sel)}년의 신생 업체로, 업력 {lim}년 미만 업체는 견적이 낮더라도 "
                         f"선정할 수 없습니다. 업력 {lim}년 이상 업체 중에서 다시 선정해 주십시오."))
    else:
        rows.append(_row("R-107", 충족, f"{sel['업체']} 설립 {sel['설립연도']} (업력 {업력(sel)}년 ≥ {lim}년)"))

    # 투자경과 = 투자 시작부터 걸리는 기간(개월 수)
    경과txt = s.get("투자경과", "").strip()
    경과 = to_number(경과txt)
    if 경과 is None or 경과 <= 0:
        rows.append(_row("R-108", 보완, f'{loc}: 투자경과 "{경과txt or "공란"}"',
                         "투자경과 칸에 투자 시작 후 걸리는 기간을 개월 수로 기재해 주십시오. (예: 5개월)"))
    else:
        시작 = re.match(r"\s*(\d{2,4})\s*[.\-/년]\s*(\d{1,2})", s.get("투자시작", ""))
        end = ""
        if 시작:
            y, m = int(시작.group(1)), int(시작.group(2))
            y = y + 2000 if y < 100 else y
            tot = y * 12 + (m - 1) + int(경과)
            end = f" → 완료 예상 {tot // 12}년 {tot % 12 + 1}월"
        rows.append(_row("R-108", 충족, f'{loc}: 투자시작 "{s.get("투자시작", "")}", 투자경과 {fmt_num(경과)}개월{end}'))

    # ---- 검토 2 유형 분류
    rows.append(_row("R-201", 판단, f'기재 유형 "{유형}" ↔ 투자효과: "{d.get("투자효과", "")[:80]}"', ai=True))
    if not 부유형:
        rows.append(_row("R-202", 해당없음, "부 유형 없음"))
    else:
        req = cfg.get("R202_부유형_필수항목도_요구")
        rows.append(_row("R-202", 판단 if req is None else 충족,
                         f'주 유형 "{유형}" / 부 유형 "{부유형}"' + (" — 부 유형 필수 항목 요구 여부 [확정 필요]" if req is None else "")))

    필수 = cfg["유형별_필수_부가항목"].get(유형, [])
    if not 필수:
        rows.append(_row("R-203", 해당없음, "투자유형 미확정 (R-103 참조)"))
    else:
        missing = [k for k in 필수 if not 부가.get(k)]
        if missing:
            rows.append(_row("R-203", 보완, f"{유형} 필수 항목 중 없음: {', '.join(missing)}",
                             f"{유형} 투자는 '{', '.join(missing)}' 항목이 필수입니다. 투자 상세 부가 칸 또는 "
                             "같은 제목의 추가 슬라이드로 작성해 주십시오."))
        else:
            rows.append(_row("R-203", 충족, ", ".join(f"{k}({출처.get(k, '')})" for k in 필수)))

    # ---- 검토 3 유형별 상세 (주 유형 기준)
    def na(*rids):
        for r in rids:
            rows.append(_row(r, 해당없음, f"투자유형 {유형 or '미기재'}"))

    if 유형 == "생산필수":
        if not d.get("투자효과"):
            rows.append(_row("R-301", 보완, "투자효과 칸 비어 있음", "투자효과 칸에 이 투자가 왜 필요한지 서술해 주십시오."))
        else:
            rows.append(_row("R-301", 판단, f'투자효과: "{d["투자효과"][:120]}"', ai=True))
        # 필수: 내년 부하 또는 조업도 수치 (현재 → 설치 후, 숫자 2개 이상)
        t = 부가.get("내년도 부하·조업도", "")
        has_kw = bool(re.search(r"부하|조업도", t))
        if has_kw and len(numbers_in(t)) >= 2:
            rows.append(_row("R-302", 충족, f'{출처.get("내년도 부하·조업도")}: "{t[:120]}" (기준치 [확정 필요])'))
        else:
            why = "없음" if not t else ("부하·조업도 언급 없음" if not has_kw else "수치 부족")
            rows.append(_row("R-302", 보완, f'내년도 부하·조업도: {why} — "{t[:80]}"',
                             "생산필수 투자는 내년 부하·조업도 정보가 필수입니다. 현재 부하(조업도)와 "
                             "설치 후 수치를 숫자로 기재해 주십시오. (예: 부하율 현재 118% → 설치 후 96%)"))
    else:
        na("R-301", "R-302")

    if 유형 == "노후교체":
        t = 부가.get("현 장비 재원 소개", "")
        src = 출처.get("현 장비 재원 소개", "")
        m = re.search(r"운용\s*연수\s*[:：]?\s*(\d+)", t) or re.search(r"(\d+)\s*년\s*(?:사용|운용|경과)", t)
        lim = cfg.get("R311_노후교체_운용연수_기준_년")
        if not m:
            rows.append(_row("R-311", 보완, f'현 장비 재원 소개: "{t[:80] or "없음"}"',
                             "현 장비 재원 소개에 운용연수를 '운용연수: N년' 형식으로 기재해 주십시오."))
        else:
            y = int(m.group(1))
            if lim is not None and y < lim:
                rows.append(_row("R-311", 판단, f'{src} "운용연수 {y}년" < 기준 {lim}년 — 조기 교체 사유 확인 필요'))
            else:
                rows.append(_row("R-311", 충족, f'{src} "운용연수 {y}년"' + ("" if lim else " (기준 [확정 필요])")))

        보유 = re.search(r"보유\s*수량\s*[:：]?\s*(\d[\d,]*)", t)
        교체 = re.search(r"교체\s*수량\s*[:：]?\s*(\d[\d,]*)", t)
        표수량 = _qty_from_name(d.get("투자명", "")) or to_number(s.get("투자수량"))
        if not (보유 and 교체):
            rows.append(_row("R-312", 보완, "보유 수량 / 교체 수량 기재 없음",
                             "현 장비 재원 소개에 '보유 수량: N대, 교체 수량: N대'를 기재해 주십시오."))
        elif 표수량 is not None and to_number(교체.group(1)) != 표수량:
            rows.append(_row("R-312", 보완, f"교체 수량 {교체.group(1)} ≠ 투자 수량 {fmt_num(표수량)}",
                             f"현 장비 재원 소개의 교체 수량({교체.group(1)})과 투자 수량({fmt_num(표수량)})이 다릅니다."))
        else:
            rows.append(_row("R-312", 충족, f"{src}: 보유 {보유.group(1)}, 교체 {교체.group(1)}"))

        # 필수: 노후된 수량 (교체 수량이 노후 수량보다 많으면 노후 아닌 장비까지 교체 → 판단 필요)
        노후 = re.search(r"노후\s*수량\s*[:：]?\s*(\d[\d,]*)", t)
        if not 노후:
            rows.append(_row("R-314", 보완, "노후 수량 기재 없음",
                             "노후교체 투자는 노후된 장비 수량이 필수입니다. 현 장비 재원 소개에 "
                             "'노후 수량: N대'를 기재해 주십시오."))
        else:
            n노후 = to_number(노후.group(1))
            n교체 = to_number(교체.group(1)) if 교체 else None
            if n교체 is not None and n교체 > n노후:
                rows.append(_row("R-314", 판단, f"{src}: 교체 수량 {fmt_num(n교체)} > 노후 수량 {fmt_num(n노후)} "
                                 "— 노후되지 않은 장비까지 교체하는지 확인 필요"))
            else:
                rows.append(_row("R-314", 충족, f"{src}: 노후 수량 {fmt_num(n노후)}"
                                 + (f", 교체 수량 {fmt_num(n교체)}" if n교체 is not None else "")))

        kw = [k for k in cfg["R313_인정_교체사유_키워드"] if k in t]
        if kw:
            rows.append(_row("R-313", 충족, f"{src}: 교체 사유 키워드 {', '.join(kw)} (인정 사유 목록 [확정 필요])"))
        else:
            rows.append(_row("R-313", 보완, "교체 사유(고장 이력, 수리비 증가, 부품 단종, 안전 문제 등) 없음",
                             "현 장비 재원 소개에 교체 사유(고장 이력, 수리비 증가, 부품 단종, 안전 문제 등)를 기재해 주십시오."))
    else:
        na("R-311", "R-312", "R-313", "R-314")

    if 유형 == "생산성 향상(자동화)":
        # 필수: 금액 · JMH · 인력 감소 세 가지 모두 숫자로
        t = 부가.get("생산성 향상 산출", "")
        src = 출처.get("생산성 향상 산출", "")
        필수효과 = {
            "금액": r"\d[\d,\.]*\s*(백만원|천만원|억|만원|원)",
            "JMH": r"\d[\d,\.]*\s*(JMH|M/H|MH)",
            "인력 감소": r"\d[\d,\.]*\s*(명|인(?!치))",
        }
        found = {k: bool(re.search(p, t)) for k, p in 필수효과.items()}
        missing = [k for k, ok in found.items() if not ok]
        if not missing:
            rows.append(_row("R-321", 충족, f'{src}: 금액·JMH·인력 감소 모두 기재 — "{t[:120]}"'))
        else:
            rows.append(_row("R-321", 보완, f'생산성 향상 산출에 {", ".join(missing)} 없음 — "{t[:80] or "없음"}"',
                             f"생산성 향상 투자는 금액, JMH, 인력 감소 효과가 모두 필수입니다. "
                             f"생산성 향상 산출에 {', '.join(missing)} 효과를 숫자로 기재해 주십시오."))
        rows.append(_row("R-322", 판단, f'{src}: "{t[:160]}" — 회수기간 기준 {cfg.get("R322_회수기간_기준_년")}년 [확정 필요]',
                         ai=bool(t)))
    else:
        na("R-321", "R-322")

    return {"신청서ID": d["신청서ID"], "파일": d["파일"], "경로": d.get("경로", ""), "원본해시": d.get("원본해시", ""), "투자명": d.get("투자명", ""),
            "투자유형": 유형, "부유형": 부유형, "투자금액": 금액, "AREA": s.get("AREA", ""),
            "계획NO": s.get("계획 NO", ""), "수량": _qty_from_name(d.get("투자명", "")) or to_number(s.get("투자수량")),
            "에이전트판단유형": "", "변환원본": d.get("변환원본"), "경고": d.get("경고", []), "rows": rows}


def summarize(r: dict) -> dict:
    cnt = {k: sum(1 for x in r["rows"] if x["판정"] == k) for k in (충족, 보완, 판단, 해당없음)}
    ai = sum(1 for x in r["rows"] if x.get("AI대기"))
    if cnt[보완]:
        state = "보완 요청"
    elif cnt[판단] == 0 and ai == 0:
        state = "검토 통과"
    else:
        state = "검토자 확인 필요"
    # R-401: 보완 0 + 판단 필요 행 전부 검토자 확인 O → 상정
    unconfirmed = [x for x in r["rows"] if x["판정"] == 판단 and norm(x.get("검토자확인")) not in 확인_OK]
    상정 = "상정" if cnt[보완] == 0 and ai == 0 and not unconfirmed else "보류"
    return {**cnt, "AI대기": ai, "종합상태": state, "회의상정": 상정}


# ---------------- 명령
def _unconvertible(p: Path, why: str) -> dict:
    """표준양식이 아니고 쓸 수 있는 변환본도 없는 신청서 — 엑셀에 사유를 남기고 보류."""
    rec = {"신청서ID": p.stem, "파일": p.name, "경로": str(p.resolve()), "원본해시": ex.file_hash(p),
           "투자명": "", "투자유형": "", "부유형": "", "투자금액": None, "AREA": "", "계획NO": "", "수량": None,
           "에이전트판단유형": "", "변환원본": None, "경고": [why],
           "rows": [_row("R-100", 보완, f"{p.name}: {why}",
                         "투자신청서를 표준양식(표준양식_수정.pptx)으로 작성해 다시 제출해 주십시오.")]}
    return rec


def cmd_check(paths: list[str]) -> None:
    """인자 없이 실행하면 inbox 전체: 표준양식은 원본, 다른 양식은 검증 통과한 변환본(out/표준화)을 점검."""
    cfg = load_config()
    if not paths:
        import standardize
        ok, bad = standardize.targets()
        paths = [str(x) for x in ok]
        for p, why in bad:
            r = _unconvertible(p, why)
            write_json(REVIEW_DIR / f"{r['신청서ID']}.json", r)
            print(f"[검토] {p.stem}: 점검 불가 — {why} → 보완 요청")
    for p in paths:
        p = Path(p)
        d = ex.extract(p) if p.suffix.lower() == ".pptx" else read_json(p)
        write_json(OUT / "extract" / f"{d['신청서ID']}.json", d)
        r = check(d, cfg)
        # 다시 check 하면 에이전트 판단(AI대기 행)은 초기화된다 — 신청서가 바뀌었을 수 있으므로.
        # 검토자가 엑셀에 적은 확인/수정 표시는 report 단계에서 다시 반영된다.
        write_json(REVIEW_DIR / f"{d['신청서ID']}.json", r)
        sm = summarize(r)
        print(f"[검토] {d['신청서ID']} {r['투자명']}: 보완 {sm[보완]} / 판단 필요 {sm[판단]} "
              f"(AI대기 {sm['AI대기']}) → {sm['종합상태']}")


def cmd_pending() -> None:
    for f in sorted(REVIEW_DIR.glob("*.json")):
        r = read_json(f)
        for x in r["rows"]:
            if x.get("AI대기"):
                print(f"{f.name}\t{x['규칙ID']}\t{x['검토항목']}\t{x['근거']}")


def cmd_judge(argv: list[str]) -> None:
    """에이전트 판단 기록: judge <신청서ID> <규칙ID> <판정> "<근거>" ["<보완요청문구>"] [--type <판단유형>]"""
    판단유형 = ""
    if "--type" in argv:
        i = argv.index("--type")
        판단유형 = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    if len(argv) < 4:
        sys.exit('사용: review.py judge <신청서ID> <규칙ID> <판정> "<근거>" ["<보완요청문구>"] [--type 유형]')
    sid, rid, 판정, 근거 = argv[:4]
    문구 = argv[4] if len(argv) > 4 else ""
    if 판정 not in (충족, 보완, 판단):
        sys.exit(f"판정은 {충족} / {보완} / {판단} 중 하나여야 합니다.")
    if 판정 == 보완 and not 문구:
        sys.exit("판정이 보완이면 신청자에게 보낼 보완 요청 문구가 필요합니다.")
    f = REVIEW_DIR / f"{sid}.json"
    r = read_json(f)
    row = next((x for x in r["rows"] if x["규칙ID"] == rid), None)
    if row is None:
        sys.exit(f"{sid} 에 {rid} 행이 없습니다.")
    if not row.get("AI대기") and row.get("판단주체") != "에이전트":
        sys.exit(f"{rid} 는 규칙(스크립트)으로 확정된 행입니다. 에이전트가 바꿀 수 없습니다.")
    row.update({"판정": 판정, "근거": 근거, "보완요청문구": 문구 if 판정 == 보완 else "",
                "AI대기": False, "판단주체": "에이전트"})
    if 판단유형:
        r["에이전트판단유형"] = 판단유형
    write_json(f, r)
    print(f"[판단] {sid} {rid} → {판정}")


def _load_human_marks() -> dict:
    """기존 검토결과.xlsx 에서 검토자가 적은 판정 수정·확인 표시를 읽는다."""
    marks = {}
    if not REPORT.exists():
        return marks
    import openpyxl
    ws = openpyxl.load_workbook(REPORT)["① 상세 검토"]
    hdr = [c.value for c in ws[1]]
    ix = {h: i for i, h in enumerate(hdr)}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[ix["신청서 ID"]]:
            marks[(row[ix["신청서 ID"]], row[ix["규칙 ID"]])] = (row[ix["판정"]], row[ix["검토자 확인"]] or "")
    return marks


def load_reviews(apply_marks: bool = True) -> list[dict]:
    marks = _load_human_marks() if apply_marks else {}
    out = []
    for f in sorted(REVIEW_DIR.glob("*.json")):
        r = read_json(f)
        for x in r["rows"]:
            m = marks.get((r["신청서ID"], x["규칙ID"]))
            if m:
                if m[1]:  # 검토자가 확인/수정 표시를 했으면 그 판정을 따른다
                    x["판정"], x["검토자확인"] = m[0], m[1]
                    x["AI대기"] = False
        r["요약"] = summarize(r)
        out.append(r)
    return out


def cmd_report() -> None:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    reviews = load_reviews()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "① 상세 검토"
    cols = ["신청서 ID", "투자명", "규칙 ID", "검토 항목", "판정", "근거", "보완 요청 문구", "AI대기", "검토자 확인"]
    ws.append(cols)
    fill = {보완: "FCE4D6", 판단: "FFF2CC", 충족: "E2EFDA", 해당없음: "EDEDED"}
    for r in reviews:
        for x in r["rows"]:
            ws.append([r["신청서ID"], r["투자명"], x["규칙ID"], x["검토항목"], x["판정"], x["근거"],
                       x["보완요청문구"], "대기" if x.get("AI대기") else "", x.get("검토자확인", "")])
            ws.cell(ws.max_row, 5).fill = PatternFill("solid", fgColor=fill.get(x["판정"], "FFFFFF"))
    dv = DataValidation(type="list", formula1='"O,수정"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"I2:I{max(ws.max_row, 2)}")
    dv2 = DataValidation(type="list", formula1=f'"{충족},{보완},{판단},{해당없음}"', allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add(f"E2:E{max(ws.max_row, 2)}")

    ws2 = wb.create_sheet("② 신청서 요약")
    ws2.append(["신청서 ID", "투자명", "기재 유형", "에이전트 판단 유형", "투자금액(백만원)", "AREA", "계획 NO",
                "충족", "보완", "판단 필요", "해당 없음", "AI대기", "종합 상태", "회의 상정"])
    for r in reviews:
        sm = r["요약"]
        ws2.append([r["신청서ID"], r["투자명"], r["투자유형"], r.get("에이전트판단유형", ""), r["투자금액"],
                    r["AREA"], r["계획NO"], sm[충족], sm[보완], sm[판단], sm[해당없음], sm["AI대기"],
                    sm["종합상태"], sm["회의상정"]])

    # ③ 잘못된 것만 모아 설명 (신청서별) — 신청자 회신·검토자 확인용
    ws4 = wb.create_sheet("③ 보완·확인 사항")
    ws4.append(["신청서 ID", "투자명", "종합 상태", "규칙 ID", "검토 항목", "판정", "무엇이 문제인가 (근거)",
                "신청자에게 보낼 문구", "검토자 확인"])
    for r in reviews:
        bad = [x for x in r["rows"] if x["판정"] in (보완, 판단)]
        if not bad:
            ws4.append([r["신청서ID"], r["투자명"], r["요약"]["종합상태"], "", "", "", "문제 없음 → 회의자료 에이전트로 전달", "", ""])
            continue
        for x in bad:
            ws4.append([r["신청서ID"], r["투자명"], r["요약"]["종합상태"], x["규칙ID"], x["검토항목"], x["판정"],
                        x["근거"], x["보완요청문구"] or "(검토자 판단 사항 — 신청자 회신 불필요)", x.get("검토자확인", "")])
            ws4.cell(ws4.max_row, 6).fill = PatternFill("solid", fgColor=fill[x["판정"]])

    ws3 = wb.create_sheet("사용법")
    for line in [
        "0) 점검 에이전트는 inbox 원본을 고치지 않습니다. 잘못된 것은 ③ 보완·확인 사항에 설명합니다.",
        "1) ① 상세 검토에서 '판단 필요' 행을 읽고 동의하면 검토자 확인 = O",
        "2) 판정을 바꾸려면 판정 칸을 고치고 검토자 확인 = 수정",
        "3) ② 신청서 요약의 회의 상정 = 상정 인 건만 회의자료 에이전트가 받습니다 (보완 0건 + 판단 필요 전부 O, R-401)",
        "4) 점검 후 신청서를 고치면 회의자료 단계에서 '재점검 필요'로 보류됩니다 → 점검을 다시 실행하세요.",
        "※ report 를 다시 돌려도 검토자 확인/판정 수정 값은 유지됩니다.",
    ]:
        ws3.append([line])

    for sh in (ws, ws2, ws4):
        for c in sh[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="D9E1F2")
        for col in sh.columns:
            width = min(60, max(8, max(len(str(c.value or "")) for c in col) * 1.6))
            sh.column_dimensions[col[0].column_letter].width = width
        for row in sh.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        sh.freeze_panes = "A2"
    ws3.column_dimensions["A"].width = 100

    REPORT.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(REPORT)
    except PermissionError:
        sys.exit(f"{REPORT} 가 엑셀에서 열려 있습니다. 닫고 다시 실행하세요.")
    from report_html import write_html
    html = write_html(reviews, OUT / "검토결과.html")
    print(f"[보고서] {html.relative_to(OUT.parent)}  ← 오류 설명 화면 (브라우저로 열기)")
    print(f"[승인]   {REPORT.relative_to(OUT.parent)}  ← 판단 필요 항목 검토자 확인(O) 표시")
    from system_html import build as build_system
    print(f"[한 화면] {build_system().relative_to(OUT.parent)}  ← 검토 결과 + 회의자료 탭")
    for r in reviews:
        sm = r["요약"]
        print(f"  {r['신청서ID']:<24} {sm['종합상태']:<10} 보완 {sm[보완]} 판단 필요 {sm[판단]} "
              f"AI대기 {sm['AI대기']} → {sm['회의상정']}")


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("check", "report", "pending", "judge"):
        sys.exit(__doc__)
    {"check": lambda: cmd_check(sys.argv[2:]), "report": cmd_report, "pending": cmd_pending,
     "judge": lambda: cmd_judge(sys.argv[2:])}[sys.argv[1]]()
