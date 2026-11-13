"""수량 칸 변경 (2026-11-13): 현수량/보유수량/필요수량 → 현수량/필요수량/투자수량.

  옛 정의: 현수량 − 보유수량 = 필요수량   (현 = 있어야 할 수, 보유 = 가진 수, 필요 = 살 수)
  새 정의: 투자수량 ≤ 필요수량 − 현수량    (현 = 가진 수, 필요 = 있어야 할 수, 투자 = 살 수)
  → 값의 의미를 그대로 옮긴다: 새 현 = 옛 보유, 새 필요 = 옛 현, 새 투자 = 옛 필요
    (옛 규칙을 지킨 신청서는 새 규칙도 지키고, 일부러 넣은 오류도 그대로 남는다)

또 표준양식의 작성 안내 문구, 샘플 04 의 옛 선정 사유 문장(원문 그대로일 때만)을 새 문구로 바꾼다.
원본을 쓰는 도구라 실행 전에 backup_20261112/ 에 복사하고, 안전장치 훅이 사용자 승인을 받는다.

사용: python scripts/migrate_quantity.py <파일.pptx> [...]
"""
from __future__ import annotations

import datetime as dt
import shutil
import sys
from pathlib import Path

from pptx import Presentation

from common import ROOT

OLD_HEAD = ["현수량", "보유수량", "필요수량"]
NEW_HEAD = ["현수량", "필요수량", "투자수량"]
OLD_GUIDE = "· 현수량 / 보유수량 / 필요수량: 노후교체는 필요수량 ≤ 현수량 (보유수량 무관), 그 외 유형은 현수량 − 보유수량 = 필요수량"
NEW_GUIDE = ("· 현수량 / 필요수량 / 투자수량: 현수량 = 지금 가진 수, 필요수량 = 있어야 할 수, 투자수량 = 이번에 살 수. "
             "노후교체는 투자수량 ≤ 필요수량, 그 외 유형은 투자수량 ≤ 필요수량 − 현수량")
OLD_REASON = "선정 사유: F사는 2023년 설립 신설 업체로 재정 위험 및 A/S 인력 부족"
NEW_REASON = ("선정 사유: 최저가를 제시한 F사(780백만원)는 2023년 설립된 업력 3년의 신생 업체로, 로봇용접기의 장기 "
              "A/S·부품 공급 지속성과 재무 안정성을 확인하기 어려워 선정에서 제외함. 차액 20백만원(2.5%)은 설치 후 "
              "10년 이상 사후 지원 리스크를 고려할 때 합리적인 수준으로 판단하여, 업력 23년의 E사를 선정함")
OLD_SEL_GUIDE = "· 선정 업체 / 견적 수: 선정한 업체명과 받은 견적 수 (2곳 이상)"
NEW_SEL_GUIDE = ("· 선정 업체 / 견적 수: 선정한 업체명과 받은 견적 수 (2곳 이상). 최저가 선정이 원칙이나 "
                 "업력 5년 미만 신생 업체는 최저가여도 선정하지 않고, 미선정 사유를 '선정 사유:'로 기재")


def set_text(cell_or_tf, text: str):
    tf = getattr(cell_or_tf, "text_frame", cell_or_tf)
    runs = [r for p in tf.paragraphs for r in p.runs]
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        tf.text = text


def migrate(path: Path) -> list[str]:
    prs = Presentation(str(path))
    done = []
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_table:
                rows = list(sh.table.rows)
                for ri, row in enumerate(rows):
                    cells = list(row.cells)
                    heads = [c.text.strip() for c in cells]
                    if heads[:3] == OLD_HEAD and ri + 1 < len(rows):
                        vals = [c.text.strip() for c in list(rows[ri + 1].cells)[:3]]
                        new_vals = [vals[1], vals[0], vals[2]]      # 새 현=옛 보유, 새 필요=옛 현, 새 투자=옛 필요
                        for i, h in enumerate(NEW_HEAD):
                            set_text(cells[i], h)
                            set_text(list(rows[ri + 1].cells)[i], new_vals[i])
                        done.append(f"수량 칸 {OLD_HEAD}={vals} → {NEW_HEAD}={new_vals}")
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for old, new, label in ((OLD_GUIDE, NEW_GUIDE, "작성 안내(수량)"),
                                            (OLD_SEL_GUIDE, NEW_SEL_GUIDE, "작성 안내(선정)"),
                                            (OLD_REASON, NEW_REASON, "선정 사유 문구")):
                        if p.text.strip() == old:
                            set_text_para(p, new)
                            done.append(label)
    if done:
        prs.save(str(path))
    return done


def set_text_para(p, text):
    runs = list(p.runs)
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


SPEC_ROWS = {   # 구현시나리오_보완.xlsx '검토 기준' — 규칙 ID: {열: 값}
    "R-102": {"E": "숫자로 기재됨 (현수량 = 지금 가진 수, 필요수량 = 있어야 할 수, 투자수량 = 이번에 살 수). "
                   "노후교체: 투자수량 ≤ 필요수량 / 그 외 유형: 투자수량 ≤ 필요수량 − 현수량",
              "G": "한도 초과·부족분 없음 표시 + 보완 요청 문구 작성",
              "H": "투자상세 표: 현수량/필요수량/투자수량",
              "K": "2026-11-13 변경: 칸을 현수량/필요수량/투자수량으로, 투자수량 한도 규칙"},
    "R-106": {"E": "최저가 업체 선정이면 충족. 최저가가 아니면 선정 사유 기재 필요. 단, 더 싼 업체가 모두 업력 5년 미만 "
                   "신생 업체면 미선정이 기준에 맞음 (사유 없으면 예시 문구를 붙여 보완 요청)",
              "G": "선정 사유 보완 요청 / 신생 업체 미선정 사유 예시 문구 제시",
              "K": "2026-11-13 변경: 신생 업체(5년 미만) 최저가 미선정 허용"},
    "R-107": {"D": "신생 업체 선정 금지",
              "E": "선정 업체 업력(올해 − 설립연도)이 5년 이상. 5년 미만 신생 업체는 견적이 낮아도 선정 불가",
              "F": "보완",
              "G": "업력 5년 이상 업체 중 재선정 요청",
              "K": "2026-11-13 확정: 기준 5년, 판단 필요 → 보완"},
}


def migrate_spec(path: Path) -> list[str]:
    """기준 엑셀의 해당 규칙 행 글자만 바꾼다 (행 추가·서식 변경 없음)."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb["검토 기준"]
    done = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        rid = str(row[1].value or "").strip()          # B열 = ID
        if rid in SPEC_ROWS:
            for col, val in SPEC_ROWS[rid].items():
                ws[f"{col}{row[0].row}"] = val
            done.append(rid)
    t = wb["유형별 필수 항목"]
    if t["C6"].value and "업력 5년" not in str(t["C6"].value):
        t["C6"] = str(t["C6"].value) + " (업력 5년 미만 신생 업체는 최저가여도 미선정)"
        done.append("유형별 견적서 비교")
    wb.save(path)
    return done


def main(paths):
    paths = [p for p in paths if not Path(p).name.startswith("~$")]      # Office 잠금 파일 제외
    bk = ROOT / "backup_20261112" / f"수량변경전_{dt.datetime.now():%H%M%S}"
    for p in map(Path, paths):
        if (p.parent / f"~${p.name}").exists():
            sys.exit(f"{p.name} 이(가) 열려 있습니다. 닫고 다시 실행하세요.")
    bk.mkdir(parents=True, exist_ok=True)
    for p in map(Path, paths):
        shutil.copy2(p, bk / p.name)
        done = migrate_spec(p) if p.suffix.lower() == ".xlsx" else migrate(p)
        print(f"[변경] {p.name}: " + ("; ".join(done) if done else "바꿀 것 없음"))
    print(f"[백업] {bk.relative_to(ROOT)}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1:])
