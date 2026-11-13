"""데모·테스트용 샘플 생성 (익명화된 가상 데이터).

  data/경영계획.xlsx, data/집행이력.xlsx
  inbox/신청서_0N_*.pptx  ← 표준양식_수정.pptx 를 채워 만든 투자신청서 5건

신청서 1~4 = '소 항목' 시트의 예시(T1~T4), 5 = 필수 항목이 빠진 보완 대상 사례.
"""
from __future__ import annotations

import copy

import openpyxl
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from common import ROOT

TEMPLATE = ROOT / "표준양식_수정.pptx"
INBOX = ROOT / "inbox"
DATA = ROOT / "data"

PLANS = [
    ["NO.1", "Z1 골리앗 크레인 신설", "생산필수", "골리앗 크레인", "Z1", 1, 5000],
    ["NO.2", "A1 안벽 연장", "생산필수", "안벽", "A1", 1, 3000],
    ["NO.3", "용접기 500대 구매", "노후교체", "용접기", "전사", 500, 500],
    ["NO.4", "A팀 로봇용접기 1대 구매", "생산성 향상(자동화)", "로봇용접기", "A팀", 1, 500],
]

APPS = [
    {
        "file": "신청서_01_Z2골리앗크레인.pptx",
        "투자명": "Z2 골리앗 크레인 신설",
        "투자효과": "- SN0000 특수선 제작 필수요소\n  . 기존장비 Capa 500ton 이상 블록다수\n"
                   "    : N100, M100, Q100 등 기존 골리앗 사용 불가로 해당 장비 신설 필요\n"
                   "  . Z1 부지 지반 보강 비용 과다로 Z2 로 위치 변경",
        "상세": {"투자시작": "27.4月", "투자경과": "5개월", "투자금액(백만원)": "3000", "투자유형": "생산필수",
                "현수량": "0", "필요수량": "1", "투자수량": "1", "부 유형": "",
                "AREA": "Z2", "계획 NO": "NO.1", "선정 업체": "A사", "견적 수": "2"},
        "부가": "- 견적서 비교\n  A사: 3,000백만원, 설립 1998년 (선정)\n  B사: 3,250백만원, 설립 2004년\n\n"
                "- 내년도 부하·조업도\n  Z도크 부하율 현재 118% → 설치 후 96%\n  500ton 이상 블록 처리량 월 12개 → 월 20개",
    },
    {
        "file": "신청서_02_용접기.pptx",
        "투자명": "용접기 300대 구매",
        "투자효과": "- 내년도 물량 증가로 용접기 가동률 부족\n  . 현 보유 용접기로는 내년 용접 물량의 85% 만 처리 가능\n"
                   "  . 300대 추가 시 물량 100% 처리",
        "상세": {"투자시작": "27.3月", "투자경과": "2개월", "투자금액(백만원)": "200", "투자유형": "생산필수",
                "현수량": "1200", "필요수량": "1500", "투자수량": "300", "부 유형": "",
                "AREA": "전사", "계획 NO": "NO.3", "선정 업체": "C사", "견적 수": "2"},
        "부가": "- 견적서 비교\n  C사: 200백만원, 설립 2001년 (선정)\n  D사: 215백만원, 설립 2010년\n\n"
                "- 내년도 부하·조업도\n  용접 부하 현재 117% → 설치 후 98%",
    },
    {
        "file": "신청서_03_GR.pptx",
        "투자명": "GR 100대 구매",
        "투자효과": "- 노후 GR 교체로 고장 정지 감소 및 안전 확보",
        "상세": {"투자시작": "27.2月", "투자경과": "1개월", "투자금액(백만원)": "50", "투자유형": "노후교체",
                "현수량": "100", "필요수량": "100", "투자수량": "100", "부 유형": "",
                "AREA": "전사", "계획 NO": "계획外", "선정 업체": "G사", "견적 수": "2"},
        "부가": "- 견적서 비교\n  G사: 50백만원, 설립 1995년 (선정)\n  H사: 56백만원, 설립 2012년\n\n"
                "- 현 장비 재원 소개 : 다음 슬라이드 참조",
        "추가": [("현 장비 재원 소개",
                  "운용연수: 15년\n보유 수량: 100대, 노후 수량: 100대, 교체 수량: 100대\n"
                  "교체 사유: 최근 1년 고장 이력 42건(전년 대비 2배), 주요 부품 단종으로 수리 불가, 수리비 연 18백만원",
                  None)],
    },
    {
        "file": "신청서_04_A팀로봇용접기.pptx",
        "투자명": "A팀 로봇용접기 1대 구매",
        "투자효과": "- A팀 소조립 반복 용접 자동화\n  . 용접 JMH 절감, 작업자 2명 재배치",
        "상세": {"투자시작": "27.5月", "투자경과": "3개월", "투자금액(백만원)": "800", "투자유형": "생산성 향상(자동화)",
                "현수량": "0", "필요수량": "1", "투자수량": "1", "부 유형": "",
                "AREA": "A팀", "계획 NO": "NO.4", "선정 업체": "E사", "견적 수": "2"},
        "부가": "- 생산성 향상 산출\n  용접 JMH 연 4,000 → 2,400 (1,600 JMH 절감)\n  인원 2명 감소(재배치)\n"
                "  연 절감액 = 1,600 JMH × 0.1백만원 + 2명 × 80백만원 = 320백만원\n  회수기간 = 800 / 320 = 2.5년\n\n"
                "- 견적서 비교 : 다음 슬라이드 참조",
        "추가": [("견적서 비교", "선정 사유: 최저가를 제시한 F사(780백만원)는 2023년 설립된 업력 3년의 신생 업체로, 로봇용접기의 장기 A/S·부품 공급 지속성과 재무 안정성을 확인하기 어려워 선정에서 제외함. 차액 20백만원(2.5%)은 설치 후 10년 이상 사후 지원 리스크를 고려할 때 합리적인 수준으로 판단하여, 업력 23년의 E사를 선정함",
                  [["업체", "금액(백만원)", "설립연도", "비고"],
                   ["E사", "800", "2003", "선정"],
                   ["F사", "780", "2023", ""]])],
    },
    {
        "file": "신청서_05_C동도장부스.pptx",
        "투자명": "C동 도장부스 신설",
        "투자효과": "",
        "상세": {"투자시작": "27.6月", "투자경과": "", "투자금액(백만원)": "1200", "투자유형": "생산필수",
                "현수량": "2", "필요수량": "", "투자수량": "", "부 유형": "",
                "AREA": "C동", "계획 NO": "", "선정 업체": "J사", "견적 수": "1"},
        "부가": "- 견적서 비교\n  J사: 1,200백만원 (선정)",
    },
]


def set_cell(cell, text: str):
    """셀 글자를 바꾸되 양식의 글자 서식(색·크기·굵기)은 유지한다.
    (서식 없이 쓰면 표 스타일의 머리글 색=흰색이 적용돼 글자가 안 보임)"""
    tf = cell.text_frame
    runs = [r for p in tf.paragraphs for r in p.runs if r.text.strip()]
    first = next((r for r in runs if not r.text.strip().startswith("※")), runs[0] if runs else None)
    rpr = first._r.find(qn("a:rPr")) if first is not None else None
    tf.text = ""
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        if rpr is not None:
            r._r.insert(0, copy.deepcopy(rpr))
            r.font.size = Pt(12) if len(text) > 40 else r.font.size


def remove_slide(prs, index):
    lst = prs.slides._sldIdLst
    sid = list(lst)[index]
    prs.part.drop_rel(sid.rId)
    lst.remove(sid)


def add_extra_slide(prs, title, body, table):
    layout = next((lo for lo in prs.slide_layouts if lo.name == "빈 화면"), prs.slide_layouts[6])
    s = prs.slides.add_slide(layout)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    tb = s.shapes.add_textbox(Emu(788416), Emu(350000), Emu(10294112), Emu(500000))
    tb.text_frame.text = title
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
    tb.text_frame.paragraphs[0].runs[0].font.bold = True
    y = 1000000
    if table:
        shp = s.shapes.add_table(len(table), len(table[0]), Emu(788416), Emu(y), Emu(8000000), Emu(380000 * len(table)))
        for ri, row in enumerate(table):
            for ci, v in enumerate(row):
                shp.table.cell(ri, ci).text = v
        y += 380000 * len(table) + 300000
    if body:
        b = s.shapes.add_textbox(Emu(788416), Emu(y), Emu(10294112), Emu(2000000))
        b.text_frame.word_wrap = True
        b.text_frame.text = body


def make_app(spec, outdir=None):
    outdir = outdir or INBOX
    prs = Presentation(str(TEMPLATE))
    remove_slide(prs, 1)  # '투자신청서 작성 안내' 슬라이드 제거
    slide = prs.slides[0]
    for shp in slide.shapes:
        if not shp.has_table:
            continue
        rows = shp.table.rows
        texts = [[c.text.strip() for c in r.cells] for r in rows]
        if any(c == "투자명" for r in texts for c in r):
            for ri, r in enumerate(texts):
                for ci, c in enumerate(r):
                    if c == "투자명":
                        set_cell(shp.table.cell(ri, ci + 1), spec["투자명"])
                    elif c == "투자효과":
                        set_cell(shp.table.cell(ri, ci + 1), spec["투자효과"])
                    elif c.replace("\n", "").replace(" ", "") == "투자상세부가":
                        set_cell(shp.table.cell(ri, ci + 1), spec["부가"])
        elif any("투자금액" in c for r in texts for c in r):
            for ri in range(0, len(texts) - 1, 2):
                for ci, key in enumerate(texts[ri]):
                    if key in spec["상세"]:
                        set_cell(shp.table.cell(ri + 1, ci), spec["상세"][key])
    for title, body, table in spec.get("추가", []):
        add_extra_slide(prs, title, body, table)
    outdir.mkdir(exist_ok=True)
    prs.save(str(outdir / spec["file"]))
    print("[샘플]", outdir / spec["file"])


def make_xlsx(path, header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    path.parent.mkdir(exist_ok=True)
    wb.save(path)
    print("[샘플]", path)


LEGACY = ROOT / "samples" / "다른양식"


def make_legacy(outdir=LEGACY):
    """표준양식이 아닌 '다른 양식' 신청서 (변환 스킬 시험용). 번호 목록 + 효과 슬라이드 + 업체 비교 표."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    blank = prs.slide_layouts[6]

    def text_slide(title, lines):
        s = prs.slides.add_slide(blank)
        tb = s.shapes.add_textbox(Emu(600000), Emu(300000), Emu(11000000), Emu(600000))
        tb.text_frame.text = title
        tb.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
        tb.text_frame.paragraphs[0].runs[0].font.bold = True
        body = s.shapes.add_textbox(Emu(600000), Emu(1100000), Emu(11000000), Emu(5000000))
        body.text_frame.word_wrap = True
        for i, line in enumerate(lines):
            p = body.text_frame.paragraphs[0] if i == 0 else body.text_frame.add_paragraph()
            p.text = line
            p.runs[0].font.size = Pt(18)
        return s

    text_slide("설비 투자 품의서", [
        "작성 부서: B도크 도장팀",
        "1. 투자 건명: B도크 도장 로봇 2대 도입",
        "2. 투자 구분: 생산성 향상(자동화)",
        "3. 착수 시기: 27.6月 / 소요 기간: 4개월",
        "4. 설치 장소: B도크",
        "5. 경영계획 번호: 계획外",
        "6. 수량: 현수량 0 / 필요수량 2 / 투자수량 2",
        "7. 투자 금액: 640백만원",
    ])
    text_slide("투자 효과", [
        "- 도장 공정 자동화로 작업자 유해 환경 노출 감소",
        "- 도장 JMH 연 6,000 → 3,600 (2,400 JMH 절감)",
        "- 인원 3명 감소(재배치)",
        "- 연 절감액 = 2,400 JMH × 0.1백만원 + 3명 × 80백만원 = 480백만원",
        "- 회수기간 = 640 / 480 = 1.3년",
    ])
    s = text_slide("업체 비교", [])
    rows = [["공급사", "견적가(백만원)", "설립", "비고"], ["K사", "640", "2005", "채택"], ["L사", "690", "2011", ""]]
    t = s.shapes.add_table(3, 4, Emu(600000), Emu(1200000), Emu(8000000), Emu(1200000)).table
    for ri, row in enumerate(rows):
        for ci, v in enumerate(row):
            t.cell(ri, ci).text = v
    body = s.shapes.add_textbox(Emu(600000), Emu(2700000), Emu(11000000), Emu(600000))
    body.text_frame.text = "채택 사유: 최저가, 견적 2곳 비교"
    outdir.mkdir(parents=True, exist_ok=True)
    path = outdir / "신청서_06_B도크도장로봇_다른양식.pptx"
    prs.save(str(path))
    print("[샘플]", path)
    return path


if __name__ == "__main__":
    # 기본: 이미 있는 파일은 건드리지 않는다 (사용자가 고친 신청서·경영계획 보호).
    # 샘플로 되돌리려면: python scripts/make_samples.py --force
    # 다른 양식 시험 샘플만 (samples/다른양식/, inbox 는 건드리지 않음): --legacy
    import sys
    if "--legacy" in sys.argv:
        make_legacy()
        sys.exit(0)
    force = "--force" in sys.argv

    def skip(path):
        if path.exists() and not force:
            print("[유지] 이미 있음 — 덮어쓰지 않음:", path)
            return True
        return False

    if not skip(DATA / "경영계획.xlsx"):
        make_xlsx(DATA / "경영계획.xlsx",
                  ["계획 NO", "투자명", "투자유형", "장비 종류", "위치(AREA)", "수량", "계획금액(백만원)"], PLANS)
    if not skip(DATA / "집행이력.xlsx"):
        make_xlsx(DATA / "집행이력.xlsx", ["계획 NO", "신청일", "투자명", "집행금액(백만원)"], [])
    for spec in APPS:
        if not skip(INBOX / spec["file"]):
            make_app(spec)
