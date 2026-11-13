"""투자 회의자료 생성 (구현시나리오_보완.xlsx 소항목 3 '투자 회의', R-401).

검토결과(out/review/*.json + 검토자가 표시한 out/검토결과.xlsx)에서 회의 상정 건만 골라
  1) 표지: 투자 심의 | 경영계획 | 실행금액 | 차이금액 | 비고  (+ 계획 미신청)
     보류(미상정) 건은 회의자료에 넣지 않는다 — 실무자 확인 후 신청자에게 전달 (콘솔·회의_요약.xlsx '보류(실무자용)' 시트)
  2) 상정된 투자신청서 원본 슬라이드를 순서대로 첨부
를 out/투자회의자료_YYYYMMDD.pptx 로, 표지 표를 out/회의_요약.xlsx 로 만든다.

사용: python scripts/meeting.py [--plan data/경영계획.xlsx] [--history data/집행이력.xlsx] [--list]
"""
from __future__ import annotations

import argparse
import copy
import datetime as dt
import io
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from common import OUT, ROOT, fmt_num, load_config, norm, write_json
from match import load_history, load_plans, match
from review import load_reviews

TEMPLATE = ROOT / "표준양식_수정.pptx"
ROWS_PER_COVER = 9
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
RED = RGBColor(0xC0, 0x00, 0x00)
BLUE = RGBColor(0x1F, 0x5F, 0xBF)
GREY = RGBColor(0x59, 0x59, 0x59)


# ---------------- pptx 유틸
def _blank_layout(prs):
    for lo in prs.slide_layouts:
        if lo.name in ("빈 화면", "Blank"):
            return lo
    return prs.slide_layouts[len(prs.slide_layouts) - 1]


def _clear_slides(prs):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        prs.part.drop_rel(sid.rId)
        lst.remove(sid)


def _new_slide(prs):
    s = prs.slides.add_slide(_blank_layout(prs))
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    return s


def copy_slide(src_slide, dst_prs):
    """다른 pptx 의 슬라이드를 도형 단위로 복사 (표·텍스트·이미지)."""
    new = _new_slide(dst_prs)
    rid_map = {}
    for rel in src_slide.part.rels.values():
        if rel.is_external:
            continue
        if rel.reltype.endswith("/image"):
            _, new_rid = new.part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
            rid_map[rel.rId] = new_rid
    for shp in src_slide.shapes:
        el = copy.deepcopy(shp._element)
        for node in el.iter():
            for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                if node.get(attr) in rid_map:
                    node.set(attr, rid_map[node.get(attr)])
        new.shapes._spTree.insert_element_before(el, "p:extLst")
    return new


def _text(slide, x, y, w, h, text, size=12, bold=False, color=None, align=None):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        if color is not None:
            r.font.color.rgb = color
        if align:
            p.alignment = align
    return tb


def _set_cell(cell, text, size=11, bold=False, color=None, align=PP_ALIGN.CENTER, fill=None):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    if color is not None:
        r.font.color.rgb = color
    p.alignment = align
    cell.margin_left = cell.margin_right = Emu(60000)
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def _diff_text(x: float) -> str:
    return "0" if x == 0 else (f"+{fmt_num(x)}" if x > 0 else f"−{fmt_num(-x)}")


def build_cover(prs, rows, 미신청, 보류, today, page, pages):
    s = _new_slide(prs)
    W = prs.slide_width
    L = 600000
    tw = W - 2 * L
    suffix = f" ({page}/{pages})" if pages > 1 else ""
    _text(s, L, 280000, tw, 500000, f"투자 심의 요약{suffix}", size=26, bold=True, color=NAVY)
    plan_sum = sum(r["경영계획"] for r in rows)
    exec_sum = sum(r["실행금액"] for r in rows)
    _text(s, L, 780000, tw, 320000,
          f"{today:%Y-%m-%d} · 상정 {len(rows)}건 · 경영계획 {fmt_num(plan_sum)} / 실행 {fmt_num(exec_sum)} "
          f"(차이 {_diff_text(exec_sum - plan_sum)}) · 단위: 백만원", size=12, color=GREY)

    cols = [("No", 0.05), ("투자 심의", 0.27), ("유형", 0.13), ("계획 NO", 0.08),
            ("경영계획", 0.09), ("실행금액", 0.09), ("차이금액", 0.09), ("비고", 0.20)]
    n = len(rows) + 2  # 헤더 + 합계
    tbl = s.shapes.add_table(n, len(cols), Emu(L), Emu(1200000), Emu(tw), Emu(380000 * n)).table
    for i, (name, ratio) in enumerate(cols):
        tbl.columns[i].width = Emu(int(tw * ratio))
        _set_cell(tbl.cell(0, i), name, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), fill=NAVY)
    for ri, r in enumerate(rows, start=1):
        note = r["비고"] + (" (확인 필요)" if r["확인필요"] else "")
        vals = [str(r["_no"]), r["투자명"], r["투자유형"], r["계획NO"], fmt_num(r["경영계획"]),
                fmt_num(r["실행금액"]), _diff_text(r["차이금액"]), note]
        for ci, v in enumerate(vals):
            color = None
            if ci == 6 and r["차이금액"] > 0:
                color = RED
            elif ci == 6 and r["차이금액"] < 0:
                color = BLUE
            align = PP_ALIGN.LEFT if ci in (1, 7) else (PP_ALIGN.RIGHT if 4 <= ci <= 6 else PP_ALIGN.CENTER)
            _set_cell(tbl.cell(ri, ci), v, size=10 if ci == 7 else 11, color=color, align=align)
    last = n - 1
    tbl.cell(last, 0).merge(tbl.cell(last, 3))
    tot_fill = RGBColor(0xE7, 0xEC, 0xF3)
    _set_cell(tbl.cell(last, 0), "합계", bold=True, fill=tot_fill)
    for ci, v in ((4, fmt_num(plan_sum)), (5, fmt_num(exec_sum)), (6, _diff_text(exec_sum - plan_sum))):
        _set_cell(tbl.cell(last, ci), v, bold=True, align=PP_ALIGN.RIGHT, fill=tot_fill)
    _set_cell(tbl.cell(last, 7), "", fill=tot_fill)

    if page == pages:
        y = 1200000 + 380000 * n + 250000
        lines = []
        if 미신청:
            lines.append("계획 미신청: " + " · ".join(
                f"{m['계획NO']} {m['투자명']} ({fmt_num(m['경영계획'])})" for m in 미신청))
        lines.append("차이 = 실행 − 계획 · (확인 필요) = 계획 매칭을 담당자가 확인해야 하는 건 · 첨부: 상정 신청서 원본")
        _text(s, L, y, tw, 900000, "\n".join(lines), size=10, color=GREY)
    return s


def _divider(prs, no, row):
    s = _new_slide(prs)
    _text(s, 600000, 2600000, prs.slide_width - 1200000, 700000, f"[{no}] {row['투자명']}",
          size=28, bold=True, color=NAVY)
    _text(s, 600000, 3300000, prs.slide_width - 1200000, 500000,
          f"{row['투자유형']} · 계획 {row['계획NO']} · 경영계획 {fmt_num(row['경영계획'])} → "
          f"실행 {fmt_num(row['실행금액'])} ({_diff_text(row['차이금액'])})"
          + (f" · {row['비고']}" if row["비고"] else ""), size=14, color=GREY)


def write_summary_xlsx(path, rows, 미신청, 보류):
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "투자 심의"
    ws.append(["No", "투자 심의", "유형", "계획 NO", "경영계획", "실행금액", "차이금액", "비고", "확인 필요", "적용 규칙"])
    for r in rows:
        ws.append([r["_no"], r["투자명"], r["투자유형"], r["계획NO"], r["경영계획"], r["실행금액"],
                   r["차이금액"], r["비고"], "확인 필요" if r["확인필요"] else "", ", ".join(r["규칙"])])
    ws2 = wb.create_sheet("계획 미신청")
    ws2.append(["계획 NO", "투자명", "유형", "경영계획", "기집행", "비고"])
    for m in 미신청:
        ws2.append([m["계획NO"], m["투자명"], m["투자유형"], m["경영계획"], m["기집행"], m["비고"]])
    ws3 = wb.create_sheet("보류(실무자용)")
    ws3.append(["신청서 ID", "투자명", "사유"])
    for b in 보류:
        ws3.append([b["신청서ID"], b["투자명"], b["사유"]])
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True)
        for col in sh.columns:
            sh.column_dimensions[col[0].column_letter].width = min(50, max(10, max(len(str(c.value or "")) for c in col) * 1.6))
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", default=str(ROOT / "data" / "경영계획.xlsx"))
    ap.add_argument("--history", default=str(ROOT / "data" / "집행이력.xlsx"))
    ap.add_argument("--list", action="store_true", help="상정/보류 목록만 출력")
    ap.add_argument("--no-divider", action="store_true", help="신청서 앞 구분 슬라이드 생략")
    args = ap.parse_args()

    cfg = load_config()
    reviews = load_reviews()
    if not reviews:
        sys.exit("out/review 에 검토 결과가 없습니다. 먼저 review.py check 를 실행하세요.")
    # 점검 에이전트가 넘긴 것(회의 상정) 중, 점검 이후 원본이 바뀐 건은 받지 않는다.
    from extract import file_hash
    보류 = []
    상정 = []
    for r in reviews:
        if r["요약"]["회의상정"] != "상정":
            continue
        p = Path(r.get("경로", ""))
        if not p.exists():
            보류.append({"신청서ID": r["신청서ID"], "투자명": r["투자명"], "사유": "원본 파일 없음 — 재점검 필요"})
        elif r.get("원본해시") and file_hash(p) != r["원본해시"]:
            보류.append({"신청서ID": r["신청서ID"], "투자명": r["투자명"], "사유": "점검 후 신청서가 바뀜 — 재점검 필요"})
        elif r.get("변환원본") and (not Path(r["변환원본"]["경로"]).exists()
                                  or file_hash(Path(r["변환원본"]["경로"])) != r["변환원본"]["해시"]):
            보류.append({"신청서ID": r["신청서ID"], "투자명": r["투자명"],
                       "사유": "변환 후 원본 신청서가 바뀜 — 다시 변환·점검 필요"})
        else:
            상정.append(r)
    for r in reviews:
        if r["요약"]["회의상정"] == "상정":
            continue
        sm = r["요약"]
        if sm["보완"]:
            why = f"보완 {sm['보완']}건"
        elif sm["AI대기"]:
            why = f"AI 판단 대기 {sm['AI대기']}건"
        else:
            why = f"검토자 확인 대기 {sm['판단 필요']}건"
        보류.append({"신청서ID": r["신청서ID"], "투자명": r["투자명"], "사유": why, "계획NO": r.get("계획NO", "")})

    print(f"[상정] {len(상정)}건: " + ", ".join(r["투자명"] for r in 상정))
    for b in 보류:
        print(f"[보류] {b['투자명']} — {b['사유']}")
    if args.list:
        return
    if not 상정:
        sys.exit("회의 상정 건이 없습니다. 검토결과.xlsx 에서 검토자 확인(O)을 표시했는지 확인하세요.")

    plans = load_plans(Path(args.plan))
    history = load_history(Path(args.history))
    res = match([{"신청서ID": r["신청서ID"], "투자명": r["투자명"], "투자유형": r["투자유형"],
                  "투자금액": r["투자금액"], "AREA": r["AREA"], "계획NO": r["계획NO"], "수량": r["수량"]}
                 for r in 상정], plans, history, show_unrequested=bool(cfg.get("M8_미신청_계획_표지표시", True)))
    rows = res["rows"]
    for i, r in enumerate(rows, start=1):
        r["_no"] = i
    # 신청서는 들어왔지만 보류된 계획은 '미신청'이 아니다 → 미신청 목록에서 빼고 보류 항목에 계획 NO 를 붙인다
    held = {norm(b["계획NO"]): b for b in 보류 if norm(b.get("계획NO", "")) not in ("", norm("계획外"))}
    for m in list(res["미신청"]):
        b = held.get(norm(m["계획NO"]))
        if b:
            res["미신청"].remove(m)
            b["사유"] += f" · 계획 {m['계획NO']}({fmt_num(m['경영계획'])}) 신청 보류"

    today = dt.date.today()
    prs = Presentation(str(TEMPLATE))
    _clear_slides(prs)
    pages = max(1, -(-len(rows) // ROWS_PER_COVER))
    for pg in range(pages):
        build_cover(prs, rows[pg * ROWS_PER_COVER:(pg + 1) * ROWS_PER_COVER], res["미신청"], 보류, today, pg + 1, pages)

    by_id = {r["신청서ID"]: r for r in 상정}
    for row in rows:
        rv = by_id[row["신청서ID"]]
        if not args.no_divider:
            _divider(prs, row["_no"], row)
        src = Presentation(rv["경로"])
        for slide in src.slides:
            texts = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
            if "투자신청서 작성 안내" in texts:
                continue
            copy_slide(slide, prs)

    OUT.mkdir(exist_ok=True)
    pptx_path = OUT / f"투자회의자료_{today:%Y%m%d}.pptx"
    try:
        prs.save(str(pptx_path))
        write_summary_xlsx(OUT / "회의_요약.xlsx", rows, res["미신청"], 보류)
    except PermissionError as e:
        sys.exit(f"파일이 열려 있어 저장하지 못했습니다: {e.filename}")
    write_json(OUT / "회의_매칭.json", res)
    from meeting_html import write_meeting_html
    html_path = write_meeting_html(rows, res["미신청"], 보류, OUT / f"투자회의자료_{today:%Y%m%d}.html", today)
    print(f"[회의자료] {html_path.relative_to(ROOT)}  ← 회의자료 (브라우저로 열기, 인쇄 시 신청서별 새 페이지)")
    print(f"[출력용]   {pptx_path.relative_to(ROOT)}  (표지 {pages}장 + 신청서 {len(rows)}건)")
    print(f"[요약표]   out/회의_요약.xlsx")
    from system_html import build as build_system
    print(f"[한 화면]  {build_system().relative_to(ROOT)}  ← 검토 결과 + 회의자료 탭")
    for r in rows:
        print(f"  {r['_no']}. {r['투자명']:<22} 계획 {fmt_num(r['경영계획']):>6} 실행 {fmt_num(r['실행금액']):>6} "
              f"차이 {_diff_text(r['차이금액']):>7}  {r['비고']}{' (확인 필요)' if r['확인필요'] else ''}")


if __name__ == "__main__":
    main()
