"""투자 회의자료 HTML (out/투자회의자료_YYYYMMDD.html). meeting.py 가 호출한다.

표지(투자 심의 요약) + 상정 신청서마다 표준양식 그대로의 내용(투자명·투자상세·투자효과·투자 상세 부가·견적서 비교·추가 슬라이드).
신청서 내용은 점검 때 읽은 값(out/extract/<ID>.json)을 그대로 보여 준다 — 고치지 않는다.
"""
from __future__ import annotations

import datetime as dt
from html import escape
from pathlib import Path

from common import OUT, fmt_num, read_json

DETAIL_ROWS = [["투자시작", "투자경과", "투자금액(백만원)", "투자유형"],
               ["현수량", "필요수량", "투자수량", "부 유형"],
               ["AREA", "계획 NO", "선정 업체", "견적 수"]]
부가순서 = ["견적서 비교", "내년도 부하·조업도", "현 장비 재원 소개", "생산성 향상 산출"]

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1d2433;--sub:#5b6474;--line:#d9dde3;--accent:#1f3a5f;--label:#fff2cc;
--head:#dbe5f1;--up:#b42318;--down:#1f5fbf;--warn:#8a5a00;--warn-bg:#fff4d6;--on:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#0b0d11;--card:#20262f;--ink:#e9ecf1;--sub:#a7b0bf;--line:#3c4553;
--accent:#8fb3e0;--label:#3a3420;--head:#22324a;--up:#f2867b;--down:#7fb0f0;--warn:#e8c06a;--warn-bg:#352a12;--on:#0b0d11}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
font:15px/1.55 "Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif}
main{max-width:1120px;margin:0 auto;padding:28px 16px 60px}
h1{font-size:26px;margin:0 0 4px;color:var(--accent)}h2{font-size:20px;margin:0}
.meta{color:var(--sub);font-size:13px;margin-bottom:18px}
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:16px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.tile b{display:block;font-size:24px}.tile span{color:var(--sub);font-size:13px}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin-bottom:18px}
.scroll{overflow-x:auto}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid var(--line);padding:7px 9px;vertical-align:top;text-align:left}
.cover th{background:var(--accent);color:var(--on);white-space:nowrap}
.cover td.n{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
.cover tr.sum td{font-weight:700;background:var(--head)}
.up{color:var(--up);font-weight:700}.down{color:var(--down);font-weight:700}
.cover tr.need td{background:var(--warn-bg)}.cover tr.need td:first-child{box-shadow:inset 4px 0 0 var(--warn)}
.chk{display:inline-block;margin-left:6px;padding:1px 7px;border-radius:2px;font-size:12px;font-weight:700;
color:var(--on);background:var(--warn)}
.notes{margin:12px 0 0;padding-left:18px;color:var(--sub);font-size:14px}.notes li{margin:3px 0}
.toc{columns:2;column-gap:28px;margin:0;padding-left:18px}.toc li{margin:3px 0}
@media (max-width:640px){.toc{columns:1}}
.app{page-break-before:always}
.apphead{display:flex;flex-wrap:wrap;gap:6px 14px;align-items:baseline;margin-bottom:10px}
.apphead .sub{color:var(--sub);font-size:14px}
.form td.l{background:var(--label);font-weight:700;width:120px;text-align:center;vertical-align:middle}
.detail th{background:var(--head);font-weight:700;text-align:center}.detail td{text-align:center}
.pre{white-space:pre-wrap;margin:0}
.sec{margin:0 0 12px}.sec h4{margin:0 0 4px;font-size:14px}
.extra{margin-top:12px}.extra h3{font-size:15px;margin:14px 0 6px}
.conv{color:var(--sub);font-size:13px;margin-top:6px}
footer{color:var(--sub);font-size:12px;margin-top:24px}
a.jump{color:var(--accent);text-decoration:none;border-bottom:1px dashed currentColor;cursor:pointer}
a.jump::after{content:" ↓";font-size:12px}
a.jump:hover{border-bottom-style:solid}
.flash{animation:flash 1.6s ease-out}
@keyframes flash{0%{box-shadow:0 0 0 4px var(--accent)}100%{box-shadow:0 0 0 0 transparent}}

.origbtn{font-size:13px;padding:3px 11px;border:1px solid var(--line);border-radius:999px;text-decoration:none;color:var(--accent);white-space:nowrap}
@media print{.origbtn{display:none}}
@media print{body{background:#fff}.card{border:none;padding:0}.tiles{display:none}a{color:inherit;text-decoration:none}}
"""


JUMP_JS = """
document.querySelectorAll('a.jump').forEach(a=>a.addEventListener('click',e=>{e.preventDefault();
 const t=document.getElementById(a.dataset.target);if(!t)return;
 t.scrollIntoView({behavior:'auto',block:'start'});t.classList.remove('flash');void t.offsetWidth;t.classList.add('flash')}));
"""


def _diff(x: float) -> str:
    return "0" if x == 0 else (f"+{fmt_num(x)}" if x > 0 else f"−{fmt_num(-x)}")


def _cls(x: float) -> str:
    return "up" if x > 0 else ("down" if x < 0 else "")


def _pre(text: str) -> str:
    return f'<p class="pre">{escape(text or "")}</p>' if (text or "").strip() else '<p class="pre">—</p>'


def _orig_btn(row: dict, d: dict) -> str:
    from system_html import orig_button
    src = Path(d["변환원본"]["경로"]) if d.get("변환원본") else Path(d.get("경로", ""))
    return orig_button(row["신청서ID"], src.resolve().as_uri()) if d.get("경로") and src.exists() else ""


def _app_section(no: int, row: dict, d: dict) -> str:
    s = d.get("상세", {})
    detail = "".join(
        "<tr>" + "".join(f"<th>{escape(k)}</th>" for k in hdr) + "</tr>"
        + "<tr>" + "".join(f"<td>{escape(s.get(k, '')) or '&nbsp;'}</td>" for k in hdr) + "</tr>"
        for hdr in DETAIL_ROWS)
    부가 = d.get("부가", {})
    keys = [k for k in 부가순서 if 부가.get(k)] + [k for k in 부가 if k not in 부가순서 and 부가.get(k)]
    부가_html = "".join(f'<div class="sec"><h4>{escape(k)}</h4>{_pre(부가[k])}</div>' for k in keys) or "<p>—</p>"
    quotes = d.get("견적", [])
    q_html = ""
    if quotes:
        q_html = ('<div class="extra"><h3>견적서 비교 (읽은 값)</h3><div class="scroll"><table class="detail">'
                  "<tr><th>업체</th><th>금액(백만원)</th><th>설립연도</th><th>선정</th></tr>"
                  + "".join(f"<tr><td>{escape(q['업체'])}</td><td>{fmt_num(q.get('금액'))}</td>"
                            f"<td>{q.get('설립연도') or '—'}</td><td>{'선정' if q.get('선정') else ''}</td></tr>"
                            for q in quotes)
                  + "</table></div>"
                  + (f'<p class="pre">선정 사유: {escape(d["선정사유"])}</p>' if d.get("선정사유") else "")
                  + "</div>")
    # 표준양식 1장 외의 추가 슬라이드 (제목 = 부가 항목명 등) — 원문 그대로
    # 제목이 부가 항목명인 추가 슬라이드는 위 '투자 상세 부가'에 이미 실렸으므로 다시 싣지 않는다
    extra = [sl for sl in d.get("슬라이드", []) if sl.get("번호") != d.get("부가_슬라이드", 1)
             and (sl.get("제목") or "").strip() not in keys]
    ex_html = "".join(f'<div class="extra"><h3>추가 슬라이드 {sl["번호"]} · {escape(sl.get("제목") or "")}</h3>'
                      f'{_pre(sl.get("텍스트"))}</div>' for sl in extra)
    conv = ""
    if d.get("변환원본"):
        conv = (f'<p class="conv">다른 양식으로 제출된 신청서를 표준양식으로 옮긴 것 (원본 '
                f'{escape(Path(d["변환원본"]["경로"]).name)}, 내용 보존 검증 통과)</p>')
    note = escape(row["비고"]) + (' <span class="chk">확인 필요</span>' if row["확인필요"] else "")
    return f"""
<section class="card app" id="app{no}">
  <div class="apphead"><h2>[{no}] {escape(row['투자명'])}</h2>{_orig_btn(row, d)}
    <span class="sub">{escape(row['투자유형'])} · 계획 {escape(row['계획NO'])} · 경영계획 {fmt_num(row['경영계획'])} → 실행 {fmt_num(row['실행금액'])}
    (<span class="{_cls(row['차이금액'])}">{_diff(row['차이금액'])}</span>){' · ' + note if row['비고'] or row['확인필요'] else ''}</span></div>
  <div class="scroll"><table class="form">
    <tr><td class="l">투자명</td><td><b>{escape(d.get('투자명', ''))}</b></td></tr>
    <tr><td class="l">투자상세</td><td><table class="detail">{detail}</table></td></tr>
    <tr><td class="l">투자효과</td><td>{_pre(d.get('투자효과'))}</td></tr>
    <tr><td class="l">투자 상세 부가</td><td>{부가_html}</td></tr>
  </table></div>
  {q_html}{ex_html}{conv}
  <p class="conv">원본 파일: {escape(d.get('파일', ''))}</p>
</section>"""


def write_meeting_html(rows: list[dict], 미신청: list[dict], 보류: list[dict], path: Path,
                       today: dt.date | None = None) -> Path:
    today = today or dt.date.today()
    plan = sum(r["경영계획"] for r in rows)
    exe = sum(r["실행금액"] for r in rows)
    body_rows = "".join(
        f'<tr{' class="need"' if r["확인필요"] else ''}><td>{r["_no"]}</td><td><a class="jump" href="#app{r["_no"]}" data-target="app{r["_no"]}" title="아래 신청서로 이동">{escape(r["투자명"])}</a></td>'
        f'<td>{escape(r["투자유형"])}</td><td>{escape(r["계획NO"])}</td>'
        f'<td class="n">{fmt_num(r["경영계획"])}</td><td class="n">{fmt_num(r["실행금액"])}</td>'
        f'<td class="n {_cls(r["차이금액"])}">{_diff(r["차이금액"])}</td>'
        f'<td>{escape(r["비고"])}{" <span class=chk>확인 필요</span>" if r["확인필요"] else ""}</td></tr>'
        for r in rows)
    notes = []
    if 미신청:
        notes.append("계획 미신청: " + " · ".join(
            f"{escape(m['계획NO'])} {escape(m['투자명'])} ({fmt_num(m['경영계획'])})" for m in 미신청))
    # 보류(미상정) 건은 회의자료에 넣지 않는다 — 실무자가 신청자에게 전달할 내용 (CLAUDE.md §5)
    notes.append("차이 = 실행 − 계획 (이번 신청분) · 단위 백만원 · <span class=chk>확인 필요</span> = 계획 매칭을 회의에서 확인할 건")
    sections = []
    for r in rows:
        f = OUT / "extract" / f"{r['신청서ID']}.json"
        d = read_json(f) if f.exists() else {"투자명": r["투자명"], "상세": {}, "부가": {}}
        sections.append(_app_section(r["_no"], r, d))
    toc = "".join(f'<li><a class="jump" href="#app{r["_no"]}" data-target="app{r["_no"]}">[{r["_no"]}] {escape(r["투자명"])}</a></li>' for r in rows)
    html = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>투자 심의 회의자료</title>
<style>{CSS}</style></head><body><main>
<h1>투자 심의 요약</h1>
<div class="meta">{today:%Y-%m-%d} · 상정 {len(rows)}건 · 단위 백만원 · 점검을 통과하고 검토자가 확인한 신청서만 실었습니다</div>
<div class="tiles">
  <div class="tile"><b>{len(rows)}</b><span>상정</span></div>
  <div class="tile"><b>{fmt_num(plan)}</b><span>경영계획 합계</span></div>
  <div class="tile"><b>{fmt_num(exe)}</b><span>실행금액 합계</span></div>
  <div class="tile"><b class="{_cls(exe - plan)}">{_diff(exe - plan)}</b><span>차이 (실행 − 계획)</span></div>
</div>
<section class="card"><div class="scroll"><table class="cover">
<tr><th>No</th><th>투자 심의</th><th>유형</th><th>계획 NO</th><th>경영계획</th><th>실행금액</th><th>차이금액</th><th>비고</th></tr>
{body_rows}
<tr class="sum"><td colspan="4">합계</td><td class="n">{fmt_num(plan)}</td><td class="n">{fmt_num(exe)}</td>
<td class="n {_cls(exe - plan)}">{_diff(exe - plan)}</td><td></td></tr>
</table></div>
<ul class="notes">{"".join(f"<li>{n}</li>" for n in notes)}</ul></section>
<section class="card"><h2 style="font-size:16px;margin:0 0 8px">첨부 — 상정 신청서</h2><ol class="toc" style="list-style:none;padding-left:0">{toc}</ol></section>
{"".join(sections)}
<footer>표지 숫자는 경영계획·집행이력과 자동 대조한 값입니다 (M-0~M-9). 신청서 내용은 점검 때 읽은 원문 그대로입니다.</footer>
</main><script>{JUMP_JS}</script></body></html>"""
    path.write_text(html, encoding="utf-8")
    return path
