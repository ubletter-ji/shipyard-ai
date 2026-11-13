"""검토 결과를 HTML 한 장으로 (out/검토결과.html). review.py report 가 호출한다.

오류 설명은 이 화면으로 보고, 검토자 확인(O) 은 out/검토결과.xlsx 에 표시한다 (meeting 이 읽는 승인 기록).
"""
from __future__ import annotations

import datetime as dt
from html import escape
from pathlib import Path

from common import fmt_num

_STATE = {
    "검토 통과": ("ok", "검토 통과"),
    "보완 요청": ("bad", "보완 요청"),
    "검토자 확인 필요": ("warn", "검토자 확인 필요"),
}

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1d2433;--sub:#5b6474;--line:#e3e6eb;--accent:#1f3a5f;
--ok:#1e7a46;--ok-bg:#e6f4ec;--bad:#b42318;--bad-bg:#fdecea;--warn:#8a5a00;--warn-bg:#fff4d6;--na:#8a93a3;--on:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#0b0d11;--card:#20262f;--ink:#e9ecf1;--sub:#a7b0bf;--line:#3c4553;
--accent:#8fb3e0;--ok:#5cc58a;--ok-bg:#16301f;--bad:#f2867b;--bad-bg:#3a1c1a;--warn:#e8c06a;--warn-bg:#352a12;--na:#6c7584;--on:#0b0d11}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
font:15px/1.55 "Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif}
main{max-width:1080px;margin:0 auto;padding:28px 16px 60px}
h1{font-size:24px;margin:0 0 4px;color:var(--accent)}.meta{color:var(--sub);font-size:13px;margin-bottom:20px}
.sum{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:14px}
.tile{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
.tile b{display:block;font-size:26px}.tile span{color:var(--sub);font-size:13px}
.tile.t-go{border-top:4px solid var(--ok)}.tile.t-go b{color:var(--ok)}
.tile.t-hold{border-top:4px solid var(--na)}
.tile.t-bad{border-top:4px solid var(--bad)}.tile.t-bad b{color:var(--bad)}
.tile.t-warn{border-top:4px solid var(--warn)}.tile.t-warn b{color:var(--warn)}
.handoff{background:var(--card);border:1px solid var(--line);border-left:4px solid var(--ok);border-radius:10px;
padding:12px 16px;margin-bottom:22px}.handoff p{margin:4px 0}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;margin-bottom:16px;overflow:hidden}
.head{display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;padding:14px 18px;border-bottom:1px solid var(--line)}
.head h2{font-size:18px;margin:0;flex:1 1 260px}.id{color:var(--sub);font-size:12px;display:block;font-weight:400}
.badge{border-radius:3px;padding:3px 11px;font-size:13px;font-weight:700;white-space:nowrap}
.badge.ok{background:var(--ok);color:var(--on)}.badge.bad{background:var(--bad);color:var(--on)}.badge.warn{background:var(--warn);color:var(--on)}.badge.pass{border:1px solid var(--na);color:var(--ink);background:transparent}
.ok{color:var(--ok);background:var(--ok-bg)}.bad{color:var(--bad);background:var(--bad-bg)}
.warn{color:var(--warn);background:var(--warn-bg)}.pass{border:1px solid var(--line);color:var(--sub)}
.facts{display:flex;flex-wrap:wrap;gap:4px 18px;padding:10px 18px;color:var(--sub);font-size:13px;border-bottom:1px solid var(--line)}
.body{padding:12px 18px 16px}.clean{color:var(--ok);margin:4px 0}
.issue{border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:10px 0}
.issue.i-bad{border-left:5px solid var(--bad)}.issue.i-warn{border-left:5px solid var(--warn)}
.issue.i-bad .t{color:var(--bad)}.issue.i-warn .t{color:var(--warn)}
.issue .t{font-weight:700}.issue .t small{font-weight:400;color:var(--sub);margin-left:6px}
.why{margin:4px 0 0}.ask{margin-top:8px;padding:8px 10px;border-radius:8px;background:var(--bg);display:flex;gap:10px;align-items:flex-start}
.ask p{margin:0;flex:1}.ask button{border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:6px;
padding:3px 10px;cursor:pointer;font:inherit;font-size:12px;white-space:nowrap}
.todo{margin-top:6px;font-size:13px;color:var(--warn)}
.warns{font-size:13px;color:var(--sub);margin:8px 0 0;padding-left:18px}
details{margin-top:12px}summary{cursor:pointer;color:var(--sub);font-size:13px}
.tbl{overflow-x:auto}table{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px}
th,td{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}
th{color:var(--sub);font-weight:600}td.v{white-space:nowrap;font-weight:700}
td.v.충족{color:var(--ok)}td.v.보완{color:var(--bad);background:var(--bad-bg)}td.v.판단{color:var(--warn);background:var(--warn-bg)}td.v.해당{color:var(--na);font-weight:400}
footer{color:var(--sub);font-size:12px;margin-top:24px}
a.jump{color:var(--accent);text-decoration:none;border-bottom:1px dashed currentColor;cursor:pointer}
a.jump::after{content:" ↓";font-size:12px}
a.jump:hover{border-bottom-style:solid}
.flash{animation:flash 1.6s ease-out}
@keyframes flash{0%{box-shadow:0 0 0 4px var(--accent)}100%{box-shadow:0 0 0 0 transparent}}

.origbtn{font-size:13px;padding:3px 11px;border:1px solid var(--line);border-radius:999px;text-decoration:none;color:var(--accent);white-space:nowrap}
"""

JS = """
document.querySelectorAll('a.jump').forEach(a=>a.addEventListener('click',e=>{e.preventDefault();
 const t=document.getElementById(a.dataset.target);if(!t)return;
 t.scrollIntoView({behavior:'auto',block:'start'});t.classList.remove('flash');void t.offsetWidth;t.classList.add('flash')}));
document.querySelectorAll('.ask button').forEach(b=>b.addEventListener('click',()=>{
const t=b.previousElementSibling.textContent;const done=()=>{b.textContent='복사됨';setTimeout(()=>b.textContent='복사',1200)};
if(navigator.clipboard){navigator.clipboard.writeText(t).then(done,()=>{})}else{const r=document.createRange();
r.selectNodeContents(b.previousElementSibling);getSelection().removeAllRanges();getSelection().addRange(r);document.execCommand('copy');done()}}));"""


def _issue(x: dict) -> str:
    kind = "i-bad" if x["판정"] == "보완" else "i-warn"
    h = [f'<div class="issue {kind}"><div class="t">{escape(x["규칙ID"])} {escape(x["검토항목"])}'
         f'<small>{escape(x["판정"])}</small></div><p class="why">{escape(x["근거"])}</p>']
    if x["판정"] == "보완" and x.get("보완요청문구"):
        h.append(f'<div class="ask"><p>{escape(x["보완요청문구"])}</p><button type="button">복사</button></div>')
    if x.get("AI대기"):
        h.append('<div class="todo">AI 판단 대기 — 점검 에이전트가 아직 판단하지 않음 (review.py pending)</div>')
    elif x["판정"] == "판단 필요":
        done = (x.get("검토자확인") or "").strip()
        h.append(f'<div class="todo">{"검토자 확인: " + escape(done) if done else "검토자 판단 필요 — 동의하면 검토결과.xlsx ① 상세 검토의 검토자 확인 칸에 O"}</div>')
    h.append("</div>")
    return "".join(h)


def _orig_btn(r: dict) -> str:
    from system_html import orig_button
    src = Path(r["변환원본"]["경로"]) if r.get("변환원본") else Path(r.get("경로", ""))
    return orig_button(r["신청서ID"], src.resolve().as_uri()) if src.exists() else ""


def _card(r: dict) -> str:
    sm = r["요약"]
    unconvertible = any(x["규칙ID"] == "R-100" and x["판정"] == "보완" for x in r["rows"])
    cls, label = ("bad", "점검 불가") if unconvertible else _STATE.get(sm["종합상태"], ("warn", sm["종합상태"]))
    go = sm["회의상정"] == "상정"
    facts = []
    if r.get("투자유형"):
        facts.append(f"유형 {escape(r['투자유형'])}" + (f" (에이전트 판단: {escape(r['에이전트판단유형'])})"
                     if r.get("에이전트판단유형") and r["에이전트판단유형"] != r["투자유형"] else ""))
    if r.get("투자금액") is not None:
        facts.append(f"금액 {fmt_num(r['투자금액'])}백만원")
    if r.get("AREA"):
        facts.append(f"AREA {escape(r['AREA'])}")
    if r.get("계획NO"):
        facts.append(f"계획 NO {escape(r['계획NO'])}")
    if r.get("변환원본"):
        facts.append(f"다른 양식 → 표준양식 변환본 (원본 {escape(Path(r['변환원본']['경로']).name)})")
    bad = [x for x in r["rows"] if x["판정"] == "보완"]
    judge = [x for x in r["rows"] if x["판정"] == "판단 필요"]
    body = []
    if not bad and not judge:
        body.append('<p class="clean">잘못된 점 없음 — 회의자료 에이전트로 전달</p>')
    else:
        if bad:
            body.append(f'<p><b>보완 {len(bad)}건</b> — 신청자에게 보낼 문구</p>' + "".join(_issue(x) for x in bad))
        if judge:
            body.append(f'<p><b>판단 필요 {len(judge)}건</b> — 검토자 확인</p>' + "".join(_issue(x) for x in judge))
    warns = [w for w in r.get("경고", []) if not w.startswith("표준양식 변환본")]
    if warns:
        body.append('<ul class="warns">' + "".join(f"<li>{escape(w)}</li>" for w in warns) + "</ul>")
    rows = "".join(
        f'<tr><td>{escape(x["규칙ID"])}</td><td>{escape(x["검토항목"])}</td>'
        f'<td class="v {escape(x["판정"].split()[0])}">{escape(x["판정"])}</td><td>{escape(x["근거"])}</td></tr>'
        for x in r["rows"])
    body.append(f'<details><summary>전체 규칙 판정 {len(r["rows"])}건 보기</summary><div class="tbl"><table>'
                f'<tr><th>규칙</th><th>항목</th><th>판정</th><th>근거</th></tr>{rows}</table></div></details>')
    return (f'<section class="card" id="{escape(r["신청서ID"])}"><div class="head"><h2>{escape(r["투자명"] or r["파일"])}'
            f'<span class="id">{escape(r["신청서ID"])}</span></h2><span class="badge {cls}">{label}</span>'
            f'<span class="badge {"ok" if go else "pass"}">{"회의 상정" if go else "보류"}</span>'
            f'{_orig_btn(r)}</div>'
            f'<div class="facts">{"<span>" + "</span><span>".join(facts) + "</span>" if facts else ""}</div>'
            f'<div class="body">{"".join(body)}</div></section>')


def write_html(reviews: list[dict], path: Path) -> Path:
    go = [r for r in reviews if r["요약"]["회의상정"] == "상정"]
    held = [r for r in reviews if r["요약"]["회의상정"] != "상정"]
    n_bad = sum(r["요약"]["보완"] for r in reviews)
    n_judge = sum(r["요약"]["판단 필요"] for r in reviews)
    # 문제 있는 신청서를 위로
    order = sorted(reviews, key=lambda r: (r["요약"]["회의상정"] == "상정", r["신청서ID"]))
    handoff = ("<p><b>회의자료 에이전트로 넘기는 신청서</b> — " +
               (", ".join(f'<a class="jump" href="#{escape(r["신청서ID"])}" data-target="{escape(r["신청서ID"])}" title="아래 검토 결과로 이동">{escape(r["투자명"])}</a>' for r in go) or "없음") +
               "</p><p><b>보류</b> — " +
               (", ".join(f'<a class="jump" href="#{escape(r["신청서ID"])}" data-target="{escape(r["신청서ID"])}" title="아래 검토 결과로 이동">{escape(r["투자명"] or r["파일"])}</a>' for r in held) or "없음") +
               "</p>")
    html = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>투자신청서 검토 결과</title>
<style>{CSS}</style></head><body><main>
<h1>투자신청서 검토 결과</h1>
<div class="meta">{dt.datetime.now():%Y-%m-%d %H:%M} · 신청서 {len(reviews)}건 · 원본 신청서는 수정하지 않았습니다</div>
<div class="sum">
<div class="tile t-go"><b>{len(go)}</b><span>회의 상정</span></div>
<div class="tile t-hold"><b>{len(held)}</b><span>보류</span></div>
<div class="tile t-bad"><b>{n_bad}</b><span>보완 항목 (신청자 회신)</span></div>
<div class="tile t-warn"><b>{n_judge}</b><span>판단 필요 (검토자 확인)</span></div>
</div>
<div class="handoff">{handoff}</div>
{"".join(_card(r) for r in order)}
<footer>판단 필요 항목에 동의하면 out/검토결과.xlsx ① 상세 검토의 '검토자 확인' 칸에 O 를 적고 저장하세요.
보완 0건이고 판단 필요가 모두 확인된 신청서만 회의자료 에이전트가 받습니다 (R-401).</footer>
</main><script>{JS}</script></body></html>"""
    path.write_text(html, encoding="utf-8")
    return path
