"""투자신청서 처리 결과 한 화면 (out/투자시스템_결과.html).

접수 → 검토 → 상정 → 회의자료 흐름 요약 + 탭 2개(① 투자검토 결과, ② 회의자료).
각 탭은 기존 화면(out/검토결과.html, out/투자회의자료_*.html)을 그대로 한 파일 안에 담는다.
review.py report 와 meeting.py 가 끝날 때마다 다시 만든다.
"""
from __future__ import annotations

import datetime as dt
from html import escape
from pathlib import Path

from common import OUT, ROOT

CSS = """
:root{--bg:#f6f7f9;--card:#fff;--ink:#1d2433;--sub:#5b6474;--line:#e3e6eb;--accent:#1f3a5f;
--ok:#1e7a46;--bad:#b42318;--warn:#8a5a00;--tab:#eef1f5}
@media (prefers-color-scheme:dark){:root{--bg:#111418;--card:#1a1f26;--ink:#e6e9ee;--sub:#9aa3b2;--line:#2b323c;
--accent:#8fb3e0;--ok:#5cc58a;--bad:#f2867b;--warn:#e8c06a;--tab:#222a34}}
*{box-sizing:border-box}html,body{margin:0;background:var(--bg);color:var(--ink);
font:15px/1.55 "Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif}
header{max-width:1160px;margin:0 auto;padding:24px 16px 0}
h1{font-size:24px;margin:0 0 2px;color:var(--accent)}.meta{color:var(--sub);font-size:13px}
.flow{display:flex;flex-wrap:wrap;align-items:stretch;gap:8px;margin:16px 0}
.step{flex:1 1 170px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.step h3{margin:0 0 4px;font-size:13px;color:var(--sub);font-weight:600}
.step b{font-size:24px}.step .s{font-size:13px;color:var(--sub)}
.arrow{align-self:center;color:var(--sub);font-size:20px}
@media (max-width:700px){.arrow{display:none}}
.ok{color:var(--ok)}.bad{color:var(--bad)}.warn{color:var(--warn)}
nav{max-width:1160px;margin:0 auto;padding:0 16px;display:flex;gap:6px;border-bottom:1px solid var(--line)}
nav button{font:inherit;font-weight:700;padding:10px 16px;border:1px solid var(--line);border-bottom:none;
border-radius:10px 10px 0 0;background:var(--tab);color:var(--sub);cursor:pointer}
nav button[aria-selected=true]{background:var(--card);color:var(--accent)}
.panel{display:none}.panel.on{display:block}header,nav{position:relative;z-index:1;background:var(--bg)}
iframe{display:block;width:100%;height:calc(100vh - 56px);min-height:520px;border:0;border-top:1px solid var(--line);background:var(--bg)}
.empty{max-width:1160px;margin:24px auto;padding:24px 16px;color:var(--sub)}

.origs{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:0 0 14px}
.origs span{color:var(--sub);font-size:13px;margin-right:4px}
.ob{font:inherit;font-size:13px;padding:4px 10px;border:1px solid var(--line);border-radius:999px;background:var(--card);
color:var(--ink);cursor:pointer}.ob:hover{border-color:var(--accent);color:var(--accent)}
.modal{position:fixed;inset:0;background:rgba(10,14,20,.6);display:none;z-index:10;padding:16px;overflow:auto}
.modal.on{display:block}
.mbox{max-width:1100px;margin:0 auto;background:var(--card);border-radius:12px;padding:16px 18px}
.mhead{display:flex;flex-wrap:wrap;gap:8px 12px;align-items:center;margin-bottom:10px}
.mhead h2{font-size:18px;margin:0;flex:1 1 300px}.mhead a{font-size:13px}
.mclose{font:inherit;border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:8px;padding:4px 12px;cursor:pointer}
.mbox h3{font-size:14px;margin:14px 0 6px;color:var(--sub)}
.mbox img{display:block;width:100%;height:auto;border:1px solid var(--line);border-radius:6px;margin:0 0 10px;background:#fff}
.mbox pre{white-space:pre-wrap;font:inherit;font-size:14px;border:1px solid var(--line);border-radius:6px;padding:10px;margin:0 0 10px}
.cmp{display:grid;grid-template-columns:1fr 1fr;gap:14px}@media (max-width:800px){.cmp{grid-template-columns:1fr}}

#start{max-width:920px;margin:0 auto;padding:40px 16px 60px}
#start h1{font-size:26px}
.drop{display:block;margin:18px 0 14px;border:2px dashed var(--line);border-radius:14px;padding:26px 16px;text-align:center;
background:var(--card);cursor:pointer;color:var(--sub)}
.drop.over,.drop:hover{border-color:var(--accent);color:var(--accent)}
.drop b{display:block;font-size:17px;color:var(--ink);margin-bottom:4px}
.drop input{display:none}
.files{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.files th,.files td{padding:9px 12px;border-bottom:1px solid var(--line);text-align:left;vertical-align:middle;font-size:14px}
.files th{color:var(--sub);font-weight:600;background:var(--tab)}
.files img{width:96px;height:54px;object-fit:cover;border:1px solid var(--line);border-radius:4px;background:#fff;display:block}
.files .ph{width:96px;height:54px;border:1px dashed var(--line);border-radius:4px}
.st{font-size:12px;font-weight:700;padding:2px 9px;border-radius:999px;white-space:nowrap}
.st.ok{background:#e6f4ec}.st.wait{color:var(--warn);background:#fff4d6}
@media (prefers-color-scheme:dark){.st.ok{background:#16301f}.st.wait{background:#352a12}}
.go{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-top:16px}
.go button{font:inherit;font-size:17px;font-weight:700;padding:12px 34px;border:0;border-radius:10px;background:var(--accent);
color:#fff;cursor:pointer}
@media (prefers-color-scheme:dark){.go button{color:#111418}}
.go .s{color:var(--sub);font-size:13px}
.steps{list-style:none;padding:0;margin:18px 0 0;display:none}.steps.on{display:block}
.steps li{padding:6px 0;color:var(--sub)}.steps li.done{color:var(--ok)}.steps li.run{color:var(--ink);font-weight:700}
.steps li::before{content:'○ '}.steps li.done::before{content:'✔ '}.steps li.run::before{content:'▶ '}
.notyet{max-width:1160px;margin:0 auto 10px;padding:10px 14px;border:1px solid var(--line);border-left:4px solid var(--warn);
border-radius:10px;background:var(--card);font-size:14px;display:none}
.back{font-size:13px;margin-left:10px}

/* ---- 업무 개선 비교 (현재 업무 vs 에이전트 적용 후) */
.cmpbtn{font:inherit;font-size:14px;font-weight:700;padding:8px 16px;border:1px solid var(--accent);border-radius:999px;
background:var(--card);color:var(--accent);cursor:pointer;white-space:nowrap}
.cmpbtn:hover{background:var(--accent);color:var(--card)}
.toprow{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between}
#compare{max-width:1160px;margin:0 auto;padding:28px 16px 48px}
#compare h1{font-size:24px;overflow-wrap:anywhere}
.lanes>div,.lanes .row>div{min-width:0;overflow-wrap:anywhere}
.lanes{display:grid;grid-template-columns:120px 1fr 1fr;gap:0;margin-top:18px;background:var(--card);
border:1px solid var(--line);border-radius:14px;overflow:hidden}
.lanes .hd{padding:12px 14px;font-weight:700;font-size:15px;border-bottom:1px solid var(--line);background:var(--tab)}
.lanes .hd.after{color:var(--accent)}
.lanes .stg,.lanes .cell{padding:12px 14px;border-bottom:1px solid var(--line);transition:opacity .25s,background .25s}
.lanes .stg{font-weight:700;color:var(--sub);font-size:14px}
.lanes .row:last-child>div{border-bottom:none}
.lanes .row{display:contents}
.cell b{display:block;font-size:15px;margin-bottom:2px}.cell .d{font-size:13px;color:var(--sub)}
.who{display:inline-block;font-size:11px;font-weight:700;padding:1px 8px;border-radius:999px;margin-bottom:4px}
.who.h{background:#fde8e6;color:#9a2a1f}.who.a{background:#e3ecf8;color:#1f3a5f}.who.f{background:#e6f4ec;color:#1e6b3e}
@media (prefers-color-scheme:dark){.who.h{background:#3a1e1b;color:#f2a49b}.who.a{background:#1d2a3b;color:#a9c6ec}
.who.f{background:#16301f;color:#8fd8ab}}
.cell.focus{background:linear-gradient(90deg,rgba(30,122,70,.10),transparent)}
.lanes.stepping .row>div{opacity:.28}.lanes.stepping .row.cur>div{opacity:1}
.lanes.stepping .row.cur .stg{color:var(--accent)}
.ctl{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:14px}
.ctl button{font:inherit;font-size:14px;font-weight:700;padding:8px 16px;border:1px solid var(--line);border-radius:10px;
background:var(--card);color:var(--ink);cursor:pointer}
.ctl button.pri{background:var(--accent);border-color:var(--accent);color:#fff}
@media (prefers-color-scheme:dark){.ctl button.pri{color:#111418}}
.ctl .s{color:var(--sub);font-size:13px}
.say{margin-top:12px;padding:12px 16px;border-left:4px solid var(--accent);background:var(--card);border-radius:10px;
font-size:15px;min-height:48px}
.sum{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;margin-top:16px}
.sum .k{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.sum .k h3{margin:0 0 4px;font-size:13px;color:var(--sub);font-weight:600}.sum .k b{font-size:22px}
.sum .k .s{font-size:12px;color:var(--sub)}
.legend{font-size:12px;color:var(--sub);margin-top:8px}
@media (max-width:760px){.lanes{grid-template-columns:1fr}.lanes .hd.stghd{display:none}
.lanes .stg{background:var(--tab);border-bottom:none;padding-bottom:4px}
.lanes .hd{display:none}.lanes .cell::before{display:block;font-size:11px;color:var(--sub);margin-bottom:2px}
.lanes .cell.bf::before{content:'현재 업무'}.lanes .cell.af::before{content:'개선 후'}}
"""

JS = r"""
const tabs=[...document.querySelectorAll('nav button')];
function show(id){tabs.forEach(b=>{const on=b.dataset.t===id;b.setAttribute('aria-selected',on);
document.getElementById(b.dataset.t).classList.toggle('on',on)})}
tabs.forEach(b=>b.addEventListener('click',()=>show(b.dataset.t)));


const ORIG=JSON.parse(document.getElementById('origdata').textContent);
const modal=document.getElementById('modal');
function imgs(list,texts){if(list.length)return list.map((s,i)=>`<img alt="슬라이드 ${i+1}" src="${s}">`).join('');
 return texts.map((t,i)=>`<pre>슬라이드 ${i+1}\n${t.replace(/[&<]/g,c=>c=='&'?'&amp;':'&lt;')}</pre>`).join('')||'<p>미리보기 없음</p>'}
function openOrig(id){const o=ORIG[id];if(!o)return;
 let body=`<div class="mhead"><h2>원본 — ${o.name}</h2><a href="${o.href}">PowerPoint 로 열기</a>
 <button class="mclose" type="button">닫기 (Esc)</button></div><p style="margin:0 0 8px;color:var(--sub);font-size:13px">${o.file}</p>`;
 if(o.conv){body+=`<div class="cmp"><div><h3>제출 원본 (다른 양식)</h3>${imgs(o.images,o.texts)}</div>
   <div><h3>표준양식 변환본 (내용 그대로)</h3>${imgs(o.conv.images,o.conv.texts)}</div></div>`}
 else body+=imgs(o.images,o.texts);
 modal.querySelector('.mbox').innerHTML=body;modal.classList.add('on');
 modal.querySelector('.mclose').onclick=closeOrig}
function closeOrig(){modal.classList.remove('on')}
modal.addEventListener('click',e=>{if(e.target===modal)closeOrig()});
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeOrig()});
document.querySelectorAll('.ob').forEach(b=>b.addEventListener('click',()=>openOrig(b.dataset.orig)));
window.addEventListener('message',e=>{if(e.data&&e.data.orig)openOrig(e.data.orig)});
if(location.hash.startsWith('#orig='))openOrig(decodeURIComponent(location.hash.slice(6)));

// ---- 첨부 → 조회
const startEl=document.getElementById('start'), resultEl=document.getElementById('result');
const added=[];
function fmtSize(n){return n>1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,Math.round(n/1024))+' KB'}
function addFiles(list){const tb=document.querySelector('.files tbody');
 [...list].forEach(f=>{if(!/\.pptx?$/i.test(f.name))return;
  if([...tb.querySelectorAll('td.fn')].some(td=>td.textContent===f.name))return;
  added.push(f.name);
  tb.insertAdjacentHTML('beforeend',`<tr><td><div class="ph"></div></td><td class="fn">${f.name.replace(/[&<]/g,'')}</td>
   <td>${fmtSize(f.size)}</td><td><span class="st wait">점검 전 — inbox 에 넣고 점검 실행 필요</span></td></tr>`)});
 document.getElementById('cnt').textContent=tb.querySelectorAll('tr').length}
const drop=document.querySelector('.drop'), inp=drop.querySelector('input');
inp.addEventListener('change',()=>addFiles(inp.files));
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('over')}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('over')}));
drop.addEventListener('drop',e=>addFiles(e.dataTransfer.files));
function showResult(){startEl.hidden=true;resultEl.hidden=false;window.scrollTo(0,0);
 const ny=document.querySelector('.notyet');
 if(added.length){ny.style.display='block';
  ny.innerHTML='<b>점검 전 신청서 '+added.length+'건</b> — '+added.map(n=>n.replace(/[&<]/g,'')).join(', ')+
  ' : 브라우저에서 첨부만 된 상태입니다. inbox 에 넣고 “invest-reviewer 로 inbox 신청서 점검해줘”를 실행하면 결과에 들어갑니다.'}}
document.getElementById('run').addEventListener('click',()=>{
 const steps=[...document.querySelectorAll('.steps li')];document.querySelector('.steps').classList.add('on');
 document.getElementById('run').disabled=true;let i=0;
 (function tick(){steps.forEach((s,k)=>s.className=k<i?'done':(k===i?'run':''));
  if(i<steps.length){i++;setTimeout(tick,450)}else setTimeout(showResult,300)})()});
document.getElementById('toStart').addEventListener('click',e=>{e.preventDefault();resultEl.hidden=true;startEl.hidden=false;
 document.getElementById('run').disabled=false;document.querySelector('.steps').classList.remove('on')});
if(/^#(result|meeting|orig=)/.test(location.hash)){startEl.hidden=true;resultEl.hidden=false;
 if(location.hash==='#meeting')requestAnimationFrame(()=>requestAnimationFrame(()=>{show('meeting');window.scrollTo(0,0)}))}

// ---- 업무 개선 비교: 버튼으로 넘어가고, 단계별로 하나씩 짚어 설명
const cmpEl=document.getElementById('compare'), lanes=document.querySelector('.lanes');
const rows=[...document.querySelectorAll('.lanes .row')], say=document.querySelector('.say');
const SAY=JSON.parse(document.getElementById('saydata').textContent);
let cur=-1, from=startEl;
function paint(){lanes.classList.toggle('stepping',cur>=0&&cur<rows.length);
 rows.forEach((r,k)=>r.classList.toggle('cur',k===cur));
 say.innerHTML=cur<0?SAY.intro:(cur<rows.length?SAY.steps[cur]:SAY.end);
 document.getElementById('cmpNext').textContent=cur<rows.length?'▶ 다음 단계':'처음부터';
 document.getElementById('cmpPos').textContent=cur<0?'전체 보기':(cur<rows.length?`${cur+1} / ${rows.length}`:'요약');
 document.querySelector('.sum').style.outline=cur>=rows.length?'2px solid var(--accent)':'none'}
function openCompare(back){from=back;startEl.hidden=true;resultEl.hidden=true;cmpEl.hidden=false;cur=-1;paint();window.scrollTo(0,0)}
function closeCompare(){cmpEl.hidden=true;from.hidden=false;window.scrollTo(0,0)}
document.querySelectorAll('[data-open=compare]').forEach(b=>b.addEventListener('click',e=>{e.preventDefault();
 openCompare(resultEl.hidden?startEl:resultEl)}));
document.getElementById('cmpBack').addEventListener('click',closeCompare);
document.getElementById('cmpNext').addEventListener('click',()=>{cur=cur>rows.length-1?-1:cur+1;paint()});
document.getElementById('cmpPrev').addEventListener('click',()=>{cur=Math.max(-1,cur-1);paint()});
rows.forEach((r,k)=>r.addEventListener('click',()=>{cur=k;paint()}));
document.addEventListener('keydown',e=>{if(cmpEl.hidden||modal.classList.contains('on'))return;
 if(e.key==='ArrowRight'||e.key===' '){e.preventDefault();document.getElementById('cmpNext').click()}
 else if(e.key==='ArrowLeft'){e.preventDefault();document.getElementById('cmpPrev').click()}
 else if(e.key==='Escape')closeCompare()});
if(location.hash==='#compare')openCompare(startEl);
"""

# 업무 개선 비교 — (단계, 현재 업무, 개선 후, 설명). who: h=담당자, a=AI 에이전트, f=사람(집중)
COMPARE = [
    ("접수",
     ("h", "담당자", "신청서 접수", "내부 결재 후 통보된 신청서를 담당자가 모아 정리"),
     ("a", "그대로", "inbox 에 접수", "신청서 pptx 를 inbox 폴더에 넣으면 끝"),
     "접수 방식은 그대로입니다. 신청서 파일을 inbox 에 넣는 것으로 시작합니다."),
    ("양식 점검",
     ("h", "담당자", "양식 확인", "칸이 빠졌는지, 다른 양식인지 하나씩 열어 확인"),
     ("a", "AI 에이전트", "양식 확인 · 표준양식 변환", "다른 양식은 내용 그대로 표준양식으로 옮기고 추가·누락 0건을 자동 검증"),
     "담당자가 파일을 하나씩 열어 보던 양식 확인을 에이전트가 합니다. 다른 양식도 내용을 바꾸지 않고 표준양식으로 옮깁니다."),
    ("내용 점검",
     ("h", "담당자", "내용 점검", "수량·금액·견적·업체 업력·유형별 필수 항목을 사람이 대조, 보완 요청 메일 작성"),
     ("a", "AI 에이전트", "규칙 점검 + 보완 문구 작성", "R-100~R-322 규칙을 스크립트가 확정, 유형 일치·필요성·산출 재계산은 AI 가 판단. 신청자에게 보낼 문장까지 작성"),
     "수량 계산, 견적 금액 대조, 신생 업체 확인처럼 사람이 표를 대조하던 일을 규칙으로 바꿨습니다. 신청자에게 보낼 보완 요청 문장도 바로 쓸 수 있게 만들어 줍니다."),
    ("현장 확인",
     ("h", "담당자", "현장 확인", "점검에 시간을 쓴 뒤 남은 시간으로 현장 확인"),
     ("f", "담당자 · 집중", "현장 확인 + 판단", "에이전트가 정리한 보완·판단 필요 목록을 들고 현장 확인. 판단 필요 행은 검토자가 확인(O)"),
     "담당자의 시간은 여기에 씁니다. 서류 대조는 에이전트가 끝냈으니, 담당자는 현장에서 실제 장비·위치·필요성을 확인하고 판단이 필요한 항목만 결정합니다."),
    ("회의자료",
     ("h", "담당자", "요약장 + 자료 세팅", "경영계획과 대조해 변경 사항을 요약장 한 장에 정리, 신청서 자료 묶기"),
     ("a", "AI 에이전트", "회의자료 자동 작성", "상정 건만 골라 경영계획과 대조(M-0~M-9)한 표지 + 신청서 원본 첨부 + 의장 브리핑"),
     "요약장과 자료 준비는 에이전트가 합니다. 계획 대비 위치·수량·금액 변경과 계획 外 건을 표지에 자동으로 표시합니다."),
    ("회의",
     ("h", "담당자", "회의 진행", ""),
     ("f", "사람", "회의 진행 · 의사결정", "'확인 필요' 표시 건(계획 外 등)을 중심으로 결정"),
     "회의와 결정은 그대로 사람이 합니다. 표지의 '확인 필요' 건에 집중하면 됩니다."),
]


def _compare_html(reviews, go) -> tuple[str, str]:
    """업무 개선 비교 화면 (HTML, 설명 문구 JSON)."""
    import json
    rows = [x for r in reviews for x in r["rows"]]
    checked = sum(1 for x in rows if x["판정"] != "해당 없음")
    fixes = sum(1 for x in rows if x["판정"] == "보완" and x.get("보완요청문구"))
    judge = sum(1 for x in rows if x["판정"] == "판단 필요")
    ai = sum(1 for x in rows if x["규칙ID"] in ("R-201", "R-301", "R-322") and x["판정"] != "해당 없음")

    def cell(c, side):
        who, label, title, desc = c
        cls = "cell " + side + (" focus" if who == "f" else "")
        return (f'<div class="{cls}"><span class="who {who}">{escape(label)}</span><b>{escape(title)}</b>'
                + (f'<div class="d">{escape(desc)}</div>' if desc else "") + "</div>")

    body = "".join(f'<div class="row"><div class="stg">{escape(s)}</div>{cell(b, "bf")}{cell(a, "af")}</div>'
                   for s, b, a, _ in COMPARE)
    human_before = sum(1 for _, b, _, _ in COMPARE if b[0] == "h")
    human_after = sum(1 for _, _, a, _ in COMPARE if a[0] == "f")
    html = f"""<section id="compare" hidden>
  <div class="toprow"><h1>업무 개선 비교 — 투자검토 · 투자회의</h1>
    <button class="cmpbtn" id="cmpBack" type="button">← 돌아가기</button></div>
  <div class="meta">현재 업무는 담당자가 모든 단계를 처리합니다. 에이전트를 적용하면 서류 점검과 회의자료는 AI 가 하고, 담당자는 <b>현장 확인</b>과 판단에 집중합니다.</div>
  <div class="lanes">
    <div class="hd stghd">단계</div><div class="hd">현재 업무</div><div class="hd after">개선 후 (AI 에이전트 적용)</div>
    {body}
  </div>
  <div class="legend"><span class="who h">담당자</span> 사람이 직접 처리 · <span class="who a">AI 에이전트</span> 에이전트가 처리 · <span class="who f">사람 · 집중</span> 사람이 집중할 일</div>
  <div class="ctl"><button id="cmpPrev" type="button">◀ 이전</button><button id="cmpNext" class="pri" type="button">▶ 다음 단계</button>
    <span class="s" id="cmpPos">전체 보기</span><span class="s">· 키보드 ← → 로 넘기기, 단계를 눌러도 됩니다</span></div>
  <div class="say"></div>
  <div class="sum">
    <div class="k"><h3>담당자가 직접 하는 단계</h3><b>{human_before} → {human_after}</b><div class="s">현장 확인 · 회의만 사람이 (나머지는 에이전트)</div></div>
    <div class="k"><h3>이번 점검 · 신청서</h3><b>{len(reviews)}건</b><div class="s">규칙 점검 {checked}행 (해당 없음 제외) · AI 판단 {ai}행</div></div>
    <div class="k"><h3>보완 요청 문구 자동 작성</h3><b>{fixes}건</b><div class="s">신청자에게 그대로 보낼 수 있는 문장</div></div>
    <div class="k"><h3>담당자 판단으로 넘긴 항목</h3><b>{judge}건</b><div class="s">판단 필요 → 현장 확인 후 검토자 확인(O)</div></div>
    <div class="k"><h3>회의 상정</h3><b>{len(go)}건</b><div class="s">표지 · 원본 첨부 · 브리핑 자동 작성</div></div>
  </div>
</section>"""
    say = {"intro": "<b>전체 보기</b> — 왼쪽은 지금 담당자가 하는 일, 오른쪽은 에이전트 적용 후입니다. "
                    "<b>▶ 다음 단계</b>를 누르면 단계별로 짚어 설명합니다.",
           "steps": [f"<b>{i + 1}. {escape(s)}</b> — {escape(t)}" for i, (s, _, _, t) in enumerate(COMPARE)],
           "end": f"<b>요약</b> — 담당자가 직접 하던 {human_before}단계 중 서류 점검과 회의자료는 에이전트가 맡고, "
                  f"담당자는 <b>현장 확인과 판단</b>에 집중합니다. 이번 점검에서도 신청서 {len(reviews)}건의 규칙 {checked}행을 "
                  f"에이전트가 점검했고, 보완 문구 {fixes}건을 만들었으며, 사람이 판단할 항목은 {judge}건으로 줄었습니다."}
    return html, json.dumps(say, ensure_ascii=False).replace("</", "<\\/")


def _orig_data(reviews) -> dict:
    """신청서별 원본 미리보기 데이터: 슬라이드 이미지(data URI) 또는 글자, 변환본이면 둘 다."""
    import base64
    import originals
    from common import read_json
    pairs = []
    for r in reviews:
        cur = Path(r.get("경로", ""))
        src = Path(r["변환원본"]["경로"]) if r.get("변환원본") else cur
        pairs.append((r, src, cur if r.get("변환원본") else None))
    originals.render([p for _, s, c in pairs for p in (s, c) if p])

    def pack(p: Path) -> dict:
        ims = ["data:image/png;base64," + base64.b64encode(i.read_bytes()).decode() for i in originals.images(p)]
        texts = []
        if not ims:
            ex = OUT / "extract" / f"{p.stem}.json"
            if ex.exists():
                texts = [s.get("텍스트", "") for s in read_json(ex).get("슬라이드", [])]
        return {"images": ims, "texts": texts}

    data = {}
    for r, src, conv in pairs:
        if not src.exists():
            continue
        d = {"name": escape(r.get("투자명") or src.stem), "file": escape(src.name), "size": src.stat().st_size,
             "href": src.resolve().as_uri(), **pack(src)}
        if conv and conv.exists():
            d["conv"] = pack(conv)
        data[r["신청서ID"]] = d
    return data


def orig_button(sid: str, href: str) -> str:
    """검토결과·회의자료 HTML 안의 '원본 보기' 버튼 (한 화면 안에서는 모달, 단독으로 열면 파일 열기)."""
    return (f'<a class="origbtn" href="{escape(href)}" data-orig="{escape(sid)}" '
            "onclick=\"if(window.parent!==window){event.preventDefault();"
            "window.parent.postMessage({orig:this.dataset.orig},'*')}\">원본 보기</a>")


def _latest_meeting() -> Path | None:
    files = sorted(OUT.glob("투자회의자료_*.html"))
    return files[-1] if files else None


def build() -> Path:
    from review import load_reviews
    reviews = load_reviews()
    inbox = [p for p in (ROOT / "inbox").glob("*.pptx") if not p.name.startswith("~$")]
    converted = sum(1 for r in reviews if r.get("변환원본"))
    states = [r["요약"]["종합상태"] for r in reviews]
    passed = states.count("검토 통과")
    fix = states.count("보완 요청")
    check = len(states) - passed - fix
    go = [r for r in reviews if r["요약"]["회의상정"] == "상정"]

    review_html = (OUT / "검토결과.html")
    meeting_html = _latest_meeting()
    review_doc = review_html.read_text(encoding="utf-8") if review_html.exists() else ""
    meeting_doc = meeting_html.read_text(encoding="utf-8") if meeting_html else ""

    def frame(doc, title):
        return f'<iframe title="{escape(title)}" srcdoc="{escape(doc, quote=True)}"></iframe>'

    meeting_panel = (frame(meeting_doc, "회의자료") if meeting_doc else
                     '<div class="empty">회의자료가 아직 없습니다. 검토자 확인(O) 후 '
                     '“invest-meeting 으로 회의자료 만들어줘”를 실행하면 여기에 표시됩니다.</div>')
    review_panel = frame(review_doc, "투자검토 결과") if review_doc else '<div class="empty">검토 결과가 없습니다.</div>'
    meeting_label = meeting_html.stem.replace("투자회의자료_", "") if meeting_html else "미생성"
    import json
    orig = _orig_data(reviews)
    chips = "".join(f'<button class="ob" type="button" data-orig="{escape(sid)}">{d["name"]}</button>'
                    for sid, d in orig.items())

    def size(n):
        return f"{n / 1048576:.1f} MB" if n > 1048576 else f"{max(1, round(n / 1024))} KB"

    def thumb(d):
        return '<img alt="" src="' + d["images"][0] + '">' if d["images"] else '<div class="ph"></div>'

    compare_html, say_json = _compare_html(reviews, go)

    file_rows = "".join(
        f'<tr><td>{thumb(d)}</td>'
        f'<td class="fn">{d["file"]}</td><td>{size(d["size"])}</td>'
        f'<td><span class="st ok">첨부됨</span> <button class="ob" type="button" data-orig="{escape(sid)}">미리보기</button></td></tr>'
        for sid, d in orig.items())

    html = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>투자신청서 처리 결과</title>
<style>{CSS}</style></head><body>
<section id="start">
  <div class="toprow"><h1>투자신청서 점검</h1>
    <button class="cmpbtn" type="button" data-open="compare">업무 개선 비교 →</button></div>
  <div class="meta">투자신청서(.pptx)를 첨부하고 조회하면 투자검토 결과와 회의자료를 보여 줍니다. 원본은 수정하지 않습니다.</div>
  <label class="drop"><input type="file" accept=".pptx,.ppt" multiple>
    <b>투자신청서 첨부</b>여기로 끌어다 놓거나 눌러서 파일 선택 (.pptx · 표준양식이 아니어도 됨)</label>
  <table class="files"><thead><tr><th>미리보기</th><th>파일</th><th>크기</th><th>상태</th></tr></thead>
    <tbody>{file_rows}</tbody></table>
  <div class="go"><button id="run" type="button">조회</button>
    <span class="s">첨부 <b id="cnt">{len(orig)}</b>건 · inbox 에 접수된 신청서</span></div>
  <ol class="steps">
    <li>① 양식 확인 — 다른 양식은 표준양식으로 변환 (내용 그대로)</li>
    <li>② 규칙 점검 — R-100~R-322</li>
    <li>③ AI 판단 — 유형 일치 · 필요성 · 산출 재계산</li>
    <li>④ 회의 상정 · 경영계획 대조 · 회의자료</li>
  </ol>
</section>
<div id="result" hidden>
<header>
  <h1>투자신청서 처리 결과 <a href="#" id="toStart" class="back">← 첨부 화면</a>
    <a href="#compare" class="back" data-open="compare">업무 개선 비교</a></h1>
  <div class="meta">{dt.datetime.now():%Y-%m-%d %H:%M} · 원본 신청서는 수정하지 않았습니다</div>
  <div class="flow">
    <div class="step"><h3>① 접수</h3><b>{len(inbox)}</b>건<div class="s">다른 양식 {converted}건 → 표준양식으로 변환</div></div>
    <div class="arrow">→</div>
    <div class="step"><h3>② 투자검토</h3><b class="ok">{passed}</b> 통과 · <b class="bad">{fix}</b> 보완 · <b class="warn">{check}</b> 확인
      <div class="s">보완 = 신청자에게 문구 전달 / 확인 = 검토자 O</div></div>
    <div class="arrow">→</div>
    <div class="step"><h3>③ 회의 상정</h3><b>{len(go)}</b>건<div class="s">보완 0 + 검토자 확인 완료 (R-401)</div></div>
    <div class="arrow">→</div>
    <div class="step"><h3>④ 회의자료</h3><b>{escape(meeting_label)}</b><div class="s">경영계획 대조 표지 + 신청서</div></div>
  </div>
  <div class="origs"><span>원본 신청서 보기</span>{chips}</div>
  <div class="notyet"></div>
</header>
<nav role="tablist">
  <button role="tab" data-t="review" aria-selected="true">① 투자검토 결과</button>
  <button role="tab" data-t="meeting" aria-selected="false">② 회의자료</button>
</nav>
<section id="review" class="panel on">{review_panel}</section>
<section id="meeting" class="panel">{meeting_panel}</section>
</div>
{compare_html}
<div id="modal" class="modal" role="dialog" aria-modal="true"><div class="mbox"></div></div>
<script id="saydata" type="application/json">{say_json}</script>
<script id="origdata" type="application/json">{json.dumps(orig, ensure_ascii=False).replace("</", "<\\/")}</script>
<script>{JS}</script>
</body></html>"""
    path = OUT / "투자시스템_결과.html"
    path.write_text(html, encoding="utf-8")
    return path


if __name__ == "__main__":
    print("[한 화면]", build().relative_to(ROOT))
