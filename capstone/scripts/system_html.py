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
--ok:#1e7a46;--bad:#b42318;--warn:#8a5a00;--tab:#eef1f5;
--fill:#1f3a5f;--on:#fff;--hi:#c2410c;--on-hi:#fff;--hi-bg:#fff1e6;--mute:#9aa1ad}
@media (prefers-color-scheme:dark){:root{--bg:#0b0d11;--card:#20262f;--ink:#e9ecf1;--sub:#a7b0bf;--line:#3c4553;
--accent:#8fb3e0;--ok:#5cc58a;--bad:#f2867b;--warn:#e8c06a;--tab:#2b3440;
--fill:#3a6aa8;--on:#fff;--hi:#f59e4b;--on-hi:#1a1206;--hi-bg:#3a2614;--mute:#7c8594}}
*{box-sizing:border-box}html,body{margin:0;background:var(--bg);color:var(--ink);
font:15px/1.55 "Pretendard","Malgun Gothic","Apple SD Gothic Neo","Noto Sans KR",sans-serif;letter-spacing:-.1px;-webkit-font-smoothing:antialiased}
header{max-width:1160px;margin:0 auto;padding:24px 16px 0}
h1{font-size:26px;font-weight:800;letter-spacing:-.6px;margin:0 0 2px;color:var(--accent)}.meta{color:var(--sub);font-size:13px}
.flow{display:flex;flex-wrap:wrap;align-items:stretch;gap:8px;margin:16px 0}
.step{flex:1 1 170px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.step h3{margin:0 0 4px;font-size:13px;color:var(--sub);font-weight:600}
.step b{font-size:24px}.step .s{font-size:13px;color:var(--sub)}
.arrow{align-self:center;color:var(--sub);font-size:20px}
.step.go{border:2px solid var(--ok)}.step.go h3{color:var(--ok)}.step.go b{color:var(--ok)}
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
#start h1{font-size:30px}
.drop{display:block;margin:18px 0 14px;border:2px dashed var(--line);border-radius:14px;padding:26px 16px;text-align:center;
background:var(--card);cursor:pointer;color:var(--sub)}
.drop.over,.drop:hover{border-color:var(--accent);color:var(--accent)}
.drop b{display:block;font-size:17px;color:var(--ink);margin-bottom:4px}
.drop input{display:none}
.files{width:100%;border-collapse:collapse;background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.files th,.files td{padding:5px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:middle;font-size:13.5px}
.files th{color:var(--sub);font-weight:600;background:var(--tab)}
.files img{width:64px;height:36px;object-fit:cover;border:1px solid var(--line);border-radius:4px;background:#fff;display:block}
.files .ph{width:64px;height:36px;border:1px dashed var(--line);border-radius:4px}
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
.back{font-size:13px;margin-left:10px;color:var(--accent)}

/* ---- 업무 개선안 (AS-IS / TO-BE) — 보고서 장표 형식 */
.cmpbtn{font:inherit;font-size:13px;font-weight:600;padding:6px 14px;border:1px solid var(--line);border-radius:3px;
background:var(--card);color:var(--ink);cursor:pointer;white-space:nowrap}
.cmpbtn:hover{border-color:var(--accent);color:var(--accent)}
.toprow{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between}
#compare{max-width:1160px;margin:0 auto;padding:24px 16px 48px}
.doc{background:var(--card);border:1px solid var(--line);padding:22px 26px 26px}
.dochead{display:flex;flex-wrap:wrap;gap:8px 16px;align-items:center;justify-content:space-between;
background:var(--fill);color:var(--on);margin:-22px -26px 0;padding:18px 26px;border-bottom:4px solid var(--hi)}
.dochead h1{font-size:27px;font-weight:800;margin:0;color:var(--on);letter-spacing:-.8px;overflow-wrap:anywhere;
display:flex;flex-wrap:wrap;align-items:center;gap:6px 12px}
.dochead h1 .ttag{font-size:14px;font-weight:800;letter-spacing:0;padding:3px 10px;background:var(--hi);color:var(--on-hi)}
.dochead .cmpbtn{background:transparent;color:var(--on);border-color:rgba(255,255,255,.55)}
.dochead .cmpbtn:hover{background:var(--on);color:var(--fill)}
.dochead .meta{font-size:12.5px;color:var(--on);opacity:.85}
.doc{counter-reset:sec}
.doc h2{counter-increment:sec;display:flex;align-items:center;gap:10px;font-size:19px;font-weight:800;
letter-spacing:-.5px;margin:28px 0 10px;padding-bottom:7px;color:var(--ink);border-bottom:2px solid var(--line)}
.doc h2::before{content:counter(sec,decimal-leading-zero);font-size:13px;font-weight:800;letter-spacing:0;
background:var(--fill);color:var(--on);padding:3px 8px;line-height:1.3}
.doc .lead{font-size:14px;margin:12px 0 0;color:var(--ink)}
.proc{display:flex;flex-direction:column;gap:8px}
.frow{display:grid;grid-template-columns:72px minmax(0,1fr);gap:8px;align-items:stretch}
.frow .lab{display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:800;letter-spacing:.3px;
background:var(--tab);color:var(--sub)}
.frow .lab.to{background:var(--fill);color:var(--on)}
.chev{display:flex;min-width:0}
.chev .p{flex:1 1 0;min-width:0;margin-left:-8px;padding:11px 8px 11px 24px;text-align:center;cursor:pointer;
background:var(--tab);color:var(--mute);overflow-wrap:anywhere;transition:transform .15s,filter .15s;
clip-path:polygon(0 0,calc(100% - 14px) 0,100% 50%,calc(100% - 14px) 100%,0 100%,14px 50%)}
.chev .p:first-child{margin-left:0;padding-left:12px;
clip-path:polygon(0 0,calc(100% - 14px) 0,100% 50%,calc(100% - 14px) 100%,0 100%)}
.chev .p b{display:block;font-size:14px;font-weight:800;letter-spacing:-.3px}
.chev .p span{display:block;font-size:11.5px;font-weight:600;margin-top:1px;opacity:.9}
.chev .p.same{background:var(--line);color:var(--ink)}
.chev .p.auto{background:var(--fill);color:var(--on)}
.chev .p.key{background:var(--hi);color:var(--on-hi)}
.chev .p.on{transform:translateY(-4px);filter:brightness(1.1) saturate(1.15)}
.chev .p.on:not(.auto):not(.key){background:var(--line);color:var(--ink)}
.tag{display:inline-block;font-size:12px;font-weight:700;padding:2px 8px;border-radius:2px;white-space:nowrap;line-height:1.5}
.tag.auto{background:var(--fill);color:var(--on)}.tag.key{background:var(--hi);color:var(--on-hi)}
.tag.same{border:1px solid var(--line);color:var(--sub);padding:1px 7px}
.legend{display:flex;flex-wrap:wrap;gap:6px 8px;align-items:center;font-size:12px;color:var(--sub);margin-top:8px}
.legend .tag{margin-left:8px}.legend .tag:first-child{margin-left:0}
.bgsec{display:grid;grid-template-columns:minmax(0,1.9fr) 28px minmax(0,1fr);align-items:stretch;gap:0}
.bgt td.prob{color:var(--bad);font-weight:600}
.arrowr{display:flex;align-items:center;justify-content:center;color:var(--fill);font-size:16px}
.dir{border:2px solid var(--fill);display:flex;flex-direction:column}
.dir .dh{background:var(--fill);color:var(--on);font-weight:700;text-align:center;padding:8px 10px;font-size:14px}
.dir ol{margin:0;padding:10px 14px 10px 32px;font-size:15px;font-weight:700;flex:1;display:flex;flex-direction:column;
justify-content:space-around;gap:6px}
.keymsg{display:grid;grid-template-columns:1fr 1.6fr 1fr;margin-top:0;background:var(--fill);color:var(--on)}
.keymsg>div{padding:12px 16px;border-right:1px solid rgba(255,255,255,.25);min-width:0}.keymsg>div:last-child{border-right:none}
.keymsg small{display:block;font-size:12.5px;font-weight:600;opacity:.85}.keymsg b{display:block;font-size:22px;font-weight:800;letter-spacing:-.5px;line-height:1.35;overflow-wrap:anywhere}
.keymsg .k{background:var(--hi);color:var(--on-hi)}
table.cmpt{width:100%;border-collapse:collapse;font-size:13.5px;table-layout:fixed}
table.cmpt th,table.cmpt td{border:1px solid var(--line);padding:8px 10px;vertical-align:top;text-align:left;overflow-wrap:anywhere}
table.cmpt th{background:var(--tab);font-weight:700;color:var(--ink);text-align:center}
table.cmpt th.to{background:var(--fill);color:var(--on)}
table.cmpt td.c{text-align:center;color:var(--sub)}
table.cmpt td.as{color:var(--sub)}table.cmpt td.stg{font-weight:700;color:var(--ink)}
table.cmpt tr.r-auto td.to{font-weight:700;color:var(--ink)}table.cmpt tr.r-auto td.c.to{color:var(--accent)}
table.cmpt tr.r-key td{background:var(--hi-bg)}table.cmpt tr.r-key td.to{font-weight:700;color:var(--hi)}
table.cmpt tr.r-same td{color:var(--mute)}table.cmpt tr.r-same td.stg{color:var(--sub)}
table.cmpt tr.on td{background:var(--tab)}table.cmpt tr.on td:first-child{box-shadow:inset 5px 0 0 var(--ink)}
table.cmpt tr.r-key.on td{background:var(--hi-bg)}
table.cmpt tr[data-step]{cursor:pointer}
table.cmpt td.tob{font-weight:700;color:var(--ink)}
table.cmpt td.num{font-weight:700;font-size:15px;color:var(--accent)}
.ctl{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-top:14px}
.ctl button{font:inherit;font-size:13px;padding:5px 14px;border:1px solid var(--line);border-radius:3px;background:var(--card);
color:var(--ink);cursor:pointer}.ctl button:hover{border-color:var(--accent)}
.ctl .s{color:var(--sub);font-size:12px;margin-left:6px}
.note{margin-top:8px;border:1px solid var(--line);border-left:4px solid var(--fill);background:var(--bg);padding:9px 12px;
font-size:14px;min-height:40px}
.foot{font-size:12px;color:var(--sub);margin-top:8px}
@media (max-width:760px){.doc{padding:16px 14px}.dochead{margin:-16px -14px 0;padding:14px}.dochead h1{font-size:22px}
.frow{grid-template-columns:44px minmax(0,1fr);gap:4px}.frow .lab{font-size:11px;letter-spacing:0}
.chev .p{padding:9px 2px 9px 13px;margin-left:-6px;
clip-path:polygon(0 0,calc(100% - 8px) 0,100% 50%,calc(100% - 8px) 100%,0 100%,8px 50%)}
.chev .p:first-child{padding-left:4px;clip-path:polygon(0 0,calc(100% - 8px) 0,100% 50%,calc(100% - 8px) 100%,0 100%)}
.chev .p b{font-size:11.5px;letter-spacing:-.5px}.chev .p span{font-size:10px}
table.cmpt{font-size:12.5px}table.cmpt th,table.cmpt td{padding:6px}
table.cmpt col.who{width:46px!important}table.cmpt col.stg{width:58px!important}table.cmpt col.chg{width:64px!important}
.bgsec{grid-template-columns:1fr}.arrowr{transform:rotate(90deg);height:26px}
.keymsg{grid-template-columns:1fr}.keymsg>div{border-right:none;border-bottom:1px solid rgba(255,255,255,.25)}
.keymsg b{font-size:17px}}

/* ---- 시작·결과 화면: 업무 개선안과 같은 장표 디자인 */
#start{max-width:1000px;padding:24px 16px 48px}
#start .dochead h1{font-size:28px}
.drop{border-radius:3px;margin:0 0 4px}
.files{border-radius:0}.files img,.files .ph{border-radius:2px}
.st{border-radius:2px}.ob{border-radius:3px}
.go{margin-top:14px}.go button{border-radius:3px;padding:11px 40px;background:var(--fill);color:var(--on);font-weight:800}
.steps.chev{display:none;margin-top:14px}.steps.chev.on{display:flex}
.steps .p{cursor:default;padding-top:10px;padding-bottom:10px}
.steps .p small{display:block;font-size:12px;font-weight:800}.steps .p b{font-size:14px}
.steps .p.run{background:var(--fill);color:var(--on);animation:pulse 1s ease-in-out infinite}
.steps .p.done{background:var(--line);color:var(--ink)}.steps .p.done small::before{content:'✔ ';color:var(--ok)}
@keyframes pulse{50%{filter:brightness(1.25)}}
#runmsg{font-weight:700;color:var(--accent)}
header{max-width:1160px;padding:24px 16px 0}
header .doc{padding-bottom:18px}
.dochead .hbtns{display:flex;gap:6px;flex-wrap:wrap}
.dochead .cmpbtn{background:transparent;color:var(--on);border-color:rgba(255,255,255,.55)}
.dochead .cmpbtn:hover{background:var(--on);color:var(--fill)}
.rflow .p{padding-top:12px;padding-bottom:12px;cursor:default}
.rflow .p small{display:block;font-size:12px;font-weight:700;opacity:.9}
.rflow .p b{font-size:26px;letter-spacing:-.5px;line-height:1.2}
.rflow .p span{font-size:12px}
.origs{margin:0}.origs span{display:none}
nav{border-bottom:3px solid var(--fill);gap:4px;margin-top:14px}
nav button{border-radius:0;font-weight:800;letter-spacing:-.3px;padding:10px 20px}
nav button[aria-selected=true]{background:var(--fill);color:var(--on);border-color:var(--fill)}
.mbox,.notyet{border-radius:3px}
@media (max-width:760px){.rflow .p b{font-size:18px}.rflow .p small,.rflow .p span{font-size:10px}
.go button{width:100%}}

/* ---- 운영 규칙 · 안전장치: 파일 내용 */
.layers.l4{grid-template-columns:repeat(4,minmax(0,1fr))}b.un{font-weight:800;font-size:14px}
@media (max-width:760px){.layers.l4{grid-template-columns:1fr}}
ul.plain{margin:0;padding-left:20px;font-size:14px}ul.plain li{margin:3px 0}ul.plain.in{padding-left:18px;font-size:13px}
.kv{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));border:1px solid var(--line);margin:0 0 10px}
.kv>div{padding:9px 12px;font-size:13.5px;border-right:1px solid var(--line);min-width:0;overflow-wrap:anywhere}
.kv>div:last-child{border-right:none}.kv b{display:block;font-size:12px;color:var(--sub);margin-bottom:3px}
.kv code,table.cmpt code{font-size:12px;background:var(--tab);padding:1px 5px;margin:1px 2px 1px 0;display:inline-block}
.rgrid{border:1px solid var(--line)}
.rg{display:grid;grid-template-columns:88px minmax(0,1fr);border-bottom:1px solid var(--line)}.rg:last-child{border-bottom:none}
.rh{background:var(--tab);font-weight:800;font-size:13px;display:flex;align-items:center;justify-content:center;text-align:center}
.rcs{display:flex;flex-wrap:wrap;gap:6px;padding:8px 10px}
.rc{border:1px solid var(--line);padding:3px 8px;font-size:13px;cursor:help;background:var(--card)}
.rc b{color:var(--accent);font-weight:800}.rc .tag{font-size:10.5px;padding:0 5px;margin-left:2px}
details.more{margin:8px 0 0}details.more summary{cursor:pointer;font-weight:700;font-size:13px;color:var(--accent);margin-bottom:6px}
table.cmpt.mt{margin-top:12px}
@media (max-width:760px){.kv{grid-template-columns:1fr}.kv>div{border-right:none;border-bottom:1px solid var(--line)}
.rg{grid-template-columns:64px minmax(0,1fr)}}
/* ---- 운영 규칙 · 안전장치 */
#guide{max-width:1160px;margin:0 auto;padding:24px 16px 48px}
.layers{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));border:1px solid var(--line)}
.layers>div{border-right:1px solid var(--line);min-width:0}.layers>div:last-child{border-right:none}
.layers h3{margin:0;padding:10px 14px;background:var(--fill);color:var(--on);font-size:15px;font-weight:800;letter-spacing:-.3px}
.layers h3 small{display:block;font-size:11.5px;font-weight:600;opacity:.85}
.layers .lx{padding:10px 14px;font-size:13.5px}
.layers .lx b{display:block;font-size:16px;font-weight:800;letter-spacing:-.4px;margin-bottom:4px}
.layers .lx code{font-size:12px}
.tag.block{background:var(--bad);color:var(--on-hi)}.tag.ask{background:var(--warn);color:var(--on-hi)}
.tag.log{border:1px solid var(--line);color:var(--sub);padding:1px 7px}
.guidechev .p{cursor:default;padding-top:11px;padding-bottom:11px}
.guidechev .p small{display:block;font-size:11.5px;font-weight:700;opacity:.9}
table.cmpt code{font-size:12px;background:var(--tab);padding:0 4px}
@media (max-width:760px){.layers{grid-template-columns:1fr}.layers>div{border-right:none;border-bottom:1px solid var(--line)}}
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
 const box=document.querySelector('.steps'), steps=[...box.querySelectorAll('.p')], msg=document.getElementById('runmsg');
 box.classList.add('on');box.scrollIntoView({behavior:'auto',block:'center'});
 if(!msg.dataset.orig)msg.dataset.orig=msg.innerHTML;
 document.getElementById('run').disabled=true;let i=0;
 (function tick(){steps.forEach((s,k)=>s.className='p'+(k<i?' done':(k===i?' run':'')));
  msg.textContent=i<steps.length?`점검 실행 중 · ${i+1} / ${steps.length}`:'점검 완료 · 결과 화면으로 이동';
  if(i<steps.length){i++;setTimeout(tick,800)}else setTimeout(showResult,700)})()});
document.getElementById('toStart').addEventListener('click',e=>{e.preventDefault();resultEl.hidden=true;startEl.hidden=false;
 document.getElementById('run').disabled=false;const box=document.querySelector('.steps');box.classList.remove('on');
 box.querySelectorAll('.p').forEach(s=>s.className='p');const msg=document.getElementById('runmsg');
 if(msg.dataset.orig)msg.innerHTML=msg.dataset.orig});
if(/^#(result|meeting|orig=)/.test(location.hash)){startEl.hidden=true;resultEl.hidden=false;
 if(location.hash==='#meeting')requestAnimationFrame(()=>requestAnimationFrame(()=>{show('meeting');window.scrollTo(0,0)}))}

// ---- 업무 개선안: 버튼으로 넘어가고, 단계별로 짚어 설명
const cmpEl=document.getElementById('compare');
const NOTE=JSON.parse(document.getElementById('saydata').textContent);
const nStep=NOTE.steps.length, noteEl=cmpEl.querySelector('.note');
let cur=-1, from=startEl;
function paint(){cmpEl.querySelectorAll('[data-step]').forEach(e=>e.classList.toggle('on',+e.dataset.step===cur));
 noteEl.innerHTML=cur<0?NOTE.intro:NOTE.steps[cur];
 document.getElementById('cmpPos').textContent=cur<0?'전체':`${cur+1} / ${nStep}`}
function openCompare(back){from=back;startEl.hidden=true;resultEl.hidden=true;cmpEl.hidden=false;cur=-1;paint();window.scrollTo(0,0)}
function closeCompare(){cmpEl.hidden=true;from.hidden=false;window.scrollTo(0,0)}
document.querySelectorAll('[data-open=compare]').forEach(b=>b.addEventListener('click',e=>{e.preventDefault();
 openCompare(resultEl.hidden?startEl:resultEl)}));
document.getElementById('cmpBack').addEventListener('click',closeCompare);
document.getElementById('cmpNext').addEventListener('click',()=>{cur=cur>=nStep-1?-1:cur+1;paint()});
document.getElementById('cmpPrev').addEventListener('click',()=>{cur=cur<=-1?nStep-1:cur-1;paint()});
document.getElementById('cmpAll').addEventListener('click',()=>{cur=-1;paint()});
cmpEl.querySelectorAll('[data-step]').forEach(e=>e.addEventListener('click',()=>{cur=+e.dataset.step;paint()}));
document.addEventListener('keydown',e=>{if(cmpEl.hidden||modal.classList.contains('on'))return;
 if(e.key==='ArrowRight'){e.preventDefault();document.getElementById('cmpNext').click()}
 else if(e.key==='ArrowLeft'){e.preventDefault();document.getElementById('cmpPrev').click()}
 else if(e.key==='Escape')closeCompare()});
if(location.hash==='#compare')openCompare(startEl);

// ---- 운영 규칙 · 안전장치
const guideEl=document.getElementById('guide');let gfrom=startEl;
function openGuide(){gfrom=[startEl,resultEl,cmpEl].find(e=>!e.hidden)||startEl;
 [startEl,resultEl,cmpEl].forEach(e=>e.hidden=true);guideEl.hidden=false;window.scrollTo(0,0)}
function closeGuide(){guideEl.hidden=true;gfrom.hidden=false;window.scrollTo(0,0)}
document.querySelectorAll('[data-open=guide]').forEach(b=>b.addEventListener('click',e=>{e.preventDefault();openGuide()}));
document.getElementById('guideBack').addEventListener('click',closeGuide);
document.addEventListener('keydown',e=>{if(!guideEl.hidden&&e.key==='Escape'&&!modal.classList.contains('on'))closeGuide()});
if(location.hash==='#guide')openGuide();
"""

# 업무 개선안 (AS-IS / TO-BE) — (단계, AS-IS 주체, AS-IS 내용, TO-BE 주체, TO-BE 내용, 단계 설명)
COMPARE = [
    ("접수", "담당자", "신청서 수합·정리", "담당자", "inbox 등록", "접수 방식은 동일."),
    ("양식 점검", "담당자", "양식·누락 육안 확인", "시스템", "표준양식 판정·자동 변환", "건별 확인 → 시스템 일괄 처리."),
    ("내용 점검", "담당자", "수량·금액·견적 수작업 대조", "시스템", "기준 일괄 점검·보완 문안 작성",
     "수작업 대조 → 검토 기준으로 자동 점검, 보완 요청 문안까지."),
    ("현장 확인", "담당자", "서류 점검 후 남는 시간에", "담당자", "점검 결과 기반 현장 확인·판단",
     "담당자 업무의 중심. 현장 확인과 판단에 집중."),
    ("회의자료", "담당자", "요약장 작성·자료 편철", "시스템", "표지·첨부·브리핑 자동 작성", "계획 대비 변경·계획 외 건 자동 표기."),
    ("투자회의", "회의체", "회의 진행", "회의체", "'확인 필요' 건 중심 심의", "진행·결정은 동일."),
]
KEY_STEP = "현장 확인"

# 추진 배경 — (구분, 한 줄 요약)
BACKGROUND = [
    ("업무량", "소수 담당자가 전사 투자 검토 → 서류 점검에 시간 대부분"),
    ("인력 구조", "현장 전문 숙련 인력 중심 → AI 도구 활용 낮음"),
    ("업무 방식", "수작업 대조·요약장 반복 → 현장 전문성이 서류에 묶임"),
]
DIRECTION = ["서류 업무는 시스템", "담당자는 현장 확인·판단", "기존 흐름 그대로 사용"]

def _compare_html(reviews, go) -> tuple[str, str]:
    """업무 개선안 화면 (HTML, 단계 설명 JSON)."""
    import json
    rows = [x for r in reviews for x in r["rows"]]
    checked = sum(1 for x in rows if x["판정"] != "해당 없음")
    fixes = sum(1 for x in rows if x["판정"] == "보완" and x.get("보완요청문구"))
    judge = sum(1 for x in rows if x["판정"] == "판단 필요")
    before = sum(1 for c in COMPARE if c[1] == "담당자")
    after = sum(1 for c in COMPARE if c[3] == "담당자")
    today = dt.date.today()

    def change(c):
        """AS-IS → TO-BE 변경 구분: (css 클래스, 표시 글자)."""
        if c[0] == KEY_STEP:
            return "key", "중점 전환"
        if c[3] == "시스템" and c[1] != "시스템":
            return "auto", "자동화"
        return "same", "변경 없음"

    def proc(side):
        cells = []
        for i, c in enumerate(COMPARE):
            who = c[1] if side == "AS-IS" else c[3]
            cls = "p" if side == "AS-IS" else "p " + change(c)[0]
            label = who + (" · 중점" if side == "TO-BE" and c[0] == KEY_STEP else "")
            cells.append(f'<div class="{cls}" data-step="{i}"><b>{escape(c[0])}</b><span>{escape(label)}</span></div>')
        return (f'<div class="frow"><div class="lab{" to" if side == "TO-BE" else ""}">{side}</div>'
                f'<div class="chev">{"".join(cells)}</div></div>')

    trs = "".join(
        f'<tr data-step="{i}" class="r-{change(c)[0]}"><td class="c stg">{escape(c[0])}</td>'
        f'<td class="as">{escape(c[2])}</td><td class="c as">{escape(c[1])}</td>'
        f'<td class="to">{escape(c[4])}</td><td class="c to">{escape(c[3])}</td>'
        f'<td class="c"><span class="tag {change(c)[0]}">{change(c)[1]}</span></td></tr>'
        for i, c in enumerate(COMPARE))
    n_auto = sum(1 for c in COMPARE if change(c)[0] == "auto")

    effect = [
        ("서류 점검", "건별 수작업 대조", "자동 일괄 점검", f"{len(reviews)}건 · {checked}항목"),
        ("보완 요청", "담당자 직접 작성", "송부용 문안 자동 작성", f"{fixes}건"),
        ("담당자 판단", "전 항목 직접 검토", "판단 필요 항목만 선별", f"{judge}건"),
        ("회의자료", "요약장 수작업", "표지·첨부·브리핑 자동", f"상정 {len(go)}건"),
        ("담당자 역할", "서류 점검 중심", "현장 확인 중심", f"수행 단계 {before} → {after}"),
    ]
    eff = "".join(f'<tr><td class="c stg">{escape(a)}</td><td class="as">{escape(b)}</td>'
                  f'<td class="tob">{escape(c)}</td><td class="c num">{escape(d)}</td></tr>'
                  for a, b, c, d in effect)
    bg_rows = "".join(f'<tr><td class="c stg">{escape(a)}</td><td class="prob">{escape(b)}</td></tr>'
                      for a, b in BACKGROUND)
    dir_items = "".join(f"<li>{escape(x)}</li>" for x in DIRECTION)
    steps_auto = " · ".join(c[0] for c in COMPARE if change(c)[0] == "auto")

    html = f"""<section id="compare" hidden>
 <div class="doc">
  <div class="dochead"><h1>투자검토 업무 개선안 <span class="ttag">AS-IS → TO-BE</span></h1>
    <div class="toprow"><span class="meta">{today:%Y.%m.%d} · 설비투자 투자검토·투자회의</span>
    <button class="cmpbtn" id="cmpBack" type="button">돌아가기</button></div></div>
  <h2>추진 배경</h2>
  <div class="bgsec">
    <table class="cmpt bgt"><colgroup><col style="width:84px"><col></colgroup>
     <thead><tr><th>구분</th><th>현황 → 문제점</th></tr></thead>
     <tbody>{bg_rows}</tbody></table>
    <div class="arrowr" aria-hidden="true">▶</div>
    <div class="dir"><div class="dh">개선 방향</div><ol>{dir_items}</ol></div>
  </div>

  <h2>핵심 변화</h2>
  <div class="keymsg">
    <div><small>담당자 수행 단계</small><b>{before} → {after}</b></div>
    <div><small>시스템 전환 ({n_auto}단계)</small><b>{escape(steps_auto)}</b></div>
    <div class="k"><small>담당자 중점 업무</small><b>{escape(KEY_STEP)}</b></div>
  </div>

  <h2>업무 절차</h2>
  <div class="proc">{proc("AS-IS")}{proc("TO-BE")}</div>
  <div class="legend"><span class="tag auto">자동화</span> <span class="tag key">중점 전환</span> <span class="tag same">변경 없음</span></div>

  <h2>단계별 비교</h2>
  <table class="cmpt"><colgroup><col class="stg" style="width:84px"><col><col class="who" style="width:64px"><col><col class="who" style="width:64px"><col class="chg" style="width:88px"></colgroup>
   <thead><tr><th>단계</th><th>AS-IS</th><th>수행</th><th class="to">TO-BE</th><th class="to">수행</th><th>변경</th></tr></thead>
   <tbody>{trs}</tbody></table>
  <div class="ctl"><button id="cmpPrev" type="button">이전</button><button id="cmpNext" type="button">다음</button>
    <button id="cmpAll" type="button">전체</button><span class="s" id="cmpPos">전체</span>
    <span class="s">← → 키로 이동</span></div>
  <div class="note"></div>

  <h2>기대 효과</h2>
  <table class="cmpt"><colgroup><col class="stg" style="width:100px"><col><col><col style="width:170px"></colgroup>
   <thead><tr><th>구분</th><th>AS-IS</th><th class="to">TO-BE</th><th>적용 실적</th></tr></thead><tbody>{eff}</tbody></table>
  <div class="foot">※ {today:%Y.%m.%d} 점검 기준 (익명화 샘플)</div>
 </div>
</section>"""
    note = {"intro": "<b>단계 설명</b> · 이전/다음으로 단계별 변경 확인",
            "steps": [f"<b>{i + 1}. {escape(c[0])}</b> · {escape(c[5])}" for i, c in enumerate(COMPARE)]}
    return html, json.dumps(note, ensure_ascii=False).replace("</", "<\\/")

def _guide_html() -> str:
    """운영 규칙 · 안전장치: CLAUDE.md · settings.json · 훅 파일에 적힌 주요 내용을 직접 읽어 보여 준다."""
    import ast
    import json
    import re

    def read(rel):
        f = ROOT / rel
        return f.read_text(encoding="utf-8") if f.exists() else ""

    def clean(s):
        return re.sub(r"\*\*|`", "", s).strip()

    md = read("CLAUDE.md")
    today = dt.date.today()

    # --- CLAUDE.md: 작업 규칙 (원문 항목)
    m = re.search(r"^## 작업 규칙[^\n]*\n(.*?)(?:\n---|\n## |\n# )", md, re.S | re.M)
    work = [clean(x[2:]) for x in (m.group(1).splitlines() if m else []) if x.startswith("- ")]
    work = [w for w in work if not w.startswith(("테스트:", "패키지:"))]
    work_html = "".join(f"<li>{escape(w)}</li>" for w in work)

    # --- CLAUDE.md: 판정 값 · 종합 상태
    verdict = re.search(r"^판정 값은 (.+?)\.", md, re.M)
    state = re.search(r"^종합 상태: (.+?)\.\s*$", md, re.M)
    verdict_txt = clean(verdict.group(1)) if verdict else ""
    state_txt = clean(state.group(1)) if state else ""

    # --- CLAUDE.md: 검토 규칙 R (ID, 항목, 조건, 미충족 시, 주체)
    rules = [tuple(clean(c) for c in r) for r in
             re.findall(r"^\| (R-\d+) \| ([^|]+)\| ([^|]+)\| ([^|]+)\| ([^|]+)\|", md, re.M)]
    groups = [("공통", "1"), ("유형", "20"), ("생산필수", "30"), ("노후교체", "31"), ("자동화", "32"), ("회의 상정", "4")]
    def grp(rid):
        n = rid[2:]
        for name, pre in sorted(groups, key=lambda g: -len(g[1])):
            if n.startswith(pre):
                return name
        return "기타"
    def who_tag(w):
        return "auto" if w == "스크립트" else ("key" if "AI" in w else "same")
    rgrid = ""
    for name, _ in groups:
        items = [r for r in rules if grp(r[0]) == name]
        if not items:
            continue
        chips = "".join(f'<span class="rc" title="{escape(r[2])}"><b>{escape(r[0])}</b> {escape(r[1])}'
                        f' <span class="tag {who_tag(r[4])}">{escape(r[4])}</span></span>' for r in items)
        rgrid += f'<div class="rg"><div class="rh">{escape(name)}</div><div class="rcs">{chips}</div></div>'
    rtable = "".join(f'<tr><td class="c stg">{escape(r[0])}</td><td>{escape(r[1])}</td><td>{escape(r[2])}</td>'
                     f'<td>{escape(r[3])}</td><td class="c">{escape(r[4])}</td></tr>' for r in rules)

    # --- CLAUDE.md: 유형별 필수 부가 항목
    req = [tuple(clean(c) for c in r) for r in
           re.findall(r"^\| (견적서 비교|내년도 부하·조업도|현 장비 재원 소개|생산성 향상 산출) \|[^|]+\|([^|]*)\|([^|]*)\|([^|]*)\|([^|]*)\|", md, re.M)]
    mark = lambda v: '<span class="tag auto">필수</span>' if v == "필수" else ""
    req_html = "".join(f'<tr><td class="stg">{escape(r[0])}</td><td class="c">{mark(r[1])}</td><td class="c">{mark(r[2])}</td>'
                       f'<td class="c">{mark(r[3])}</td><td class="c">{escape(r[4])}</td></tr>' for r in req)

    # --- CLAUDE.md: 경영계획 매칭 M
    mrows = [tuple(clean(c) for c in r) for r in re.findall(r"^\| (M-\d) \| ([^|]+)\| ([^|]+)\| ([^|]+)\|", md, re.M)]
    m_html = "".join(f'<tr><td class="c stg">{escape(r[0])}</td><td>{escape(r[1])}</td><td>{escape(r[2])}</td>'
                     f'<td>{escape(r[3])}</td></tr>' for r in mrows)

    # --- settings.json: 권한 deny
    try:
        deny = json.loads(read(".claude/settings.json")).get("permissions", {}).get("deny", [])
    except ValueError:
        deny = []
    dgroups = [("원본 편집 금지", [d for d in deny if d.startswith("Edit(")]),
               ("비밀 파일 읽기 금지", [d for d in deny if d.startswith("Read(")]),
               ("위험 명령 금지", [d for d in deny if d.startswith("Bash(")])]
    deny_html = "".join(f'<tr><td class="stg">{escape(n)}</td><td>' + " ".join(f"<code>{escape(x)}</code>" for x in xs)
                        + "</td></tr>" for n, xs in dgroups if xs)

    # --- guard.py: 설명(docstring) · 보호 대상 · 비밀 값 종류
    gsrc = read(".claude/hooks/guard.py")
    gdoc = ast.get_docstring(ast.parse(gsrc)) or "" if gsrc else ""
    def section(doc, start, stop):
        s = doc.find(start)
        e = doc.find(stop, s + 1) if s >= 0 else -1
        return doc[s + len(start):e if e > 0 else None] if s >= 0 else ""
    blocks = re.findall(r"^\s*\d\. (.+?)(?=^\s*\d\. |\Z)", section(gdoc, "막는 것", "허용"), re.S | re.M)
    blocks = [re.sub(r"\s+", " ", x).strip() for x in blocks]
    handle = [re.sub(r"\s+", " ", x).strip() for x in re.findall(r"^\s*- (.+)$", section(gdoc, "원본 쓰기는", "막는 것"), re.M)]
    allow = re.search(r"^허용: (.+)$", gdoc, re.M)
    def const(name):
        mm = re.search(rf"^{name} = (\[.*?\])", gsrc, re.M | re.S)
        try:
            return ast.literal_eval(mm.group(1)) if mm else []
        except (ValueError, SyntaxError):
            return []
    prot = [f"{d}/" for d in const("PROTECTED_DIRS")] + ["backup_*/"] + const("PROTECTED_FILES")
    secrets = list(dict.fromkeys(re.findall(r'^\s*\(r".*?", "(.+?)"\),?$', gsrc, re.M)))
    block_html = "".join(f'<tr><td class="c stg">{i}</td><td>{escape(x)}</td></tr>' for i, x in enumerate(blocks, 1))
    handle_html = "".join(f"<li>{escape(x)}</li>" for x in handle)

    # --- inbox_watch.py: 감시 대상 · 동작
    wsrc = read(".claude/hooks/inbox_watch.py")
    wdoc = ast.get_docstring(ast.parse(wsrc)) or "" if wsrc else ""
    wlines = [re.sub(r"\s+", " ", x).strip() for x in re.findall(r"^\s*(Subagent\w+\s+→.+)$", wdoc, re.M)]
    wm = re.search(r"^WATCH = (\[.*?\])", wsrc, re.M)
    fm = re.search(r"^FILES = (\[.*?\])", wsrc, re.M)
    watch = ([f"{d}/" for d in ast.literal_eval(wm.group(1))] if wm else []) + (ast.literal_eval(fm.group(1)) if fm else [])
    w_html = "".join(f"<li>{escape(x)}</li>" for x in wlines)

    # --- 에이전트 · 스킬 (.claude/agents/*.md, .claude/skills/*/SKILL.md)
    def front(src):
        mm = re.match(r"---\n(.*?)\n---\n(.*)", src, re.S)
        if not mm:
            return {}, src
        meta = dict(re.findall(r"^(\w+): (.+)$", mm.group(1), re.M))
        return meta, mm.group(2)

    def bullets(body, heading, n=3):
        mm = re.search(rf"^## {heading}[^\n]*\n(.*?)(?=^## |^# |\Z)", body, re.S | re.M)
        items = [clean(x[2:]) for x in (mm.group(1).splitlines() if mm else []) if x.startswith("- ")]
        return items[:n]

    def first_sentence(s):
        """설명 앞 두 문장 (너무 길면 200자)."""
        s = clean(s)
        parts = re.split(r"(?<=\.) ", s)
        out = " ".join(parts[:2])
        return out if len(out) <= 200 else out[:200] + "…"

    units = []
    for f in sorted((ROOT / ".claude" / "agents").glob("*.md")):
        meta, body = front(f.read_text(encoding="utf-8"))
        rules_ = bullets(body, "절대 규칙") or bullets(body, "원칙")
        units.append(("에이전트", meta.get("name", f.stem), first_sentence(meta.get("description", "")),
                      meta.get("tools", ""), rules_, f".claude/agents/{f.name}"))
    for f in sorted((ROOT / ".claude" / "skills").glob("*/SKILL.md")):
        meta, body = front(f.read_text(encoding="utf-8"))
        units.append(("스킬", meta.get("name", f.parent.name), first_sentence(meta.get("description", "")),
                      "(호출한 에이전트의 도구 사용)", bullets(body, "원칙"), f".claude/skills/{f.parent.name}/SKILL.md"))
        steps = [clean(x) for x in re.findall(r"^\d+\. (.+)$", body.split("## 절차", 1)[-1].split("\n## ", 1)[0], re.M)]
    unit_html = "".join(
        f'<tr><td class="c stg">{escape(k)}</td><td><b class="un">{escape(n)}</b><br><code>{escape(src)}</code></td>'
        f'<td>{escape(d)}</td><td class="c">{escape(tl)}</td>'
        f'<td><ul class="plain in">{"".join(f"<li>{escape(r)}</li>" for r in rs)}</ul></td></tr>'
        for k, n, d, tl, rs, src in units)
    skill_steps = "".join(f'<div class="p {"auto" if i % 2 else "same"}"><small>STEP {i}</small><b>{escape(re.split(r" — | → | \(|: |,", s)[0].strip()[:24])}</b></div>'
                          for i, s in enumerate(locals().get("steps", []), 1))

    codes = lambda xs: " ".join(f"<code>{escape(x)}</code>" for x in xs)
    return f"""<section id="guide" hidden>
 <div class="doc">
  <div class="dochead"><div><h1>운영 규칙 · 안전장치 <span class="ttag">CLAUDE.md · 훅</span></h1>
    <div class="meta">{today:%Y.%m.%d} · CLAUDE.md · settings.json · 훅 파일에 적힌 주요 내용 (파일에서 직접 읽어 표시)</div></div>
    <button class="cmpbtn" id="guideBack" type="button">돌아가기</button></div>

  <h2>구성 한눈에</h2>
  <div class="layers l4">
    <div><h3><small>① 규칙</small>CLAUDE.md</h3><div class="lx"><b>무엇을 해야 하나</b>검토 규칙 {len(rules)}개 · 매칭 {len(mrows)}개 · 작업 규칙</div></div>
    <div><h3><small>② 실행</small>에이전트 · 스킬</h3><div class="lx"><b>누가 하나</b>에이전트 {sum(1 for u in units if u[0] == "에이전트")}개 · 스킬 {sum(1 for u in units if u[0] == "스킬")}개</div></div>
    <div><h3><small>③ 권한</small>settings.json · deny</h3><div class="lx"><b>절대 못 하는 것</b>deny {len(deny)}개</div></div>
    <div><h3><small>④ 훅</small>guard · inbox_watch</h3><div class="lx"><b>실행 전후 자동 검사</b>막기 · 원본 변경 감지</div></div>
  </div>

  <h2>CLAUDE.md · 작업 규칙</h2>
  <ul class="plain">{work_html}</ul>

  <h2>CLAUDE.md · 검토 규칙 ({len(rules)}개)</h2>
  <div class="kv"><div><b>판정 값</b>{escape(verdict_txt)}</div><div><b>종합 상태</b>{escape(state_txt)}</div></div>
  <div class="rgrid">{rgrid}</div>
  <div class="legend"><span class="tag auto">스크립트</span> review.py 가 확정 <span class="tag key">AI</span> 점검 에이전트 판단
    <span class="tag same">사람</span> 검토자 · 항목에 마우스를 올리면 판정 기준</div>
  <details class="more"><summary>판정 기준 전체 보기 ({len(rules)}개)</summary>
   <table class="cmpt"><colgroup><col class="stg" style="width:70px"><col style="width:120px"><col><col style="width:22%"><col style="width:80px"></colgroup>
    <thead><tr><th>ID</th><th>항목</th><th>조건 (판정 기준)</th><th>미충족 시</th><th>주체</th></tr></thead><tbody>{rtable}</tbody></table></details>
  <table class="cmpt mt"><colgroup><col style="width:150px"><col><col><col><col style="width:120px"></colgroup>
   <thead><tr><th>유형별 필수 부가 항목</th><th>생산필수</th><th>노후교체</th><th>자동화</th><th>연결 규칙</th></tr></thead><tbody>{req_html}</tbody></table>

  <h2>CLAUDE.md · 경영계획 매칭 ({len(mrows)}개)</h2>
  <table class="cmpt"><colgroup><col class="stg" style="width:64px"><col><col><col></colgroup>
   <thead><tr><th>ID</th><th>상황</th><th>판정</th><th>표지 비고</th></tr></thead><tbody>{m_html}</tbody></table>

  <h2>에이전트 · 스킬</h2>
  <table class="cmpt"><colgroup><col class="stg" style="width:76px"><col style="width:230px"><col><col style="width:120px"><col style="width:34%"></colgroup>
   <thead><tr><th>종류</th><th>이름 · 파일</th><th>하는 일</th><th>도구</th><th>핵심 규칙</th></tr></thead><tbody>{unit_html}</tbody></table>
  <div class="legend">invest-standardize 스킬 절차</div>
  <div class="chev guidechev">{skill_steps}</div>

  <h2>settings.json · 권한 deny</h2>
  <table class="cmpt"><colgroup><col style="width:150px"><col></colgroup><tbody>{deny_html}</tbody></table>

  <h2>guard.py · 실행 직전 안전장치 (PreToolUse)</h2>
  <div class="kv"><div><b>보호 대상 (원본)</b>{codes(prot)}</div><div><b>허용</b>{escape(allow.group(1) if allow else "")}</div></div>
  <table class="cmpt"><colgroup><col class="stg" style="width:50px"><col></colgroup>
   <thead><tr><th>#</th><th>막는 것</th></tr></thead><tbody>{block_html}</tbody></table>
  <div class="kv"><div><b>원본 쓰기 처리</b><ul class="plain in">{handle_html}</ul></div>
    <div><b>비밀 값 감지</b>{codes(secrets)}</div></div>

  <h2>inbox_watch.py · 원본 변경 감지 (SubagentStart / Stop)</h2>
  <div class="kv"><div><b>감시 대상</b>{codes(watch)}</div><div><b>동작</b><ul class="plain in">{w_html}</ul></div></div>
  <div class="foot">※ 표시 내용은 화면을 만든 시점({today:%Y.%m.%d})의 CLAUDE.md · .claude/settings.json · .claude/hooks/*.py 에서 읽음</div>
 </div>
</section>"""

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
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>{CSS}</style></head><body>
<section id="start"><div class="doc">
  <div class="dochead"><div><h1>투자신청서 점검 <span class="ttag">투자검토 · 투자회의</span></h1>
    <div class="meta">신청서(.pptx) 첨부 → 조회 → 투자검토 결과·회의자료 · 원본은 수정하지 않음</div></div>
    <div class="hbtns"><button class="cmpbtn" type="button" data-open="compare">업무 개선안 보기</button>
    <button class="cmpbtn" type="button" data-open="guide">운영 규칙 · 안전장치</button></div></div>
  <h2>신청서 첨부</h2>
  <label class="drop"><input type="file" accept=".pptx,.ppt" multiple>
    <b>투자신청서 첨부</b>여기로 끌어다 놓거나 눌러서 파일 선택 (.pptx · 표준양식이 아니어도 됨)</label>
  <h2>접수 목록</h2>
  <table class="files"><thead><tr><th>미리보기</th><th>파일</th><th>크기</th><th>상태</th></tr></thead>
    <tbody>{file_rows}</tbody></table>
  <div class="go"><button id="run" type="button">조회</button>
    <span class="s" id="runmsg">첨부 <b id="cnt">{len(orig)}</b>건 · inbox 에 접수된 신청서</span></div>
  <div class="steps chev">
    <div class="p"><small>① 양식 확인</small><b>표준양식 변환</b><span>내용 그대로</span></div>
    <div class="p"><small>② 규칙 점검</small><b>R-100~R-322</b><span>스크립트 확정</span></div>
    <div class="p"><small>③ AI 판단</small><b>유형·필요성·산출</b><span>판단 필요 선별</span></div>
    <div class="p"><small>④ 회의자료</small><b>상정·계획 대조</b><span>표지 + 신청서</span></div>
  </div>
</div></section>
<div id="result" hidden>
<header><div class="doc">
  <div class="dochead"><div><h1>투자신청서 처리 결과 <span class="ttag">{dt.datetime.now():%Y.%m.%d %H:%M}</span></h1>
    <div class="meta">원본 신청서는 수정하지 않았습니다</div></div>
    <div class="hbtns"><button class="cmpbtn" id="toStart" type="button">← 첨부 화면</button>
    <button class="cmpbtn" type="button" data-open="compare">업무 개선안</button>
    <button class="cmpbtn" type="button" data-open="guide">운영 규칙</button></div></div>
  <h2>처리 현황</h2>
  <div class="chev rflow">
    <div class="p same"><small>① 접수</small><b>{len(inbox)}건</b><span>다른 양식 {converted}건 변환</span></div>
    <div class="p auto"><small>② 투자검토</small><b>{passed} · {fix} · {check}</b><span>통과 · 보완 · 확인</span></div>
    <div class="p key"><small>③ 회의 상정</small><b>{len(go)}건</b><span>보완 0 + 검토자 확인 (R-401)</span></div>
    <div class="p auto"><small>④ 회의자료</small><b>{escape(meeting_label)}</b><span>계획 대조 표지 + 신청서</span></div>
  </div>
  <h2>원본 신청서</h2>
  <div class="origs"><span>원본 신청서 보기</span>{chips}</div>
  <div class="notyet"></div>
</div></header>
<nav role="tablist">
  <button role="tab" data-t="review" aria-selected="true">① 투자검토 결과</button>
  <button role="tab" data-t="meeting" aria-selected="false">② 회의자료</button>
</nav>
<section id="review" class="panel on">{review_panel}</section>
<section id="meeting" class="panel">{meeting_panel}</section>
</div>
{compare_html}
{_guide_html()}
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
