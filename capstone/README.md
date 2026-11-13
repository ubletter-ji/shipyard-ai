# 투자신청서 점검·회의자료 AI 에이전트 (캡스톤)

조선소 설비투자 업무의 **투자검토 → 투자회의** 단계를 Claude Code 에이전트 2개로 자동화한다.
업무 규칙(검토 규칙·계획 매칭·표지 형식·관리 방법)은 `CLAUDE.md` 「업무 규칙」에 정리되어 있다. 빈 신청서 양식은 `표준양식_수정.pptx`, 사람이 보는 규칙 표는 `구현시나리오_보완.xlsx` 이고, 규칙을 바꾸면 CLAUDE.md §8 순서대로 함께 고친다.

```
inbox/*.pptx ─▶ [invest-standardize 스킬] ─▶ [invest-reviewer] ─▶ out/검토결과.html  (잘못된 점 설명)
 투자신청서       다른 양식이면 표준양식으로     규칙 R-100~R-322      out/검토결과.xlsx  (검토자 O 승인)
                 (내용 그대로, 자동 검증)        + AI 판단 3종                   │ 사람 확인
                                                                                ▼
                                    out/투자회의자료_날짜.html ◀─ [invest-meeting] 경영계획 대조 M-0~M-9
                                    (표지 + 상정 신청서 원본)       상정 건만 받음 (R-401)
```

## 1. 준비 (처음 한 번)

| 필요한 것 | 확인 |
|---|---|
| Python 3.10 이상 | `python --version` |
| Claude Code | 이 폴더에서 `claude` 실행 |
| (선택) PowerPoint | 결과 pptx 를 열어 보는 용도 — 없어도 동작 |

```bash
pip install -r requirements.txt
python scripts/selfcheck.py --demo
```
`selfcheck` 는 환경·테스트(정답셋 T1~T6 포함)·훅 설정을 점검하고, `--demo` 는 임시 폴더의 사본으로
'다른 양식 변환 → 규칙 점검 → HTML 화면'까지 돌려 본다. 원본은 바뀌지 않는다(마지막 줄에서 확인).
모두 `통과` 면 준비 끝. 샘플이 없으면 `python scripts/make_samples.py` (기존 파일은 덮어쓰지 않음).

> Windows 콘솔에서 한글이 깨지면 `set PYTHONIOENCODING=utf-8` (PowerShell: `$env:PYTHONIOENCODING="utf-8"`).

## 2. 사용 (Claude Code 안에서)

1. **신청서 넣기** — `inbox/` 에 투자신청서 .pptx 를 둔다. 시험할 때는 샘플 한두 개를 PowerPoint 로 열어
   일부러 틀리게 고치고 **저장 후 닫는다**. (다른 양식 시험: `samples/다른양식/` 파일을 inbox 로 복사)
2. **점검** — `invest-reviewer 로 inbox 신청서 점검해줘`
   - 결과 화면: `out/검토결과.html` (브라우저로 열기) — 신청서별 잘못된 점, 신청자에게 보낼 문구(복사 버튼)
3. **사람 확인** — `out/검토결과.xlsx` ① 상세 검토에서 '판단 필요' 행을 읽고 동의하면 검토자 확인 = `O`, 저장
4. **회의자료** — `invest-meeting 으로 회의자료 만들어줘`
   - `out/투자회의자료_YYYYMMDD.html` (브라우저로 열기 · 출력용 .pptx 함께) (표지: 투자 심의 | 경영계획 | 실행금액 | 차이금액 | 비고 + 신청서 원본)
   - `out/회의_요약.xlsx`, `out/회의_브리핑.md`
5. 점검 뒤에 신청서를 고쳤다면 4번에서 '재점검 필요'로 보류된다 → 2번부터 다시.
6. **한 화면으로 보기**: `out/투자시스템_결과.html` — 처음엔 **첨부 화면**(inbox 에 접수된 신청서가 첨부된 상태, 파일 추가 가능) → **조회**를 누르면 접수→검토→상정→회의자료 흐름과 ① 투자검토 결과 / ② 회의자료 탭 (점검·회의자료를 만들 때마다 자동 갱신).
   신청서 이름이나 카드의 **원본 보기**를 누르면 원본 슬라이드가 그대로 뜬다 — 다른 양식이면 제출 원본과 표준양식 변환본을 나란히 (PowerPoint 가 있으면 슬라이드 이미지, 없으면 글자). `#orig=신청서ID` 로 바로 열 수도 있다.

샘플을 처음 상태로 되돌리기: 사용자가 직접 `! python scripts/make_samples.py --force`
(Claude 가 실행하면 안전장치가 막는다 — 사용자가 넣은 오류를 지우기 때문)

## 3. 시연 (평가용 순서)

| 단계 | 하는 일 | 기대 결과 |
|---|---|---|
| ① | `python scripts/selfcheck.py --demo` | 모든 항목 통과, 원본 무변경 |
| ② | 샘플 02 의 투자수량을 필요수량 − 현수량보다 크게, 샘플 04 의 선정 업체를 신생 업체(F사)로 바꿔 저장 | — |
| ③ | `invest-reviewer 로 점검해줘` | HTML 에 02 = R-102 보완, 04 = R-107 보완(신생 업체), 05 = 보완 7건 |
| ④ | Claude 에게 `inbox 의 02 고쳐줘` 라고 시켜 봄 | 안전장치가 승인을 묻거나(대화 세션) 막음(에이전트) |
| ⑤ | 엑셀에서 판단 필요 행에 O → `invest-meeting 으로 회의자료` | 통과 건만 표지·첨부 (보류 건은 회의자료에 없음 — 에이전트가 실무자에게 따로 보고) |

자세한 시나리오·사람이 확인한 지점·회고: [`docs/시연_회고.md`](docs/시연_회고.md)
에이전트·스킬·훅·데이터 명세: [`docs/명세.md`](docs/명세.md)

## 4. 안전장치 요약

| 장치 | 위치 | 하는 일 |
|---|---|---|
| 권한 deny | `.claude/settings.json` | inbox·data·samples·백업·기준 문서 편집, `.env` 읽기, force push 금지 |
| PreToolUse 훅 | `.claude/hooks/guard.py` | 원본 쓰기 → 에이전트는 차단 / 대화 세션은 사용자 승인. 위험 명령·비밀 값 저장 → 차단 |
| Subagent 훅 | `.claude/hooks/inbox_watch.py` | 에이전트 실행 전후 원본 지문 비교, 바뀌면 경고 |
| 스크립트 검증 | `review.py judge`, `standardize.py` | 규칙 판정은 에이전트가 못 바꿈 / 변환은 추가·누락 0건이어야 통과 |
| 원본 지문 | `meeting.py` | 점검 뒤 바뀐 신청서는 회의자료에 넣지 않음 |
| 저장소 위생 | `.gitignore`, `tests/test_safety.py` | 산출물·백업·.env 제외, 비밀 값 패턴 검사 |

## 5. 폴더

```
.claude/agents/      invest-reviewer.md, invest-meeting.md   에이전트 (역할·도구 권한)
.claude/skills/      invest-standardize/SKILL.md              다른 양식 → 표준양식 변환 절차
.claude/hooks/       guard.py, inbox_watch.py                 안전장치 훅
.claude/settings.json                                         훅 연결 + 권한 deny
config/기준값.json   엑셀의 [확정 필요] 수치 (임시값 표시)
scripts/             extract · review · standardize · match · meeting · report_html · make_samples · migrate_quantity · selfcheck
tests/               test_rules(규칙·정답셋) · test_safety(훅·비밀값) · test_migrate(양식 변환)
data/                경영계획.xlsx, 집행이력.xlsx (실데이터로 교체)
inbox/               점검할 투자신청서
samples/다른양식/     변환 시험용
docs/                명세.md, 시연_회고.md
out/                 산출물 (저장소 제외)
```

## 6. 문제 해결

| 증상 | 해결 |
|---|---|
| `…이(가) 열려 있습니다` | PowerPoint/Excel 에서 파일을 닫고 다시 |
| `[안전장치] …` 메시지로 막힘 | 원본·위험 작업이라 막힌 것. 정말 필요하면 사용자가 터미널에서 직접 실행하거나 승인 |
| 회의자료에 신청서가 없음 | `out/검토결과.html` 보류 사유 확인 — 보완 남음 / 검토자 확인 O 없음 / 점검 후 원본 변경 |
| 다른 양식이 'R-100 보완' | 변환 검증 실패(원문 추가·누락). `out/표준화/<이름>.json` 의 누락·추가 목록 확인 |

## 7. 웹 공개 (Vercel)

Vercel 은 이 폴더(`capstone/`)를 루트로 연결하되 **정적 결과 화면만** 올린다 — 파이썬 스크립트는 서버 앱이 아니므로 올리지 않는다.
- `vercel.json`: 빌드 없음, `web/` 폴더만 배포 · `.vercelignore`: `web/`·`vercel.json` 외 제외
- 화면 갱신: 점검·회의자료를 만든 뒤 `python scripts/build_web.py` → `web/index.html` 을 커밋·푸시
  (PC 원본 파일 링크는 지우고 '익명화 샘플 시연' 안내를 붙임. 실제 신청서 결과는 올리지 말 것)
