# qc-kit

조선소 주간 품질검사(QC) 기록을 집계하고, 감사하고, 회의 문장으로 만드는 플러그인입니다.

## 들어 있는 것

| 종류 | 이름 | 하는 일 |
|---|---|---|
| 스킬 | `/qc-kit:weekly-ship [기간]` | qc-aggregator → qc-auditor → qc-writer 순서로 집계·감사·회의 문장. 반려는 최대 2번 |
| 에이전트 | qc-aggregator, qc-auditor, qc-writer | `/weekly-ship`이 부르는 세 역할. 산출물은 `out/` 아래 |
| 훅 | `hooks/hooks.json` → `protect-data.js` | `data/` 원본 수정, `.env`·`*.pem`·`*.key` 읽기·수정 차단, `git push` 확인 |
| 훅 스크립트 | `hooks/scripts/guard_out_only.py` | 세 에이전트의 쓰기를 `out/` 아래로 제한 (에이전트 frontmatter에서 연결) |

## 설치

```text
/plugin marketplace add ubletter-ji/shipyard-ai
/plugin install qc-kit@shipyard-ai
/reload-plugins
```

## 필요한 것

- Node.js (`protect-data.js`), Python 3 (`guard_out_only.py`)

## 검사

```powershell
claude plugin validate ./plugins/qc-kit
```

## 바꿀 때

- 원본: 스킬·에이전트·`guard_out_only.py`는 `labs/현업미션/data-셋째주/.claude`, `protect-data.js`는 `weekly-qc/.claude/hooks`.
- 고친 뒤 `plugin.json`의 `version`을 올립니다.
