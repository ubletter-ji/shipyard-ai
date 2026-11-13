# qna-helper

사내 기술 QnA 게시판을 Claude Code에서 바로 찾아 답하게 하는 플러그인입니다. 게시판에 있는 답만 근거로 쓰고, 질문 번호와 링크를 붙입니다. 게시판에 없으면 지어내지 않고 게시판에 질문을 올리라고 안내합니다.

## 들어 있는 것

| 파일 | 하는 일 |
|---|---|
| `server/qna-mcp.js` | MCP 서버(stdio). 외부 패키지 없이 Node.js 18 이상에서 동작 |
| `skills/ask/SKILL.md` | 답하는 규칙. 게시판만 근거, 검색어를 바꿔 세 번까지, 채택 답변에 숫자를 더하지 않음 |
| `.mcp.json` | 서버 실행 설정. 게시판 주소를 `QNA_API`로 넘김 |
| `.claude-plugin/plugin.json` | 플러그인 정보와 설정값(`qna_api`, 게시판 주소) |

도구는 두 개입니다.

- `search_questions(query)`: `GET /search?q=` 호출. 검색어는 핵심 용어 1~3개
- `get_question(id)`: `GET /questions/:id` 호출. 질문 원문과 답변(채택 여부 포함)

## 설치

```text
/plugin install qna-helper --marketplace ubletter-ji/shipyard-ai
```

Claude Code 2.1.275 이상에서 한 번에 됩니다. 그보다 낮은 버전에서는 마켓플레이스를 먼저 추가합니다.

```text
/plugin marketplace add ubletter-ji/shipyard-ai
/plugin install qna-helper@shipyard-ai
```

설치할 때 게시판 주소(예: `http://localhost:3001`)를 묻습니다. 나중에 바꾸려면 `/plugin configure qna-helper`를 씁니다.

필요한 것: Node.js 18 이상, 게시판 서버에 접속할 수 있는 네트워크.

## 써 보기

```text
선미 블록 개정 도면은 생산 도면에 언제 반영돼?
```

도구나 스킬 사용을 묻는 창이 뜨면 허락합니다.

## 개발할 때

플러그인 폴더 **밖에서** 엽니다. 이 폴더 안에서 열면 루트의 `.mcp.json`이 프로젝트 서버로 따로 읽혀 `qna`가 연결 실패로 보입니다.

```powershell
cd C:\work
claude --plugin-dir C:\work\qna-helper
```

서버만 따로 시험하기:

```powershell
'{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | & { $env:QNA_API = "http://localhost:3001"; node server\qna-mcp.js }
```

검사: `claude plugin validate .`

## 막히면

| 증상 | 확인할 것 |
|---|---|
| `/mcp`에서 `qna`가 실패 | `node`가 PATH에 있는지, 게시판 주소 |
| 답에 `게시판 조회 실패`가 나옴 | 게시판 주소(`/plugin configure qna-helper`)와 게시판 서버가 켜져 있는지 |
| 검색 결과가 늘 0건 | 검색어가 너무 김. 게시판은 모든 단어가 들어간 질문만 찾음 |
