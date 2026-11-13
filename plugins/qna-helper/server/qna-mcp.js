#!/usr/bin/env node
// 사내 QnA 게시판 MCP 서버 (stdio, 줄 단위 JSON-RPC 2.0, 외부 패키지 없음)
// 게시판 주소: 환경 변수 QNA_API (예: http://localhost:3001)
// 주의: stdout은 프로토콜 전용이다. 로그는 stderr로만 쓴다.
'use strict';

const readline = require('node:readline');

const SERVER_INFO = { name: 'qna', version: '1.0.0' };
const SUPPORTED_PROTOCOLS = ['2025-06-18', '2025-03-26', '2024-11-05'];
const BASE = (process.env.QNA_API || '').trim().replace(/\/+$/, '');
const TIMEOUT_MS = 10000;

const TOOLS = [
  {
    name: 'search_questions',
    description:
      '사내 기술 QnA 게시판에서 질문을 검색한다(제목, 본문, 호선 번호). ' +
      '검색어는 핵심 용어 1~3개로 짧게 쓴다(예: "선미 도면 반영"). 게시판은 모든 단어가 들어간 질문만 찾으므로 문장을 통째로 넣으면 0건이 나온다. ' +
      '결과가 없으면 검색어를 바꿔 다시 찾는다. ' +
      '결과는 제목 목록일 뿐이다. 답하기 전에 get_question으로 원문과 답변을 반드시 읽는다.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: '핵심 용어 1~3개, 공백으로 구분 (1~100자)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'get_question',
    description:
      '질문 번호로 게시판 질문의 원문과 답변 전체(채택 여부 포함)를 읽는다. ' +
      '답하기 전에 get_question으로 원문을 읽는다. 답할 때는 채택 답변을 근거로 하고, 채택되지 않은 답변은 그렇다고 밝힌다. ' +
      '결과의 링크를 근거로 붙인다.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'integer', minimum: 1, description: '질문 번호 (search_questions 결과의 #번호)' },
      },
      required: ['id'],
    },
  },
];

// ---------- 게시판 호출 ----------

class BoardError extends Error {}

async function callBoard(path) {
  if (!BASE) throw new BoardError('QNA_API 환경 변수가 비어 있습니다. 게시판 주소(예: http://localhost:3001)를 설정하세요.');
  let res;
  try {
    res = await fetch(BASE + path, {
      headers: { accept: 'application/json' },
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
  } catch (err) {
    throw new BoardError(`게시판에 연결하지 못했습니다 (${BASE}): ${err.cause?.code || err.name || err.message}`);
  }
  let json;
  try {
    json = await res.json();
  } catch {
    throw new BoardError(`게시판 응답을 읽지 못했습니다 (HTTP ${res.status})`);
  }
  if (!res.ok || json?.error) {
    const e = json?.error || {};
    const err = new BoardError(`${e.code || 'HTTP_' + res.status}: ${e.message || '알 수 없는 오류'}`);
    err.status = res.status;
    throw err;
  }
  return json.data;
}

const link = (id) => `${BASE}/q/${id}`;
const day = (iso) => (iso ? String(iso).slice(0, 10) : '-');

function formatSummary(q) {
  const state = q.hasAccepted ? '채택 답변 있음' : q.answerCount > 0 ? `답변 ${q.answerCount}개(채택 없음)` : '답변 없음';
  const meta = [q.dept, q.hullNo, q.stage].filter(Boolean).join(' / ');
  return `#${q.id} ${q.title}\n  ${meta} · ${state} · ${day(q.createdAt)}\n  링크: ${link(q.id)}`;
}

async function searchQuestions(args) {
  const query = typeof args?.query === 'string' ? args.query.trim() : '';
  if (!query) return errorResult('query가 비어 있습니다. 핵심 용어 1~3개를 넣으세요.');
  if (query.length > 100) return errorResult('query는 100자 이하입니다. 핵심 용어 1~3개로 줄이세요.');

  const data = await callBoard(`/search?q=${encodeURIComponent(query)}`);
  const items = data?.items || [];
  if (items.length === 0) {
    return textResult(
      `검색 결과 없음: "${query}" (0건)\n` +
        '게시판에 이 검색어가 모두 들어간 질문이 없습니다. 더 짧거나 다른 용어로 다시 찾으세요. ' +
        '여러 번 찾아도 없으면 게시판에 없는 내용입니다. 지어내지 말고 게시판에 질문을 올리라고 안내하세요.'
    );
  }
  return textResult(
    `검색 "${query}": ${items.length}건 (채택 질문 먼저)\n\n` +
      items.map(formatSummary).join('\n\n') +
      '\n\n답하기 전에 get_question으로 원문과 답변을 읽으세요.'
  );
}

async function getQuestion(args) {
  const id = Number(args?.id);
  if (!Number.isInteger(id) || id < 1) return errorResult('id는 1 이상의 정수인 질문 번호입니다.');

  let q;
  try {
    q = await callBoard(`/questions/${id}`);
  } catch (err) {
    if (err.status === 404) return errorResult(`질문 #${id}이(가) 게시판에 없습니다.`);
    throw err;
  }

  const answers = q.answers || [];
  const lines = [
    `#${q.id} ${q.title}`,
    `링크: ${link(q.id)}`,
    `부서: ${q.dept || '-'} · 호선: ${q.hullNo || '-'} · 공정: ${q.stage || '-'} · 작성: ${q.authorName || '-'} (${day(q.createdAt)})`,
    '',
    '[질문 본문]',
    q.body || '',
    '',
  ];
  if (answers.length === 0) {
    lines.push('[답변] 아직 답변이 없습니다. 답을 지어내지 말고 이 질문에 답변이 없다고 알리세요.');
  } else {
    if (!q.acceptedAnswerId) lines.push('※ 채택된 답변이 없습니다. 아래 답변은 모두 채택되지 않은 답변입니다.', '');
    for (const a of answers) {
      lines.push(`[답변 #${a.id}${a.accepted ? ' · 채택' : ' · 미채택'}] ${a.authorName || '-'} (${day(a.createdAt)})`, a.body || '', '');
    }
  }
  return textResult(lines.join('\n').trimEnd());
}

const HANDLERS = { search_questions: searchQuestions, get_question: getQuestion };

const textResult = (text) => ({ content: [{ type: 'text', text }] });
const errorResult = (text) => ({ content: [{ type: 'text', text }], isError: true });

// ---------- JSON-RPC ----------

function send(msg) {
  process.stdout.write(JSON.stringify(msg) + '\n');
}
const reply = (id, result) => send({ jsonrpc: '2.0', id, result });
const replyError = (id, code, message) => send({ jsonrpc: '2.0', id, error: { code, message } });

async function handle(msg) {
  const { id, method, params } = msg;
  const isRequest = id !== undefined && id !== null;

  switch (method) {
    case 'initialize':
      return reply(id, {
        protocolVersion: SUPPORTED_PROTOCOLS.includes(params?.protocolVersion)
          ? params.protocolVersion
          : SUPPORTED_PROTOCOLS[0],
        capabilities: { tools: {} },
        serverInfo: SERVER_INFO,
        instructions: '사내 기술 QnA 게시판 검색. search_questions로 찾고, 답하기 전에 get_question으로 원문을 읽는다. 게시판에 없으면 지어내지 않는다.',
      });
    case 'ping':
      return reply(id, {});
    case 'tools/list':
      return reply(id, { tools: TOOLS });
    case 'tools/call': {
      const handler = HANDLERS[params?.name];
      if (!handler) return replyError(id, -32602, `알 수 없는 도구: ${params?.name}`);
      try {
        return reply(id, await handler(params.arguments || {}));
      } catch (err) {
        // 게시판 오류는 도구 결과(isError)로 돌려 모델이 읽을 수 있게 한다
        return reply(id, errorResult(`게시판 조회 실패: ${err.message}`));
      }
    }
    default:
      if (method?.startsWith('notifications/')) return; // initialized, cancelled 등은 응답 없음
      if (isRequest) return replyError(id, -32601, `지원하지 않는 메서드: ${method}`);
  }
}

const rl = readline.createInterface({ input: process.stdin, crlfDelay: Infinity });
rl.on('line', (line) => {
  if (!line.trim()) return;
  let msg;
  try {
    msg = JSON.parse(line);
  } catch {
    return send({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'JSON 파싱 오류' } });
  }
  const p = handle(msg).catch((err) => {
    process.stderr.write(`[qna-mcp] ${err.stack || err}\n`);
    if (msg.id !== undefined && msg.id !== null) replyError(msg.id, -32603, '서버 내부 오류');
  });
  pending.add(p);
  p.finally(() => pending.delete(p));
});
// stdin이 닫혀도 처리 중인 요청의 응답은 마저 보낸 뒤 끝낸다
const pending = new Set();
rl.on('close', () => Promise.allSettled([...pending]).then(() => process.exit(0)));

process.stderr.write(`[qna-mcp] 시작. QNA_API=${BASE || '(비어 있음)'}\n`);
