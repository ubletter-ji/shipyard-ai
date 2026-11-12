#!/usr/bin/env node
// PreToolUse 가드 훅
//
// 파일 도구 (경로를 정확히 판단)
//   deny  Edit/Write/MultiEdit/NotebookEdit : ./data/**, ./.env, ./.env.*
//   deny  Read/Grep                         : ./.env, ./.env.*, **/*.pem, **/*.key
// 셸 도구 Bash/PowerShell (명령 문자열 휴리스틱)
//   deny  data/ 를 언급하면서 쓰기·삭제·이동으로 보이는 명령
//   deny  .env / .env.* / *.pem / *.key 파일을 언급하는 명령
//   ask   git push
//
// settings.json permissions 의 다음 규칙을 훅으로 옮긴 것:
//   ask:  Bash(git push *), PowerShell(git push *)
//   deny: Read(./.env), Read(./.env.*), Read(**/*.pem), Read(**/*.key),
//         Edit(./.env), Edit(./.env.*)  (+ 기존 data/ 보호)

const path = require('path');

function readStdin() {
  return new Promise((resolve) => {
    let buf = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', (c) => (buf += c));
    process.stdin.on('end', () => resolve(buf));
  });
}

function decide(decision, reason) {
  process.stdout.write(
    JSON.stringify({
      hookSpecificOutput: {
        hookEventName: 'PreToolUse',
        permissionDecision: decision,
        permissionDecisionReason: reason,
      },
    })
  );
  process.exit(0);
}

const norm = (p) => path.resolve(p).replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();

const DATA_REASON =
  'data/ 는 검사 원본 기록이라 수정할 수 없습니다 (집계 규칙 R9: 기록을 바꾸지 않음). ' +
  '판정 보정이 필요하면 summary.xlsx 에 별도 칸으로 표시하세요.';
const SECRET_REASON = '비밀 파일(.env, .env.*, *.pem, *.key)은 읽거나 수정할 수 없습니다.';

(async () => {
  let input;
  try {
    input = JSON.parse(await readStdin());
  } catch {
    process.exit(0);
  }

  const root = norm(process.env.CLAUDE_PROJECT_DIR || input.cwd || process.cwd());
  const base = input.cwd || root;
  const tool = input.tool_name || '';
  const ti = input.tool_input || {};

  const abs = (p) => norm(path.isAbsolute(p) ? p : path.join(base, p));
  const under = (t, dir) => t === dir || t.startsWith(dir + '/');
  const isData = (t) => under(t, root + '/data');
  // ./.env, ./.env.*  (프로젝트 루트만)
  const isRootEnv = (t) => {
    if (path.posix.dirname(t) !== root) return false;
    const b = path.posix.basename(t);
    return b === '.env' || b.startsWith('.env.');
  };
  // **/*.pem, **/*.key  (어디서든)
  const isKeyFile = (t) => /\.(pem|key)$/.test(t);

  // ── 파일 쓰기 도구 ────────────────────────────────
  if (['Write', 'Edit', 'MultiEdit', 'NotebookEdit'].includes(tool)) {
    const fp = ti.file_path || ti.notebook_path;
    if (!fp) process.exit(0);
    const t = abs(fp);
    if (isData(t)) decide('deny', `${DATA_REASON} (대상: ${fp})`);
    if (isRootEnv(t)) decide('deny', `${SECRET_REASON} (대상: ${fp})`);
    process.exit(0);
  }

  // ── 파일 읽기 도구 ────────────────────────────────
  if (tool === 'Read' || tool === 'Grep') {
    const fp = tool === 'Read' ? ti.file_path : ti.path;
    if (fp) {
      const t = abs(fp);
      if (isRootEnv(t) || isKeyFile(t)) decide('deny', `${SECRET_REASON} (대상: ${fp})`);
    }
    // Grep glob 필터로 비밀 파일만 골라 읽는 경우
    if (tool === 'Grep' && typeof ti.glob === 'string' && /(^|[\/{,])(\.env(\.[^\/,}]*)?|\*\.(pem|key))([},]|$)/i.test(ti.glob)) {
      decide('deny', `${SECRET_REASON} (glob: ${ti.glob})`);
    }
    process.exit(0);
  }

  // ── 셸 도구 ───────────────────────────────────────
  if (tool === 'Bash' || tool === 'PowerShell') {
    const cmd = String(ti.command || '');

    // .env / .env.* / *.pem / *.key 를 파일로 언급 (process.env, obj.keys 등은 제외)
    const B = `(^|[\\s"'=(/\\\\:])`;
    const E = `(?=$|[\\s"'),;|&<>])`;
    const envRe = new RegExp(`${B}\\.env(\\.[\\w.-]+)?${E}`, 'i');
    const keyRe = new RegExp(`${B}[\\w.*-]+\\.(pem|key)${E}`, 'i');
    if (envRe.test(cmd) || keyRe.test(cmd)) {
      decide('deny', `${SECRET_REASON} (셸 명령에서 비밀 파일 언급)`);
    }

    // data/ 쓰기
    const mentionsData = /(^|[\s"'=(/\\])data[\\/]/i.test(cmd);
    const writey =
      /\.save\s*\(|\bto_excel\b|\bwrite(_text|_bytes)?\s*\(|open\([^)]*['"][wa]b?\+?['"]|(^|[^2&])>{1,2}\s*["']?[^&\s]|\b(rm|mv|cp|del|rmdir|tee|truncate|touch)\b|sed\s+-i|\b(Remove-Item|Move-Item|Copy-Item|Rename-Item|Set-Content|Add-Content|Out-File|Clear-Content|New-Item)\b|git\s+(checkout|restore|rm|mv|reset|stash)\b/i;
    if (mentionsData && writey.test(cmd)) {
      decide('deny', `${DATA_REASON} (data/ 를 수정할 수 있는 셸 명령으로 판단됨)`);
    }

    // git push → 사용자 확인
    if (/(^|[;&|\s(])git(\s+-[^\s]+(\s+[^\s-][^\s]*)?)*\s+push\b/i.test(cmd)) {
      decide('ask', 'git push 는 원격 저장소에 게시되므로 확인이 필요합니다.');
    }
    process.exit(0);
  }

  process.exit(0);
})();
