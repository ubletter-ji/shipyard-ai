"""PreToolUse 안전장치 훅 — 보호 경로 쓰기·위험 명령·비밀 값 저장을 막는다.

Claude Code 가 도구 실행 직전에 stdin 으로 {"tool_name", "tool_input", ...} JSON 을 넘긴다.
막을 때는 이유를 stderr 에 쓰고 exit 2 (Claude 에게 이유가 전달되고 도구는 실행되지 않음).

원본 쓰기는 누가 하느냐에 따라 다르게 처리한다 (사람이 확인하는 지점)
  - 에이전트(invest-reviewer / invest-meeting 등 서브에이전트) → 무조건 차단 (exit 2)
  - 메인 세션(사용자와 대화 중인 Claude)            → 사용자 승인 요청 (permissionDecision: ask)
    사용자가 "이 신청서 고쳐줘"처럼 직접 요청한 수정만 승인하면 된다.

막는 것
  1. 원본 쓰기: inbox/ (사용자가 오류를 넣은 신청서), data/ (경영계획·집행이력), backup_*/,
     samples/, 기준 문서 2개(구현시나리오_보완.xlsx, 표준양식_수정.pptx)
     — Write/Edit 도구, 셸의 rm·mv·cp(대상)·리다이렉트·Set-Content 등, 파이썬 .save() 로 inbox 저장
  2. 샘플 초기화 `make_samples.py --force` (사용자 inbox 를 덮어씀) → 사용자가 `!` 로 직접 실행
  3. 위험 명령: 프로젝트·드라이브 통째 삭제, git push --force, git reset --hard, git clean -f
  4. 비밀 값: API 키·토큰·비밀번호 형태의 문자열을 파일에 쓰기
허용: out/ (산출물), 스크립트·문서 수정, 원본 읽기
"""
from __future__ import annotations

import json
import os
import re
import shlex
import sys
from pathlib import Path

PROJECT = Path(os.environ.get("CLAUDE_PROJECT_DIR") or Path(__file__).resolve().parents[2]).resolve()

CWD = PROJECT     # main() 에서 훅 입력의 cwd 로 바꾼다
PROTECTED_DIRS = ["inbox", "data", "samples"]
PROTECTED_GLOB_DIRS = [re.compile(r"^backup_[^/]*$")]
PROTECTED_FILES = ["구현시나리오_보완.xlsx", "표준양식_수정.pptx"]

SECRET_PATTERNS = [
    (r"sk-ant-[A-Za-z0-9_\-]{20,}", "Anthropic API 키"),
    (r"sk-[A-Za-z0-9]{32,}", "API 키"),
    (r"AKIA[0-9A-Z]{16}", "AWS 액세스 키"),
    (r"gh[pousr]_[A-Za-z0-9]{30,}", "GitHub 토큰"),
    (r"xox[baprs]-[A-Za-z0-9\-]{10,}", "Slack 토큰"),
    (r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----", "개인 키"),
    (r"(?i)\b(password|passwd|pwd|secret|api[_-]?key|token)\s*[:=]\s*['\"][^'\"\s]{6,}['\"]", "비밀번호·키 값"),
]


def rel(p: str) -> str | None:
    """프로젝트 기준 상대경로 (밖이면 None)."""
    if not p:
        return None
    p = p.strip().strip("'\"")
    if re.match(r"^/[a-zA-Z]/", p):            # Git Bash: /c/work/... → C:/work/...
        p = p[1] + ":" + p[2:]
    path = Path(p)
    if not path.is_absolute():
        path = CWD / path            # 상대경로는 명령이 실행되는 폴더 기준
    try:
        r = path.resolve().relative_to(PROJECT)
    except (ValueError, OSError):
        return None
    return r.as_posix()


def protected(p: str) -> str | None:
    r = rel(p)
    if r is None or r in ("", "."):
        return None
    first = r.split("/")[0]
    if first in PROTECTED_DIRS or any(g.match(first) for g in PROTECTED_GLOB_DIRS):
        return first + "/"
    if r in PROTECTED_FILES:
        return r
    return None


def mentions_protected(cmd: str) -> str | None:
    """명령 속 경로 조각 중 실제로 원본을 가리키는 것이 있는지 (작업 폴더 기준으로 풀어서 판단).
    out/demo/inbox 같은 사본, 이름만 같은 파일은 원본이 아니다."""
    names = PROTECTED_DIRS + PROTECTED_FILES + ["backup_"]
    for tok in re.split(r"[\s'\"=(),;|&<>]+", cmd):
        tok = tok.strip()
        if not tok or not any(n in tok for n in names):
            continue
        tok = re.sub(r"[*?].*$", "", tok)        # 글롭은 앞부분 폴더로 판단 (inbox/*.pptx → inbox/)
        hit = protected(tok)
        if hit:
            return hit
    return None


IS_AGENT = False   # main() 에서 설정: 서브에이전트의 도구 호출이면 True


def guard_original(msg: str):
    """원본 쓰기: 에이전트면 차단, 메인 세션이면 사용자에게 승인을 묻는다."""
    if IS_AGENT:
        block(msg + " (에이전트는 원본을 수정할 수 없습니다)")
    out = {"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": "ask",
                                  "permissionDecisionReason": f"[안전장치·사람 확인] {msg} 사용자가 요청한 수정이 맞으면 승인하세요."}}
    sys.stdout.buffer.write(json.dumps(out, ensure_ascii=False).encode("utf-8"))
    sys.stdout.flush()
    sys.exit(0)


def block(msg: str):
    # Windows 파이썬은 기본이 cp949 라 한글이 깨진다 → 항상 UTF-8 바이트로 쓴다
    sys.stderr.buffer.write(f"[안전장치] {msg}\n".encode("utf-8"))
    sys.stderr.flush()
    sys.exit(2)


def check_secret(text: str, where: str):
    for pat, name in SECRET_PATTERNS:
        if re.search(pat, text or ""):
            block(f"{where} 에 {name} 로 보이는 값이 있습니다. 비밀 값은 저장소 파일에 쓰지 말고 환경변수(.env, 저장소 제외)로 두세요.")


# ---------------- 파일 도구
def check_file_tool(tool: str, ti: dict):
    path = ti.get("file_path") or ti.get("notebook_path") or ""
    hit = protected(path)
    if hit:
        guard_original(f"{tool} → {path}: '{hit}' 은 원본(읽기 전용)입니다. 산출물은 out/ 에 쓰세요.")
    text = ti.get("content") or ti.get("new_string") or ti.get("new_source") or ""
    for e in ti.get("edits") or []:
        text += "\n" + (e.get("new_string") or "")
    check_secret(text, path)


# ---------------- 셸 명령
WRITE_ANY_ARG = re.compile(r"^(rm|rmdir|del|erase|rd|unlink|shred|truncate|touch|tee|chmod|attrib|"
                           r"remove-item|set-content|add-content|clear-content|out-file|new-item|rename-item|ren)$", re.I)
WRITE_LAST_ARG = re.compile(r"^(cp|copy|xcopy|robocopy|mv|move|copy-item|move-item|install|rsync)$", re.I)


def _split(cmd: str) -> list[list[str]]:
    parts = re.split(r"\s*(?:&&|\|\||;|\||\n)\s*", cmd)
    out = []
    for p in parts:
        try:
            toks = shlex.split(p, posix=True)
        except ValueError:
            toks = p.split()
        if toks:
            out.append(toks)
    return out


def check_shell(cmd: str):
    c = cmd.replace("\\", "/")

    # 3. 위험 명령
    if re.search(r"git\s+push\b[^\n]*(--force\b|-f\b|--force-with-lease)", c):
        block("git push --force 는 금지입니다.")
    if re.search(r"git\s+reset\s+--hard", c) or re.search(r"git\s+clean\s+-[a-z]*f", c):
        block("git reset --hard / git clean -f 는 작업을 되돌릴 수 없어 금지입니다. 사용자가 직접 실행하세요.")
    if re.search(r"(rm\s+-[a-z]*r[a-z]*f?|rm\s+-[a-z]*f[a-z]*r|remove-item[^\n]*-recurse|rd\s+/s|rmdir\s+/s)"
                 r"\s+[\"']?(/|~|\.|\.\.|\*|[a-z]:/?|\$claude_project_dir)[\"']?(\s|$)", c, re.I):
        block("프로젝트·드라이브 통째 삭제 명령은 금지입니다.")

    # 재귀 삭제의 대상이 루트·홈·프로젝트 자체이면 (옵션 위치와 무관하게)
    for toks in _split(c):
        name = Path(toks[0]).name.lower()
        flags = " ".join(t.lower() for t in toks[1:] if t.startswith("-") or t.startswith("/"))
        recursive = (name == "rm" and re.search(r"-[a-z]*r", flags)) or \
                    (name in ("remove-item", "ri", "del") and "-recurse" in flags) or \
                    (name in ("rd", "rmdir") and "/s" in flags)
        if not recursive:
            continue
        for t in toks[1:]:
            if t.startswith("-") or (t.startswith("/") and len(t) <= 3 and name in ("rd", "rmdir")):
                continue
            tt = t.strip("'\"").rstrip("/")
            if tt in ("", "~", ".", "..", "*", "$claude_project_dir") or re.fullmatch(r"[a-zA-Z]:", tt) \
                    or re.fullmatch(r"/[a-zA-Z]?", tt) or rel(t) == ".":
                block(f"'{' '.join(toks[:4])}…': 루트·홈·프로젝트 전체를 지우는 명령은 금지입니다.")

    # 2. 샘플 초기화
    if re.search(r"make_samples\.py[^\n]*--force", c):
        block("make_samples.py --force 는 사용자가 오류를 넣은 inbox 를 덮어씁니다. "
              "초기화가 필요하면 사용자가 직접 `! python scripts/make_samples.py --force` 로 실행하세요.")

    hit = mentions_protected(c)
    if not hit:
        return
    # 1-0. 원본을 고치는 전용 도구(마이그레이션 스크립트)도 사람 확인을 거친다
    if re.search(r"scripts/migrate_[a-z_]+\.py", c):
        guard_original(f"원본 변환 스크립트가 '{hit}' 를 수정합니다 (실행 전 backup_20261112/ 에 백업).")
    # 1-a. 파이썬 등으로 보호 경로에 저장
    if re.search(r"\.save(as)?\(|write_bytes|write_text|open\([^)]*['\"][wa]b?['\"]|shutil\.(copy|move|rmtree)|os\.(remove|unlink|rename)", c, re.I):
        guard_original(f"명령이 '{hit}' 를 언급하면서 파일을 저장·삭제합니다. 원본은 읽기 전용입니다 — 결과는 out/ 에 쓰세요.")
    # 1-b. 리다이렉트 대상
    for m in re.finditer(r"(?<![0-9&])>{1,2}\s*[\"']?([^\s\"';|&<>]+)", c):
        if protected(m.group(1)):
            guard_original(f"'{m.group(1)}' 로 리다이렉트(>)합니다. 원본은 읽기 전용입니다.")
    # 1-c. 쓰기 명령의 대상 인자
    for toks in _split(c):
        name = Path(toks[0]).name.lower()
        args = [t for t in toks[1:] if not t.startswith("-")]
        if name == "sed" and any(t.startswith("-i") for t in toks[1:]):
            targets = args[1:]
        elif WRITE_ANY_ARG.match(name):
            targets = args
        elif WRITE_LAST_ARG.match(name):
            targets = args[-1:]
        else:
            continue
        for t in targets:
            h = protected(t)
            if h:
                guard_original(f"'{toks[0]} … {t}': '{h}' 은 원본(읽기 전용)입니다. 산출물은 out/ 에 쓰세요.")


def main():
    try:
        data = json.loads(sys.stdin.buffer.read().decode("utf-8") or "{}")
    except json.JSONDecodeError:
        sys.exit(0)
    global IS_AGENT, CWD
    if data.get("cwd"):
        CWD = Path(data["cwd"])
    # 서브에이전트의 도구 호출이면 훅 입력에 agent_id / agent_type 이 들어온다
    IS_AGENT = bool(data.get("agent_id") or data.get("agent_type"))
    tool = data.get("tool_name", "")
    ti = data.get("tool_input") or {}
    if tool in ("Write", "Edit", "MultiEdit", "NotebookEdit"):
        check_file_tool(tool, ti)
    elif tool in ("Bash", "PowerShell"):
        check_shell(ti.get("command", ""))
    sys.exit(0)


if __name__ == "__main__":
    main()
