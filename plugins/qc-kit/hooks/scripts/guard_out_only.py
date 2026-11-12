"""PreToolUse guard for qc-aggregator: writes only allowed under <project>/out/."""
import json, os, re, sys

data = json.load(sys.stdin)
tool = data.get("tool_name", "")
inp = data.get("tool_input", {}) or {}
root = os.path.realpath(os.environ.get("CLAUDE_PROJECT_DIR") or data.get("cwd") or os.getcwd())
out_dir = os.path.join(root, "out")


def inside_out(p: str) -> bool:
    if not p:
        return False
    full = os.path.realpath(p if os.path.isabs(p) else os.path.join(root, p))
    return os.path.normcase(full).startswith(os.path.normcase(out_dir) + os.sep) or \
        os.path.normcase(full) == os.path.normcase(out_dir)


def block(msg: str):
    print(f"[qc-aggregator guard] {msg} — 쓰기는 out/ 아래만 허용됩니다.", file=sys.stderr)
    sys.exit(2)


if tool in ("Write", "Edit", "MultiEdit", "NotebookEdit"):
    path = inp.get("file_path") or inp.get("notebook_path") or ""
    if not inside_out(path):
        block(f"{tool} 대상 경로 차단: {path}")

elif tool in ("Bash", "PowerShell"):
    cmd = inp.get("command", "")
    # 원본 파일 변경/삭제 명령 차단
    if re.search(r"\b(rm|mv|del|rmdir|Remove-Item|Move-Item|Rename-Item|Set-Content|Clear-Content|truncate)\b", cmd):
        block(f"파일 변경/삭제 명령 차단: {cmd[:120]}")
    # 리다이렉트(>, >>, tee, cp, Copy-Item, Out-File) 대상이 out/ 이 아니면 차단
    for m in re.finditer(r"(?:>>?|\btee\s+(?:-a\s+)?|\bOut-File\s+(?:-FilePath\s+)?)\s*[\"']?([^\s\"';|&]+)", cmd):
        target = m.group(1)
        if target.startswith("&") or target in ("/dev/null", "$null", "nul", "NUL"):
            continue
        if not inside_out(target):
            block(f"리다이렉트 대상 차단: {target}")
    for m in re.finditer(r"\b(?:cp|Copy-Item)\b\s+(?:-\S+\s+)*\S+\s+[\"']?([^\s\"';|&]+)", cmd):
        if not inside_out(m.group(1)):
            block(f"복사 대상 차단: {m.group(1)}")

sys.exit(0)
