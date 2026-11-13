"""한 번에 확인: 환경 → 테스트 → 안전장치 설정 → (선택) 규칙 시연.

  python scripts/selfcheck.py          # 환경·테스트·설정 점검
  python scripts/selfcheck.py --demo   # + 임시 폴더에서 inbox 사본으로 변환·점검·HTML 생성 (원본 무변경)

에이전트(AI 판단)를 쓰지 않는 확정적 확인이다. 에이전트 시연은 README '시연' 절차를 따른다.
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OK, NG = "통과", "실패"
results: list[tuple[str, str, str]] = []


def rec(name, ok, detail=""):
    results.append((name, OK if ok else NG, detail))


def hashes():
    out = {}
    for d in ("inbox", "data"):
        for p in sorted((ROOT / d).glob("*")):
            if p.is_file() and not p.name.startswith("~$"):
                out[str(p)] = hashlib.sha256(p.read_bytes()).hexdigest()
    for f in ("구현시나리오_보완.xlsx", "표준양식_수정.pptx"):
        out[f] = hashlib.sha256((ROOT / f).read_bytes()).hexdigest()
    return out


def main():
    before = hashes()
    rec("Python 3.10+", sys.version_info >= (3, 10), sys.version.split()[0])
    for mod in ("openpyxl", "pptx"):
        try:
            __import__(mod)
            rec(f"패키지 {mod}", True)
        except ImportError:
            rec(f"패키지 {mod}", False, "pip install -r requirements.txt")

    try:
        s = json.loads((ROOT / ".claude" / "settings.json").read_text(encoding="utf-8"))
        hooks = s.get("hooks", {})
        rec("훅 설정 (PreToolUse·SubagentStart·SubagentStop)",
            all(k in hooks for k in ("PreToolUse", "SubagentStart", "SubagentStop")))
        rec("권한 deny 규칙", len(s.get("permissions", {}).get("deny", [])) > 0)
    except Exception as e:  # noqa: BLE001
        rec(".claude/settings.json", False, str(e))
    for f in ("guard.py", "inbox_watch.py"):
        rec(f"훅 스크립트 {f}", (ROOT / ".claude" / "hooks" / f).exists())
    for f in ("invest-reviewer.md", "invest-meeting.md"):
        rec(f"에이전트 {f}", (ROOT / ".claude" / "agents" / f).exists())
    rec("스킬 invest-standardize", (ROOT / ".claude" / "skills" / "invest-standardize" / "SKILL.md").exists())

    r = subprocess.run([sys.executable, "-m", "unittest", "discover", "-s", "tests"], cwd=ROOT,
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    tail = [l for l in r.stderr.splitlines() if l.startswith(("Ran ", "OK", "FAILED"))]
    rec("단위 테스트 (규칙·정답셋·변환·안전장치)", r.returncode == 0, " / ".join(tail))
    if r.returncode != 0:
        print(r.stderr[-3000:])

    if "--demo" in sys.argv:
        demo()

    rec("원본 무변경 (inbox·data·기준 문서)", hashes() == before)

    w = max(len(n) for n, _, _ in results)
    print("\n" + "=" * (w + 30))
    for n, st, d in results:
        print(f"{'✔' if st == OK else '✘'} {n.ljust(w)}  {st}  {d}")
    print("=" * (w + 30))
    sys.exit(0 if all(st == OK for _, st, _ in results) else 1)


def demo():
    """임시 폴더에 프로젝트 사본(+ 다른 양식 1건)을 만들어 변환→규칙 점검→HTML. 원본은 읽기만."""
    tmp = Path(tempfile.mkdtemp(prefix="invest_demo_"))
    for d in ("scripts", "config", "data"):
        shutil.copytree(ROOT / d, tmp / d, ignore=shutil.ignore_patterns("__pycache__"))
    shutil.copy2(ROOT / "표준양식_수정.pptx", tmp)
    (tmp / "inbox").mkdir()
    for p in (ROOT / "inbox").glob("*.pptx"):
        if not p.name.startswith("~$"):
            shutil.copy2(p, tmp / "inbox")
    run = lambda *a: subprocess.run([sys.executable, *a], cwd=tmp, capture_output=True, text=True,
                                    encoding="utf-8", errors="replace")
    run("scripts/make_samples.py", "--legacy")
    legacy = next((tmp / "samples" / "다른양식").glob("*.pptx"))
    shutil.copy2(legacy, tmp / "inbox")
    m = json.loads((ROOT / "tests" / "fixtures" / "신청서_06_매핑.json").read_text(encoding="utf-8"))
    m["원본"] = f"inbox/{legacy.name}"
    (tmp / "out" / "표준화").mkdir(parents=True)
    mp = tmp / "out" / "표준화" / f"{legacy.stem}.매핑.json"
    mp.write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
    b = run("scripts/standardize.py", "build", str(mp))
    rec("시연: 다른 양식 → 표준양식 변환 (추가·누락 0)", b.returncode == 0, b.stdout.strip().splitlines()[0] if b.stdout else b.stderr[-200:])
    c = run("scripts/review.py", "check")
    rep = run("scripts/review.py", "report")
    html = tmp / "out" / "검토결과.html"
    rec("시연: 규칙 점검 + HTML 결과 화면", html.exists(), str(html))
    print("\n[시연 점검 결과 — AI 판단 전, 규칙만]\n" + c.stdout + rep.stdout)


if __name__ == "__main__":
    main()
