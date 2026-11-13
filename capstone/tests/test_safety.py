"""안전장치 테스트: 훅(.claude/hooks/guard.py, inbox_watch.py)과 저장소 비밀 값 검사.

실행: python -m unittest discover -s tests -v
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GUARD = ROOT / ".claude" / "hooks" / "guard.py"
WATCH = ROOT / ".claude" / "hooks" / "inbox_watch.py"
P = str(ROOT).replace("\\", "\\\\")


def run_guard(tool, agent=None, **tool_input):
    data = {"tool_name": tool, "tool_input": tool_input}
    if agent:
        data.update(agent_id="a1", agent_type=agent)     # 서브에이전트의 도구 호출
    payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
    env = {**os.environ, "CLAUDE_PROJECT_DIR": str(ROOT), "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run([sys.executable, str(GUARD)], input=payload, capture_output=True, env=env)
    return r.returncode, r.stderr.decode("utf-8", "replace"), r.stdout.decode("utf-8", "replace")


def run_guard_cwd(tool, cwd, agent=None, **tool_input):
    data = {"tool_name": tool, "tool_input": tool_input, "cwd": str(cwd)}
    if agent:
        data.update(agent_id="a1", agent_type=agent)
    env = {**os.environ, "CLAUDE_PROJECT_DIR": str(ROOT), "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run([sys.executable, str(GUARD)], input=json.dumps(data, ensure_ascii=False).encode("utf-8"),
                       capture_output=True, env=env)
    return r.returncode, r.stderr.decode("utf-8", "replace"), r.stdout.decode("utf-8", "replace")


class TestGuardBlocks(unittest.TestCase):
    """원본 쓰기: 에이전트 → 차단(exit 2) / 메인 세션 → 사용자 승인 요청(ask).
    위험 명령·비밀 값·샘플 강제 초기화 → 누구든 차단."""

    def blocked(self, tool, agent=None, **ti):
        code, err, _ = run_guard(tool, agent=agent, **ti)
        self.assertEqual(code, 2, f"{tool} {ti} 가 막히지 않음")
        self.assertIn("[안전장치]", err)

    def original(self, tool, **ti):
        self.blocked(tool, agent="invest-reviewer", **ti)            # 에이전트는 차단
        code, _, out = run_guard(tool, **ti)                          # 메인 세션은 승인 요청
        self.assertEqual(code, 0)
        d = json.loads(out)["hookSpecificOutput"]
        self.assertEqual(d["permissionDecision"], "ask", f"{tool} {ti}")
        self.assertIn("사람 확인", d["permissionDecisionReason"])

    def test_원본_파일도구_쓰기(self):
        self.original("Write", file_path="inbox/신청서_01_Z2골리앗크레인.pptx", content="x")
        self.original("Edit", file_path=str(ROOT / "data" / "경영계획.xlsx"), old_string="a", new_string="b")
        self.original("Write", file_path="구현시나리오_보완.xlsx", content="x")
        self.original("Write", file_path="표준양식_수정.pptx", content="x")
        self.original("Write", file_path="backup_20261112/x.txt", content="x")
        self.original("Write", file_path="samples/다른양식/x.pptx", content="x")

    def test_원본_셸_쓰기(self):
        self.original("Bash", command="rm inbox/신청서_02_용접기.pptx")
        self.original("Bash", command="cp samples/다른양식/x.pptx inbox/")
        self.original("Bash", command="mv out/a.pptx inbox/a.pptx")
        self.original("Bash", command="echo hi > data/경영계획.xlsx")
        self.original("Bash", command="sed -i s/a/b/ inbox/a.txt")
        self.original("Bash", command="python -c \"p=Presentation('inbox/a.pptx'); p.save('inbox/a.pptx')\"")
        self.original("PowerShell", command=f"Remove-Item {ROOT}\\inbox\\a.pptx")
        self.original("PowerShell", command=f"Copy-Item x.pptx {ROOT}\\표준양식_수정.pptx -Force")
        self.original("PowerShell", command=f"$p = $app.Presentations.Open('{ROOT}\\표준양식_수정.pptx'); $p.SaveAs('x')")

    def test_샘플_강제초기화(self):
        self.blocked("Bash", command="python scripts/make_samples.py --force")

    def test_위험_명령(self):
        self.blocked("Bash", command="git push --force origin main")
        self.blocked("Bash", command="git reset --hard HEAD~3")
        self.blocked("Bash", command="git clean -fd")
        self.blocked("Bash", command="rm -rf .")
        self.blocked("Bash", command="rm -rf /")
        self.blocked("PowerShell", command="Remove-Item -Recurse -Force C:\\")

    def test_비밀값_쓰기(self):
        self.blocked("Write", file_path="config/x.json", content='api_key = "abcd1234efgh5678"')
        self.blocked("Write", file_path="README.md", content="sk-ant-api03-" + "A" * 30)
        self.blocked("Edit", file_path="scripts/a.py", old_string="x", new_string="AKIA" + "ABCDEFGHIJKLMNOP")
        self.blocked("Write", file_path="k.pem", content="-----BEGIN RSA PRIVATE KEY-----\nabc")


class TestGuardAllows(unittest.TestCase):
    """정상 작업 → exit 0 (과잉 차단하지 않음)"""

    def allowed(self, tool, **ti):
        for agent in (None, "invest-reviewer"):
            code, err, out = run_guard(tool, agent=agent, **ti)
            self.assertEqual(code, 0, f"{tool} {ti} 가 잘못 막힘: {err}")
            self.assertEqual(out.strip(), "", f"{tool} {ti} 에 불필요한 승인 요청: {out}")

    def test_산출물_스크립트_문서(self):
        self.allowed("Write", file_path="out/검토결과.html", content="<html>")
        self.allowed("Write", file_path="out/표준화/a.매핑.json", content="{}")
        self.allowed("Edit", file_path="scripts/review.py", old_string="a", new_string="b")
        self.allowed("Write", file_path="docs/명세.md", content="# 명세\npassword 는 저장소에 두지 않는다")

    def test_사본_폴더는_원본이_아니다(self):
        demo = ROOT / "out" / "demo"
        for agent in (None, "invest-reviewer"):
            code, err, out = run_guard_cwd("Bash", demo, agent,
                                           command="python -c \"import pathlib; pathlib.Path('../_tmp/x.json').write_text('1')\" inbox")
            self.assertEqual((code, out.strip()), (0, ""), err)
            code, err, out = run_guard_cwd("Bash", ROOT, agent, command="cp x.pptx out/demo/inbox/x.pptx")
            self.assertEqual((code, out.strip()), (0, ""), err)
        code, _, _ = run_guard_cwd("Bash", demo, "invest-reviewer", command="cp x.pptx ../../inbox/x.pptx")
        self.assertEqual(code, 2)                                   # 사본 폴더에서도 진짜 원본을 가리키면 차단

    def test_읽기와_파이프라인(self):
        self.allowed("Read", file_path="inbox/a.pptx")
        self.allowed("Bash", command="python scripts/review.py check")
        self.allowed("Bash", command="python scripts/standardize.py dump inbox/a.pptx")
        self.allowed("Bash", command="cp inbox/a.pptx out/_tmp/")
        self.allowed("Bash", command="sha256sum inbox/*.pptx > out/_tmp/inbox.sha")
        self.allowed("Bash", command="python scripts/make_samples.py --legacy")
        self.allowed("Bash", command="rm -rf out")
        self.allowed("Bash", command="git push origin feature/x")


class TestInboxWatch(unittest.TestCase):
    """에이전트 실행 전후 원본 변경 감지"""

    def test_변경_감지(self):
        with tempfile.TemporaryDirectory() as tmp:
            proj = Path(tmp)
            (proj / "inbox").mkdir()
            (proj / "data").mkdir()
            f = proj / "inbox" / "a.pptx"
            f.write_bytes(b"one")
            (proj / "inbox" / "~$a.pptx").write_bytes(b"lock")   # 잠금 파일은 무시
            env = {**os.environ, "CLAUDE_PROJECT_DIR": str(proj), "PYTHONIOENCODING": "utf-8"}
            run = lambda m: subprocess.run([sys.executable, str(WATCH), m], capture_output=True, env=env)
            run("snapshot")
            self.assertEqual(run("verify").stdout.strip(), b"")             # 그대로면 조용
            (proj / "inbox" / "~$a.pptx").write_bytes(b"lock2")
            self.assertEqual(run("verify").stdout.strip(), b"")             # 잠금 파일 변화는 무시
            f.write_bytes(b"two")
            out = json.loads(run("verify").stdout.decode("utf-8"))
            self.assertIn("inbox/a.pptx", out["systemMessage"])


class TestSettings(unittest.TestCase):
    def test_훅_설정이_연결됨(self):
        s = json.loads((ROOT / ".claude" / "settings.json").read_text(encoding="utf-8"))
        pre = s["hooks"]["PreToolUse"][0]
        self.assertTrue({"Write", "Edit", "Bash", "PowerShell"} <= set(pre["matcher"].split("|")))
        self.assertIn("guard.py", pre["hooks"][0]["command"])
        start = s["hooks"]["SubagentStart"][0]["hooks"][0]["command"]
        stop = s["hooks"]["SubagentStop"][0]["hooks"][0]["command"]
        self.assertTrue("inbox_watch.py" in start and start.endswith("snapshot"))
        self.assertTrue("inbox_watch.py" in stop and stop.endswith("verify"))
        self.assertTrue(any("inbox" in r for r in s["permissions"]["deny"]))

    def test_에이전트_도구_권한_분리(self):
        def tools(name):
            head = (ROOT / ".claude" / "agents" / f"{name}.md").read_text(encoding="utf-8").split("---")[1]
            return {t.strip() for t in re.search(r"tools:\s*(.+)", head).group(1).split(",")}
        self.assertNotIn("Write", tools("invest-reviewer"))   # 점검은 판정만, 파일 직접 작성 없음
        self.assertNotIn("Edit", tools("invest-reviewer"))
        self.assertNotIn("Edit", tools("invest-meeting"))     # 회의자료는 새 파일만 (브리핑)


class TestNoSecretsInRepo(unittest.TestCase):
    def test_저장소에_비밀값_없음(self):
        sys.path.insert(0, str(GUARD.parent))
        import guard
        skip = {"out", "backup_20261112", "__pycache__", ".state"}
        hits = []
        for p in ROOT.rglob("*"):
            if not p.is_file() or set(p.relative_to(ROOT).parts) & skip or p.name.startswith("backup_"):
                continue
            if p.suffix.lower() not in {".py", ".md", ".json", ".txt", ".html", ".yml", ".yaml", ".env", ".cfg", ".ini"}:
                continue
            if p.name == "test_safety.py":      # 이 파일은 일부러 가짜 키 패턴을 담고 있음
                continue
            text = p.read_text(encoding="utf-8", errors="ignore")
            for pat, name in guard.SECRET_PATTERNS:
                if re.search(pat, text):
                    hits.append(f"{p.relative_to(ROOT)}: {name}")
        self.assertEqual(hits, [])
        self.assertFalse(any(ROOT.glob(".env*")) and not (ROOT / ".gitignore").read_text(encoding="utf-8").count(".env"))


if __name__ == "__main__":
    unittest.main()
