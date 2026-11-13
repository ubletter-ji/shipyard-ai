"""에이전트 실행 전후 원본 변경 감지 훅 (막는 guard.py 와 별도의 '감지' 층).

  SubagentStart → snapshot : inbox/, data/, 기준 문서 2개의 SHA-256 을 기록
  SubagentStop  → verify   : 다시 계산해 달라졌으면 사용자에게 경고(systemMessage)

에이전트(invest-reviewer / invest-meeting)는 원본을 읽기만 해야 한다. guard.py 가 알려진 쓰기 경로를 막고,
이 훅은 그 밖의 경로(예: 스크립트 안에서의 저장)로 원본이 바뀐 경우까지 잡아낸다.
"""
from __future__ import annotations

import hashlib
import json
import os
import sys
from pathlib import Path

PROJECT = Path(os.environ.get("CLAUDE_PROJECT_DIR") or Path(__file__).resolve().parents[2]).resolve()
STATE = PROJECT / ".claude" / "hooks" / ".state" / "originals.json"
WATCH = ["inbox", "data"]
FILES = ["구현시나리오_보완.xlsx", "표준양식_수정.pptx"]


def fingerprint() -> dict:
    out = {}
    for d in WATCH:
        for p in sorted((PROJECT / d).glob("*")):
            if p.is_file() and not p.name.startswith("~$"):      # ~$ = PowerPoint/Excel 잠금 파일
                out[p.relative_to(PROJECT).as_posix()] = hashlib.sha256(p.read_bytes()).hexdigest()
    for f in FILES:
        p = PROJECT / f
        if p.exists():
            out[f] = hashlib.sha256(p.read_bytes()).hexdigest()
    return out


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    if mode == "snapshot":
        STATE.parent.mkdir(parents=True, exist_ok=True)
        STATE.write_text(json.dumps(fingerprint(), ensure_ascii=False), encoding="utf-8")
    elif mode == "verify":
        if not STATE.exists():
            return
        before = json.loads(STATE.read_text(encoding="utf-8"))
        now = fingerprint()
        changed = sorted(k for k in before.keys() | now.keys() if before.get(k) != now.get(k))
        if changed:
            msg = ("[안전장치] 에이전트 실행 중 원본이 바뀌었습니다: " + ", ".join(changed[:8])
                   + (" …" if len(changed) > 8 else "")
                   + " — 사용자가 직접 고친 게 아니라면 backup_* 에서 복구하고 점검을 다시 실행하세요.")
            sys.stdout.buffer.write(json.dumps({"systemMessage": msg}, ensure_ascii=False).encode("utf-8"))


if __name__ == "__main__":
    main()
