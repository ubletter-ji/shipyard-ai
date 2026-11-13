"""원본 신청서 미리보기: pptx 슬라이드를 PNG 로 (out/원본미리보기/<지문>/sN.png).

- PowerPoint(Windows COM)로 **읽기 전용**으로 열어 슬라이드를 이미지로 내보낸다. 원본은 바꾸지 않는다.
- 같은 파일(SHA-256 지문)이면 다시 그리지 않는다.
- PowerPoint 가 없으면 빈 목록을 돌려준다 → 화면은 슬라이드 글자로 대신 보여 준다.
"""
from __future__ import annotations

import hashlib
import shutil
import subprocess
from pathlib import Path

from common import OUT

CACHE = OUT / "원본미리보기"

PS = r"""
$ErrorActionPreference = 'Stop'
$items = Get-Content -LiteralPath $args[0] -Encoding UTF8
$app = New-Object -ComObject PowerPoint.Application
try {
  foreach ($line in $items) {
    if (-not $line) { continue }
    $src, $dst = $line -split "`t"
    New-Item -ItemType Directory -Force $dst | Out-Null
    $p = $app.Presentations.Open($src, -1, 0, 0)   # ReadOnly, 창 없이
    $i = 0
    foreach ($s in $p.Slides) { $i++; $s.Export((Join-Path $dst "s$i.png"), 'PNG', 1280, 720) }
    $p.Close()
    Set-Content -LiteralPath (Join-Path $dst 'done') -Value $i
  }
} finally { $app.Quit() }
"""


def _key(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()[:16]


def _dir(p: Path) -> Path:
    return CACHE / _key(p)


def images(p: Path) -> list[Path]:
    d = _dir(p)
    if not (d / "done").exists():
        return []
    return sorted(d.glob("s*.png"), key=lambda x: int(x.stem[1:]))


def render(paths: list[Path]) -> None:
    """아직 그리지 않은 파일만 PowerPoint 를 한 번 띄워 한꺼번에 그린다."""
    todo = [p for p in dict.fromkeys(Path(x) for x in paths) if p.exists() and not images(p)]
    if not todo or not shutil.which("powershell"):
        return
    CACHE.mkdir(parents=True, exist_ok=True)
    lst = CACHE / "_todo.txt"
    lst.write_text("\n".join(f"{p.resolve()}\t{_dir(p).resolve()}" for p in todo), encoding="utf-8-sig")
    ps1 = CACHE / "_render.ps1"
    ps1.write_text(PS, encoding="utf-8-sig")            # BOM: Windows PowerShell 5.1 이 UTF-8 로 읽게
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1), str(lst)],
                       capture_output=True, timeout=300)
    except (OSError, subprocess.TimeoutExpired):
        pass
