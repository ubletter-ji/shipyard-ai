"""웹 배포용 정적 페이지 만들기: out/투자시스템_결과.html → web/index.html

Vercel 은 web/ 폴더만 정적 사이트로 올린다 (vercel.json, .vercelignore).
- PC 의 원본 파일을 가리키는 file:/// 링크는 웹에서 열리지 않으므로 지운다 (원본 보기 = 슬라이드 이미지는 그대로).
- 상단에 '익명화 샘플로 만든 시연 화면' 안내를 붙인다.

사용: python scripts/build_web.py [결과 html 경로]   (기본: out/투자시스템_결과.html)
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "out" / "투자시스템_결과.html"
DST = ROOT / "web" / "index.html"

NOTICE = ('<div style="background:#fff4d6;color:#5b4300;padding:8px 16px;font-size:13px;text-align:center">'
          '공개 시연 화면 — 업체명·금액을 바꾼 <b>익명화 샘플</b> 신청서로 만든 결과입니다. '
          '실제 점검은 Claude Code 에이전트(invest-reviewer · invest-meeting)가 수행합니다.</div>')
HIDE_LOCAL = "<style>.mhead a,.origbtn[href='#']{display:none}</style>"


def main():
    if not SRC.exists():
        sys.exit(f"{SRC} 가 없습니다. 먼저 점검·회의자료를 만들어 결과 화면을 생성하세요.")
    html = SRC.read_text(encoding="utf-8")
    html, n = re.subn(r"file:///[^\"'\s<>&]*", "#", html)          # 로컬 파일 링크 제거 (srcdoc 안 포함)
    html = html.replace("</head>", HIDE_LOCAL + "</head>", 1)
    html = html.replace("<body>", "<body>" + NOTICE, 1)
    leaks = re.findall(r"[A-Za-z]:\\\\[^\"'<>]*|[A-Za-z]:/(?:work|Users)[^\"'<>]*", html)
    if leaks:
        sys.exit(f"로컬 경로가 남아 있습니다: {leaks[:3]}")
    DST.parent.mkdir(exist_ok=True)
    DST.write_text(html, encoding="utf-8")
    print(f"[웹] {DST.relative_to(ROOT)}  ({len(html) // 1024} KB, 로컬 링크 {n}개 제거)")


if __name__ == "__main__":
    main()
