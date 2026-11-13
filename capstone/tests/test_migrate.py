"""수량 칸 변환(scripts/migrate_quantity.py) 사전 검증 — 원본은 읽기만, 변환은 임시 폴더의 사본에서.

옛 양식(현수량/보유수량/필요수량)인 신청서가 있으면 그 사본을 변환해:
  - 칸 이름이 현수량/필요수량/투자수량 으로 바뀌는지
  - 값이 의미대로 옮겨져 R-102 판정이 옛 규칙의 판정과 같은지 (일부러 넣은 오류도 유지)
를 확인한다. 이미 새 양식이면 건너뛴다.
"""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))


def old_r102(유형, 현, 보유, 필요):
    """옛 규칙 (2026-11-12): 노후교체 필요 ≤ 현 / 그 외 현 − 보유 = 필요"""
    return 필요 <= 현 if 유형 == "노후교체" else 현 - 보유 == 필요


class TestMigrateQuantity(unittest.TestCase):
    def test_inbox_사본_변환(self):
        import extract
        import migrate_quantity as mq
        import review
        from common import load_config, to_number
        cfg = load_config()
        olds = []
        for p in sorted((ROOT / "inbox").glob("*.pptx")):
            if p.name.startswith("~$"):
                continue
            d = extract.extract(p)
            if "보유수량" in d["상세"]:
                olds.append((p, d))
        if not olds:
            self.skipTest("inbox 가 이미 새 양식(현수량/필요수량/투자수량)")
        with tempfile.TemporaryDirectory() as tmp:
            for p, d in olds:
                cp = Path(tmp) / p.name
                shutil.copy2(p, cp)
                mq.migrate(cp)
                nd = extract.extract(cp)
                self.assertNotIn("보유수량", nd["상세"], p.name)
                self.assertIn("투자수량", nd["상세"], p.name)
                s = d["상세"]
                q = [to_number(s.get(k)) for k in ("현수량", "보유수량", "필요수량")]
                r = {x["규칙ID"]: x["판정"] for x in review.check(nd, cfg)["rows"]}
                if all(v is not None for v in q):
                    want = "충족" if old_r102(s.get("투자유형", "").strip(), *q) else "보완"
                    self.assertEqual(r["R-102"], want, f"{p.name}: 옛 판정과 다름")
                else:
                    self.assertEqual(r["R-102"], "보완", p.name)

    def test_기준엑셀_사본_변환(self):
        import openpyxl
        import migrate_quantity as mq
        with tempfile.TemporaryDirectory() as tmp:
            cp = Path(tmp) / "spec.xlsx"
            shutil.copy2(ROOT / "구현시나리오_보완.xlsx", cp)
            done = mq.migrate_spec(cp)
            self.assertTrue({"R-102", "R-106", "R-107"} <= set(done))
            ws = openpyxl.load_workbook(cp)["검토 기준"]
            rows = {str(r[1].value).strip(): r for r in ws.iter_rows() if r[1].value}
            self.assertIn("투자수량 ≤ 필요수량 − 현수량", rows["R-102"][4].value)
            self.assertIn("5년 미만", rows["R-107"][4].value)

    def test_표준양식_사본_변환(self):
        import extract
        import migrate_quantity as mq
        with tempfile.TemporaryDirectory() as tmp:
            cp = Path(tmp) / "tpl.pptx"
            shutil.copy2(ROOT / "표준양식_수정.pptx", cp)
            mq.migrate(cp)
            d = extract.extract(cp)
            if "투자수량" not in d["상세"]:
                self.fail("표준양식 상세 표에 투자수량 칸이 없음")
            self.assertNotIn("보유수량", d["상세"])


if __name__ == "__main__":
    unittest.main()
