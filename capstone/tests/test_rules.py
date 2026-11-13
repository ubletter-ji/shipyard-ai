"""정답셋 테스트: 구현시나리오_보완.xlsx '계획 매칭 규칙' T1~T6 + 샘플 신청서 검토 결과.

실행: python -m unittest discover -s tests -v   (프로젝트 루트에서)
"""
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

from match import match  # noqa: E402

PLANS = [
    {"계획NO": "NO.1", "투자명": "Z1 골리앗 크레인 신설", "투자유형": "생산필수", "장비종류": "골리앗 크레인",
     "AREA": "Z1", "수량": 1, "계획금액": 5000},
    {"계획NO": "NO.2", "투자명": "A1 안벽 연장", "투자유형": "생산필수", "장비종류": "안벽",
     "AREA": "A1", "수량": 1, "계획금액": 3000},
    {"계획NO": "NO.3", "투자명": "용접기 500대 구매", "투자유형": "노후교체", "장비종류": "용접기",
     "AREA": "전사", "수량": 500, "계획금액": 500},
    {"계획NO": "NO.4", "투자명": "A팀 로봇용접기 1대 구매", "투자유형": "생산성 향상(자동화)", "장비종류": "로봇용접기",
     "AREA": "A팀", "수량": 1, "계획금액": 500},
]


def app(name, 유형, 금액, area, no, 수량):
    return {"투자명": name, "투자유형": 유형, "투자금액": 금액, "AREA": area, "계획NO": no, "수량": 수량}


class TestPlanMatching(unittest.TestCase):
    def one(self, a, history=None):
        return match([a], PLANS, history)["rows"][0]

    def test_T1_위치변경(self):
        r = self.one(app("Z2 골리앗 크레인 신설", "생산필수", 3000, "Z2", "NO.1", 1))
        self.assertIn("M-2", r["규칙"])
        self.assertEqual(r["비고"], "계획 Z1 → 실행 Z2 로 변경")
        self.assertEqual(r["차이금액"], -2000)

    def test_T2_수량_유형변경(self):
        r = self.one(app("용접기 300대 구매", "생산필수", 200, "전사", "NO.3", 300))
        self.assertTrue({"M-3", "M-9"} <= set(r["규칙"]))
        self.assertEqual(r["비고"], "계획 500대 → 실행 300대로 변경 / 유형 노후교체 → 생산필수 변경")
        self.assertEqual(r["차이금액"], -300)
        self.assertTrue(r["확인필요"])

    def test_T3_계획외(self):
        r = self.one(app("GR 100대 구매", "노후교체", 50, "전사", "계획外", 100))
        self.assertIn("M-4", r["규칙"])
        self.assertEqual(r["비고"], "계획 外")
        self.assertEqual((r["경영계획"], r["차이금액"]), (0, 50))

    def test_T4_계획초과(self):
        r = self.one(app("A팀 로봇용접기 1대 구매", "생산성 향상(자동화)", 800, "A팀", "NO.4", 1))
        self.assertIn("M-6", r["규칙"])
        self.assertEqual(r["비고"], "계획 금액 초과")
        self.assertEqual(r["차이금액"], 300)

    def test_T5_미신청(self):
        res = match([app("Z2 골리앗 크레인 신설", "생산필수", 3000, "Z2", "NO.1", 1)], PLANS)
        m = {x["계획NO"]: x for x in res["미신청"]}
        self.assertIn("NO.2", m)
        self.assertEqual(m["NO.2"]["비고"], "미신청")
        self.assertEqual(m["NO.2"]["규칙"], ["M-8"])

    def test_T6_분할신청(self):
        hist = [{"계획NO": "NO.3", "집행금액": 200, "투자명": "용접기 300대 구매"}]
        r = self.one(app("용접기 200대 구매 (2차)", "노후교체", 150, "전사", "NO.3", 200), hist)
        self.assertEqual(r["규칙"], ["M-0", "M-7"])
        self.assertEqual(r["비고"], "누적 350 / 계획 500 (잔여 150)")
        self.assertEqual(r["차이금액"], -350)  # 차이는 이번 신청분만

    def test_M1_계획대로(self):
        r = self.one(app("A1 안벽 연장", "생산필수", 3000, "A1", "NO.2", None))
        self.assertIn("M-1", r["규칙"])
        self.assertEqual(r["비고"], "")

    def test_M5_후보복수(self):
        plans = PLANS + [{"계획NO": "NO.9", "투자명": "Z3 골리앗 크레인 신설", "투자유형": "생산필수",
                          "장비종류": "골리앗 크레인", "AREA": "Z3", "수량": 1, "계획금액": 4000}]
        r = match([app("Z2 골리앗 크레인 신설", "생산필수", 3000, "Z2", "", 1)], plans)["rows"][0]
        self.assertIn("M-5", r["규칙"])
        self.assertTrue(r["비고"].startswith("후보 NO.1, NO.9"))

    def test_계획NO_미기재시_제안(self):
        r = self.one(app("A팀 로봇용접기 1대 구매", "생산성 향상(자동화)", 400, "A팀", "", 1))
        self.assertEqual(r["계획NO"], "NO.4")  # '용접기'가 아니라 더 구체적인 '로봇용접기'와 매칭
        self.assertIn("NO.4 제안", r["비고"])
        self.assertTrue(r["확인필요"])


class TestSampleReview(unittest.TestCase):
    """샘플 신청서 5건의 규칙 판정 (make_samples.APPS 기준, inbox 와 무관)."""

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    @classmethod
    def setUpClass(cls):
        # inbox 는 사용자가 고칠 수 있으므로, 샘플 정의(APPS)로 임시 폴더에 고정 샘플을 만들어 검사한다.
        import contextlib
        import io
        import tempfile
        import extract
        import make_samples
        import review
        from common import load_config
        cfg = load_config()
        cls._tmp = tempfile.TemporaryDirectory()
        folder = Path(cls._tmp.name)
        with contextlib.redirect_stdout(io.StringIO()):
            for spec in make_samples.APPS:
                make_samples.make_app(spec, folder)
        cls.res = {}
        for p in sorted(folder.glob("신청서_0*.pptx")):
            r = review.check(extract.extract(p), cfg)
            cls.res[p.stem] = ({x["규칙ID"]: x["판정"] for x in r["rows"]}, review.summarize(r))

    def test_정상_신청서는_보완_0건(self):
        for k in ("신청서_01_Z2골리앗크레인", "신청서_02_용접기", "신청서_03_GR", "신청서_04_A팀로봇용접기"):
            self.assertEqual(self.res[k][1]["보완"], 0, k)

    def test_추가슬라이드_항목_인식(self):
        rows, _ = self.res["신청서_03_GR"]
        self.assertEqual((rows["R-203"], rows["R-311"], rows["R-312"], rows["R-313"]), ("충족",) * 4)

    def test_최저가_아니어도_선정사유_있으면_충족(self):
        rows, _ = self.res["신청서_04_A팀로봇용접기"]
        self.assertEqual(rows["R-106"], "충족")

    def test_누락_신청서는_보완요청(self):
        rows, sm = self.res["신청서_05_C동도장부스"]
        self.assertEqual(sm["종합상태"], "보완 요청")
        for rid in ("R-102", "R-105", "R-107", "R-203", "R-301", "R-302"):
            self.assertEqual(rows[rid], "보완", rid)

    def test_AI판단_대기중이면_상정_안됨(self):
        for _, sm in self.res.values():
            self.assertEqual(sm["회의상정"], "보류")


class TestTypeRequiredInfo(unittest.TestCase):
    """유형별 필수 정보: 생산필수=내년 부하·조업도 / 노후교체=현 장비 재원+노후 수량 / 자동화=금액·JMH·인력 감소."""

    def run_check(self, 유형, 부가):
        import review
        from common import load_config
        d = {"신청서ID": "t", "파일": "t.pptx", "투자명": "장비 10대 구매", "투자효과": "필요",
             "상세": {"투자유형": 유형, "투자금액(백만원)": "100", "현수량": "0", "필요수량": "1", "투자수량": "1",
                    "선정 업체": "A사", "견적 수": "2"},
             "부가": 부가, "부가_출처": {k: "슬라이드 1" for k in 부가},
             "견적": [{"업체": "A사", "금액": 100, "설립연도": 2000, "선정": True},
                    {"업체": "B사", "금액": 110, "설립연도": 2000, "선정": False}]}
        return {x["규칙ID"]: x for x in review.check(d, load_config())["rows"]}

    def test_생산필수_부하조업도_없으면_보완(self):
        r = self.run_check("생산필수", {"내년도 부하·조업도": "처리량 월 12개 → 20개"})
        self.assertEqual(r["R-302"]["판정"], "보완")
        r = self.run_check("생산필수", {"내년도 부하·조업도": "조업도 현재 92% → 설치 후 80%"})
        self.assertEqual(r["R-302"]["판정"], "충족")

    def test_노후교체_노후수량_없으면_보완(self):
        base = "운용연수: 15년\n보유 수량: 10대, 교체 수량: 10대\n고장 이력 증가"
        r = self.run_check("노후교체", {"현 장비 재원 소개": base})
        self.assertEqual(r["R-314"]["판정"], "보완")
        r = self.run_check("노후교체", {"현 장비 재원 소개": base + "\n노후 수량: 10대"})
        self.assertEqual(r["R-314"]["판정"], "충족")

    def test_노후교체_교체가_노후보다_많으면_판단필요(self):
        t = "운용연수: 15년\n보유 수량: 10대, 노후 수량: 6대, 교체 수량: 10대\n고장 이력 증가"
        self.assertEqual(self.run_check("노후교체", {"현 장비 재원 소개": t})["R-314"]["판정"], "판단 필요")

    def test_자동화_세가지_모두_필수(self):
        r = self.run_check("생산성 향상(자동화)", {"생산성 향상 산출": "연 절감 320백만원, 인원 2명 감소"})
        self.assertEqual(r["R-321"]["판정"], "보완")
        self.assertIn("JMH", r["R-321"]["보완요청문구"])
        r = self.run_check("생산성 향상(자동화)",
                           {"생산성 향상 산출": "1,600 JMH 절감, 인원 2명 감소, 연 320백만원"})
        self.assertEqual(r["R-321"]["판정"], "충족")

    def test_다른_유형에는_적용_안됨(self):
        r = self.run_check("생산필수", {"내년도 부하·조업도": "부하 118% → 96%"})
        self.assertEqual((r["R-314"]["판정"], r["R-321"]["판정"]), ("해당 없음", "해당 없음"))


class TestQuantityRelation(unittest.TestCase):
    """R-102 (2026-11-13): 현수량/필요수량/투자수량.
    노후교체 = 투자 ≤ 필요 / 그 외 = 투자 ≤ 필요 − 현."""

    def r102(self, 유형, 현, 필요, 투자):
        import review
        from common import load_config
        d = {"신청서ID": "t", "파일": "t.pptx", "투자명": "장비", "투자효과": "",
             "상세": {"투자유형": 유형, "현수량": str(현), "필요수량": str(필요), "투자수량": str(투자)},
             "부가": {}, "부가_출처": {}, "견적": []}
        return next(x for x in review.check(d, load_config())["rows"] if x["규칙ID"] == "R-102")

    def test_노후교체_필요수량까지(self):
        self.assertEqual(self.r102("노후교체", 100, 100, 100)["판정"], "충족")
        self.assertEqual(self.r102("노후교체", 100, 60, 40)["판정"], "충족")    # 필요 이하면 됨 (현수량 무관)
        r = self.r102("노후교체", 100, 60, 61)
        self.assertEqual(r["판정"], "보완")
        self.assertIn("필요수량(60)", r["보완요청문구"])

    def test_그외_부족분까지(self):
        for t in ("생산필수", "생산성 향상(자동화)"):
            self.assertEqual(self.r102(t, 1200, 1500, 300)["판정"], "충족", t)
            self.assertEqual(self.r102(t, 1200, 1500, 250)["판정"], "충족", t)   # 부족분 이하
        r = self.r102("생산필수", 1200, 1500, 301)
        self.assertEqual(r["판정"], "보완")
        self.assertIn("1,500", r["보완요청문구"])
        self.assertIn("300", r["보완요청문구"])

    def test_부족분이_없으면_보완(self):
        r = self.r102("생산필수", 5, 5, 1)
        self.assertEqual(r["판정"], "보완")
        self.assertIn("추가로 투자할 수량이 없습니다", r["보완요청문구"])

    def test_투자수량_0이하_보완(self):
        self.assertEqual(self.r102("노후교체", 10, 10, 0)["판정"], "보완")


class TestNewCompanyQuote(unittest.TestCase):
    """R-106/R-107 (2026-11-13): 업력 5년 미만 신생 업체는 최저가여도 미선정."""

    def rows(self, quotes, 사유=""):
        import datetime as dt
        import review
        from common import load_config
        d = {"신청서ID": "t", "파일": "t.pptx", "투자명": "장비", "투자효과": "",
             "상세": {"투자유형": "생산필수", "투자금액(백만원)": str(next(q["금액"] for q in quotes if q["선정"]))},
             "부가": {}, "부가_출처": {}, "견적": quotes, "선정사유": 사유}
        r = review.check(d, load_config(), today=dt.date(2026, 11, 13))
        return {x["규칙ID"]: x for x in r["rows"]}

    E = {"업체": "E사", "금액": 800, "설립연도": 2003, "선정": True}
    F = {"업체": "F사", "금액": 780, "설립연도": 2023, "선정": False}

    def test_최저가_신생_미선정_사유있으면_충족(self):
        r = self.rows([self.E, self.F], 사유="F사는 신생 업체라 제외")
        self.assertEqual(r["R-106"]["판정"], "충족")
        self.assertIn("신생 업체라 미선정(기준 부합)", r["R-106"]["근거"])
        self.assertEqual(r["R-107"]["판정"], "충족")

    def test_최저가_신생_미선정_사유없으면_예시문구_제시(self):
        r = self.rows([self.E, self.F])
        self.assertEqual(r["R-106"]["판정"], "보완")
        t = r["R-106"]["보완요청문구"]
        self.assertIn("F사(2023년 설립, 업력 3년)", t)
        self.assertIn("설립 5년이 되지 않은 신생 업체", t)
        self.assertIn("업력 23년의 E사를 선정", t)

    def test_신생_업체를_선정하면_보완(self):
        F_sel = {**self.F, "선정": True}
        E_not = {**self.E, "선정": False}
        r = self.rows([E_not, F_sel])
        self.assertEqual(r["R-106"]["판정"], "충족")          # 최저가 선정 자체는 맞지만
        self.assertEqual(r["R-107"]["판정"], "보완")          # 신생이라 선정 불가
        self.assertIn("업력 5년 이상", r["R-107"]["보완요청문구"])

    def test_5년_경계(self):
        G = {"업체": "G사", "금액": 700, "설립연도": 2021, "선정": True}   # 업력 5년 → 신생 아님
        r = self.rows([G, {**self.E, "선정": False}])
        self.assertEqual(r["R-107"]["판정"], "충족")


class TestInvestPeriod(unittest.TestCase):
    """R-108: 투자경과 = 투자 시작 후 걸리는 개월 수."""

    def r108(self, 시작, 경과):
        import review
        from common import load_config
        d = {"신청서ID": "t", "파일": "t.pptx", "투자명": "장비", "투자효과": "",
             "상세": {"투자유형": "생산필수", "투자시작": 시작, "투자경과": 경과},
             "부가": {}, "부가_출처": {}, "견적": []}
        return next(x for x in review.check(d, load_config())["rows"] if x["규칙ID"] == "R-108")

    def test_개월수_있으면_충족_완료예상_표시(self):
        r = self.r108("27.4月", "5개월")
        self.assertEqual(r["판정"], "충족")
        self.assertIn("완료 예상 2027년 9월", r["근거"])
        self.assertIn("2028년 3월", self.r108("27.11月", "4개월")["근거"])  # 연도 넘김 (11월+4개월)

    def test_공란이면_보완(self):
        self.assertEqual(self.r108("27.4月", "")["판정"], "보완")
        self.assertEqual(self.r108("27.4月", "미정")["판정"], "보완")


class TestTemplateGuidance(unittest.TestCase):
    """표준양식의 안내 문구만 남은 칸은 '작성된 내용'으로 보지 않는다."""

    def test_빈_표준양식은_부가항목_없음(self):
        import extract
        d = extract.extract(ROOT / "표준양식_수정.pptx")
        self.assertEqual({k: v for k, v in d["부가"].items() if v}, {})

    def test_표준양식에_새_필수정보_안내(self):
        from pptx import Presentation
        text = " ".join(sh.text_frame.text if sh.has_text_frame else
                        " ".join(c.text for r in sh.table.rows for c in r.cells) if sh.has_table else ""
                        for s in Presentation(str(ROOT / "표준양식_수정.pptx")).slides for sh in s.shapes)
        for phrase in ("노후 수량", "인력 감소 (3가지 모두)", "내년 부하·조업도를 현재"):
            self.assertIn(phrase, text)


class TestSourceFingerprint(unittest.TestCase):
    """점검 후 원본이 바뀌면 지문(해시)이 달라진다 → 회의자료 단계에서 재점검 필요로 보류."""

    def test_원본_변경_감지(self):
        import shutil
        import tempfile
        import extract
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "신청서.pptx"
            shutil.copy(ROOT / "표준양식_수정.pptx", p)
            before = extract.extract(p)["원본해시"]
            self.assertEqual(before, extract.file_hash(p))
            with open(p, "ab") as f:
                f.write(b"\0")
            self.assertNotEqual(before, extract.file_hash(p))


class TestStandardize(unittest.TestCase):
    """다른 양식 → 표준양식: 내용 그대로면 통과, 지어내거나 빠뜨리면 실패."""

    @classmethod
    def setUpClass(cls):
        import contextlib
        import io
        import json
        import tempfile
        import make_samples
        cls._tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls._tmp.name)
        with contextlib.redirect_stdout(io.StringIO()):
            src = make_samples.make_legacy(cls.dir / "src")
        m = json.loads((ROOT / "tests" / "fixtures" / "신청서_06_매핑.json").read_text(encoding="utf-8"))
        m["원본"] = str(src)
        cls.mapping = m

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def write(self, m, name="m.json"):
        import json
        p = self.dir / name
        p.write_text(json.dumps(m, ensure_ascii=False), encoding="utf-8")
        return p

    def test_다른양식_감지(self):
        import standardize
        self.assertFalse(standardize.is_standard(Path(self.mapping["원본"])))
        self.assertTrue(standardize.is_standard(ROOT / "표준양식_수정.pptx"))

    def test_올바른_매핑은_통과하고_점검기로_읽힌다(self):
        import extract
        import review
        import standardize
        from common import load_config
        rec = standardize.build(self.write(self.mapping), self.dir / "out")
        self.assertTrue(rec["통과"], rec)
        d = extract.extract(Path(rec["변환본"]))
        self.assertEqual(d["투자명"], "B도크 도장 로봇 2대 도입")
        self.assertEqual(d["변환원본"]["해시"], extract.file_hash(Path(self.mapping["원본"])))
        rows = {x["규칙ID"]: x for x in review.check(d, load_config())["rows"]}
        self.assertEqual(rows["R-100"]["판정"], "충족")
        self.assertEqual(rows["R-104"]["판정"], "충족")   # 640백만원 = K사 640
        self.assertEqual(rows["R-106"]["판정"], "충족")   # '채택' 표시 인식, 최저가
        self.assertEqual(rows["R-321"]["판정"], "충족")

    def test_지어낸_글은_실패(self):
        import copy
        import standardize
        m = copy.deepcopy(self.mapping)
        m["상세"]["견적 수"] = "2곳"          # 원문에 '2곳' 이라는 칸 값은 없음 (세어서 넣음)
        m["투자효과"] = "도장 품질 향상"       # 원문에 없는 문장
        res = standardize.verify(m)
        self.assertFalse(res["통과"])
        self.assertIn("도장 품질 향상", res["추가"])

    def test_빠뜨린_글은_실패(self):
        import copy
        import standardize
        m = copy.deepcopy(self.mapping)
        m["기타원문"] = []                    # '작성 부서' 줄을 버림
        res = standardize.verify(m)
        self.assertFalse(res["통과"])
        self.assertTrue(any("작성 부서" in x for x in res["누락"]))

    def test_단위를_빼면_누락(self):
        import copy
        import standardize
        m = copy.deepcopy(self.mapping)
        m["상세"]["투자금액(백만원)"] = "640"    # '백만원' 을 버림 — 단위도 내용이다
        res = standardize.verify(m)
        self.assertFalse(res["통과"])
        self.assertTrue(any("백만원" in x for x in res["누락"]))


class TestRuleDocs(unittest.TestCase):
    """업무 규칙의 기준은 CLAUDE.md. 에이전트·스킬 문서의 규칙 행은 CLAUDE.md 와 같아야 하고,
    코드에 있는 규칙 ID 는 모두 문서에 적혀 있어야 한다. 에이전트·스킬은 엑셀·pptx 를 '참고'하지 않는다."""

    DOCS = [".claude/agents/invest-reviewer.md", ".claude/agents/invest-meeting.md",
            ".claude/skills/invest-standardize/SKILL.md"]

    @staticmethod
    def rule_rows(text):
        return {l.split("|")[1].strip(): l.strip() for l in text.splitlines()
                if re.match(r"^\| (R-\d{3}|M-\d|T\d) \|", l)}

    def test_규칙_행이_CLAUDE_md_와_같다(self):
        base = self.rule_rows((ROOT / "CLAUDE.md").read_text(encoding="utf-8"))
        for d in self.DOCS:
            for rid, line in self.rule_rows((ROOT / d).read_text(encoding="utf-8")).items():
                self.assertIn(rid, base, f"{d}: {rid} 가 CLAUDE.md 에 없음")
                self.assertEqual(line, base[rid], f"{d}: {rid} 행이 CLAUDE.md 와 다름")

    def test_코드의_규칙ID가_문서에_있다(self):
        import review
        base = self.rule_rows((ROOT / "CLAUDE.md").read_text(encoding="utf-8"))
        reviewer = self.rule_rows((ROOT / ".claude/agents/invest-reviewer.md").read_text(encoding="utf-8"))
        for rid in list(review.규칙명) + ["R-401"]:
            self.assertIn(rid, base, f"CLAUDE.md 에 {rid} 없음")
            self.assertIn(rid, reviewer, f"invest-reviewer.md 에 {rid} 없음")
        meeting = self.rule_rows((ROOT / ".claude/agents/invest-meeting.md").read_text(encoding="utf-8"))
        for mid in [f"M-{i}" for i in range(10)] + [f"T{i}" for i in range(1, 7)] + ["R-401"]:
            self.assertIn(mid, meeting, f"invest-meeting.md 에 {mid} 없음")

    def test_수치기준이_config와_같다(self):
        from common import load_config
        cfg = load_config()
        text = (ROOT / "CLAUDE.md").read_text(encoding="utf-8")
        self.assertIn(f"업력 {cfg['R107_신생업체_기준_년']}년", text)
        self.assertIn(f"견적 {cfg['R105_최소_견적수']}곳 이상", text)
        self.assertIn(f"교체 기준 {cfg['R311_노후교체_운용연수_기준_년']}년", text)
        self.assertIn(f"회수기간 {cfg['R322_회수기간_기준_년']}년", text)

    def test_에이전트_스킬은_기준파일을_참고하지_않는다(self):
        for d in self.DOCS:
            t = (ROOT / d).read_text(encoding="utf-8")
            for name in ("구현시나리오_보완.xlsx", "표준양식_수정.pptx"):
                self.assertNotIn(name, t, d)


class TestMeetingHtml(unittest.TestCase):
    """회의자료 HTML: 표지 숫자·합계, 상정 신청서 섹션(표준양식 칸), 보류 건은 섹션 없이 사유만."""

    def test_회의자료_HTML(self):
        import json
        import tempfile
        import meeting_html as mh
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            (out / "extract").mkdir()
            d = {"신청서ID": "A", "파일": "A.pptx", "투자명": "Z2 골리앗 크레인 신설", "투자효과": "기존 골리앗 사용 불가",
                 "상세": {"투자시작": "27.4月", "투자금액(백만원)": "3000", "현수량": "0", "필요수량": "1", "투자수량": "1"},
                 "부가": {"견적서 비교": "A사: 3,000백만원, 설립 1998년 (선정)"}, "견적": [], "슬라이드": [{"번호": 1, "제목": "", "텍스트": ""}]}
            (out / "extract" / "A.json").write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")
            rows = [{"_no": 1, "신청서ID": "A", "투자명": d["투자명"], "투자유형": "생산필수", "계획NO": "NO.1",
                     "경영계획": 5000.0, "실행금액": 3000.0, "차이금액": -2000.0, "비고": "계획 Z1 → 실행 Z2 로 변경", "확인필요": False}]
            보류 = [{"신청서ID": "B", "투자명": "C동 도장부스 신설", "사유": "보완 7건"}]
            old = mh.OUT
            mh.OUT = out
            try:
                html = mh.write_meeting_html(rows, [{"계획NO": "NO.2", "투자명": "A1 안벽 연장", "경영계획": 3000}],
                                             보류, out / "m.html").read_text(encoding="utf-8")
            finally:
                mh.OUT = old
        self.assertIn("투자 심의 요약", html)
        self.assertIn("−2,000", html)                       # 차이금액
        self.assertIn("계획 Z1 → 실행 Z2 로 변경", html)
        self.assertIn('id="app1"', html)                    # 상정 신청서 섹션
        self.assertIn("<th>투자수량</th>", html)             # 표준양식 상세 표 칸
        self.assertIn("기존 골리앗 사용 불가", html)         # 원문 그대로
        self.assertNotIn("C동 도장부스", html)                 # 보류 건은 회의자료에 없음 (실무자 보고로만)
        self.assertNotIn("보류", html)
        self.assertNotIn('id="app2"', html)
        self.assertIn("A1 안벽 연장", html)                  # 계획 미신청


class TestSystemHtml(unittest.TestCase):
    def test_한화면_탭_두개(self):
        import system_html as sh
        src = (ROOT / "scripts" / "system_html.py").read_text(encoding="utf-8")
        self.assertIn("① 투자검토 결과", src)
        self.assertIn("② 회의자료", src)
        self.assertTrue(callable(sh.build))


class TestOriginalView(unittest.TestCase):
    def test_원본보기_버튼(self):
        import system_html as sh
        b = sh.orig_button("신청서_02", "file:///C:/x/신청서_02.pptx")
        self.assertIn('data-orig="신청서_02"', b)
        self.assertIn("postMessage", b)               # 한 화면 안에서는 모달
        self.assertIn('href="file:///C:/x/', b)        # 단독으로 열면 파일 열기

    def test_미리보기_없으면_빈목록(self):
        import tempfile
        import originals
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "x.pptx"
            p.write_bytes(b"not rendered")
            self.assertEqual(originals.images(p), [])  # 그린 적 없으면 글자 대체로


class TestStartScreen(unittest.TestCase):
    def test_첨부_조회_화면(self):
        src = (ROOT / "scripts" / "system_html.py").read_text(encoding="utf-8")
        for s in ('id="start"', 'id="run"', ">조회<", 'type="file"', 'id="result" hidden', "점검 전 — inbox 에 넣고 점검 실행 필요"):
            self.assertIn(s, src)


class TestMeetingGate(unittest.TestCase):
    """R-401: 보완 0건 + 판단 필요 행 전부 검토자 확인 O 일 때만 상정."""

    def r(self, *rows):
        return {"rows": [{"규칙ID": f"R-{i}", "판정": p, "AI대기": False, "검토자확인": c}
                         for i, (p, c) in enumerate(rows)]}

    def test_판단필요_미확인이면_보류(self):
        from review import summarize
        self.assertEqual(summarize(self.r(("충족", ""), ("판단 필요", "")))["회의상정"], "보류")

    def test_판단필요_검토자O면_상정(self):
        from review import summarize
        self.assertEqual(summarize(self.r(("충족", ""), ("판단 필요", "O")))["회의상정"], "상정")

    def test_보완있으면_검토자O여도_보류(self):
        from review import summarize
        self.assertEqual(summarize(self.r(("보완", "O"), ("판단 필요", "O")))["회의상정"], "보류")


if __name__ == "__main__":
    unittest.main()
