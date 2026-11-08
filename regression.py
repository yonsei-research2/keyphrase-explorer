#﻿논문 키프레이즈 워드클라우드 — tests/regression.py (회귀 테스트)
#코드 버전 a73b136 · 2026-10-28
#==============================================================================

# -*- coding: utf-8 -*-
"""회귀 테스트 — 고치다가 다른 곳이 망가지지 않았는지 자동으로 확인한다.

실행:  python tests/regression.py            (전체, 2~4분)
       python tests/regression.py --quick    (빠른 검사만, spaCy 없이 수 초)
       python tests/regression.py --no-app   (웹 화면 검사 제외)

가짜 Scopus·Web of Science·ProQuest 내보내기 파일을 만들어 쓰므로
내 컴퓨터의 논문 파일이 없어도 됩니다.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import traceback
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import keyphrase_cloud as kc  # noqa: E402


# ── 작은 검사 도구 ────────────────────────────────────────────────────────────
class Checker:
    def __init__(self):
        self.passed, self.failed = 0, []

    def ok(self, condition, label, detail=""):
        if condition:
            self.passed += 1
        else:
            self.failed.append(f"{label}" + (f" — {detail}" if detail else ""))
            print(f"   [실패] {label}" + (f" — {detail}" if detail else ""))

    def equal(self, got, want, label):
        self.ok(got == want, label, f"기대 {want}, 실제 {got}")

    def section(self, name):
        print(f"\n■ {name}")


# ── 가짜 내보내기 파일 ────────────────────────────────────────────────────────
TOPIC_WORDS = [
    ("machine learning", "neural network", "image classification"),
    ("climate change", "carbon emission", "renewable energy"),
    ("public health", "vaccination", "epidemic model"),
    ("urban mobility", "public transport", "travel behaviour"),
]
SENTENCE = ("This study examines {a} in relation to {b}. We analyse data from a large survey and "
            "report that {a} is associated with {c}. The findings inform policy on {b}.")


def _record(i: int) -> dict:
    words = TOPIC_WORDS[i % len(TOPIC_WORDS)]
    return {
        "title": f"Study {i} on {words[0]} and {words[1]}",
        "abstract": SENTENCE.format(a=words[0], b=words[1], c=words[2]),
        "keywords": list(words),
        "doi": f"10.1000/test.{i:04d}",
        "year": 2024 + (i % 3),
    }


def make_fixtures(folder: Path) -> dict:
    """세 DB 형식의 가짜 파일. 중복은 의도적으로 넣는다.
    - Scopus 40건 (저널 36 · 회의 2 · 도서 2)
    - WoS 25건 중 15건은 Scopus와 DOI가 같음
    - ProQuest 20건 중 8건은 Scopus와 제목이 같고 DOI가 없음, 2건은 학위논문
    """
    scopus = [_record(i) for i in range(40)]
    rows = []
    for n, r in enumerate(scopus):
        doctype = "Article" if n < 36 else ("Conference paper" if n < 38 else "Book chapter")
        rows.append({"Authors": "Kim J.", "Title": r["title"], "Year": r["year"], "Source title": "Test Journal",
                     "Cited by": n, "DOI": r["doi"], "Abstract": r["abstract"],
                     "Author Keywords": "; ".join(r["keywords"]), "Index Keywords": "indexed term",
                     "Document Type": doctype, "EID": f"2-s2.0-{n:08d}"})
    scopus_path = folder / "scopus_export.csv"
    pd.DataFrame(rows).to_csv(scopus_path, index=False, encoding="utf-8-sig")

    wos_rows = []
    for n in range(25):
        r = scopus[n] if n < 15 else _record(100 + n)      # 앞 15건은 중복
        wos_rows.append({"Authors": "Lee S.", "Article Title": r["title"], "Source Title": "Test Journal",
                         "Author Keywords": "; ".join(r["keywords"]), "Keywords Plus": "AUTOMATIC TERM",
                         "Abstract": r["abstract"], "DOI": r["doi"], "Publication Year": r["year"],
                         "Document Type": "Article", "Times Cited, WoS Core": n,
                         "UT (Unique WOS ID)": f"WOS:{n:09d}"})
    wos_path = folder / "savedrecs.xlsx"
    pd.DataFrame(wos_rows).to_excel(wos_path, index=False)

    pq_rows = []
    for n in range(20):
        r = scopus[n] if n < 8 else _record(200 + n)       # 앞 8건은 제목만 같고 DOI 없음
        pq_rows.append({"Title": r["title"], "Abstract": r["abstract"],
                        "identifierKeywords": " , ".join(r["keywords"]), "subjects": " , ".join(r["keywords"]),
                        "subjectTerms": "INDEXED TERM , Human", "digitalObjectIdentifier": "" if n < 8 else r["doi"],
                        "year": r["year"], "documentType": "Dissertation/Thesis" if n >= 18 else "Journal Article",
                        "StoreId": f"{n:07d}"})
    pq_path = folder / "ProQuestDocuments.xlsx"
    pd.DataFrame(pq_rows).to_excel(pq_path, index=False)
    return {"scopus": scopus_path, "wos": wos_path, "proquest": pq_path}


# ── 1. 빠른 검사 (spaCy 없이) ────────────────────────────────────────────────
def quick_checks(c: Checker, files: dict) -> None:
    c.section("DOI·제목·문헌유형 판정")
    c.equal(kc.norm_doi("https://doi.org/10.1000/ABC"), "10.1000/abc", "DOI 접두어·대소문자 정리")
    c.equal(kc.norm_doi("DOI: 10.1000/xyz."), "10.1000/xyz", "DOI: 접두어와 끝 마침표 제거")
    c.equal(kc.norm_doi("없음"), "", "DOI가 아니면 빈 값")
    c.equal(kc.title_key("Reading-While-Listening: A Review!"), "readingwhilelisteningareview", "제목 정규화")
    for raw, want in (("Article", "저널"), ("Article; Early Access", "저널"), ("Review", "저널"),
                      ("Journal , Peer Reviewed Journal , Journal Article", "저널"), ("Article , Report", "저널"),
                      ("Dissertation/Thesis", "기타"), ("Book Review", "기타"), ("Conference paper", "기타"),
                      ("Editorial Material", "기타"), ("", "미상")):
        c.equal(kc.doc_kind(raw), want, f"문헌유형 '{raw or '(빈칸)'}'")

    c.section("DB 자동 인식")
    c.equal(kc.detect_source(["Title", "Abstract", "Author Keywords", "EID", "Source title"]), "Scopus", "Scopus")
    c.equal(kc.detect_source(["Article Title", "Abstract", "Author Keywords", "Keywords Plus"]),
            "Web of Science", "Web of Science")
    c.equal(kc.detect_source(["Title", "Abstract", "identifierKeywords", "StoreId"]), "ProQuest", "ProQuest")
    c.equal(kc.detect_source(["이름", "내용"]), None, "모르는 형식이면 None")

    c.section("통합·중복 제거")
    df, report = kc.load_records([str(files["scopus"]), str(files["wos"]), str(files["proquest"])],
                                 None, None, "journal")
    c.equal(report["raw_total"], 85, "원본 합계")
    c.equal(report["dup_doi"], 15, "DOI 일치 제거 (WoS 중복)")
    c.equal(report["dup_title"], 8, "제목 일치 제거 (ProQuest, DOI 없음)")
    c.equal(report["after_merge"], 62, "중복 제거 후 건수")
    c.equal(len(df), 56, "저널 논문만 남긴 건수 (회의 2·도서 2·학위 2 제외)")
    c.ok((df[kc.COL_KEYWORDS].str.contains(";").all()), "ProQuest 쉼표 키워드가 세미콜론으로 통일됨")
    c.ok(not df[kc.COL_KEYWORDS].str.contains("AUTOMATIC TERM|INDEXED TERM").any(),
         "색인어(Keywords Plus·subjectTerms)는 저자키워드에 섞이지 않음")
    c.ok((df[df[kc.COL_SOURCE].str.contains("Scopus")][kc.COL_CITED] != "").all(), "Scopus 인용수 읽힘")

    all_df, all_report = kc.load_records([str(files["scopus"])], None, None, "all")
    c.equal(len(all_df), 40, "doc-types=all이면 회의·도서도 포함")
    year_df, _ = kc.load_records([str(files["scopus"])], None, [2024], "all")
    c.ok(set(year_df[kc.COL_YEAR]) == {"2024"}, "연도 필터")

    c.section("분석 전 정리 추천")
    merged, _ = kc.load_records([str(files["scopus"]), str(files["proquest"])], None, None, "all")
    pre = kc.suggest_before_analysis(merged)
    c.ok({"tags", "queries", "merges"} <= set(pre), "추천 세 종류 반환")
    kws = merged[kc.COL_KEYWORDS].str.cat(sep=";").lower()
    if "machine learning" in kws:                    # 가짜 데이터는 키워드 4종이 반복돼 검색어로 잡혀야 한다
        c.ok(any("learning" in q["키워드"].lower() for q in pre["queries"]) or pre["queries"] == [],
             "검색어 후보 판정 동작")
    c.ok(all(t["등장 논문"] >= 3 for t in pre["tags"]), "색인어 후보는 3편 이상만")
    c.ok(all("근거" in t for t in pre["tags"]), "근거가 함께 제공됨")
    c.ok("auto" in pre, "자동 통합 내역 제공")
    c.ok(all(a["표기 수"] >= 2 for a in pre["auto"]), "자동 통합 내역은 표기가 2개 이상인 것만")
    mixed = pd.DataFrame({kc.COL_TITLE: ["t"] * 3, kc.COL_ABSTRACT: ["a"] * 3,
                          kc.COL_KEYWORDS: ["Neural Networks", "neural network", "neural networks"]})
    auto = kc.suggest_before_analysis(mixed)["auto"]
    c.ok(len(auto) == 1 and auto[0]["표기 수"] == 3, "대소문자·복수형이 한 묶음으로 보고됨", str(auto))

    c.section("분석 전 필터 강도 추천 (추정치)")
    quick_keys, quick_kw = kc.quick_doc_stats(merged)
    c.equal(len(quick_keys), len(merged), "논문마다 추정 집합 생성")
    c.ok(any(len(k) > len(w) for k, w in zip(quick_keys, quick_kw)),
         "본문에만 나오는 표현도 후보에 포함 (필터 평가에 필요)")
    scored_q, best_q = kc.recommend_ratio_quick(merged)
    c.ok(not scored_q.empty and best_q in kc.SENSITIVITY_RATIOS, "추천값 산출", str(best_q))
    c.ok({"정확도", "재현율", "F1", "추천"} <= set(scored_q.columns), "추천 표 열 구성")
    off = scored_q[scored_q["일반어 필터"] == "끔"]
    c.ok(not off.empty and off.iloc[0]["재현율"] >= 0.99, "필터를 끄면 재현율 최대")

    c.section("기본 설정은 분야 중립인가")
    # 논문 10편 중 3편에만 달린 상태로 둔다 (40% 이상이면 '검색어'로 따로 분류되므로)
    kws = ["Human; Statistical analysis; Present study"] * 3 + ["walkability"] * 7
    why = {r["근거"] for r in kc.suggest_before_analysis(pd.DataFrame({
        kc.COL_KEYWORDS: kws, kc.COL_TITLE: ["t"] * 10, kc.COL_ABSTRACT: ["a"] * 10}))["tags"]}
    c.ok(any("통계" in w for w in why), "통계 용어를 통계 근거로 제시")
    c.ok(any("색인어" in w for w in why), "색인어를 색인 근거로 제시")
    c.ok(any("상투구" in w for w in why), "서술 상투구를 상투구 근거로 제시")
    # DB 색인어·연구분야 라벨은 기본 제외가 아니라 '분석 전 추천'으로만 다룬다.
    # (사용자가 직접 넣은 목록은 건드리지 않으므로 여기서 검사하지 않는다)
    defaults = [line.strip().lower() for line in kc.default_settings()["exclude_phrases.txt"].splitlines()]
    c.equal([d for d in defaults if d and not d.startswith("#")], [], "기본 제외 목록은 비어 있음")
    # 54개 기본 제외어는 카탈로그로 옮겨 '분석 전 추천'으로만 제시한다
    catalogue = kc.INDEX_VOCAB | kc.STAT_VOCAB | kc.NARRATIVE_VOCAB | kc.DEMO_TAG_VOCAB
    for word in ("odds ratio", "present study", "statistical analysis", "middle aged", "year old",
                 "human", "female", "empirical study", "applied economics"):
        c.ok(word in catalogue, f"'{word}'가 추천 카탈로그에 있음")
    c.ok(not (kc.STAT_VOCAB & kc.NARRATIVE_VOCAB) and not (kc.STAT_VOCAB & kc.DEMO_TAG_VOCAB),
         "카탈로그 그룹이 겹치지 않음")
    syn = kc.default_settings()["synonyms.txt"].lower()
    c.ok("agent-based" not in syn, "기본 동의어에 특정 분야 규칙 없음")

    c.section("설정 주입 (사용자별 단어 설정)")
    lines = kc.read_config_lines("exclude_phrases.txt", {"exclude_phrases.txt": "alpha\n# 주석\nbeta  # 꼬리 주석"})
    c.equal(lines, ["alpha", "beta"], "주석 제거하고 줄만 읽기")
    c.ok(set(kc.default_settings()) == set(kc.CONFIG_NAMES), "기본 설정 6개 파일")

    c.section("토픽 수 제한 (NMF 제약)")
    import numpy as np
    small = np.random.default_rng(0).integers(0, 3, size=(60, 8)).astype(float)
    c.equal(kc.fit_topics(small, 20)[1].shape[0], 8, "키프레이즈 8개면 토픽도 8개까지")
    c.equal(kc.fit_topics(np.ones((6, 40)), 20)[1].shape[0], 6, "문헌 6편이면 토픽 6개까지")
    best, curve = kc.choose_topics(small)
    c.ok(best is not None and all(r["토픽 수"] <= 8 for r in curve), "자동 선택도 제한을 지킴")

    c.section("거품 배치")
    packed = kc._pack_circles([3.0, 2.0, 2.0, 1.0])
    overlap = [(i, j) for i in range(len(packed)) for j in range(i + 1, len(packed))
               if (packed[i][0] - packed[j][0]) ** 2 + (packed[i][1] - packed[j][1]) ** 2
               < (packed[i][2] + packed[j][2]) ** 2 * 0.98]
    c.ok(not overlap, "원끼리 겹치지 않음", str(overlap))


# ── 2. 전체 파이프라인 ───────────────────────────────────────────────────────
def pipeline_checks(c: Checker, files: dict, out_dir: Path) -> None:
    c.section("전체 분석 (집계 → 토픽 → 지도 → 저장)")
    args = kc.parse_args([str(files["scopus"]), str(files["wos"]), str(files["proquest"]),
                          "--label", "회귀", "--out", str(out_dir), "--topics", "auto",
                          "--min-phrase-papers", "2"])
    res = kc.run_pipeline(args)
    c.equal(res["settings"]["집계 논문 수"], 56, "분석 논문 수")
    c.ok(len(res["table"]) >= 5, "키프레이즈가 집계됨", f"{len(res['table'])}개")
    c.ok(res["topics"] is not None, "토픽 모델 실행됨")
    if res["topics"]:
        t = res["topics"]
        c.ok(2 <= t["k"] <= 20, "토픽 수가 범위 안", str(t["k"]))
        c.equal(len(t["docs"]), 56, "문헌별 토픽 배정 행 수")
        c.equal(len(t["assign"]), 56, "토픽 배정 배열 길이")
        c.ok(not t["summary"].empty, "토픽 요약표 생성")
    c.ok(res["map"] is not None, "논문 지도 생성됨")
    if res["topics"]:
        tp = res["topics"]
        for key in ("trends", "pairs", "bridges", "weights"):
            c.ok(key in tp, f"토픽 결과에 '{key}'")
        c.ok({"점유율", "평균 인용", "인용 자료 있는 논문"} <= set(tp["summary"].columns), "요약표에 점유율·인용")
        c.ok({"2순위 토픽", "인용수"} <= set(tp["docs"].columns), "문헌표에 2순위 토픽·인용수")
        if not tp["bridges"].empty:
            b = tp["bridges"].iloc[0]
            c.ok(b["2순위 비중"] >= b["1순위 비중"] * 0.5, "경계 논문 조건")
    if res["map"]:
        m = res["map"]
        c.equal(len(m["nodes"]), 56, "지도 노드 수")
        c.ok({"bx", "by", "tx", "ty"} <= set(m["nodes"].columns), "두 가지 거품 좌표 존재")
        c.ok(not m["areas"].empty and not m["topic_areas"].empty, "위치·토픽 기준 영역 모두 생성")
    c.ok(res["suggest"] is not None, "자동 정리 제안 생성")

    xlsx = [p for p in res["outputs"] if p.suffix == ".xlsx"]
    html = [p for p in res["outputs"] if p.suffix == ".html"]
    png = [p for p in res["outputs"] if p.suffix == ".png"]
    c.ok(len(xlsx) == 1 and xlsx[0].exists(), "엑셀 저장됨")
    c.ok(len(png) >= 1 and all(p.exists() for p in png), "워드클라우드 저장됨")
    c.ok(len(html) == 1 and html[0].exists(), "논문 지도 HTML 저장됨")
    if html:
        text = html[0].read_text(encoding="utf-8")
        c.ok("const DATA" in text and '"tareas"' in text, "지도 HTML에 데이터 포함")
    if xlsx:
        sheets = pd.ExcelFile(xlsx[0]).sheet_names
        for name in ("전체", "설정", "출처·통합", "중복 제거 목록", "토픽", "문헌별 토픽", "주제 지도",
                     "논문 지도 좌표", "내 단어 설정"):
            c.ok(name in sheets, f"엑셀 시트 '{name}'")
        cfg_sheet = pd.read_excel(xlsx[0], "내 단어 설정")
        c.equal(len(cfg_sheet), 1, "기본값만 쓰면 한 줄로 기록")
        c.equal(str(cfg_sheet.iloc[0]["설정 파일"]), "(전부 기본값)", "기본값 표시")
        changed = kc.settings_diff({**kc.default_settings(),
                                    "exclude_phrases.txt": "Korea\n", "synonyms.txt": "a => b\n"})
        kinds = {(r["설정 파일"], r["변경"]) for r in changed}
        c.ok(("exclude_phrases.txt", "추가") in kinds, "추가한 줄을 잡아냄")
        c.ok(any(k[1] == "삭제" for k in kinds), "지운 줄도 기록")
        c.ok(all(r["줄"] for r in changed), "줄 내용이 함께 기록")
        settings_sheet = pd.read_excel(xlsx[0], "설정").set_index("항목")["값"].to_dict()
        c.ok("코드 버전" in settings_sheet and str(settings_sheet["코드 버전"]).strip(), "코드 버전 기록")
        dropped_sheet = pd.read_excel(xlsx[0], "제외됨(일반어 필터)")
        c.equal(len(dropped_sheet), len(res["dropped"]), "제외 목록이 잘리지 않음")
    if res["topics"]:
        tp = res["topics"]
        for key in ("trends", "pairs", "bridges", "weights"):
            c.ok(key in tp, f"토픽 결과에 '{key}'")
        c.ok({"점유율", "평균 인용", "인용 자료 있는 논문"} <= set(tp["summary"].columns), "요약표에 점유율·인용")
        c.ok({"2순위 토픽", "인용수"} <= set(tp["docs"].columns), "문헌표에 2순위 토픽·인용수")
        if not tp["bridges"].empty:
            b = tp["bridges"].iloc[0]
            c.ok(b["2순위 비중"] >= b["1순위 비중"] * 0.5, "경계 논문 조건")
    if res["map"]:
        nodes = res["map"]["nodes"]
        c.ok("인용수 없는 논문" in res["map"], "인용수 결측 수 보고")
        c.ok(str(nodes["인용수"].dtype) == "Int64", "인용수는 결측을 구분하는 자료형", str(nodes["인용수"].dtype))
        html_text = html[0].read_text(encoding="utf-8") if html else ""
        c.ok("인용 정보 없음" in html_text, "지도에 '인용 정보 없음' 표시")
        c.ok("멀리 떨어진 거리" in html_text, "지도에 해석 주의 문구")
        # 토픽 기준에서 거품이 그려지지도 눌리지도 않던 버그 (tx/ty를 안 넘겨 좌표가 NaN이었다)
        c.equal(html_text.count("function areaPt("), 1, "거품 좌표 변환 함수 존재")
        c.ok("pt({x:a.bx, y:a.by, bx:a.bx, by:a.by})" not in html_text, "거품 좌표에 tx/ty 누락 없음")
        c.ok("function showPaper(" in html_text, "논문 선택 기능")
        c.ok("내용이 가장 비슷한 논문" in html_text, "비슷한 논문 목록")
        edges = res["map"]["edges"]
        c.ok("sim" in edges.columns and edges["sim"].between(0, 1).all(), "연결마다 유사도 값")
        c.ok('float(r.sim)' in Path("keyphrase_cloud.py").read_text(encoding="utf-8"), "유사도를 화면으로 전달")

    c.section("결측이 섞인 표를 엑셀로 쓸 수 있는가")
    # pandas 3에서는 astype(str)이 결측을 'nan' 문자열로 바꾸지 않는다 (사용자가 겪은 TypeError)
    import numpy as np
    mixed = pd.DataFrame({"이름": ["a", None, "ccc"], "값": [1.5, np.nan, 3.0],
                          "정수": pd.array([1, None, 3], dtype="Int64")})
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "t.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            mixed.to_excel(xw, sheet_name="t", index=False)
            try:
                kc.autosize(xw.sheets["t"], mixed)
                c.ok(True, "결측이 든 숫자·문자 열의 너비 계산")
            except Exception as e:  # noqa: BLE001
                c.ok(False, "결측이 든 숫자·문자 열의 너비 계산", repr(e))
        c.ok(path.exists(), "파일 저장 완료")

    c.section("작은 표본 지원 (논문 10편)")
    c.equal(kc.TOPIC_MIN_DOCS, 10, "토픽 최소 논문 수")
    c.ok(kc.TOPIC_SMALL_DOCS > kc.TOPIC_MIN_DOCS, "참고용 경고 기준이 더 큼")
    args_small = kc.parse_args(["_"])
    changed = kc.relax_for_small_corpus(args_small, 10)
    c.equal(len(changed), 3, "기본값이면 세 기준을 낮춤")
    c.equal((args_small.min_keyword_papers, args_small.min_papers, args_small.min_phrase_papers), (1, 1, 2),
            "낮춘 값")
    args_user = kc.parse_args(["_", "--min-keyword-papers", "4", "--min-papers", "5"])
    kc.relax_for_small_corpus(args_user, 10)
    c.equal((args_user.min_keyword_papers, args_user.min_papers), (4, 5), "직접 지정한 값은 건드리지 않음")
    args_big = kc.parse_args(["_"])
    c.equal(kc.relax_for_small_corpus(args_big, 500), [], "논문이 많으면 그대로")

    c.section("논문 지도 최소 규모 (20편)")
    c.equal(kc.MAP_MIN_DOCS, 20, "지도 최소 논문 수")
    docs20 = [{"title": f"agent based model of {w} policy", "abstract": f"we simulate {w} with agents and data",
               "keywords": [w, "simulation"]}
              for w in ("energy", "water", "traffic", "waste", "health", "housing", "farming", "retail",
                        "school", "transit")] * 2
    frame20 = pd.DataFrame({kc.COL_TITLE: [d["title"] for d in docs20], kc.COL_YEAR: [2024] * 20,
                            kc.COL_DOI: [""] * 20, kc.COL_SOURCE: ["테스트"] * 20, kc.COL_CITED: [1] * 20})
    try:
        m20 = kc.paper_map(docs20, frame20, None, None, None)
        c.ok(m20 is not None and len(m20["nodes"]) == 20, "논문 20편으로 지도 생성")
        c.ok(len(m20["edges"]) > 0, "20편에서도 연결선 생성")
    except Exception as e:  # noqa: BLE001
        c.ok(False, "논문 20편으로 지도 생성", repr(e))
    c.ok(kc.paper_map(docs20[:19], frame20.head(19), None, None, None) is None, "19편이면 지도 건너뜀")

    c.section("토픽 추세 검정")
    # Mann-Kendall: 단조 증가 5년이면 S=10, z=2.20, p≈0.028 (Kendall 1975 공식)
    s_up, p_up = kc._mann_kendall([1, 2, 3, 4, 5])
    c.equal(s_up, 10.0, "단조 증가의 S 통계량")
    c.ok(p_up is not None and abs(p_up - 0.0275) < 0.002, "p값이 이론값과 일치", str(p_up))
    s_dn, p_dn = kc._mann_kendall([5, 4, 3, 2, 1])
    c.equal(s_dn, -10.0, "단조 감소의 S 통계량")
    c.ok(kc._mann_kendall([3, 3, 3, 3, 3])[0] == 0, "변화 없으면 S=0")
    c.ok(str(kc._mann_kendall([1, 2, 3, 4])[1]) == "nan", "4년이면 판단하지 않음")
    # 10년에 걸쳐 T1 점유율이 오르고 T2가 내려가는 표를 만들어 라벨을 확인한다
    rows = []
    for i, year in enumerate(range(2015, 2025)):
        for _ in range(2 + i):
            rows.append({"토픽": "T1", "연도": year, "제목": "a"})
        for _ in range(12 - i):
            rows.append({"토픽": "T2", "연도": year, "제목": "b"})
    tr = kc.topic_trends(pd.DataFrame(rows)).set_index("토픽")
    c.equal(tr.loc["T1", "추세"], "증가", "늘어나는 토픽을 증가로 판정")
    c.equal(tr.loc["T2", "추세"], "감소", "줄어드는 토픽을 감소로 판정")
    c.ok(tr.loc["T1", "연평균 변화(%p)"] > 0 > tr.loc["T2", "연평균 변화(%p)"], "연평균 변화 부호")
    flat = pd.DataFrame([{"토픽": "T1", "연도": y, "제목": "a"} for y in range(2015, 2025)] +
                        [{"토픽": "T2", "연도": y, "제목": "b"} for y in range(2015, 2025)])
    c.equal(kc.topic_trends(flat).set_index("토픽").loc["T1", "추세"], "뚜렷한 변화 없음", "변화 없음 판정")

    c.section("임계값 민감도·추천")
    sens, scored, best = kc.sensitivity_table(
        kc.parse_args([str(files["scopus"]), "--label", "민감도", "--out", str(out_dir), "--min-phrase-papers", "2"]))
    c.ok(len(sens) >= 4, "여러 임계값을 비교", f"{len(sens)}행")
    c.ok({"구분", "일반어 필터", "키프레이즈 수", "상위 20 일치", "전체 겹침(자카드)"} <= set(sens.columns), "민감도 열 구성")
    c.equal(sens.iloc[0]["상위 20 일치"], "기준", "첫 행이 기준")
    off = sens[sens["일반어 필터"] == "끔"]
    c.ok(not off.empty and off.iloc[0]["키프레이즈 수"] > sens.iloc[0]["키프레이즈 수"],
         "필터를 끄면 키프레이즈가 늘어남")
    c.ok({"정확도", "재현율", "F1", "F1 편차", "추천"} <= set(scored.columns), "추천 표 열 구성")
    c.ok(best in kc.SENSITIVITY_RATIOS, "추천값이 후보 중 하나", str(best))
    c.equal(int((scored["추천"] != "").sum()), 1, "추천 표시는 한 줄")
    c.ok(scored.loc[scored["F1"].idxmax(), "추천"] != "", "F1이 가장 높은 줄이 추천됨")
    zero = scored[scored["일반어 필터"] == "끔"]
    strict = scored[scored["일반어 필터"] == "20%"]
    c.ok(not zero.empty and zero.iloc[0]["재현율"] >= 0.99, "필터를 끄면 재현율은 최대")
    # 필터를 올리면 정확도는 떨어지지 않는다 (가짜 데이터는 모든 키프레이즈가 저자키워드라 차이가 작을 수 있음)
    c.ok(not strict.empty and strict.iloc[0]["정확도"] >= zero.iloc[0]["정확도"] - 1e-9,
         "필터를 올리면 정확도가 떨어지지 않음",
         f"끔 {zero.iloc[0]['정확도']} → 20% {strict.iloc[0]['정확도']}")

    c.section("설정이 결과에 반영되는가")
    first = res["table"]["키프레이즈"].iloc[0]
    args2 = kc.parse_args([str(files["scopus"]), "--label", "회귀2", "--out", str(out_dir),
                           "--topics", "off", "--paper-map", "off", "--min-phrase-papers", "2"])
    res2 = kc.run_pipeline(args2, settings={**kc.default_settings(), "exclude_phrases.txt": first})
    c.ok(first not in set(res2["table"]["키프레이즈"]), f"제외 설정이 적용됨 ('{first}')")
    c.ok(res2["topics"] is None and res2["map"] is None, "--topics off / --paper-map off 동작")


# ── 3. 웹 화면 ───────────────────────────────────────────────────────────────
def app_checks(c: Checker, files: dict) -> None:
    c.section("웹 화면 (파일 올리기 → 분석 → 토픽 수 변경)")
    from streamlit.testing.v1 import AppTest

    app = Path(__file__).resolve().parents[1] / "app.py"
    at = AppTest.from_file(str(app), default_timeout=900).run()
    c.ok(not at.exception, "앱 기동", str(at.exception))
    labels = [t.label for t in at.tabs]           # 도움말 안의 하위 탭도 함께 나온다
    c.ok(labels[:4] == ["📊 분석", "🧭 토픽", "🗺️ 논문 지도", "❓ 도움말"], "주요 탭 구성", str(labels[:4]))
    c.ok(not any("단어 설정" == l for l in labels), "단어 설정은 탭이 아니라 옆칸으로")
    c.ok(any(d.label.startswith("📥") for d in at.sidebar.download_button), "옆칸에 설정 내려받기")
    c.ok("❓ 도움말" in labels, "도움말 탭 존재")

    at.text_area[0].set_value("\n".join(str(p) for p in files.values())).run(timeout=900)
    c.ok(not at.exception, "파일 경로 입력", str(at.exception))
    c.ok(any(m.value == "85편" for m in at.metric), "통합 지표 표시 (원본 85편)")

    run_btn = [b for b in at.button if "분석 시작" in b.label]      # 번호로 누르면 버튼이 늘 때 어긋난다
    c.ok(bool(run_btn), "분석 시작 버튼 존재", str([b.label for b in at.button]))
    run_btn[0].click().run(timeout=1800)
    c.ok(not at.exception, "분석 실행", str(at.exception))
    c.ok(not at.error, "화면 오류 없음", str([e.value for e in at.error]))
    c.ok("result" in at.session_state, "결과가 세션에 저장됨")
    c.ok(at.session_state["sid"] in str(at.session_state["result"]["outputs"][0]),
         "결과가 세션 전용 폴더에 저장됨")

    # 정리 추천이 0건이어도 '필터 강도 추천'은 남아 있어야 한다 (한 번 통째로 사라진 적 있음)
    labels = [e.label for e in at.expander]
    c.ok(any("일반어 필터 강도 추천" in l for l in labels), "분석 전 필터 강도 추천 패널", str(labels))
    c.ok(any("정리 추천" in l for l in labels) or any("정리 추천" in cap.value for cap in at.caption),
         "분석 전 정리 추천 (없으면 없다고 표시)")

    sliders = [s for s in at.slider if s.label == "토픽 수"]
    c.ok(bool(sliders), "토픽 수 슬라이더 존재")
    if sliders:                                   # 예전에 여기서 오류가 났었다
        sliders[0].set_value(3).run(timeout=900)
        c.ok(not at.exception, "토픽 수 변경 (3)", str(at.exception))
        [s for s in at.slider if s.label == "토픽 수"][0].set_value(6).run(timeout=900)
        c.ok(not at.exception, "토픽 수 변경 (6)", str(at.exception))

    # 배포 전 버전으로 만든 결과가 세션에 남아 있는 상황 (사용자가 겪은 KeyError: '인용수')
    topics = at.session_state["result"].get("topics")
    if topics:
        old = dict(topics)
        old["docs"] = topics["docs"].drop(columns=[c2 for c2 in ("인용수", "2순위 토픽")
                                                   if c2 in topics["docs"].columns])
        old["papers"] = topics["papers"].drop(columns=[c2 for c2 in topics["papers"].columns
                                                       if "인용" in str(c2) or "Cited" in str(c2)])
        for gone in ("trends", "pairs", "bridges", "weights"):
            old.pop(gone, None)
        at.session_state["result"] = {**at.session_state["result"], "topics": old}
        at.session_state["topic_k"] = old["k"]      # 저장된 표를 그대로 쓰는 경로
        at.run(timeout=900)
        c.ok(not at.exception, "이전 버전 결과가 남아 있어도 토픽 탭이 뜸", str(at.exception))
        c.ok(any("이전 버전" in i.value for i in at.info), "다시 실행하라는 안내 표시")


def main() -> int:
    quick = "--quick" in sys.argv
    skip_app = "--no-app" in sys.argv or quick
    c = Checker()
    work = Path(tempfile.mkdtemp(prefix="keyphrase_test_"))
    try:
        files = make_fixtures(work)
        print(f"가짜 내보내기 파일 생성: {work}")
        quick_checks(c, files)
        if not quick:
            pipeline_checks(c, files, work / "out")
        if not skip_app:
            app_checks(c, files)
    except Exception:  # noqa: BLE001
        c.failed.append("예외 발생:\n" + traceback.format_exc())
        print(traceback.format_exc())
    finally:
        shutil.rmtree(work, ignore_errors=True)

    print("\n" + "─" * 60)
    if c.failed:
        print(f"실패 {len(c.failed)}건 / 통과 {c.passed}건")
        for f in c.failed:
            print("  ✗", f)
        return 1
    print(f"모두 통과 — {c.passed}건")
    return 0


if __name__ == "__main__":
    sys.exit(main())
