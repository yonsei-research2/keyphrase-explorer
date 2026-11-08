#﻿논문 키프레이즈 워드클라우드 — keyphrase_cloud.py (분석 엔진)
#코드 버전 a73b136 · 2026-10-28
#==============================================================================

# -*- coding: utf-8 -*-
"""
Scopus 논문 키프레이즈 빈도 집계 + 워드클라우드 생성기

제목(Title) · 초록(Abstract) · 저자키워드(Author Keywords)에서 키프레이즈를 뽑아
노출빈도를 집계하고, 엑셀 표와 워드클라우드 이미지를 만듭니다.

  - 단수/복수형 통합 : models / model, analyses / analysis, mice / mouse, LLMs / LLM ...
  - 다단어 구문 인식 : "large language model", "quality of life" 등을 하나의 키프레이즈로 집계
      1) 코퍼스 전체의 저자키워드를 구문 사전으로 사용
      2) 제목/초록의 명사구(형용사·명사 연속) 중 여러 논문에 반복 등장하는 구문을 자동 추출
      3) config/custom_phrases.txt 에 적은 구문
    → 제목/초록은 사전에 대해 '가장 긴 구문 우선'으로 매칭하고, 남은 명사는 단어로 집계
  - 논문 내 약어 정의("large language models (LLMs)")를 찾아 약어를 원래 표현으로 치환

사용 예)
  python keyphrase_cloud.py "YRS 2025_scopus.xlsx"
  python keyphrase_cloud.py "YRS 2026_scopus.xlsx" --label YRS2026 --year 2026
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import pickle
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
CONFIG_DIR = BASE_DIR / "config"
CACHE_DIR = BASE_DIR / ".cache"
CACHE_VERSION = 1
SPACY_MODEL = "en_core_web_sm"

COL_TITLE, COL_ABSTRACT, COL_KEYWORDS = "Title", "Abstract", "Author Keywords"
SECTION_LABELS = {"title": "제목", "abstract": "초록", "keywords": "저자키워드"}

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

_logger = None


def set_logger(fn) -> None:
    """진행 메시지를 받을 함수 지정 (웹 화면용). None이면 콘솔 출력."""
    global _logger
    _logger = fn


def log(*parts) -> None:
    msg = " ".join(str(p) for p in parts)
    if _logger:
        _logger(msg)
    else:
        print(msg)


# ─────────────────────────────────────────────────────────────────────────────
# 1. 데이터 읽기 (Scopus · Web of Science · ProQuest)
#    DB마다 다른 열 이름을 공통 형식으로 맞추고, DOI·제목으로 중복 문헌을 제거해
#    통합 목록을 만든 뒤 분석을 시작한다.
# ─────────────────────────────────────────────────────────────────────────────
COL_YEAR, COL_DOI, COL_SOURCE, COL_DOCTYPE, COL_FILE = "Year", "DOI", "출처", "문헌유형", "원본파일"
COL_RAWTYPE, COL_CITED = "원본 문헌유형", "인용수"

SOURCE_PROFILES = {
    "Scopus": {
        "mark": ("EID", "Source title", "Index Keywords"),
        "title": ("Title",), "abstract": ("Abstract",), "keywords": ("Author Keywords",),
        "doi": ("DOI",), "year": ("Year",), "doctype": ("Document Type",),
        "cited": ("Cited by",), "sep": ";",
    },
    "Web of Science": {
        "mark": ("UT (Unique WOS ID)", "Keywords Plus", "Article Title", "UT"),
        "title": ("Article Title", "TI"), "abstract": ("Abstract", "AB"),
        "keywords": ("Author Keywords", "DE"), "doi": ("DOI", "DI"),
        "year": ("Publication Year", "PY"), "doctype": ("Document Type", "DT"),
        "cited": ("Times Cited, WoS Core", "Times Cited, All Databases", "TC"), "sep": ";",
    },
    "ProQuest": {
        "mark": ("StoreId", "digitalObjectIdentifier", "identifierKeywords"),
        "title": ("Title",), "abstract": ("Abstract",),
        "keywords": ("identifierKeywords", "subjects"), "doi": ("digitalObjectIdentifier",),
        "year": ("year", "pubdate"), "doctype": ("documentType", "ArticleType"),
        "cited": (), "sep": ",",
    },
}

DOI_PREFIX_RE = re.compile(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)", re.I)
TITLE_KEY_RE = re.compile(r"[^0-9a-z]+")
# 저널 논문이 아닌 유형 (학위논문·도서·회의록 등). 아래 단어가 하나라도 있으면 '기타'
NON_JOURNAL_WORDS = ("dissertation", "thesis", "book", "conference", "proceeding", "meeting",
                     "editorial", "erratum", "correction", "retract", "patent", "news", "학위")
JOURNAL_WORDS = ("article", "review", "journal", "논문")


def norm_doi(value) -> str:
    s = DOI_PREFIX_RE.sub("", str(value or "").strip().lower()).rstrip(" .;,")
    return s if s.startswith("10.") else ""


def title_key(value) -> str:
    """대소문자·구두점·공백을 지운 제목 (중복 판정용)."""
    return TITLE_KEY_RE.sub("", str(value or "").lower())


def doc_kind(*values) -> str:
    """문헌 유형을 '저널' / '기타' / '미상'으로 분류."""
    text = " , ".join(str(v or "").lower() for v in values).strip(" ,")
    if not text:
        return "미상"
    if any(w in text for w in NON_JOURNAL_WORDS):
        return "기타"
    return "저널" if any(w in text for w in JOURNAL_WORDS) else "기타"


def _sheet_frames(path: Path, sheet: str | None):
    """파일 안의 시트를 (시트이름, DataFrame)으로 하나씩 돌려준다."""
    suffix = path.suffix.lower()
    if suffix == ".csv":
        yield "(csv)", pd.read_csv(path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
        return
    if suffix == ".xls":  # 구형 엑셀 (WoS savedrecs.xls, ProQuest 내보내기)
        try:
            sheets = pd.read_excel(path, sheet_name=sheet, dtype=str, engine="xlrd")
        except ImportError:
            raise SystemExit("구형 .xls 파일을 읽으려면 xlrd가 필요합니다:  python -m pip install xlrd")
        if isinstance(sheets, pd.DataFrame):
            sheets = {sheet or "(sheet1)": sheets}
        for name, frame in sheets.items():
            yield name, frame.fillna("")
        return

    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet and sheet not in wb.sheetnames:
            raise SystemExit(f"시트 '{sheet}'을(를) 찾을 수 없습니다. 시트 목록: {wb.sheetnames}")
        for name in [sheet] if sheet else wb.sheetnames:
            rows = wb[name].iter_rows(values_only=True)
            header = next(rows, None)
            if not header:
                continue
            header = ["" if h is None else str(h).strip() for h in header]
            if not any(header):
                continue
            width = max(i for i, h in enumerate(header) if h) + 1
            data, blank_run = [], 0
            for row in rows:
                row = list(row[:width]) + [None] * (width - len(row))
                if all(v is None or str(v).strip() == "" for v in row):
                    blank_run += 1
                    if blank_run >= 1000:  # 서식만 남은 빈 행이 계속되면 중단
                        break
                    continue
                blank_run = 0
                data.append(["" if v is None else str(v) for v in row])
            yield name, pd.DataFrame(data, columns=header[:width])
    finally:
        wb.close()


def detect_source(columns) -> str | None:
    """열 이름을 보고 어느 DB에서 내보낸 파일인지 판단."""
    cols = {str(c).strip() for c in columns}
    best, best_hits = None, 0
    for name, p in SOURCE_PROFILES.items():
        if not any(c in cols for c in p["title"]):
            continue
        hits = sum(1 for c in p["mark"] if c in cols)
        if hits > best_hits:
            best, best_hits = name, hits
    return best


def _col(df: pd.DataFrame, names) -> pd.Series:
    for n in names:
        if n in df.columns:
            return df[n].fillna("").astype(str).str.strip()
    return pd.Series([""] * len(df), index=df.index, dtype=object)


def canonical_frame(df: pd.DataFrame, source: str, file_name: str) -> pd.DataFrame:
    """DB별 열을 공통 형식(Title / Abstract / Author Keywords / DOI / Year)으로 변환."""
    p = SOURCE_PROFILES[source]
    df = df.loc[:, ~df.columns.duplicated()]
    kws = _col(df, p["keywords"])
    if p["sep"] != ";":  # ProQuest는 쉼표 구분 → 세미콜론으로 통일
        kws = kws.map(lambda s: "; ".join(k.strip() for k in s.split(p["sep"]) if k.strip()))
    raw_types = [_col(df, (c,)) for c in p["doctype"]]
    out = pd.DataFrame({
        COL_TITLE: _col(df, p["title"]),
        COL_ABSTRACT: _col(df, p["abstract"]),
        COL_KEYWORDS: kws,
        COL_DOI: _col(df, p["doi"]).map(norm_doi),
        COL_YEAR: _col(df, p["year"]).str.extract(r"(\d{4})", expand=False).fillna(""),
        COL_SOURCE: source,
        COL_FILE: file_name,
        COL_CITED: _col(df, p.get("cited", ())).str.extract(r"(\d+)", expand=False).fillna(""),
    })
    out[COL_RAWTYPE] = [" , ".join(v for v in vals if v) for vals in zip(*raw_types)] if raw_types else ""
    out[COL_DOCTYPE] = out[COL_RAWTYPE].map(doc_kind)
    return out


def _completeness(row) -> int:
    """정보가 더 충실한 레코드를 고르기 위한 점수."""
    return 2 * bool(row[COL_ABSTRACT]) + 2 * bool(row[COL_KEYWORDS]) + bool(row[COL_DOI]) + bool(row[COL_YEAR])


def merge_records(frames: list[pd.DataFrame]):
    """DOI → 제목 순으로 같은 문헌을 묶고, 빈 칸은 같은 묶음의 다른 레코드로 채운다."""
    df = pd.concat(frames, ignore_index=True)
    df = df[(df[COL_TITLE].str.strip() != "") | (df[COL_ABSTRACT].str.strip() != "")].reset_index(drop=True)
    keys = df[COL_TITLE].map(title_key)
    scores = df.apply(_completeness, axis=1)

    groups: dict[int, list[int]] = {}
    by_doi: dict[str, int] = {}
    by_title: dict[str, list[int]] = {}
    reasons: dict[int, str] = {}

    for i in range(len(df)):
        doi, key, year = df.at[i, COL_DOI], keys.iat[i], df.at[i, COL_YEAR]
        gid, why = (by_doi.get(doi), "DOI 일치") if doi and doi in by_doi else (None, "")
        if gid is None and key:
            for cand in by_title.get(key, []):
                years = [df.at[j, COL_YEAR] for j in groups[cand] if df.at[j, COL_YEAR]]
                if not year or not years or min(abs(int(year) - int(y)) for y in years) <= 1:
                    gid, why = cand, "제목 일치"
                    break
        if gid is None:
            gid, why = i, ""
            groups[gid] = []
        groups[gid].append(i)
        if why:
            reasons[i] = why
        if doi:
            by_doi.setdefault(doi, gid)
        if key and gid not in by_title.setdefault(key, []):
            by_title[key].append(gid)

    merged, removed = [], []
    for idxs in groups.values():
        best = max(idxs, key=lambda j: (scores.iat[j], -j))
        row = df.loc[best].to_dict()
        for col in (COL_ABSTRACT, COL_KEYWORDS, COL_DOI, COL_YEAR):
            if not row[col]:
                for j in idxs:
                    if df.at[j, col]:
                        row[col] = df.at[j, col]
                        break
        row[COL_SOURCE] = " + ".join(sorted({df.at[j, COL_SOURCE] for j in idxs}))
        row[COL_FILE] = " + ".join(sorted({df.at[j, COL_FILE] for j in idxs}))
        merged.append(row)
        # 묶음 전체의 판정 근거 (먼저 들어온 레코드에는 근거가 기록되지 않으므로 묶음 기준으로 표시)
        group_why = ("DOI 일치" if any(reasons.get(j) == "DOI 일치" for j in idxs)
                     else "제목 일치" if any(reasons.get(j) == "제목 일치" for j in idxs) else "중복")
        for j in idxs:
            if j != best:
                removed.append({
                    "제목": df.at[j, COL_TITLE][:150], "DOI": df.at[j, COL_DOI] or "(없음)",
                    "빠진 레코드의 출처": df.at[j, COL_SOURCE], "원본파일": df.at[j, COL_FILE],
                    "판정 근거": reasons.get(j) or group_why, "남긴 레코드의 출처": df.at[best, COL_SOURCE],
                })
    return pd.DataFrame(merged).reset_index(drop=True), removed


def load_records(paths: list[str], sheet: str | None = None, years: list[int] | None = None,
                 doc_types: str = "journal"):
    """여러 DB의 내보내기 파일을 읽어 중복을 제거한 통합 목록과 처리 보고를 돌려준다."""
    frames, file_rows = [], []
    for p in map(Path, paths):
        if not p.exists():
            raise SystemExit(f"파일이 없습니다: {p}")
        log(f"[1/7] 데이터 읽는 중: {p.name}")
        matched = False
        for sheet_name, raw in _sheet_frames(p, sheet):
            source = detect_source(raw.columns)
            if not source:
                continue
            frame = canonical_frame(raw, source, p.name)
            frames.append(frame)
            file_rows.append({"원본파일": p.name, "DB": source, "시트": sheet_name, "원본 건수": len(frame)})
            log(f"      {source} 형식으로 읽음 ({len(frame):,}건)")
            matched = True
            break
        if not matched:
            raise SystemExit(
                f"{p.name}: Scopus·Web of Science·ProQuest 중 어느 형식인지 알 수 없습니다. "
                f"제목 열 이름이 Title / Article Title 중 하나인지 확인하세요.")
    if not frames:
        raise SystemExit("읽을 수 있는 파일이 없습니다.")

    merged, removed = merge_records(frames)
    raw_total = sum(r["원본 건수"] for r in file_rows)
    dup_doi = sum(1 for r in removed if r["판정 근거"] == "DOI 일치")
    dup_title = sum(1 for r in removed if r["판정 근거"] == "제목 일치")
    after_merge = len(merged)

    if years:
        wanted = {str(y) for y in years}
        merged = merged[merged[COL_YEAR].isin(wanted)]
    after_year = len(merged)

    type_counts = dict(Counter(merged[COL_DOCTYPE]))
    if doc_types == "journal":
        merged = merged[merged[COL_DOCTYPE] != "기타"]  # 학위논문·도서·회의록 등 제외
    after_type = len(merged)

    log(f"      통합 {raw_total:,}건 → 중복 {len(removed):,}건 제거 → {after_merge:,}건"
        + (f" → 연도 필터 {after_year:,}건" if years else "")
        + (f" → 저널 논문만 {after_type:,}건" if doc_types == "journal" else ""))
    report = {
        "files": file_rows, "raw_total": raw_total, "dup_doi": dup_doi, "dup_title": dup_title,
        "dup_total": len(removed), "after_merge": after_merge, "after_year": after_year,
        "after_type": after_type, "removed": removed, "type_counts": type_counts,
        "by_source": dict(Counter(merged[COL_SOURCE])), "doc_types": doc_types,
    }
    return merged.reset_index(drop=True), report


# ─────────────────────────────────────────────────────────────────────────────
# 2. 텍스트 전처리
# ─────────────────────────────────────────────────────────────────────────────
HTML_RE = re.compile(r"<[^>]+>")
COPYRIGHT_RE = re.compile(r"(?:©|\bCopyright\b\s*(?:©\s*)?\(?\d{4}|\(C\)\s*\d{4}).*$", re.S | re.I)
CITATION_RE = re.compile(r"\((?:[A-Z][A-Za-z.&]*\s){1,6}\d{4};[^)]*\)")
STAT_PAREN_RE = re.compile(
    r"\s*\((?=[^()]*(?:[=<>%±≤≥]|\d\s*[,;–-]\s*\d|\d\.\d|\b(?:CI|OR|HR|IQR|SD|RR|aOR|aHR|vs)\b))[^()]*\)"
)
HEADING_RE = re.compile(r"(?:^|(?<=[.!?;]))\s*([A-Z][A-Za-z /&,\-]{0,60}?)\s*:(?!\d)")
HEADING_WORDS = set("""
background backgrounds aim aims objective objectives purpose purposes introduction method methods
methodology material materials result results conclusion conclusions discussion finding findings
design setting settings participant participants patient patients measure measures main outcome
outcomes intervention interventions exposure exposures evidence summary interpretation significance
importance context rationale trial registration identifier funding limitation limitations
implication implications practical social originality value research study hypothesis and of the
clinical relevance key point points question questions highlights statement problem approach
observation observations case presentation description data source sources selection extraction
synthesis eligibility criteria review level graphical abstract number new what is known this adds
subjects analysis statistical experimental procedures principal major brief scope perspective
perspectives future direction directions novelty impact recommendations benefits
""".split())
EMPTY_ABSTRACT = "[no abstract available]"


def strip_headings(text: str) -> str:
    def repl(m):
        words = re.findall(r"[A-Za-z]+", m.group(1).lower())
        return " " if words and all(w in HEADING_WORDS for w in words) else m.group(0)

    return HEADING_RE.sub(repl, text)


def clean_abstract(text: str) -> str:
    if text.strip().lower() == EMPTY_ABSTRACT:
        return ""
    text = HTML_RE.sub("", text)
    text = COPYRIGHT_RE.sub("", text)
    text = CITATION_RE.sub("", text)
    return strip_headings(text)


# ── 약어 정의 탐지 (Schwartz & Hearst 2003 알고리즘) ──
SHORT_FORM_RE = re.compile(r"\(\s*([A-Za-z][\w\-+/&.]{1,11})\s*\)")
LONG_FORM_STOP = {"the", "a", "an", "of", "and", "in", "on", "for", "to", "with", "by", "as", "or"}


def _long_form(short: str, text: str) -> str | None:
    s, l = len(short) - 1, len(text) - 1
    while s >= 0:
        c = short[s].lower()
        if not c.isalnum():
            s -= 1
            continue
        while l >= 0 and (text[l].lower() != c or (s == 0 and l > 0 and text[l - 1].isalnum())):
            l -= 1
        if l < 0:
            return None
        l -= 1
        s -= 1
    start = text.rfind(" ", 0, l + 1) + 1
    return text[start:].strip()


def find_acronyms(texts) -> dict[str, str]:
    amap: dict[str, str] = {}
    for text in texts:
        for m in SHORT_FORM_RE.finditer(text):
            short = m.group(1).rstrip(".")
            if short.lower() == short or len(short) < 2:
                continue
            before = text[: m.start()]
            before = before[max(before.rfind(ch) for ch in ".;:!?()[],") + 1 :]
            words = before.split()
            cand = " ".join(words[-min(len(short) + 5, len(short) * 2) :])
            long = _long_form(short, cand)
            if not long:
                continue
            lw = long.split()
            while lw and lw[0].lower() in LONG_FORM_STOP:
                lw.pop(0)
            long = " ".join(lw)
            if not long or len(long) <= len(short) or long.lower() == short.lower():
                continue
            amap.setdefault(short, long)
            if short.endswith("s") and short[:-1] != short[:-1].lower():
                amap.setdefault(short[:-1], long)
            else:
                amap.setdefault(short + "s", long)
    return amap


def apply_acronyms(text: str, amap: dict[str, str]) -> str:
    if not amap or not text:
        return text
    alt = "|".join(re.escape(k) for k in sorted(amap, key=len, reverse=True))
    text = re.sub(r"\s*\(\s*(?:" + alt + r")\s*\)", "", text)
    return re.sub(r"(?<![\w\-])(" + alt + r")(?![\w\-])", lambda m: amap[m.group(1)], text)


def preprocess(df: pd.DataFrame) -> list[dict]:
    docs = []
    for title, abstract, keywords in zip(df[COL_TITLE], df[COL_ABSTRACT], df[COL_KEYWORDS]):
        title = HTML_RE.sub("", title)
        abstract = clean_abstract(abstract)
        kws = [HTML_RE.sub("", k).strip() for k in keywords.split(";") if k.strip()]
        amap = find_acronyms([title, abstract, *kws])
        title = STAT_PAREN_RE.sub("", apply_acronyms(title, amap))
        abstract = STAT_PAREN_RE.sub("", apply_acronyms(abstract, amap))
        kws = [re.sub(r"\s+\([^()]*\)", "", apply_acronyms(k, amap)).strip(" .,") for k in kws]
        docs.append({
            "title": " ".join(title.split()),
            "abstract": " ".join(abstract.split()),
            "keywords": [k for k in kws if k],
        })
    return docs


# ─────────────────────────────────────────────────────────────────────────────
# 3. 품사 태깅 (spaCy) → 단위(unit) 열로 변환
#    unit = (표면형, 분류)  분류: N=명사  A=형용사  O=기타 단어  B=구두점(경계)
# ─────────────────────────────────────────────────────────────────────────────
HYPHENS = {"-", "‐", "‑", "–"}
POSSESSIVE_TOKENS = {"'s", "’s", "'", "’"}


def _token_class(tok) -> str:
    if tok.is_space or tok.is_punct or tok.pos_ in ("PUNCT", "SYM", "SPACE", "X") and not tok.is_alpha:
        return "B"
    if tok.pos_ in ("NOUN", "PROPN"):
        return "N"
    if tok.pos_ == "ADJ":
        return "A"
    return "O"


def doc_to_units(doc) -> list[tuple[str, str]]:
    toks = list(doc)
    units: list[list[str]] = []
    i = 0
    while i < len(toks):
        t = toks[i]
        if t.text in POSSESSIVE_TOKENS:
            i += 1
            continue
        if (t.text in HYPHENS and units and units[-1][1] != "B" and i + 1 < len(toks)
                and not toks[i - 1].whitespace_ and not t.whitespace_ and re.search(r"\w", toks[i + 1].text)):
            nxt_cls = _token_class(toks[i + 1])
            prev_cls = units[-1][1]
            units[-1][0] += "-" + toks[i + 1].text
            units[-1][1] = nxt_cls if nxt_cls in "NA" else (prev_cls if prev_cls in "NA" else "O")
            i += 2
            continue
        units.append([t.text, _token_class(t)])
        i += 1
    return [(s, c) for s, c in units]


def tag_documents(docs: list[dict], processes: int, use_cache: bool, progress=None) -> None:
    import spacy

    try:
        nlp = spacy.load(SPACY_MODEL, exclude=["parser", "ner", "lemmatizer", "senter"])
    except OSError:
        raise SystemExit(f"spaCy 모델이 없습니다. 먼저 실행하세요:  python -m spacy download {SPACY_MODEL}")

    h = hashlib.sha1(f"{CACHE_VERSION}|{SPACY_MODEL}|{nlp.meta.get('version')}".encode())
    for d in docs:
        h.update(d["title"].encode("utf-8", "replace") + b"\x00" + d["abstract"].encode("utf-8", "replace") + b"\x01")
    cache_file = CACHE_DIR / f"tags_{h.hexdigest()[:20]}.pkl"

    if use_cache and cache_file.exists():
        log("[3/7] 품사 태깅: 캐시 사용")
        cached = pickle.loads(cache_file.read_bytes())
    else:
        texts = [t for d in docs for t in (d["title"], d["abstract"])]
        log(f"[3/7] 품사 태깅 중 (spaCy, {len(texts):,}개 텍스트) — 수 분 걸릴 수 있습니다")
        stream = nlp.pipe(texts, batch_size=64, n_process=processes)
        if progress is None:
            from tqdm import tqdm

            stream = tqdm(stream, total=len(texts), unit="text", ncols=80)
        cached = []
        for n, doc in enumerate(stream, 1):
            cached.append(doc_to_units(doc))
            if progress and (n % 200 == 0 or n == len(texts)):
                progress(n, len(texts))
        CACHE_DIR.mkdir(exist_ok=True)
        cache_file.write_bytes(pickle.dumps(cached, protocol=pickle.HIGHEST_PROTOCOL))

    for i, d in enumerate(docs):
        d["title_units"], d["abstract_units"] = cached[2 * i], cached[2 * i + 1]


# ─────────────────────────────────────────────────────────────────────────────
# 4. 정규화 (단수형 · 소문자 · 하이픈/소유격 통일)
# ─────────────────────────────────────────────────────────────────────────────
POSSESSIVE_RE = re.compile(r"(?<=\w)['’]s\b")
WORD_RE = re.compile(r"\d+(?:\.\d+)+|[^\W_]+")


def split_words(s: str) -> list[str]:
    return WORD_RE.findall(POSSESSIVE_RE.sub("", s))


class Normalizer:
    """단어를 단수형 소문자 키로 변환. 사전에 없는 단어는 코퍼스에 단수형이 실제 등장할 때만 변환."""

    def __init__(self, vocab: Counter):
        from lemminflect import getLemma

        self.getLemma, self.vocab, self.cache = getLemma, vocab, {}

    def word(self, w: str) -> str:
        r = self.cache.get(w)
        if r is None:
            r = self.cache[w] = self._word(w)
        return r

    def _word(self, w: str) -> str:
        stem = w[:-1]
        if len(w) >= 3 and w[-1] == "s" and stem == stem.upper() and sum(c.isupper() for c in stem) >= 2:
            return stem.lower()  # LLMs → llm, MSCs → msc
        lw = w.lower()
        if (len(w) >= 2 and w.isupper()) or not lw.isalpha() or len(lw) <= 3:
            return lw  # AIDS, SARS 같은 약어는 그대로
        cands = self.getLemma(lw, upos="NOUN", lemmatize_oov=False)
        if cands:
            return lw if lw in cands else max(cands, key=lambda c: self.vocab.get(c, 0))
        if lw.endswith("s") and not lw.endswith(("ss", "us", "is")):
            oov = self.getLemma(lw, upos="NOUN", lemmatize_oov=True)
            cand = oov[0].lower() if oov else lw
            if cand != lw and self.vocab.get(cand, 0) > 0:
                return cand
        return lw

    def key(self, s: str) -> str:
        return " ".join(self.word(w) for w in split_words(s))


def raw_key(s: str) -> str:
    return " ".join(w.lower() for w in split_words(s))


# ─────────────────────────────────────────────────────────────────────────────
# 5. 설정 파일
# ─────────────────────────────────────────────────────────────────────────────
CONFIG_NAMES = ("stopwords.txt", "modifiers.txt", "trailing_generic.txt",
                "exclude_phrases.txt", "custom_phrases.txt", "synonyms.txt")


def code_version() -> str:
    """결과를 어떤 코드로 만들었는지 기록하기 위한 버전 (git 커밋)."""
    import subprocess

    try:
        out = subprocess.run(["git", "-C", str(BASE_DIR), "rev-parse", "--short", "HEAD"],
                             capture_output=True, text=True, timeout=5)
        if out.returncode == 0 and out.stdout.strip():
            dirty = subprocess.run(["git", "-C", str(BASE_DIR), "status", "--porcelain"],
                                   capture_output=True, text=True, timeout=5)
            mark = " (수정본)" if dirty.stdout.strip() else ""
            return out.stdout.strip() + mark
    except Exception:  # noqa: BLE001  (git이 없거나 저장소가 아니어도 분석은 계속)
        pass
    return "(알 수 없음)"


def default_settings() -> dict[str, str]:
    """프로그램에 들어 있는 기본 설정 파일 내용 (사용자별 설정의 출발점)."""
    out = {}
    for name in CONFIG_NAMES:
        path = CONFIG_DIR / name
        out[name] = path.read_text(encoding="utf-8-sig") if path.exists() else ""
    return out


def settings_diff(word_settings: dict) -> list[dict]:
    """기본값과 달라진 줄만 뽑는다 (결과 파일에 남길 재현 정보)."""
    base = default_settings()
    rows = []
    for name in CONFIG_NAMES:
        def lines(text):
            return [ln.strip() for ln in (text or "").splitlines()
                    if ln.strip() and not ln.lstrip().startswith("#")]

        before, after = lines(base.get(name, "")), lines(word_settings.get(name, ""))
        for line in after:
            if line not in before:
                rows.append({"설정 파일": name, "변경": "추가", "줄": line})
        for line in before:
            if line not in after:
                rows.append({"설정 파일": name, "변경": "삭제", "줄": line})
    return rows


def read_config_lines(name: str, settings: dict | None = None) -> list[str]:
    """설정 한 파일의 유효한 줄 목록. settings가 있으면 파일 대신 그 내용을 쓴다."""
    if settings is not None and name in settings:
        text = settings[name] or ""
    else:
        path = CONFIG_DIR / name
        if not path.exists():
            return []
        text = path.read_text(encoding="utf-8-sig")
    lines = []
    for line in text.splitlines():
        line = re.split(r"\s#|^#", line, maxsplit=1)[0].strip()
        if line:
            lines.append(line)
    return lines


class Config:
    def __init__(self, norm: Normalizer, settings: dict | None = None):
        from spacy.lang.en.stop_words import STOP_WORDS

        k = norm.key
        self.function_words = {w.lower() for w in STOP_WORDS}
        self.generic = {k(x) for x in read_config_lines("stopwords.txt", settings)}
        self.modifiers = {k(x) for x in read_config_lines("modifiers.txt", settings)}
        self.trailing = {k(x) for x in read_config_lines("trailing_generic.txt", settings)}
        self.exclude = {k(x) for x in read_config_lines("exclude_phrases.txt", settings)}
        self.custom = {k(x) for x in read_config_lines("custom_phrases.txt", settings)}
        self.synonyms = {}
        for line in read_config_lines("synonyms.txt", settings):
            if "=>" in line:
                left, right = line.split("=>", 1)
                for variant in left.split("|"):
                    if k(variant) and k(right):
                        self.synonyms[k(variant)] = k(right)
        self.weak = self.generic | self.modifiers | self.function_words

    def accept(self, key: str) -> bool:
        if not key or key in self.exclude:
            return False
        words = key.split()
        if len(words) == 1:
            w = words[0]
            return (w not in self.weak and len(w) >= 2 and any(c.isalpha() for c in w))
        return not all(w in self.weak or not any(c.isalpha() for c in w) for w in words)


# ─────────────────────────────────────────────────────────────────────────────
# 6. 구문 사전 구축 & 집계
# ─────────────────────────────────────────────────────────────────────────────
def unit_runs(keys, classes, cfg: Config):
    """형용사·명사가 연속된 구간(명사구 후보)의 인덱스 목록."""
    run = []
    for i, (k, c) in enumerate(zip(keys, classes)):
        if c in "NA" and k and k not in cfg.function_words and any(ch.isalpha() for ch in k):
            run.append(i)
        else:
            if len(run) >= 2:
                yield run
            run = []
    if len(run) >= 2:
        yield run


def build_phrase_dict(docs, cfg: Config, min_papers: int, max_units: int):
    ngram_papers = Counter()
    for d in docs:
        seen = set()
        for sec in ("title", "abstract"):
            keys, classes = d[f"{sec}_keys"], d[f"{sec}_classes"]
            for run in unit_runs(keys, classes, cfg):
                for a in range(len(run)):
                    if keys[run[a]] in cfg.modifiers:
                        continue
                    for b in range(a + 2, min(a + max_units, len(run)) + 1):
                        last = run[b - 1]
                        if classes[last] != "N" or keys[last] in cfg.trailing or keys[last] in cfg.modifiers:
                            continue
                        seen.add(" ".join(keys[j] for j in run[a:b]))
        ngram_papers.update(seen)

    auto = {p for p, n in ngram_papers.items() if n >= min_papers and cfg.accept(p)}
    from_keywords = {k for d in docs for k in d["keyword_keys"] if 2 <= len(k.split()) <= 8}
    phrases = (auto | from_keywords | cfg.custom) - cfg.exclude
    log(f"      구문 사전: 자동추출 {len(auto):,} + 저자키워드 {len(from_keywords):,} + 사용자 {len(cfg.custom):,}"
          f" → 총 {len(phrases):,}개")
    return phrases


def match_units(keys, classes, surfaces, phrases, first_words, max_words):
    """가장 긴 구문 우선 매칭. 사전에 없는 명사는 단어로 반환."""
    i, n = 0, len(keys)
    while i < n:
        if classes[i] == "B" or not keys[i]:
            i += 1
            continue
        best = None
        if keys[i].split(" ", 1)[0] in first_words:
            joined, words = "", 0
            for j in range(i, n):
                if classes[j] == "B" or not keys[j]:
                    break
                joined = keys[j] if j == i else f"{joined} {keys[j]}"
                words += keys[j].count(" ") + 1
                if words > max_words:
                    break
                if (j > i or " " in joined) and joined in phrases:
                    best = (j, joined)
        if best:
            j, key = best
            yield key, " ".join(surfaces[i:j + 1])
            i = j + 1
            continue
        if classes[i] == "N":
            yield keys[i], surfaces[i]
        i += 1


def count_phrases(docs, cfg: Config, phrases, sections):
    """집계 결과와 함께, 논문별 키프레이즈 목록(토픽 모델 입력)을 돌려준다."""
    first_words = {p.split(" ", 1)[0] for p in phrases}
    max_words = max((len(p.split()) for p in phrases), default=1)
    stats = defaultdict(lambda: {"title": 0, "abstract": 0, "keywords": 0, "papers": 0, "kw_papers": 0})
    forms = defaultdict(Counter)
    doc_keys: list[Counter] = []
    doc_kw: list[set] = []        # 논문별로 '저자키워드로 쓴' 항목 (임계값 검증용)

    for d in docs:
        doc_seen, doc_kw_seen = set(), set()
        items = []
        for sec in ("title", "abstract"):
            if sec in sections:
                items += [(sec, k, s) for k, s in match_units(
                    d[f"{sec}_keys"], d[f"{sec}_classes"], d[f"{sec}_surfaces"], phrases, first_words, max_words)]
        if "keywords" in sections:
            items += [("keywords", k, s) for k, s in zip(d["keyword_keys"], d["keywords"])]
        here = Counter()
        for sec, key, surface in items:
            key = cfg.synonyms.get(key, key)
            if not cfg.accept(key):
                continue
            stats[key][sec] += 1
            forms[key][surface] += 1
            here[key] += 1
            doc_seen.add(key)
            if sec == "keywords":
                doc_kw_seen.add(key)
        for key in doc_seen:
            stats[key]["papers"] += 1
        for key in doc_kw_seen:
            stats[key]["kw_papers"] += 1
        doc_keys.append(here)
        doc_kw.append(set(doc_kw_seen))
    return stats, forms, doc_keys, doc_kw


# ── 자동 정리 제안 ───────────────────────────────────────────────────────────
# 업로드한 파일 내용만 보고 '합칠 표기 · 색인 태그 · 검색어' 후보를 뽑는다.
# 분야마다 정답이 달라 자동으로 지우지 않고, 근거와 함께 보여 주고 사용자가 고르게 한다.
SUFFIX_RE = re.compile(r"(lling|ling|ing|lled|led|ed|isations|izations|isation|ization|ions|ion|s)$")
QUERY_TERM_RATE = 0.40   # 이 비율 이상 논문에 나오면 '검색어 자체'로 본다
TAG_MIN_KEYWORD = 10     # 색인 태그 판정: 저자키워드 등장 최소 횟수
TAG_BODY_RATIO = 0.1     # 본문 등장이 저자키워드 등장의 이 비율 이하면 색인 태그로 본다


def _stem_key(key: str) -> str:
    return " ".join(SUFFIX_RE.sub("", w) for w in key.split())


def suggest_cleanup(table: pd.DataFrame, stats: dict, n_docs: int) -> dict:
    """합칠 표기 / 색인 태그 / 검색어 후보를 근거와 함께 돌려준다."""
    groups = defaultdict(list)
    merges, tags, queries = [], [], []
    for kp, key, papers in zip(table["키프레이즈"], table["정규화 키"], table["논문수"]):
        groups[_stem_key(key)].append((kp, key, int(papers)))
        s = stats.get(key)
        if not s:
            continue
        body = s["title"] + s["abstract"]
        if s["keywords"] >= TAG_MIN_KEYWORD and body <= TAG_BODY_RATIO * s["keywords"]:
            tags.append({"키프레이즈": kp, "저자키워드 등장": s["keywords"], "제목·초록 등장": body,
                         "근거": "저자키워드에만 나오고 본문에는 거의 없음 → DB 색인 태그로 보임"})
        if n_docs and papers / n_docs >= QUERY_TERM_RATE:
            queries.append({"키프레이즈": kp, "등장 논문 비율": f"{100 * papers / n_docs:.1f}%", "논문수": int(papers),
                            "근거": "대부분의 논문에 나와 변별력이 없음 → 검색어 자체로 보임"})
    for members in groups.values():
        if len(members) < 2:
            continue
        members.sort(key=lambda x: -x[2])
        merges.append({"대표 표기": members[0][0], "합칠 표기": ", ".join(m[0] for m in members[1:]),
                       "합친 논문수": sum(m[2] for m in members), "근거": "어미만 다른 같은 말"})
    merges.sort(key=lambda m: -m["합친 논문수"])
    tags.sort(key=lambda t: -t["저자키워드 등장"])
    queries.sort(key=lambda q: -q["논문수"])
    return {"merges": merges, "tags": tags, "queries": queries}


# ── 분석 전 정리 추천 ────────────────────────────────────────────────────────
# 올린 파일의 저자키워드만 보고 "이건 빼는 게 좋겠다"를 근거와 함께 제안한다.
# 품사 태깅 전이라 수 초면 끝난다. 분야마다 정답이 달라 자동으로 지우지 않는다.
#
# 아래 목록은 DB가 자동으로 붙이는 색인어 카탈로그다. 기본 제외 목록에 넣지 않고
# '추천'에만 쓴다 — 성별·연령·공학처럼 분야에 따라 핵심 주제어일 수 있기 때문이다.
INDEX_VOCAB = {
    # PsycINFO·ERIC 방법·인구 태그
    "human", "humans", "male", "female", "animal", "empirical study", "quantitative study",
    "qualitative study", "literature review", "longitudinal study", "followup study",
    "experimental replication", "adulthood", "childhood", "adolescence", "school age",
    "preschool age", "infancy", "neonatal", "young adulthood", "thirties", "middle age", "aged",
    # 문헌 유형이 키워드 열에 섞인 경우
    "article", "journal article", "peer reviewed journal", "dissertation", "thesis", "report",
    # 연구분야 분류 라벨 (ANZSRC 등)
    "information and computing sciences", "medical and health sciences", "mathematical sciences",
    "psychology and cognitive sciences", "clinical sciences", "biological sciences",
    "physical sciences", "built environment and design", "studies in human society",
    "applied economics", "public health and health services",
    "artificial intelligence and image processing", "simulation and modeling",
}
STAT_VOCAB = {
    # 통계·방법 보고 용어 — 주제어가 아니라 '어떻게 분석했는지'를 가리킨다
    "confidence interval", "odds ratio", "adjusted odds ratio", "hazard ratio", "adjusted hazard ratio",
    "relative risk", "risk ratio", "standard deviation", "standard error", "interquartile range",
    "p value", "statistical significance", "statistical analysis", "regression analysis",
    "multivariate analysis", "multivariable analysis", "logistic regression analysis",
    "area under the curve",
}

NARRATIVE_VOCAB = {
    # 논문 서술 상투구 — 어느 분야에서나 나오고 주제를 구분하지 못한다
    "present study", "current study", "previous study", "this study", "further study", "future study",
    "future research", "recent study", "control group", "treatment group", "study population",
    "study design", "study period", "study group", "clinical study", "large number", "wide range",
    "high level", "low level", "significant difference", "key role", "important role", "crucial role",
    "recent year", "long term", "short term",
}

DEMO_TAG_VOCAB = {
    # 인구 집단 색인어 — 분야에 따라 주제어일 수 있어 추천으로만 제시
    "aged", "middle aged", "adult", "human", "female", "male", "animal", "man", "woman", "year old",
}
AGE_BAND_RE = re.compile(r"\(\s*(?:birth|\d+)[^)]*(?:yrs|years|mo)\b[^)]*\)", re.I)
JEL_CODE_RE = re.compile(r"^[A-Z]\d{2}$")
FIELD_LABEL_RE = re.compile(r"\b(Sciences|Studies|Technologies)\s*$")
PRE_QUERY_RATE = 0.40


def _quick_normalizer(df: pd.DataFrame) -> "Normalizer":
    """품사 태깅 없이 표기 통일 규칙만 쓰기 위한 정규화기 (분석 전 추천용).
    본 분석과 같은 Normalizer를 쓰므로 '자동으로 합쳐지는지'를 정확히 판정할 수 있다."""
    vocab = Counter()
    for col in (COL_TITLE, COL_ABSTRACT, COL_KEYWORDS):
        if col not in df.columns:      # 키워드만 넘어오는 경우도 있다
            continue
        for text in df[col]:
            vocab.update(w.lower() for w in split_words(str(text)))
    return Normalizer(vocab)


def suggest_before_analysis(df: pd.DataFrame) -> dict:
    """통합된 목록의 저자키워드를 보고 제외·통합 후보를 추천한다 (분석 전)."""
    n = max(1, len(df))
    papers = Counter()
    display = {}
    raw_papers = Counter()        # 원래 표기 그대로의 등장 논문 수 (자동 통합 내역용)
    for cell in df[COL_KEYWORDS]:
        seen, raw_seen = set(), set()
        for raw in str(cell).split(";"):
            kw = raw.strip()
            if not kw:
                continue
            key = kw.lower()
            display.setdefault(key, kw)
            seen.add(key)
            raw_seen.add(kw)
        papers.update(seen)
        raw_papers.update(raw_seen)

    tags, queries = [], []
    for key, count in papers.most_common():
        label = display[key]
        if count / n >= PRE_QUERY_RATE:
            queries.append({"키워드": label, "등장 논문": count, "비율": f"{100 * count / n:.1f}%",
                            "근거": "논문 대부분에 달려 있어 변별력이 없음 (검색어로 보임)"})
            continue
        why = ""
        if key in INDEX_VOCAB or key in DEMO_TAG_VOCAB:
            why = "DB가 자동으로 붙이는 색인어 (분야에 따라 주제어일 수 있으니 확인 후 선택)"
        elif key in STAT_VOCAB:
            why = "통계·방법 보고 용어 (주제어가 아님)"
        elif key in NARRATIVE_VOCAB:
            why = "어느 논문에나 나오는 서술 상투구"
        elif AGE_BAND_RE.search(label):
            why = "연령 구간 태그 (PsycINFO 등)"
        elif JEL_CODE_RE.match(label):
            why = "분류 코드 (JEL 등)"
        elif FIELD_LABEL_RE.search(label):
            why = "연구분야 분류 라벨"
        if why and count >= 3:
            tags.append({"키워드": label, "등장 논문": count, "근거": why})

    norm = _quick_normalizer(df)

    def rough(text: str) -> str:
        """본 분석이 실제로 합치는 형태 (대소문자·하이픈·소유격·단수형)."""
        return norm.key(text)

    groups = defaultdict(dict)
    for key, count in papers.items():
        if count >= 2:
            bucket = groups[_stem_key(raw_key(key))]
            same = rough(key)                     # 이미 자동으로 합쳐지는 표기는 하나로 묶어 둔다
            if same not in bucket or bucket[same][1] < count:
                bucket[same] = (display[key], count)
    merges = []
    for bucket in groups.values():
        members = list(bucket.values())
        if len(members) < 2:                      # 자동 통합만으로 끝나는 묶음은 추천하지 않는다
            continue
        members.sort(key=lambda x: -x[1])
        merges.append({"대표 표기": members[0][0], "합칠 표기": ", ".join(m[0] for m in members[1:]),
                       "합친 논문": sum(m[1] for m in members), "근거": "철자·어미만 다른 같은 말"})
    merges.sort(key=lambda m: -m["합친 논문"])

    # 프로그램이 알아서 합치는 표기 (대소문자·하이픈·소유격·단수형) — 확인용, 손댈 필요 없음
    auto_groups = defaultdict(Counter)
    for raw, count in raw_papers.items():
        auto_groups[norm.key(raw)][raw] = count
    auto = []
    for key, variants in auto_groups.items():
        if len(variants) < 2:
            continue
        ordered = variants.most_common()
        auto.append({"대표 표기": ordered[0][0], "합쳐진 표기": ", ".join(v for v, _ in ordered[1:9]),
                     "표기 수": len(variants), "합친 논문": sum(variants.values())})
    auto.sort(key=lambda a: -a["합친 논문"])
    return {"tags": tags[:40], "queries": queries[:10], "merges": merges[:30], "auto": auto[:60]}


# ── 임계값 민감도 ────────────────────────────────────────────────────────────
# "필터 기준을 바꾸면 결과가 얼마나 달라지는가"를 표로 만든다.
# 논문 방법론에 "임계값 5%를 썼고, 2~10% 사이에서 상위 결과는 거의 같았다"처럼 쓰기 위한 것.
SENSITIVITY_RATIOS = (0.0, 0.02, 0.05, 0.10, 0.20)
SENSITIVITY_PHRASE_PAPERS = (2, 3, 5)


QUICK_MAX_WORDS = 4      # 추정에 쓸 최대 구문 길이
QUICK_MIN_BODY = 3       # 제목·초록 표현은 이만큼의 논문에 나와야 후보로 본다


def quick_doc_stats(df: pd.DataFrame, settings: dict | None = None, max_terms: int = 6000):
    """품사 태깅 없이 논문별 (등장 표현, 저자키워드) 집합을 추정한다.
    후보 = 저자키워드(2편 이상) + 제목·초록에서 3편 이상 반복되는 표현.
    본문 전용 표현까지 후보에 넣어야 '일반어 필터가 무엇을 거르는지' 평가할 수 있다.
    본 분석은 품사와 '가장 긴 구문 우선'을 쓰므로 수치는 조금 다르다 (추정치)."""
    from spacy.lang.en.stop_words import STOP_WORDS

    norm = _quick_normalizer(df)
    weak = {w.lower() for w in STOP_WORDS}
    weak |= {norm.key(x) for x in read_config_lines("stopwords.txt", settings)}
    weak |= {norm.key(x) for x in read_config_lines("modifiers.txt", settings)}
    texts = (df[COL_TITLE].fillna("").astype(str) + ". "
             + df[COL_ABSTRACT].fillna("").astype(str)).tolist()

    def grams(text: str) -> set:
        """제목·초록에서 뽑은 표현 후보 (기능어로 시작·끝나지 않는 1~4단어)."""
        words = [norm.word(w) for w in split_words(text)]
        out = set()
        for size in range(1, QUICK_MAX_WORDS + 1):
            for i in range(len(words) - size + 1):
                chunk = words[i:i + size]
                if chunk[0] in weak or chunk[-1] in weak:
                    continue
                out.add(" ".join(chunk))
        return out

    kw_sets, kw_papers, body_papers = [], Counter(), Counter()
    for cell, text in zip(df[COL_KEYWORDS], texts):
        keys = {k for k in (norm.key(x) for x in str(cell).split(";") if x.strip()) if k}
        kw_sets.append(keys)
        kw_papers.update(keys)
        body_papers.update(grams(text))

    candidates = {k for k, c in kw_papers.items() if c >= 2}
    candidates |= {k for k, c in body_papers.items() if c >= QUICK_MIN_BODY}
    if len(candidates) > max_terms:        # 너무 많으면 자주 나오는 것만
        ranked = sorted(candidates, key=lambda k: -(body_papers[k] + kw_papers[k]))
        candidates = set(ranked[:max_terms])
    if not candidates:
        return [], []

    doc_keys, doc_kw = [], []
    for text, keys in zip(texts, kw_sets):
        here = (grams(text) | keys) & candidates
        doc_keys.append(Counter(here))
        doc_kw.append(keys & candidates)
    return doc_keys, doc_kw


def recommend_ratio_quick(df: pd.DataFrame, args=None, settings: dict | None = None):
    """분석 전 임계값 추천 (추정치). 반환: (비교표, 추천값)"""
    doc_keys, doc_kw = quick_doc_stats(df, settings)
    if not doc_keys:
        return pd.DataFrame(), None
    if args is None:
        args = parse_args(["_"])
    return recommend_ratio(doc_keys, doc_kw, args)


def _half_stats(doc_keys, doc_kw, index):
    """고른 논문들만으로 '등장 논문 수'와 '저자키워드로 쓴 논문 수'를 다시 센다."""
    papers, kw_papers = Counter(), Counter()
    for i in index:
        papers.update(doc_keys[i].keys())
        kw_papers.update(doc_kw[i])
    return papers, kw_papers


def recommend_ratio(doc_keys, doc_kw, args, ratios=SENSITIVITY_RATIOS, splits: int = 5, seed: int = 42):
    """일반어 필터 임계값을 데이터로 고른다.

    논문을 무작위 절반(A·B)으로 나눠 **A만 보고** 필터를 적용하고, 그 결과를 **B의 저자키워드**로 채점한다.
    - 정확도 = 남긴 항목 중 B에서도 저자들이 주제어로 쓴 비율
    - 재현율 = B에서 주제어로 쓰인 항목 중 남긴 비율
    F1이 가장 높은 값을 추천한다. 고르는 쪽과 채점하는 쪽이 달라 '자기 기준으로 자기를 채점'하는 문제를 피한다.
    """
    import numpy as np

    rng = np.random.default_rng(seed)
    n = len(doc_keys)
    scores = {r: {"정확도": [], "재현율": [], "F1": [], "키프레이즈 수": []} for r in ratios}
    for _ in range(splits):
        order = rng.permutation(n)
        a_idx, b_idx = order[: n // 2], order[n // 2:]
        papers_a, kw_a = _half_stats(doc_keys, doc_kw, a_idx)
        _, kw_b = _half_stats(doc_keys, doc_kw, b_idx)
        truth = {k for k, c in kw_b.items() if c >= 2}          # B에서 실제로 쓰인 주제어
        pool = {k for k, c in papers_a.items() if c >= args.min_papers}
        for ratio in ratios:
            kw_min = 0 if ratio == 0 else args.min_keyword_papers
            kept = {k for k in pool
                    if kw_a[k] >= kw_min and (papers_a[k] == 0 or kw_a[k] / papers_a[k] >= ratio)}
            tp = len(kept & truth)
            precision = tp / len(kept) if kept else 0.0
            recall = tp / len(truth & pool) if (truth & pool) else 0.0
            f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
            scores[ratio]["정확도"].append(precision)
            scores[ratio]["재현율"].append(recall)
            scores[ratio]["F1"].append(f1)
            scores[ratio]["키프레이즈 수"].append(len(kept))

    rows = [{
        "일반어 필터": "끔" if r == 0 else f"{r * 100:.0f}%",
        "_ratio": r,
        "키프레이즈 수(절반 기준)": int(np.mean(s["키프레이즈 수"])),
        "정확도": round(float(np.mean(s["정확도"])), 3),
        "재현율": round(float(np.mean(s["재현율"])), 3),
        "F1": round(float(np.mean(s["F1"])), 3),
        "F1 편차": round(float(np.std(s["F1"])), 3),
    } for r, s in scores.items()]
    table = pd.DataFrame(rows).sort_values("_ratio", ignore_index=True)
    best = float(table.loc[table["F1"].idxmax(), "_ratio"])
    table["추천"] = ["← 추천" if abs(r - best) < 1e-9 else "" for r in table["_ratio"]]
    return table.drop(columns="_ratio"), best


def sensitivity_table(args, settings: dict | None = None, progress=None) -> pd.DataFrame:
    """일반어 필터 비율과 자동 구문 기준을 바꿔가며 결과가 얼마나 달라지는지 비교."""
    df, docs, cfg, _ = prepare_docs(args, settings, progress)
    rows = []
    base_keys: set = set()
    base_top: list = []

    keep = {}

    def measure(label, ratio, kw_papers, phrase_papers):
        nonlocal base_keys, base_top
        phrases = build_phrase_dict(docs, cfg, phrase_papers, args.max_phrase_units)
        stats, forms, doc_keys, doc_kw = count_phrases(docs, cfg, phrases, set(args.sections))
        keep.setdefault("doc", (doc_keys, doc_kw))
        tweaked = argparse.Namespace(**vars(args))
        tweaked.min_keyword_ratio, tweaked.min_keyword_papers = ratio, kw_papers
        tweaked.min_phrase_papers = phrase_papers
        table, dropped = build_table(stats, forms, cfg, tweaked)
        keys = set(table["정규화 키"])
        top20 = list(table["정규화 키"].head(20))
        is_base = not rows
        if is_base:
            base_keys, base_top = keys, top20
        rows.append({
            "구분": label,
            "일반어 필터": "끔" if ratio == 0 or kw_papers == 0 else f"{ratio * 100:.0f}%",
            "자동 구문 최소 논문": phrase_papers,
            "키프레이즈 수": len(table),
            "구문(2단어 이상)": int((table["단어수"] >= 2).sum()),
            "필터로 제외": len(dropped),
            "상위 20 일치": "기준" if is_base else f"{len(set(top20) & set(base_top))}/20",
            "전체 겹침(자카드)": "기준" if is_base else round(len(keys & base_keys) / max(1, len(keys | base_keys)), 3),
            "상위 5": ", ".join(table["키프레이즈"].head(5)),
        })

    base_ratio, base_kw, base_phrase = args.min_keyword_ratio, args.min_keyword_papers, args.min_phrase_papers
    log("[민감도] 현재 설정으로 기준값 계산")
    measure(f"기준 (현재 설정)", base_ratio, base_kw, base_phrase)
    for ratio in SENSITIVITY_RATIOS:
        if abs(ratio - base_ratio) < 1e-9:
            continue
        log(f"[민감도] 일반어 필터 {ratio * 100:.0f}%")
        measure("일반어 필터 변경", ratio, 0 if ratio == 0 else base_kw, base_phrase)
    for papers in SENSITIVITY_PHRASE_PAPERS:
        if papers == base_phrase:
            continue
        log(f"[민감도] 자동 구문 최소 논문 {papers}편")
        measure("구문 기준 변경", base_ratio, base_kw, papers)

    table = pd.DataFrame(rows)
    doc_keys, doc_kw = keep["doc"]
    log("[민감도] 임계값 추천 (절반으로 고르고 나머지 절반으로 채점)")
    scored, best = recommend_ratio(doc_keys, doc_kw, args)
    return table, scored, best


def display_form(key: str, forms: Counter) -> tuple[str, str]:
    by_lower = defaultdict(Counter)
    for f, c in forms.items():
        by_lower[f.lower()][f] += c
    singular = [l for l in by_lower if raw_key(l) == key]
    best = max(singular or by_lower, key=lambda l: sum(by_lower[l].values()))
    variants = by_lower[best]
    shown = best if variants.get(best, 0) >= 0.25 * sum(variants.values()) else variants.most_common(1)[0][0]
    ordered = sorted(by_lower, key=lambda l: -sum(by_lower[l].values()))
    merged = "; ".join(by_lower[l].most_common(1)[0][0] for l in ordered[:8])
    return shown, merged


# ─────────────────────────────────────────────────────────────────────────────
# 7. 토픽 모델 (NMF) · 주제 지도 (공동출현 군집)
#    집계된 키프레이즈를 입력으로 삼아 논문을 몇 개의 주제로 묶는다.
#    - 토픽 수는 NPMI 응집도가 가장 높은 값을 추천하되, 사람이 바꿀 수 있게 한다.
#      (응집도 지표는 사람 판단과 완전히 일치하지 않는다 — Röder et al. 2015)
#    - 주제 지도는 공동출현 네트워크를 Louvain으로 묶고 Callon의 중심성·밀도로 배치한다.
# ─────────────────────────────────────────────────────────────────────────────
TOPIC_MIN_DOCS = 10      # 이보다 적으면 토픽 모델을 돌리지 않는다
TOPIC_SMALL_DOCS = 50    # 이보다 적으면 결과를 참고용으로만 본다 (경고 표시)
SMALL_CORPUS = 50        # 이보다 적으면 '2편 이상 반복' 같은 조건이 거의 다 걸러내므로 기준을 낮춘다
TOPIC_DOCS_PER_TOPIC = 5  # 토픽 하나당 최소 이만큼의 논문은 있어야 한다
TOPIC_TOP_TERMS = 8      # 응집도 계산·표시에 쓰는 토픽당 키프레이즈 수


def build_matrix(doc_keys: list[Counter], keys: list[str]):
    """논문 × 키프레이즈 빈도 행렬 (numpy)."""
    import numpy as np

    index = {k: i for i, k in enumerate(keys)}
    m = np.zeros((len(doc_keys), len(keys)), dtype=np.float32)
    for i, counter in enumerate(doc_keys):
        for k, n in counter.items():
            j = index.get(k)
            if j is not None:
                m[i, j] = n
    return m


def _npmi_scores(present, components, top: int = TOPIC_TOP_TERMS) -> list[float]:
    """토픽별 NPMI 응집도 (상위 키프레이즈 쌍의 평균)."""
    import numpy as np

    n_docs = present.shape[0]
    freq = present.sum(axis=0)
    scores = []
    for comp in components:
        idx = np.argsort(comp)[::-1][:top]
        pairs = []
        for a_i in range(len(idx)):
            for b_i in range(a_i + 1, len(idx)):
                a, b = idx[a_i], idx[b_i]
                both = int((present[:, a] & present[:, b]).sum())
                if both == 0 or not freq[a] or not freq[b]:
                    pairs.append(-1.0)
                    continue
                p_ab, p_a, p_b = both / n_docs, freq[a] / n_docs, freq[b] / n_docs
                if p_ab >= 1.0:      # 모든 문헌에 함께 나오면 NPMI는 1 (log 0 방지)
                    pairs.append(1.0)
                    continue
                pairs.append(math.log(p_ab / (p_a * p_b)) / -math.log(p_ab))
        scores.append(float(np.mean(pairs)) if pairs else 0.0)
    return scores


def fit_topics(matrix, k: int):
    """NMF 한 번 실행 → (문헌×토픽 가중치, 토픽×키프레이즈 가중치, 토픽별 응집도).
    토픽 수는 문헌 수·키프레이즈 수를 넘을 수 없다 (NMF 제약)."""
    import numpy as np
    from sklearn.decomposition import NMF
    from sklearn.feature_extraction.text import TfidfTransformer

    k = max(2, min(int(k), min(matrix.shape)))
    x = TfidfTransformer().fit_transform(matrix)
    model = NMF(n_components=k, init="nndsvda", random_state=42, max_iter=600)
    weights = model.fit_transform(x)
    scores = _npmi_scores(np.asarray(matrix) > 0, model.components_)
    return weights, model.components_, scores


def choose_topics(matrix, ks=(5, 8, 10, 12, 15, 20)):
    """토픽 수 후보별 응집도를 재어 가장 높은 값을 고른다."""
    import numpy as np

    curve = []
    best = None
    # 문헌 수·키프레이즈 수를 넘을 수 없고, 토픽당 논문이 너무 적어지지 않게도 막는다
    limit = min(min(matrix.shape), max(2, matrix.shape[0] // TOPIC_DOCS_PER_TOPIC))
    for k in sorted({min(int(k), limit) for k in ks if k >= 2}):
        if k < 2:
            continue
        weights, components, scores = fit_topics(matrix, k)
        k = int(components.shape[0])
        mean = float(np.mean(scores))
        curve.append({"토픽 수": k, "응집도(NPMI)": round(mean, 4)})
        log(f"      토픽 수 {k:>2}개 → 응집도 {mean:+.3f}")
        if best is None or mean > best[0]:
            best = (mean, k, weights, components, scores)
    return best, curve


def topic_tables(weights, components, scores, keys: list[str], labels: dict, df: pd.DataFrame):
    """토픽별 요약표와 문헌별 토픽 배정표."""
    import numpy as np

    assign = weights.argmax(axis=1)
    share = weights.max(axis=1) / np.where(weights.sum(axis=1) == 0, 1, weights.sum(axis=1))
    years = df[COL_YEAR].to_numpy()
    cited = (pd.to_numeric(df[COL_CITED], errors="coerce").to_numpy()   # 결측은 평균에서 제외
             if COL_CITED in df.columns else np.full(len(df), np.nan))
    rows = []
    for t in range(components.shape[0]):
        top_idx = np.argsort(components[t])[::-1][:TOPIC_TOP_TERMS]
        mine = assign == t
        yrs = [int(y) for y in years[mine] if y]
        cite = cited[mine]
        cite = cite[~pd.isna(cite)]
        rows.append({
            "토픽": f"T{t + 1}", "논문수": int(mine.sum()),
            "점유율": f"{100 * mine.sum() / len(assign):.1f}%",
            "평균 인용": round(float(cite.mean()), 1) if len(cite) else None,
            "인용 자료 있는 논문": int(len(cite)),
            "상위 키프레이즈": ", ".join(labels.get(keys[j], keys[j]) for j in top_idx),
            "응집도(NPMI)": round(scores[t], 3),
            "연도 중앙값": int(np.median(yrs)) if yrs else None,
            "연도 범위": f"{min(yrs)}~{max(yrs)}" if yrs else "",
        })
    summary = pd.DataFrame(rows).sort_values("논문수", ascending=False, ignore_index=True)
    summary["연도 중앙값"] = summary["연도 중앙값"].astype("Int64")  # 빈칸과 숫자가 섞이지 않게

    second = (np.argsort(weights, axis=1)[:, -2] if weights.shape[1] > 1 else assign)
    docs_tab = pd.DataFrame({
        "토픽": [f"T{t + 1}" for t in assign],
        "토픽 비중": np.round(share, 3),
        "2순위 토픽": [f"T{t + 1}" for t in second],
        "제목": df[COL_TITLE].to_numpy(),
        "연도": years, "인용수": pd.Series(cited).astype("Int64").to_numpy(),
        "DOI": df[COL_DOI].to_numpy(), "출처": df[COL_SOURCE].to_numpy(),
    }).sort_values(["토픽", "토픽 비중"], ascending=[True, False], ignore_index=True)
    return summary, docs_tab, assign



TREND_MIN_YEARS = 5      # 추세 검정에 필요한 최소 연도 수
TREND_RECENT = 3         # '최근'으로 볼 연도 수


def _mann_kendall(values: list[float]) -> tuple[float, float]:
    """Mann-Kendall 추세 검정 (Mann 1945; Kendall 1975). 반환: (S 통계량, 양측 p값).
    적은 표본·비정규 분포에서도 쓸 수 있는 비모수 검정이라 연도별 점유율에 적합하다."""
    n = len(values)
    if n < TREND_MIN_YEARS:
        return 0.0, float("nan")
    s = sum(1 if values[j] > values[i] else -1 if values[j] < values[i] else 0
            for i in range(n - 1) for j in range(i + 1, n))
    ties = Counter(values)
    var = (n * (n - 1) * (2 * n + 5) - sum(c * (c - 1) * (2 * c + 5) for c in ties.values())) / 18
    if var <= 0:                      # 모든 해가 같은 값 → 추세가 없는 것이지, 못 재는 것이 아니다
        return float(s), 1.0
    z = (s - (1 if s > 0 else -1 if s < 0 else 0)) / math.sqrt(var)
    p = 2 * (1 - 0.5 * (1 + math.erf(abs(z) / math.sqrt(2))))
    return float(s), round(p, 4)


def topic_trends(docs_tab: pd.DataFrame) -> pd.DataFrame:
    """토픽별 연도 점유율 추세. 편수만 보면 문헌 전체가 늘어난 효과와 섞이므로 점유율로 본다."""
    import numpy as np

    d = docs_tab.assign(연도=pd.to_numeric(docs_tab["연도"], errors="coerce")).dropna(subset=["연도"])
    if d.empty:
        return pd.DataFrame()
    d["연도"] = d["연도"].astype(int)
    per_year = d.pivot_table(index="연도", columns="토픽", values="제목", aggfunc="count").fillna(0)
    share = per_year.div(per_year.sum(axis=1).replace(0, 1), axis=0)
    years = share.index.to_numpy()
    recent = sorted(years)[-TREND_RECENT:]
    rows = []
    for topic in share.columns:
        vals = share[topic].to_numpy()
        s_stat, p = _mann_kendall([round(float(v), 6) for v in vals])
        slope = float(np.polyfit(years, vals, 1)[0]) * 100 if len(years) > 1 else 0.0
        before = share.loc[~share.index.isin(recent), topic]
        now = share.loc[share.index.isin(recent), topic]
        if math.isnan(p):
            label = "판단 보류 (연도 5년 미만)"
        elif p >= 0.05:
            label = "뚜렷한 변화 없음"
        else:
            label = "증가" if s_stat > 0 else "감소"
        rows.append({
            "토픽": topic, "논문수": int(per_year[topic].sum()),
            "전체 점유율": f"{100 * per_year[topic].sum() / per_year.to_numpy().sum():.1f}%",
            f"최근 {TREND_RECENT}년 점유율": f"{100 * now.mean():.1f}%" if len(now) else "",
            "그 이전": f"{100 * before.mean():.1f}%" if len(before) >= 2 else "",
            "연평균 변화(%p)": round(slope, 2),
            "추세": label, "p값": p if not math.isnan(p) else None,
        })
    out = pd.DataFrame(rows).sort_values("논문수", ascending=False, ignore_index=True)
    return out


def topic_bridges(weights, assign, df: pd.DataFrame, top_n: int = 40):
    """두 토픽에 걸친 논문 (1순위와 2순위 비중 차이가 작은 것). 분야 경계에서 나오는 연구를 찾는 데 쓴다."""
    import numpy as np

    if weights.shape[1] < 2:
        return pd.DataFrame(), pd.DataFrame()
    order = np.argsort(weights, axis=1)[:, ::-1]
    first, second = order[:, 0], order[:, 1]
    total = np.where(weights.sum(axis=1) == 0, 1, weights.sum(axis=1))
    w1, w2 = weights[np.arange(len(first)), first] / total, weights[np.arange(len(first)), second] / total
    pairs = Counter()
    for a, b, r in zip(first, second, w2 / np.where(w1 == 0, 1, w1)):
        if r >= 0.5:                       # 2순위가 1순위의 절반 이상이면 '걸친 논문'
            pairs[tuple(sorted((int(a), int(b))))] += 1
    pair_tab = pd.DataFrame([{"토픽 쌍": f"T{a + 1} ↔ T{b + 1}", "걸친 논문": n}
                             for (a, b), n in pairs.most_common(20)])
    papers = pd.DataFrame({
        "1순위": [f"T{t + 1}" for t in first], "2순위": [f"T{t + 1}" for t in second],
        "1순위 비중": np.round(w1, 3), "2순위 비중": np.round(w2, 3),
        "제목": df[COL_TITLE].to_numpy(), "연도": df[COL_YEAR].to_numpy(), "DOI": df[COL_DOI].to_numpy(),
    })
    papers = papers[papers["2순위 비중"] >= papers["1순위 비중"] * 0.5]
    papers = papers.sort_values("2순위 비중", ascending=False, ignore_index=True).head(top_n)
    return pair_tab, papers


def theme_map(doc_keys: list[Counter], labels: dict, keep_keys: set, min_cooccur: int = 2,
              min_papers: int = 3, top_n: int = 300):
    """공동출현 네트워크를 Louvain으로 묶고 Callon의 중심성·밀도를 계산.
    결과에 채택된 키프레이즈 중 많이 등장한 상위 top_n개만 사용한다."""
    import networkx as nx

    paper_count = Counter()
    for counter in doc_keys:
        paper_count.update(k for k in counter if k in keep_keys)
    keep = {k for k, n in paper_count.most_common(top_n) if n >= min_papers}
    graph = nx.Graph()
    for counter in doc_keys:
        terms = sorted(k for k in counter if k in keep)
        for a_i in range(len(terms)):
            for b_i in range(a_i + 1, len(terms)):
                a, b = terms[a_i], terms[b_i]
                graph.add_edge(a, b, weight=graph.get_edge_data(a, b, {"weight": 0})["weight"] + 1)
    graph.remove_edges_from([(a, b) for a, b, w in graph.edges(data="weight") if w < min_cooccur])
    graph.remove_nodes_from([n for n, deg in graph.degree() if deg == 0])
    if graph.number_of_nodes() < 4:
        return pd.DataFrame(), graph

    rows = []
    for nodes in nx.community.louvain_communities(graph, weight="weight", seed=42):
        nodes = set(nodes)
        inner = sum(d["weight"] for a, b, d in graph.edges(nodes, data=True) if a in nodes and b in nodes)
        outer = sum(d["weight"] for a, b, d in graph.edges(nodes, data=True) if (a in nodes) != (b in nodes))
        top = sorted(nodes, key=lambda n: graph.degree(n, weight="weight"), reverse=True)[:6]
        rows.append({
            "군집": "", "키프레이즈 수": len(nodes),
            "중심성": round(outer / len(nodes), 1), "밀도": round(inner / len(nodes), 1),
            "대표 키프레이즈": ", ".join(labels.get(t, t) for t in top),
        })
    out = pd.DataFrame(rows).sort_values("중심성", ascending=False, ignore_index=True)
    out["군집"] = [f"C{i + 1}" for i in range(len(out))]
    mid_c, mid_d = out["중심성"].median(), out["밀도"].median()
    out["사분면"] = [
        ("핵심 주제" if c >= mid_c and d >= mid_d else "기반 주제" if c >= mid_c else
         "독립 주제(니치)" if d >= mid_d else "주변 주제")
        for c, d in zip(out["중심성"], out["밀도"])]
    return out, graph


def run_topics(doc_keys, table: pd.DataFrame, df: pd.DataFrame, topics: str):
    """토픽 모델 전체 실행. topics = 'off' | 'auto' | 숫자 문자열"""
    if topics == "off" or len(df) < TOPIC_MIN_DOCS:
        return None
    keys = list(table["정규화 키"])
    labels = dict(zip(table["정규화 키"], table["키프레이즈"]))
    if len(keys) < 6:
        log(f"      [토픽] 키프레이즈가 {len(keys)}개뿐이라 건너뜁니다 (6개 이상 필요).")
        return None
    try:
        matrix = build_matrix(doc_keys, keys)
    except ImportError:
        log("      [토픽] numpy/scikit-learn이 없어 건너뜁니다.")
        return None
    log("[5/7] 토픽 모델 계산 중")
    if len(df) < TOPIC_SMALL_DOCS:
        log(f"      [주의] 논문 {len(df)}편은 토픽 모델에 적은 편입니다. "
            f"{TOPIC_SMALL_DOCS}편 미만에서는 결과를 참고용으로만 보세요.")
    try:
        if topics == "auto":
            best, curve = choose_topics(matrix)
            if best is None:
                return None
            _, k, weights, components, scores = best
            log(f"      추천 토픽 수: {k}개")
        else:
            weights, components, scores = fit_topics(matrix, max(2, int(topics)))
            k = int(components.shape[0])
            curve = []
        summary, docs_tab, assign = topic_tables(weights, components, scores, keys, labels, df)
        themes, _ = theme_map(doc_keys, labels, set(keys))
        trends = topic_trends(docs_tab)
        pair_tab, bridges = topic_bridges(weights, assign, df)
    except ImportError as e:
        log(f"      [토픽] 라이브러리가 없어 건너뜁니다 ({e}). pip install scikit-learn networkx")
        return None
    return {"k": k, "summary": summary, "docs": docs_tab, "curve": pd.DataFrame(curve),
            "themes": themes, "trends": trends, "pairs": pair_tab, "bridges": bridges,
            "matrix": matrix, "keys": keys, "labels": labels, "assign": assign, "weights": weights,
            "papers": df[[COL_TITLE, COL_YEAR, COL_DOI, COL_SOURCE, COL_CITED]].reset_index(drop=True)}


# ─────────────────────────────────────────────────────────────────────────────
# 8. 논문 지도 (초록 TF-IDF 유사도 + 토픽 색칠)
#    인용 정보 없이 '내용이 비슷한 논문끼리 가깝게' 배치한다.
#    Open Knowledge Maps와 같은 계열(텍스트 유사도 기반 지도).
# ─────────────────────────────────────────────────────────────────────────────
MAP_MIN_DOCS = 20     # 이보다 적으면 지도를 그리지 않는다
MAP_NEIGHBORS = 5     # 논문 하나당 이어줄 가장 비슷한 논문 수
MAP_MIN_SIM = 0.12    # 이보다 덜 비슷하면 선을 긋지 않는다


def paper_map(docs: list[dict], df: pd.DataFrame, topics: dict | None, doc_keys=None, labels=None,
              neighbors: int = MAP_NEIGHBORS):
    """논문별 2차원 좌표·이웃 연결선·토픽 영역.
    좌표는 t-SNE(내용이 비슷한 논문끼리 뭉치게)로 잡고, 문헌이 많으면 힘기반 배치로 대체한다."""
    if len(docs) < MAP_MIN_DOCS:
        return None
    import networkx as nx
    import numpy as np
    from sklearn.decomposition import TruncatedSVD
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.manifold import TSNE
    from sklearn.neighbors import NearestNeighbors
    from sklearn.preprocessing import normalize

    log("      논문 지도: 초록 유사도 계산 중")
    texts = [" ".join([d["title"], d["abstract"], "; ".join(d["keywords"])]) for d in docs]
    # 논문이 적으면 '3편 이상에 나온 단어'만으로는 남는 것이 거의 없다
    tfidf = TfidfVectorizer(stop_words="english", min_df=3 if len(docs) >= SMALL_CORPUS else 2,
                            max_df=0.5, sublinear_tf=True, ngram_range=(1, 2), max_features=60000)
    x = tfidf.fit_transform(texts)
    dims = min(60, max(2, min(x.shape) - 1))   # 논문 수·단어 수 둘 다 넘을 수 없다
    reduced = normalize(TruncatedSVD(n_components=dims, random_state=42).fit_transform(x))

    k = min(neighbors + 1, len(docs))
    dist, idx = NearestNeighbors(n_neighbors=k, metric="cosine").fit(reduced).kneighbors(reduced)
    graph = nx.Graph()
    graph.add_nodes_from(range(len(docs)))
    for i in range(len(docs)):
        for j, d in zip(idx[i][1:], dist[i][1:]):
            sim = 1 - float(d)
            if sim >= MAP_MIN_SIM and i != int(j):
                graph.add_edge(*sorted((i, int(j))), weight=sim)
    edges = [{"a": a, "b": b, "sim": round(w, 3)} for a, b, w in graph.edges(data="weight")]

    log(f"      논문 지도: {len(docs):,}편 · 연결 {len(edges):,}개 배치 중")
    if len(docs) <= 6000:
        coords = TSNE(n_components=2, metric="cosine", init="pca", random_state=42,
                      perplexity=min(30, max(5, len(docs) // 40))).fit_transform(reduced)
    else:  # 문헌이 아주 많으면 t-SNE가 느려 힘기반 배치를 쓴다
        pos = nx.spring_layout(graph, weight="weight", seed=42, iterations=50,
                               k=1.2 / math.sqrt(max(len(docs), 2)))
        coords = np.array([pos[i] for i in range(len(docs))])
    span = np.ptp(coords, axis=0)
    span[span == 0] = 1
    coords = (coords - coords.min(axis=0)) / span

    topic_of = ["(토픽 없음)"] * len(docs)
    if topics is not None and topics.get("assign") is not None and len(topics["assign"]) == len(docs):
        topic_of = [f"T{t + 1}" for t in topics["assign"]]
    # 인용수가 '0'인 것과 '출처에 자료가 없는 것'을 구분한다 (ProQuest에는 인용수 열이 없음)
    cited = pd.to_numeric(df[COL_CITED], errors="coerce").astype("Int64").reset_index(drop=True)
    nodes = pd.DataFrame({
        "x": coords[:, 0].round(4), "y": coords[:, 1].round(4),
        "토픽": topic_of, "연도": df[COL_YEAR].to_numpy(), "인용수": cited,
        "제목": df[COL_TITLE].to_numpy(), "DOI": df[COL_DOI].to_numpy(), "출처": df[COL_SOURCE].to_numpy(),
    })

    missing_cited = int(pd.isna(cited).sum())
    areas, bubble = map_areas(coords, doc_keys, labels)
    if bubble is not None:
        nodes["bx"] = bubble[:, 0].round(4)
        nodes["by"] = bubble[:, 1].round(4)

    # 토픽 기준 거품 (지도의 '원 기준: 토픽' 보기)
    topic_areas = pd.DataFrame()
    if topics is not None and topics.get("assign") is not None and len(topics["assign"]) == len(docs):
        head = {row["토픽"]: row["상위 키프레이즈"].split(",")[0].strip()
                for _, row in topics["summary"].iterrows()}
        names = {t: f"{f'T{t + 1}'} {head.get(f'T{t + 1}', '')}".strip() for t in set(topics["assign"])}
        topic_areas, topic_bubble = _bubble_layout(coords, topics["assign"], names)
        if topic_bubble is not None:
            nodes["tx"] = topic_bubble[:, 0].round(4)
            nodes["ty"] = topic_bubble[:, 1].round(4)
    return {"nodes": nodes, "edges": pd.DataFrame(edges), "areas": areas, "topic_areas": topic_areas,
            "인용수 없는 논문": missing_cited}


def _bubble_layout(coords, assign, names: dict):
    """묶음(assign)마다 겹치지 않는 거품을 만들고, 그 안에 논문을 원래 배치 모양대로 흩뿌린다."""
    import numpy as np

    groups = [(g, np.where(np.asarray(assign) == g)[0]) for g in sorted(set(assign))]
    groups = [(g, m) for g, m in groups if len(m) >= 3]
    if not groups:
        return pd.DataFrame(), None
    packed = _pack_circles([math.sqrt(len(m)) for _, m in groups])
    bubble = np.zeros((len(coords), 2), dtype=float)
    rows = []
    for (g, member), (bx, by, br) in zip(groups, packed):
        local = coords[member] - coords[member].mean(axis=0)
        reach = float(np.percentile(np.linalg.norm(local, axis=1), 95)) or 1.0
        bubble[member] = np.array([bx, by]) + local / reach * br * 0.78
        rows.append({"영역": names.get(g, str(g)), "논문수": int(len(member)),
                     "_bx": bx, "_by": by, "_br": br})
    corners = np.vstack([[[r["_bx"] - r["_br"], r["_by"] - r["_br"]] for r in rows],
                         [[r["_bx"] + r["_br"], r["_by"] + r["_br"]] for r in rows]])
    low = corners.min(axis=0)
    scale = float(np.ptp(corners, axis=0).max()) or 1.0
    bubble = (bubble - low) / scale
    for row in rows:
        row["거품 x"] = round((row.pop("_bx") - low[0]) / scale, 4)
        row["거품 y"] = round((row.pop("_by") - low[1]) / scale, 4)
        row["거품 반지름"] = round(row.pop("_br") / scale, 4)
    return pd.DataFrame(rows), bubble


def _pack_circles(sizes, gap: float = 1.06):
    """크기순으로 원을 가운데부터 겹치지 않게 놓는다 (Open Knowledge Maps식 거품 배치)."""
    import numpy as np

    order = np.argsort(sizes)[::-1]
    placed = []  # (x, y, r)
    for rank, i in enumerate(order):
        r = float(sizes[i])
        if rank == 0:
            placed.append((0.0, 0.0, r))
            continue
        best = None
        for angle_step in range(0, 360, 6):          # 기존 원 둘레를 돌며 빈자리를 찾는다
            for px, py, pr in placed:
                a = math.radians(angle_step)
                x, y = px + (pr + r) * gap * math.cos(a), py + (pr + r) * gap * math.sin(a)
                if any((x - qx) ** 2 + (y - qy) ** 2 < ((r + qr) * gap * 0.99) ** 2 for qx, qy, qr in placed):
                    continue
                d = x * x + y * y                    # 가운데에 가까운 자리를 선호
                if best is None or d < best[0]:
                    best = (d, x, y)
        placed.append((best[1], best[2], r) if best else (0.0, 0.0, r))
    out = [None] * len(sizes)
    for (x, y, r), i in zip(placed, order):
        out[i] = (x, y, r)
    return out


def map_areas(coords, doc_keys: list[Counter] | None, labels: dict | None, max_areas: int = 12):
    """지도 위에서 실제로 뭉친 덩어리를 영역으로 잡고, 그 영역에서만 유독 자주 나오는
    키프레이즈로 이름을 붙인다 (Open Knowledge Maps가 영역 제목을 만드는 방식)."""
    import numpy as np
    from sklearn.cluster import KMeans

    n = len(coords)
    k = int(min(max_areas, max(3, round(n / 250))))
    if n < 3 * k:
        return pd.DataFrame(), None
    km = KMeans(n_clusters=k, random_state=42, n_init=10).fit(coords)
    overall = Counter()
    if doc_keys:
        for counter in doc_keys:
            overall.update(counter.keys())
    rows = []
    for c in range(k):
        member = np.where(km.labels_ == c)[0]
        if len(member) < 5:
            continue
        center = coords[member].mean(axis=0)
        d = np.linalg.norm(coords[member] - center, axis=1)
        name = f"영역 {c + 1}"
        if doc_keys:
            inside = Counter()
            for i in member:
                inside.update(doc_keys[i].keys())
            # 영역 안 비율 ÷ 전체 비율 = 그 영역의 특징어
            scored = [(cnt / len(member) * math.log(1 + cnt) / (overall[key] / n + 1e-9), key)
                      for key, cnt in inside.items() if cnt >= max(3, len(member) * 0.08)]
            scored.sort(reverse=True)
            picked = [(labels or {}).get(key, key) for _, key in scored[:2]]
            if picked:
                name = " · ".join(picked)
        rows.append({"영역": name, "x": round(float(center[0]), 4), "y": round(float(center[1]), 4),
                     "반지름": round(float(np.percentile(d, 60)), 4), "논문수": int(len(member)),
                     "_cluster": int(c)})
    if not rows:
        return pd.DataFrame(), None

    # 거품 배치: 논문 수의 제곱근을 반지름으로 삼아 겹치지 않게 놓는다
    names = {r["_cluster"]: r["영역"] for r in rows}
    packed, bubble = _bubble_layout(coords, km.labels_, names)
    spread = pd.DataFrame(rows).drop(columns="_cluster")
    merged = spread.merge(packed[["영역", "거품 x", "거품 y", "거품 반지름"]], on="영역", how="left")
    return merged, bubble


MAP_COLORS = ["#1f5fbf", "#d1343f", "#e07b00", "#1e9e5a", "#7b3fb8", "#0f8a9d", "#c2185b", "#5d7f00",
              "#b45f06", "#34495e", "#e04e9a", "#2e86de", "#8a6d3b", "#17807a", "#9b2d86", "#4a6fa5",
              "#a04000", "#2d7f4f", "#6c3483", "#117864"]


def paper_map_html(mapping: dict, title: str, topic_names: dict | None = None) -> str:
    """혼자서도 열리는 HTML 한 장.
    보기 세 가지 — 주제 영역(Open Knowledge Maps식) · 연결망 · 밀도(VOSviewer식 열지도)."""
    nodes, edges = mapping["nodes"], mapping["edges"]
    areas = mapping.get("areas")
    topics = sorted(nodes["토픽"].unique())
    colors = {t: MAP_COLORS[i % len(MAP_COLORS)] for i, t in enumerate(topics)}
    payload = {
        # 인용수가 없는 논문은 c = -1 (화면에서 중립 크기로 그리고 '정보 없음'으로 표시)
        "nodes": [{"x": float(r.x), "y": float(r.y), "t": r.토픽,
                   "c": -1 if pd.isna(r.인용수) else int(r.인용수), "y2": str(r.연도),
                   "ti": str(r.제목)[:160], "d": str(r.DOI), "s": str(r.출처),
                   "bx": float(getattr(r, "bx", r.x)), "by": float(getattr(r, "by", r.y)),
                   "tx": float(getattr(r, "tx", getattr(r, "bx", r.x))),
                   "ty": float(getattr(r, "ty", getattr(r, "by", r.y)))}
                  for r in nodes.itertuples()],
        "edges": [[int(r.a), int(r.b), float(r.sim)] for r in edges.itertuples()],
        "areas": ([{"t": a["영역"], "x": float(a["x"]), "y": float(a["y"]), "r": float(a["반지름"]),
                    "n": int(a["논문수"]), "bx": float(a["거품 x"]), "by": float(a["거품 y"]),
                    "br": float(a["거품 반지름"])}
                   for a in areas.to_dict("records")] if areas is not None and not areas.empty else []),
        "tareas": ([{"t": a["영역"], "n": int(a["논문수"]), "bx": float(a["거품 x"]), "by": float(a["거품 y"]),
                     "br": float(a["거품 반지름"])}
                    for a in mapping["topic_areas"].to_dict("records")]
                   if mapping.get("topic_areas") is not None and not mapping["topic_areas"].empty else []),
        "colors": colors,
        "names": {t: (topic_names or {}).get(t, t) for t in topics},
    }
    data = json.dumps(payload, ensure_ascii=False)
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>{title} — 논문 지도</title>
<style>
 :root {{ color-scheme: light; }}
 body {{ margin:0; font-family:"Malgun Gothic","맑은 고딕",system-ui,sans-serif; background:#fff; color:#0f2438; }}
 #wrap {{ position:relative; width:100%; height:100vh; min-height:520px; overflow:hidden; }}
 canvas {{ display:block; width:100%; height:100%; cursor:grab; }}
 canvas.drag {{ cursor:grabbing; }}
 #panel {{ position:absolute; top:12px; left:12px; background:rgba(255,255,255,.94); border:1px solid #d9e1eb;
   border-radius:8px; padding:10px 12px; font-size:12.5px; width:150px; max-height:88%; overflow:auto; }}
 #panel h1 {{ font-size:13.5px; margin:0 0 8px; }}
 #q {{ width:100%; box-sizing:border-box; padding:5px 7px; border:1px solid #d9e1eb; border-radius:5px; font:inherit; margin-bottom:8px; }}
 #views {{ display:flex; gap:4px; margin-bottom:8px; }}
 #views button {{ flex:1; font:inherit; font-size:11px; padding:4px 1px; border:1px solid #d9e1eb; background:#fff;
   border-radius:5px; cursor:pointer; color:#3e5165; white-space:nowrap; }}
 #views button.on {{ background:#0a4f9e; border-color:#0a4f9e; color:#fff; }}
 #basis {{ display:flex; gap:4px; align-items:center; margin-bottom:8px; font-size:11px; color:#6a7b8e; }}
 #basis span {{ flex:none; }}
 #basis button {{ flex:1; font:inherit; font-size:11px; padding:3px 1px; border:1px solid #d9e1eb; background:#fff;
   border-radius:5px; cursor:pointer; color:#3e5165; }}
 #basis button.on {{ background:#eef4fb; border-color:#9bbde4; color:#0a4f9e; font-weight:600; }}
 .lg {{ display:flex; align-items:center; gap:6px; padding:2px 0; cursor:pointer; }}
 .sw {{ width:11px; height:11px; border-radius:3px; flex:none; }}
 .lg.off {{ opacity:.35; }}
 #tip {{ position:absolute; pointer-events:none; background:#0f2438; color:#fff; padding:7px 9px; border-radius:6px;
   font-size:12px; max-width:330px; display:none; line-height:1.45; }}
 #list {{ position:absolute; top:12px; right:12px; width:330px; max-height:90%; overflow:auto; display:none;
   background:rgba(255,255,255,.97); border:1px solid #d9e1eb; border-radius:8px; padding:12px 14px; font-size:12.5px; }}
 #list h2 {{ font-size:13.5px; margin:0 0 4px; }}
 #list .sub {{ color:#6a7b8e; font-size:11.5px; margin-bottom:8px; }}
 #list .p {{ padding:7px 0; border-top:1px solid #eef2f7; line-height:1.45; }}
 #list .meta {{ color:#6a7b8e; font-size:11.5px; }}
 #list a {{ color:#0a4f9e; }}
 #close {{ float:right; cursor:pointer; color:#6a7b8e; border:0; background:none; font-size:15px; }}
 #hint {{ position:absolute; right:12px; bottom:12px; font-size:11.5px; color:#6a7b8e; background:rgba(255,255,255,.9); padding:4px 8px; border-radius:5px; }}
</style></head><body>
<div id="wrap"><canvas id="cv"></canvas>
 <div id="panel"><h1>{title}</h1>
  <div id="views">
    <button data-v="bubbles" class="on">거품</button>
    <button data-v="areas">분포</button>
    <button data-v="network">연결망</button>
    <button data-v="density">밀도</button>
  </div>
  <div id="basis"><span>원 기준</span>
    <button data-b="topic" class="on">토픽</button><button data-b="location">위치</button>
  </div>
  <input id="q" placeholder="제목 검색…" autocomplete="off">
  <div id="legend"></div></div>
 <div id="tip"></div><div id="list"></div>
 <div id="hint">논문을 누르면 <b>비슷한 논문</b> · 거품을 누르면 목록 · 휠 확대 · <b>가까운 점끼리 비슷함</b>(멀리 떨어진 거리·덩어리 크기는 의미 없음)</div></div>
<script>
const DATA = {data};
const cv = document.getElementById('cv'), ctx = cv.getContext('2d'), tip = document.getElementById('tip');
let view = {{s:1, dx:0, dy:0}}, hidden = new Set(), query = '', hover = -1, W = 0, H = 0,
    state = 'bubbles', focus = -1, basis = 'topic', sel = -1;
let rel = new Set();   // 고른 논문과 이어진 논문들
const maxCite = Math.max(1, ...DATA.nodes.map(n => n.c));
// 인용수가 없는 논문(c = -1)은 중립 크기로 — '인용 0'과 구분한다
function size(n) {{ return n.c < 0 ? 4.2 : 2.6 + 7 * Math.sqrt(n.c / maxCite); }}
function citeText(n) {{ return n.c < 0 ? '인용 정보 없음' : '인용 ' + n.c; }}
function resize() {{
  const r = cv.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
  W = r.width; H = r.height; cv.width = W * dpr; cv.height = H * dpr; ctx.setTransform(dpr,0,0,dpr,0,0); draw();
}}
const PAD_L = 178, PAD = 44;   // 왼쪽은 범례 패널만큼 비워 둔다
function bubbleAreas() {{ return (basis === 'topic' && DATA.tareas.length) ? DATA.tareas : DATA.areas; }}
// 거품은 자기 좌표(bx,by)에 그린다 — 기준이 토픽이어도 같은 값을 tx,ty로 넘겨야 변환이 성립한다
function areaPt(a) {{ return pt({{x:a.bx, y:a.by, bx:a.bx, by:a.by, tx:a.bx, ty:a.by}}); }}
// 논문별 이웃 (가장 비슷한 논문 순) — 논문을 고르면 이 목록을 보여준다
const NB = DATA.nodes.map(() => []);
DATA.edges.forEach(([a,b,w]) => {{ NB[a].push([b,w]); NB[b].push([a,w]); }});
NB.forEach(l => l.sort((p,q) => q[1]-p[1]));
function pt(n) {{
  const bubble = state === 'bubbles', topic = basis === 'topic' && DATA.tareas.length;
  let x = n.x, y = n.y;
  if (bubble) {{ x = topic ? n.tx : n.bx; y = topic ? n.ty : n.by; }}
  return [PAD_L + x * (W - PAD_L - PAD) * view.s + view.dx, PAD + (1-y) * (H - 2*PAD) * view.s + view.dy];
}}
function span() {{ return (W - PAD_L - PAD) * view.s; }}
function visible(n) {{ return !hidden.has(n.t) && (!query || n.ti.toLowerCase().includes(query)); }}
function hex(c, a) {{ const n = parseInt(c.slice(1),16); return 'rgba('+(n>>16)+','+((n>>8)&255)+','+(n&255)+','+a+')'; }}
function drawAreas() {{            // Open Knowledge Maps식 주제 영역 (지도에서 뭉친 덩어리)
  DATA.areas.forEach(a => {{
    const c = pt(a), rr = a.r * (W-84) * view.s * 1.1;
    const g = ctx.createRadialGradient(c[0],c[1],rr*0.1,c[0],c[1],rr);
    g.addColorStop(0,'rgba(104,132,166,0.13)'); g.addColorStop(1,'rgba(104,132,166,0.01)');
    ctx.beginPath(); ctx.arc(c[0],c[1],rr,0,6.2832); ctx.fillStyle = g; ctx.fill();
    ctx.setLineDash([5,5]); ctx.lineWidth=1; ctx.strokeStyle='rgba(104,132,166,0.33)'; ctx.stroke(); ctx.setLineDash([]);
  }});
}}
function drawDensity() {{          // VOSviewer식 밀도 보기
  const off = document.createElement('canvas'); off.width=W; off.height=H;
  const o = off.getContext('2d');
  DATA.nodes.forEach(n => {{
    if (!visible(n)) return;
    const p = pt(n), r = (16 + 20*Math.sqrt(Math.max(0, n.c)/maxCite)) * Math.max(0.6, Math.min(2.2, view.s));
    const g = o.createRadialGradient(p[0],p[1],0,p[0],p[1],r);
    g.addColorStop(0,'rgba(255,255,255,0.10)'); g.addColorStop(1,'rgba(255,255,255,0)');
    o.fillStyle = g; o.beginPath(); o.arc(p[0],p[1],r,0,6.2832); o.fill();
  }});
  const img = o.getImageData(0,0,W,H), d = img.data;
  let peak = 1;                       // 가장 밀집한 곳을 기준으로 정규화 (포화 방지)
  for (let i=3;i<d.length;i+=4) if (d[i] > peak) peak = d[i];
  const ramp = [[255,255,255],[214,232,246],[150,205,224],[126,201,160],[214,219,110],[240,176,80],[214,90,72]];
  for (let i=0;i<d.length;i+=4) {{
    const v = Math.min(1, d[i+3]/peak), f = v*(ramp.length-1), k = Math.min(ramp.length-2, Math.floor(f)), m = f-k;
    d[i]   = ramp[k][0]+(ramp[k+1][0]-ramp[k][0])*m;
    d[i+1] = ramp[k][1]+(ramp[k+1][1]-ramp[k][1])*m;
    d[i+2] = ramp[k][2]+(ramp[k+1][2]-ramp[k][2])*m;
    d[i+3] = 255;
  }}
  ctx.putImageData(img,0,0);
}}
function fits(boxes, x, y, w, h) {{   // VOSviewer처럼 겹치는 글자는 그리지 않는다
  for (const b of boxes) if (x < b[2] && x+w > b[0] && y < b[3] && y+h > b[1]) return false;
  boxes.push([x,y,x+w,y+h]); return true;
}}
function drawLabels() {{
  const boxes = [];
  ctx.textAlign = 'center';
  {{                                    // 영역 이름 먼저 (큰 글씨)
    const bubble = state === 'bubbles';
    (bubble ? bubbleAreas() : DATA.areas).slice().sort((a,b)=>b.n-a.n).forEach(a => {{
      const c = bubble ? areaPt(a) : pt(a);
      const rr = bubble ? a.br * span() : 0;
      ctx.font = (bubble ? '600 ' + Math.max(11, Math.min(17, rr/4.2)) + 'px' : '600 13.5px') + ' "Malgun Gothic",sans-serif';
      const txt = a.t + ' (' + a.n + ')', w = ctx.measureText(txt).width;
      const ty = bubble ? c[1] - rr + 16 : c[1] - 6;
      if (!fits(boxes, c[0]-w/2-4, ty-14, w+8, 19)) return;
      ctx.lineWidth = 3.5; ctx.strokeStyle = 'rgba(255,255,255,.92)'; ctx.strokeText(txt, c[0], ty);
      ctx.fillStyle = '#0f2438'; ctx.fillText(txt, c[0], ty);
    }});
  }}
  if (view.s < 1.8) return;            // 확대해야 논문 제목이 나온다
  ctx.font = '11px "Malgun Gothic",sans-serif';
  DATA.nodes.map((n,i)=>[n,i]).sort((a,b)=>b[0].c-a[0].c).slice(0,400).forEach(([n]) => {{
    if (!visible(n)) return;
    const p = pt(n); if (p[0]<-50||p[0]>W+50||p[1]<-30||p[1]>H+30) return;
    const txt = n.ti.length > 46 ? n.ti.slice(0,46)+'…' : n.ti;
    const w = ctx.measureText(txt).width, s = size(n);
    if (!fits(boxes, p[0]-w/2-3, p[1]+s, w+6, 15)) return;
    ctx.lineWidth = 3; ctx.strokeStyle = 'rgba(255,255,255,.92)'; ctx.strokeText(txt, p[0], p[1]+s+11);
    ctx.fillStyle = '#28384a'; ctx.fillText(txt, p[0], p[1]+s+11);
  }});
}}
function drawBubbles() {{          // Open Knowledge Maps식: 겹치지 않는 거품 + 그 안의 논문
  bubbleAreas().forEach((a,i) => {{
    const c = areaPt(a), rr = a.br * span();
    ctx.beginPath(); ctx.arc(c[0],c[1],rr,0,6.2832);
    ctx.fillStyle = (focus === i) ? 'rgba(207,226,245,0.95)' : 'rgba(228,237,246,0.75)'; ctx.fill();
    ctx.lineWidth = focus === i ? 2 : 1; ctx.strokeStyle = focus === i ? '#0a4f9e' : 'rgba(146,170,196,0.75)'; ctx.stroke();
  }});
}}
function areaAt(mx,my) {{
  if (state !== 'bubbles') return -1;
  let hit = -1, best = 1e9;
  bubbleAreas().forEach((a,i) => {{
    const c = areaPt(a), rr = a.br * span();
    const d = (c[0]-mx)**2 + (c[1]-my)**2;
    if (d <= rr*rr && d < best) {{ best = d; hit = i; }}
  }});
  return hit;
}}
function showList(i) {{
  focus = i;
  if (i >= 0) {{ sel = -1; rel = new Set(); }}
  const box = document.getElementById('list');
  if (i < 0) {{ box.style.display='none'; draw(); return; }}
  const a = bubbleAreas()[i], c = [a.bx, a.by], rr = a.br, topic = basis === 'topic' && DATA.tareas.length;
  const inside = DATA.nodes.map((n,j)=>[n,j]).filter(([n]) => {{
    const px = topic ? n.tx : n.bx, py = topic ? n.ty : n.by;
    return (px-c[0])**2 + (py-c[1])**2 <= rr*rr && visible(n);
  }});
  inside.sort((p,q) => q[0].c - p[0].c);
  box.innerHTML = '<button id="close">×</button><h2>' + a.t + '</h2>' +
    '<div class="sub">논문 ' + inside.length + '편 · 인용 많은 순</div>' +
    inside.slice(0,60).map(([n]) =>
      '<div class="p">' + n.ti.replace(/</g,'&lt;') +
      '<div class="meta">' + (DATA.names[n.t]||n.t) + ' · ' + n.y2 + ' · ' + citeText(n) +
      (n.d ? ' · <a href="https://doi.org/'+n.d+'" target="_blank" rel="noopener">DOI</a>' : '') + '</div></div>').join('');
  box.style.display = 'block';
  document.getElementById('close').onclick = () => showList(-1);
  draw();
}}
function showPaper(i) {{
  const box = document.getElementById('list');
  if (i < 0) {{ sel = -1; rel = new Set(); box.style.display='none'; draw(); return; }}
  sel = i; focus = -1;
  const n = DATA.nodes[i], nb = NB[i].filter(([j]) => visible(DATA.nodes[j]));
  rel = new Set(nb.map(([j]) => j));
  box.innerHTML = '<button id="close">×</button><h2>' + n.ti.replace(/</g,'&lt;') + '</h2>' +
    '<div class="sub">' + (DATA.names[n.t]||n.t) + ' · ' + n.y2 + ' · ' + citeText(n) +
    (n.d ? ' · <a href="https://doi.org/'+n.d+'" target="_blank" rel="noopener">DOI</a>' : '') + '</div>' +
    '<div class="sub">내용이 가장 비슷한 논문 ' + nb.length + '편 — 제목을 누르면 그 논문 기준으로 바뀝니다</div>' +
    (nb.length ? nb.slice(0,40).map(([j,w]) => {{
      const m = DATA.nodes[j];
      return '<div class="p" data-i="'+j+'">' + m.ti.replace(/</g,'&lt;') +
        '<div class="meta">유사도 ' + Math.round(w*100) + '% · ' + (DATA.names[m.t]||m.t) + ' · ' + m.y2 +
        (m.d ? ' · <a href="https://doi.org/'+m.d+'" target="_blank" rel="noopener">DOI</a>' : '') + '</div></div>';
    }}).join('') : '<div class="p">이 논문과 충분히 비슷한 논문이 없습니다 (유사도 12% 미만)</div>');
  box.style.display = 'block';
  document.getElementById('close').onclick = () => showPaper(-1);
  box.querySelectorAll('.p[data-i]').forEach(d => d.addEventListener('click', ev => {{
    if (ev.target.tagName === 'A') return;         // DOI 링크는 그대로 열리게
    showPaper(parseInt(d.dataset.i));
  }}));
  draw();
}}
function draw() {{
  ctx.clearRect(0,0,W,H);
  if (state === 'density') {{ drawDensity(); drawLabels(); return; }}
  if (state === 'bubbles') drawBubbles();
  if (state === 'areas') drawAreas();
  if (state === 'network') {{
    ctx.lineWidth = 0.5; ctx.strokeStyle = 'rgba(120,140,165,0.30)'; ctx.beginPath();
    for (const [a,b] of DATA.edges) {{
      const na = DATA.nodes[a], nb = DATA.nodes[b];
      if (!visible(na) || !visible(nb)) continue;
      const pa = pt(na), pb = pt(nb);
      ctx.moveTo(pa[0],pa[1]); ctx.lineTo(pb[0],pb[1]);
    }}
    ctx.stroke();
  }}
  if (sel >= 0) {{                     // 고른 논문에서 이웃으로 선을 그어 관계를 보여준다
    const ps = pt(DATA.nodes[sel]);
    ctx.lineWidth = 1.4; ctx.strokeStyle = 'rgba(10,79,158,0.45)'; ctx.beginPath();
    rel.forEach(j => {{ const q = pt(DATA.nodes[j]); ctx.moveTo(ps[0],ps[1]); ctx.lineTo(q[0],q[1]); }});
    ctx.stroke();
  }}
  DATA.nodes.forEach((n,i) => {{
    if (!visible(n)) return;
    const p = pt(n), on = sel < 0 || i === sel || rel.has(i);
    const mag = i===sel ? 2.0 : (i===hover ? 1.6 : (rel.has(i) ? 1.25 : 1));
    ctx.beginPath(); ctx.arc(p[0],p[1], size(n) * mag, 0, 6.2832);
    ctx.fillStyle = DATA.colors[n.t] || '#888';
    ctx.globalAlpha = on ? (i===hover || i===sel ? 1 : .85) : .12; ctx.fill(); ctx.globalAlpha = 1;
    if (i===sel) {{ ctx.lineWidth=2.5; ctx.strokeStyle='#0a4f9e'; ctx.stroke(); }}
    else if (i===hover) {{ ctx.lineWidth=1.5; ctx.strokeStyle='#0f2438'; ctx.stroke(); }}
  }});
  drawLabels();
}}
function nearest(mx,my) {{
  let best=-1, bd=14*14;
  DATA.nodes.forEach((n,i) => {{ if(!visible(n)) return; const p=pt(n); const d=(p[0]-mx)**2+(p[1]-my)**2; if(d<bd) {{bd=d;best=i;}} }});
  return best;
}}
cv.addEventListener('mousemove', e => {{
  const r = cv.getBoundingClientRect(), mx = e.clientX-r.left, my = e.clientY-r.top;
  if (drag) {{ view.dx += mx-last[0]; view.dy += my-last[1]; last=[mx,my]; draw(); return; }}
  const i = nearest(mx,my);
  if (i !== hover) {{ hover = i; draw(); }}
  if (i >= 0) {{
    const n = DATA.nodes[i];
    tip.innerHTML = '<b>'+n.ti.replace(/</g,'&lt;')+'</b><br>'+(DATA.names[n.t]||n.t)+' · '+n.y2+' · '+citeText(n)+'<br>'+n.s+(n.d?('<br>'+n.d):'');
    tip.style.display='block'; tip.style.left=Math.min(mx+14, r.width-340)+'px'; tip.style.top=(my+14)+'px';
  }} else tip.style.display='none';
}});
let drag=false, last=[0,0], down=[0,0];
cv.addEventListener('mousedown', e => {{ drag=true; cv.classList.add('drag'); const r=cv.getBoundingClientRect(); last=[e.clientX-r.left,e.clientY-r.top]; down=last.slice(); }});
window.addEventListener('mouseup', () => {{ drag=false; cv.classList.remove('drag'); }});
cv.addEventListener('click', e => {{
  const r = cv.getBoundingClientRect(), mx = e.clientX-r.left, my = e.clientY-r.top;
  if (Math.abs(mx-down[0]) > 4 || Math.abs(my-down[1]) > 4) return;   // 드래그는 클릭으로 보지 않음
  const j = nearest(mx,my);
  if (j >= 0) {{ showPaper(j); return; }}            // 논문을 누르면 그 논문 기준으로
  const a = areaAt(mx,my);
  if (a >= 0) {{ showPaper(-1); showList(a); return; }}
  showPaper(-1); showList(-1);                     // 빈 곳을 누르면 해제
}});
cv.addEventListener('wheel', e => {{
  e.preventDefault();
  const r=cv.getBoundingClientRect(), mx=e.clientX-r.left, my=e.clientY-r.top, f=e.deltaY<0?1.12:1/1.12;
  view.dx = mx - (mx-view.dx)*f; view.dy = my - (my-view.dy)*f; view.s *= f; draw();
}}, {{passive:false}});
document.getElementById('q').addEventListener('input', e => {{ query=e.target.value.trim().toLowerCase(); draw(); }});
document.querySelectorAll('#views button').forEach(b => b.addEventListener('click', () => {{
  state = b.dataset.v;
  document.querySelectorAll('#views button').forEach(x => x.classList.toggle('on', x === b));
  document.getElementById('basis').style.display = state === 'bubbles' ? 'flex' : 'none';
  if (state !== 'bubbles') showList(-1); else draw();
  draw();
}}));
document.querySelectorAll('#basis button').forEach(b => b.addEventListener('click', () => {{
  basis = b.dataset.b;
  document.querySelectorAll('#basis button').forEach(x => x.classList.toggle('on', x === b));
  showList(-1);   // 기준이 바뀌면 열려 있던 목록은 닫는다
  draw();
}}));
if (!DATA.tareas.length) document.getElementById('basis').style.display = 'none';
const legend = document.getElementById('legend');
Object.keys(DATA.colors).forEach(t => {{
  const n = DATA.nodes.filter(x => x.t===t).length;
  const row = document.createElement('div'); row.className='lg';
  row.innerHTML = '<span class="sw" style="background:'+DATA.colors[t]+'"></span><span>'+(DATA.names[t]||t)+' ('+n+')</span>';
  row.onclick = () => {{ hidden.has(t) ? hidden.delete(t) : hidden.add(t); row.classList.toggle('off'); draw(); }};
  legend.appendChild(row);
}});
window.addEventListener('resize', resize); resize();
</script></body></html>"""


# ─────────────────────────────────────────────────────────────────────────────
# 9. 출력 (엑셀 + 워드클라우드)
# ─────────────────────────────────────────────────────────────────────────────
# 다채로운 색 (흰 배경에서 작은 글자도 읽히는 명도)
PALETTE = ["#1f5fbf", "#d1343f", "#e07b00", "#1e9e5a", "#7b3fb8", "#0f8a9d",
           "#c2185b", "#5d7f00", "#b45f06", "#34495e", "#e04e9a", "#2e86de"]
# 글자 각도 후보 (0이 여러 번 들어 있으면 가로가 그만큼 자주 나옴)
ANGLE_SETS = {
    "mixed": [0, 0, 0, 0, 30, -30, 45, -45, 60, -60, 90],
    "right": [0, 0, 90],
    "horizontal": [0],
}


def default_font() -> str | None:
    for p in ("C:/Windows/Fonts/seguisb.ttf", "C:/Windows/Fonts/arialbd.ttf", "C:/Windows/Fonts/arial.ttf",
              "/Library/Fonts/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        if Path(p).exists():
            return p
    return None


def _font_file(args) -> str:
    path = args.font or default_font()
    if path:
        return path
    from wordcloud.wordcloud import FONT_PATH  # wordcloud 패키지에 들어 있는 기본 글꼴

    return FONT_PATH


def _color_picker(args, rng):
    if args.colormap:
        import matplotlib

        cmap = matplotlib.colormaps[args.colormap]
        return lambda: tuple(int(255 * c) for c in cmap(rng.uniform(0.05, 0.95))[:3])
    palette, order = PALETTE, []

    def pick():  # 팔레트를 섞어 한 바퀴씩 돌려 색이 고르게 나오도록
        if not order:
            order.extend(rng.sample(palette, len(palette)))
        return order.pop()

    return pick


def _word_sprite(text: str, font, angle: int, cell: int):
    """글자를 angle만큼 돌린 알파 이미지와, 겹침 검사용 격자 마스크(주변 1칸 여백 포함)."""
    import numpy as np
    from PIL import Image, ImageDraw

    left, top, right, bottom = font.getbbox(text)
    pad = 2 * cell
    sprite = Image.new("L", (right - left + 2 * pad, bottom - top + 2 * pad), 0)
    ImageDraw.Draw(sprite).text((pad - left, pad - top), text, font=font, fill=255)
    if angle:
        sprite = sprite.rotate(angle, resample=Image.Resampling.BICUBIC, expand=True)
    ink = np.asarray(sprite) > 32
    gh, gw = -(-ink.shape[0] // cell), -(-ink.shape[1] // cell)
    padded = np.zeros((gh * cell, gw * cell), dtype=bool)
    padded[: ink.shape[0], : ink.shape[1]] = ink
    grid = padded.reshape(gh, cell, gw, cell).any(axis=(1, 3))
    mask = grid.copy()
    mask[1:] |= grid[:-1]
    mask[:-1] |= grid[1:]
    mask[:, 1:] |= grid[:, :-1]
    mask[:, :-1] |= grid[:, 1:]
    return sprite, mask


def make_wordcloud(freqs: dict[str, float], path: Path, args) -> None:
    """워드클라우드 이미지 생성.
    wordcloud 패키지는 0°/90°만 지원하므로, 여러 각도 배치를 위해 직접 배치한다.
    글자 크기는 wordcloud 패키지와 같은 방식(relative_scaling), 빈 자리는 가운데에 가까운 곳부터 채운다."""
    import random

    import numpy as np
    from PIL import Image, ImageFont

    items = sorted(freqs.items(), key=lambda kv: -kv[1])
    if not items:
        return
    rng, noise = random.Random(42), np.random.default_rng(42)
    cell, min_size, rel = 4, 9, 0.45
    width, height = args.width, args.height
    gh, gw = height // cell, width // cell
    occupied = np.zeros((gh, gw))
    occupied_f = np.fft.rfft2(occupied)
    font_file, fonts = _font_file(args), {}

    def font(size):
        if size not in fonts:
            fonts[size] = ImageFont.truetype(font_file, size)
        return fonts[size]

    def find_spot(mask):
        """겹치지 않는 위치 중 가운데에 가장 가까운 곳 (FFT 상관으로 모든 위치를 한 번에 검사)."""
        h, w = mask.shape
        if h > gh or w > gw:
            return None
        kernel = np.zeros((gh, gw))
        kernel[:h, :w] = mask
        overlap = np.fft.irfft2(occupied_f * np.conj(np.fft.rfft2(kernel)), s=(gh, gw))
        free = overlap[: gh - h + 1, : gw - w + 1] < 0.5
        if not free.any():
            return None
        ys = (np.arange(gh - h + 1) + h / 2 - gh / 2) / gh
        xs = (np.arange(gw - w + 1) + w / 2 - gw / 2) / gw
        score = ys[:, None] ** 2 + xs[None, :] ** 2 + noise.uniform(0, 0.004, free.shape)
        score[~free] = np.inf
        y, x = np.unravel_index(np.argmin(score), score.shape)
        return int(y), int(x)

    size = int(height * 0.5)  # 1위 항목이 가로로 화면에 들어가는 크기에서 시작
    while size > min_size:
        left, top, right, bottom = font(size).getbbox(items[0][0])
        if right - left <= width * 0.6 and bottom - top <= height * 0.28:
            break
        size = int(size * 0.92)

    angles, pick_color = ANGLE_SETS[args.angles], _color_picker(args, rng)
    canvas = Image.new("RGB", (width, height), args.background)
    last_freq = items[0][1]
    for rank, (word, freq) in enumerate(items):
        if rank:
            size = int(round((rel * freq / last_freq + (1 - rel)) * size))
        last_freq = freq
        angle = 0 if rank < 3 else rng.choice(angles)  # 상위 3개는 가로로
        placed = False
        while size >= min_size and not placed:
            for a in dict.fromkeys([angle, 0]):  # 자리가 없으면 가로로 한 번 더 시도
                sprite, mask = _word_sprite(word, font(size), a, cell)
                spot = find_spot(mask)
                if spot:
                    y, x = spot
                    canvas.paste(pick_color(), (x * cell, y * cell), sprite)
                    occupied[y:y + mask.shape[0], x:x + mask.shape[1]] += mask
                    occupied_f = np.fft.rfft2(occupied)
                    placed = True
                    break
            else:
                size = int(size * 0.9)
        if not placed:
            break  # 가장 작은 글자도 들어갈 자리가 없음
    canvas.save(str(path))


def autosize(ws, df: pd.DataFrame) -> None:
    from openpyxl.utils import get_column_letter

    for i, col in enumerate(df.columns, start=1):
        sample = [len(str(v)) for v in df[col].head(300).tolist()]   # 결측이 섞여도 안전하게
        width = max([len(str(col)) * 2] + sample) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(max(width, 8), 70)
    ws.freeze_panes = "A2"


def build_table(stats, forms, cfg: Config, args) -> tuple[pd.DataFrame, pd.DataFrame]:
    """키워드성 필터: 등장 논문 중 저자키워드로 쓴 논문 비율이 기준 이상인 항목만 채택.
    ('risk', 'valuable insight' 처럼 초록에 흔하지만 연구 주제어로는 쓰이지 않는 일반어 제거)"""
    rows = []
    for key, s in stats.items():
        if s["papers"] < args.min_papers:
            continue
        ratio = s["kw_papers"] / s["papers"]
        passed = key in cfg.custom or (s["kw_papers"] >= args.min_keyword_papers and ratio >= args.min_keyword_ratio)
        shown, merged = display_form(key, forms[key])
        rows.append({
            "키프레이즈": shown, "단어수": len(key.split()),
            "총빈도": s["title"] + s["abstract"] + s["keywords"], "논문수": s["papers"],
            "제목": s["title"], "초록": s["abstract"], "저자키워드": s["keywords"],
            "저자키워드 사용률": round(ratio, 3), "병합된 표기": merged, "정규화 키": key, "_pass": passed,
        })
    sort_col, other = ("총빈도", "논문수") if args.weight == "count" else ("논문수", "총빈도")
    table = pd.DataFrame(rows).sort_values([sort_col, other], ascending=False, ignore_index=True)
    kept = table[table["_pass"]].drop(columns="_pass").reset_index(drop=True)
    dropped = table[~table["_pass"]].drop(columns="_pass").reset_index(drop=True)
    return kept, dropped


def write_outputs(table: pd.DataFrame, dropped: pd.DataFrame, settings: dict, args, out_dir: Path,
                  label: str, report: dict | None = None, topics: dict | None = None,
                  mapping: dict | None = None, word_settings: dict | None = None) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    weight_col = "총빈도" if args.weight == "count" else "논문수"
    xlsx = out_dir / f"{label}_키프레이즈_빈도.xlsx"
    sheets = {
        "전체": table,
        "구문(2단어 이상)": table[table["단어수"] >= 2],
        "단어(1단어)": table[table["단어수"] == 1],
        "제외됨(일반어 필터)": dropped,      # 전부 기록 (살릴 항목을 찾을 수 있도록)
    }
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        for name, t in sheets.items():
            t = t.reset_index(drop=True)
            t.insert(0, "순위", range(1, len(t) + 1))
            t.to_excel(xw, sheet_name=name, index=False)
            autosize(xw.sheets[name], t)
        s = pd.DataFrame(list(settings.items()), columns=["항목", "값"])
        s.to_excel(xw, sheet_name="설정", index=False)
        autosize(xw.sheets["설정"], s)
        if word_settings:
            # 기본값은 코드 버전으로 특정되므로, '기본값과 달라진 줄'만 기록한다
            diff = settings_diff(word_settings)
            words = pd.DataFrame(diff or [{"설정 파일": "(전부 기본값)", "변경": "", "줄": ""}])
            words.to_excel(xw, sheet_name="내 단어 설정", index=False)
            autosize(xw.sheets["내 단어 설정"], words)
        if report:
            src = pd.DataFrame(report["files"])
            src.loc[len(src)] = {"원본파일": "합계", "DB": "", "시트": "", "원본 건수": report["raw_total"]}
            flow = pd.DataFrame([
                {"단계": "원본 합계", "건수": report["raw_total"]},
                {"단계": "중복 제거 (DOI 일치)", "건수": -report["dup_doi"]},
                {"단계": "중복 제거 (제목 일치)", "건수": -report["dup_title"]},
                {"단계": "통합 목록", "건수": report["after_merge"]},
                {"단계": "연도 필터 적용", "건수": report["after_year"]},
                {"단계": "문헌 유형 필터 적용", "건수": report["after_type"]},
            ])
            src = pd.concat([src, pd.DataFrame([{}]), flow.rename(columns={"단계": "원본파일", "건수": "원본 건수"})],
                            ignore_index=True).fillna("")
            src.to_excel(xw, sheet_name="출처·통합", index=False)
            autosize(xw.sheets["출처·통합"], src)
            rem = pd.DataFrame(report["removed"])
            if not rem.empty:
                rem.to_excel(xw, sheet_name="중복 제거 목록", index=False)
                autosize(xw.sheets["중복 제거 목록"], rem)
        if topics:
            for name, frame in (("토픽", topics["summary"]), ("토픽 추세", topics.get("trends")),
                                ("토픽 간 연결", topics.get("pairs")),
                                ("경계 논문", topics.get("bridges")), ("문헌별 토픽", topics["docs"]),
                                ("주제 지도", topics["themes"])):
                if frame is not None and not frame.empty:
                    frame.to_excel(xw, sheet_name=name, index=False)
                    autosize(xw.sheets[name], frame)
        if mapping:
            nodes = mapping["nodes"]
            nodes.to_excel(xw, sheet_name="논문 지도 좌표", index=False)
            autosize(xw.sheets["논문 지도 좌표"], nodes)

    outputs = [xlsx]
    if mapping:
        names = {}
        if topics is not None:
            names = {row["토픽"]: f"{row['토픽']} {row['상위 키프레이즈'].split(',')[0]}"
                     for _, row in topics["summary"].iterrows()}
        html = out_dir / f"{label}_논문지도.html"
        html.write_text(paper_map_html(mapping, label, names), encoding="utf-8")
        outputs.append(html)
    for suffix, t in (("전체", sheets["전체"]), ("구문", sheets["구문(2단어 이상)"])):
        top = t.head(args.top)
        freqs = defaultdict(float)
        for name, w in zip(top["키프레이즈"], top[weight_col]):
            freqs[name] += w
        if freqs:
            png = out_dir / f"{label}_워드클라우드_{suffix}.png"
            make_wordcloud(dict(freqs), png, args)
            outputs.append(png)
    return outputs


# ─────────────────────────────────────────────────────────────────────────────
# main
# ─────────────────────────────────────────────────────────────────────────────
def parse_args(argv=None):
    ap = argparse.ArgumentParser(description="Scopus 논문 키프레이즈 빈도 집계 + 워드클라우드")
    ap.add_argument("inputs", nargs="+", help="Scopus 내보내기 파일 (.xlsx / .csv). 여러 개 지정 가능")
    ap.add_argument("--sheet", help="읽을 시트 이름 (기본: Title/Abstract/Author Keywords 열이 있는 첫 시트)")
    ap.add_argument("--year", type=int, nargs="*", help="특정 발행연도만 집계 (예: --year 2026)")
    ap.add_argument("--doc-types", choices=["journal", "all"], default="journal",
                    help="journal=저널 논문만(기본), all=학위논문·회의록 등 전부 포함")
    ap.add_argument("--label", help="결과 파일 이름 앞부분 (기본: 입력 파일 이름)")
    ap.add_argument("--out", default=str(BASE_DIR / "output"), help="결과 저장 폴더")
    ap.add_argument("--sections", nargs="+", choices=["title", "abstract", "keywords"],
                    default=["title", "abstract", "keywords"], help="집계할 섹션")
    ap.add_argument("--weight", choices=["count", "papers"], default="count",
                    help="워드클라우드 크기 기준: count=총 노출빈도(기본), papers=등장 논문 수")
    ap.add_argument("--top", type=int, default=100, help="워드클라우드에 표시할 키프레이즈 수 (20~100 권장)")
    ap.add_argument("--min-phrase-papers", type=int, default=3,
                    help="제목/초록에서 자동 구문으로 인정할 최소 등장 논문 수")
    ap.add_argument("--max-phrase-units", type=int, default=4, help="자동 구문의 최대 단어(토큰) 수")
    ap.add_argument("--min-papers", type=int, default=2, help="엑셀 표에 넣을 최소 등장 논문 수")
    ap.add_argument("--min-keyword-ratio", type=float, default=0.05,
                    help="키워드성 필터: 등장 논문 중 저자키워드로 사용한 논문 비율 하한 (0이면 필터 끔)")
    ap.add_argument("--min-keyword-papers", type=int, default=2,
                    help="키워드성 필터: 저자키워드로 사용한 최소 논문 수 (0이면 필터 끔)")
    ap.add_argument("--topics", default="auto",
                    help="토픽 모델: auto=응집도로 토픽 수 자동 선택(기본), off=실행 안 함, 숫자=토픽 수 지정")
    ap.add_argument("--paper-map", choices=["on", "off"], default="on",
                    help="논문 지도(초록 유사도 배치) HTML 생성 여부")
    ap.add_argument("--sensitivity", action="store_true",
                    help="임계값을 바꿔가며 결과가 얼마나 달라지는지 비교표를 만든다 (논문 방법론용)")
    ap.add_argument("--processes", type=int, default=1, help="spaCy 병렬 프로세스 수")
    ap.add_argument("--no-cache", action="store_true", help="품사 태깅 캐시를 사용하지 않음")
    ap.add_argument("--width", type=int, default=2000)
    ap.add_argument("--height", type=int, default=1200)
    ap.add_argument("--background", default="white")
    ap.add_argument("--font", help="워드클라우드 글꼴(.ttf) 경로")
    ap.add_argument("--angles", choices=list(ANGLE_SETS), default="mixed",
                    help="워드클라우드 글자 각도: mixed=다양한 각도(기본), right=가로+세로, horizontal=가로만")
    ap.add_argument("--colormap", help="matplotlib 컬러맵 이름 (예: viridis). 지정하면 기본 12색 팔레트 대신 사용")
    return ap.parse_args(argv)


def prepare_docs(args, settings: dict | None = None, progress=None):
    """파일 읽기 → 전처리 → 품사 태깅 → 표기 통일까지. (집계 직전 상태)
    민감도 분석이 같은 준비 과정을 다시 쓰기 위해 분리해 두었다."""
    df, report = load_records(args.inputs, args.sheet, args.year, args.doc_types)
    if df.empty:
        raise SystemExit("집계할 논문이 없습니다.")
    log(f"      논문 {len(df):,}편 (중복 {report['dup_total']:,}건 제외)")

    log("[2/7] 전처리 중 (저작권 문구·구조화 초록 제목·약어 처리)")
    docs = preprocess(df)
    tag_documents(docs, args.processes, not args.no_cache, progress)

    log("[4/7] 키프레이즈 추출·집계 중")
    vocab = Counter()
    for d in docs:
        for s, _ in d["title_units"] + d["abstract_units"]:
            vocab.update(w.lower() for w in split_words(s))
        for k in d["keywords"]:
            vocab.update(w.lower() for w in split_words(k))
    norm = Normalizer(vocab)
    cfg = Config(norm, settings)
    for d in docs:
        for sec in ("title", "abstract"):
            units = d.pop(f"{sec}_units")
            d[f"{sec}_surfaces"] = [s for s, _ in units]
            d[f"{sec}_classes"] = [c for _, c in units]
            d[f"{sec}_keys"] = [norm.key(s) if c != "B" else "" for s, c in units]
        d["keyword_keys"] = [norm.key(k) for k in d["keywords"]]
    return df, docs, cfg, report


def relax_for_small_corpus(args, n_docs: int) -> list[str]:
    """논문이 적으면 '여러 편에서 반복될 것'을 요구하는 기준이 거의 전부를 걸러낸다.
    예: 10편에서는 저자키워드를 2편 이상에서 단 표현이 거의 없어 표가 3개까지 줄어든다.
    기본값을 그대로 둔 경우에만 1편 기준으로 낮추고, 무엇을 바꿨는지 돌려준다."""
    if n_docs >= SMALL_CORPUS:
        return []
    changed = []
    for name, default, small in (("min_keyword_papers", 2, 1), ("min_papers", 2, 1),
                                 ("min_phrase_papers", 3, 2)):
        if getattr(args, name, None) == default:
            setattr(args, name, small)
            changed.append(f"{name} {default}→{small}")
    if changed:
        log(f"      [작은 표본] 논문 {n_docs}편이라 반복 기준을 낮춥니다: {', '.join(changed)}")
    return changed


def run_pipeline(args, progress=None, settings: dict | None = None) -> dict:
    """전체 분석 실행. CLI(main)와 웹 화면(app.py)이 함께 사용.
    settings를 주면 config 폴더 대신 그 내용(사용자별 단어 설정)을 쓴다."""
    label = args.label or Path(args.inputs[0]).stem
    started = dt.datetime.now()
    word_settings = settings or default_settings()   # 사용자별 단어 설정 (아래 meta와 혼동 금지)
    df, docs, cfg, report = prepare_docs(args, word_settings, progress)
    relaxed = relax_for_small_corpus(args, len(df))

    phrases = build_phrase_dict(docs, cfg, args.min_phrase_papers, args.max_phrase_units)
    stats, forms, doc_keys, doc_kw = count_phrases(docs, cfg, phrases, set(args.sections))
    table, dropped = build_table(stats, forms, cfg, args)
    log(f"      키프레이즈 {len(table):,}개 채택, 일반어 필터로 {len(dropped):,}개 제외")

    suggestions = suggest_cleanup(table, stats, len(df))
    topics = run_topics(doc_keys, table, df, getattr(args, "topics", "auto"))
    mapping = None
    if getattr(args, "paper_map", "on") != "off":
        log("[6/7] 논문 지도 만드는 중")
        try:
            mapping = paper_map(docs, df, topics, doc_keys,
                                dict(zip(table["정규화 키"], table["키프레이즈"])))
        except ImportError as e:
            log(f"      [지도] 라이브러리가 없어 건너뜁니다 ({e}).")

    log("[7/7] 결과 저장 중")
    meta = {
        "입력 파일": " | ".join(f"{r['원본파일']}({r['DB']} {r['원본 건수']:,}건)" for r in report["files"]),
        "시트": args.sheet or "(자동 선택)",
        "연도 필터": ", ".join(map(str, args.year)) if args.year else "(없음)",
        "문헌 유형": "저널 논문만" if args.doc_types == "journal" else "전부 포함",
        "집계 논문 수": len(df),
        "통합 전 원본 건수": report["raw_total"],
        "중복 제거 (DOI 일치)": report["dup_doi"],
        "중복 제거 (제목 일치)": report["dup_title"],
        "초록 없는 논문": sum(1 for d in docs if not d["abstract"]),
        "저자키워드 없는 논문": sum(1 for d in docs if not d["keywords"]),
        "집계 섹션": ", ".join(SECTION_LABELS[s] for s in args.sections),
        "워드클라우드 기준": "총빈도(노출 횟수)" if args.weight == "count" else "논문수",
        "워드클라우드 표시 개수": args.top,
        "워드클라우드 각도": args.angles,
        "워드클라우드 색상": args.colormap or "다채로운 색 (12색)",
        "자동 구문 최소 논문 수": args.min_phrase_papers,
        "자동 구문 최대 단어 수": args.max_phrase_units,
        "표 포함 최소 논문 수": args.min_papers,
        "키워드성 필터: 저자키워드 사용률 하한": args.min_keyword_ratio,
        "키워드성 필터: 저자키워드 사용 최소 논문 수": args.min_keyword_papers,
        "구문 사전 크기": len(phrases),
        "실행 일시": started.strftime("%Y-%m-%d %H:%M"),
        "작은 표본 조정": ", ".join(relaxed) if relaxed else "없음",
        "코드 버전": code_version(),
        "spaCy 모델": SPACY_MODEL,
    }
    if topics:
        meta["토픽 수"] = topics["k"]
    if mapping:
        meta["논문 지도"] = f"{len(mapping['nodes']):,}편 · 연결 {len(mapping['edges']):,}개"
    outputs = write_outputs(table, dropped, meta, args, Path(args.out) / label, label, report, topics, mapping,
                            word_settings)
    return {
        "label": label, "table": table, "dropped": dropped, "settings": meta, "outputs": outputs,
        "report": report, "topics": topics, "map": mapping, "suggest": suggestions, "relaxed": relaxed,
        "weight_col": "총빈도" if args.weight == "count" else "논문수",
        "seconds": (dt.datetime.now() - started).seconds,
    }


def main():
    args = parse_args()
    if args.sensitivity:
        table, scored, best = sensitivity_table(args)
        label = args.label or Path(args.inputs[0]).stem
        out_dir = Path(args.out) / label
        out_dir.mkdir(parents=True, exist_ok=True)
        path = out_dir / f"{label}_민감도.xlsx"
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            for name, frame in (("민감도", table), ("임계값 추천", scored)):
                frame.to_excel(xw, sheet_name=name, index=False)
                autosize(xw.sheets[name], frame)
        print("\n임계값 민감도")
        print(table.to_string(index=False, max_colwidth=46))
        print("\n임계값 추천 — 논문 절반으로 고르고 나머지 절반의 저자키워드로 채점")
        print(scored.to_string(index=False))
        print(f"\n추천 일반어 필터: {'끔' if best == 0 else f'{best * 100:.0f}%'}"
              f"  (현재 설정 {args.min_keyword_ratio * 100:.0f}%)")
        print("저장:", path)
        return
    res = run_pipeline(args)
    table, weight_col = res["table"], res["weight_col"]
    print(f"\n상위 30개 키프레이즈 ({weight_col} 기준)")
    for i, r in table.head(30).iterrows():
        print(f"  {i + 1:>3}. {r['키프레이즈']:<40} 총빈도 {r['총빈도']:>6,}  논문 {r['논문수']:>5,}")
    print("\n저장된 파일:")
    for p in res["outputs"]:
        print("  -", p)
    print(f"완료 ({res['seconds']}초)")


if __name__ == "__main__":
    main()
