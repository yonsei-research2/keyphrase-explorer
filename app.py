#﻿논문 키프레이즈 워드클라우드 — app.py (웹 화면)
#코드 버전 a73b136 · 2026-10-28
#==============================================================================

# -*- coding: utf-8 -*-
"""
웹 화면 — '웹화면_실행.bat'을 더블클릭하면 브라우저에서 열립니다.
(직접 실행: python -m streamlit run app.py)
"""
import json
import os
import re
import shutil
import time
import uuid
from pathlib import Path

import pandas as pd
import streamlit as st

import keyphrase_cloud as kc

UPLOAD_DIR = kc.BASE_DIR / "uploads"
KEEP_HOURS = 12  # 이보다 오래된 다른 세션의 파일은 지운다 (서버에 남지 않도록)
SECTION_OPTIONS = {"제목": "title", "초록": "abstract", "저자키워드": "keywords"}
ANGLE_OPTIONS = {"mixed": "다양한 각도", "right": "가로 + 세로", "horizontal": "가로만"}
CONFIG_FILES = {
    "제외할 키프레이즈": ("exclude_phrases.txt", "결과에서 **완전히 뺄** 항목을 한 줄에 하나씩 적습니다. (예: `Korea`)"),
    "같은 뜻으로 합치기": ("synonyms.txt", "형식: `변형표기 => 대표표기` — 여러 개는 `|`로 구분 (예: `SARS-CoV-2 | COVID-19 pandemic => COVID-19`)"),
    "항상 포함할 구문": ("custom_phrases.txt", "하나의 구문으로 인식시키고, 일반어 필터와 관계없이 **항상 결과에 포함**할 항목"),
    "1단어 일반어": ("stopwords.txt", "1단어로는 집계하지 않을 일반 단어 (다른 단어와 결합한 구문의 일부로는 집계됨)"),
    "자동 구문 수식어": ("modifiers.txt", "자동 추출 구문의 **맨 앞**에 올 수 없는 수식어 (`novel`, `significant` …)"),
    "자동 구문 끝 일반명사": ("trailing_generic.txt", "자동 추출 구문의 **맨 끝**에 올 수 없는 일반 명사 (`model`, `approach` …)"),
}

st.set_page_config(page_title="논문 키프레이즈 워드클라우드", page_icon="☁️", layout="wide")


# ── 도우미 ────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="파일 읽고 통합하는 중…")
def inspect_files(paths: tuple, stamps: tuple, doc_types: str):
    """여러 DB 파일을 읽어 중복을 제거한 통합 목록과 처리 보고를 돌려준다."""
    df, report = kc.load_records(list(paths), None, None, doc_types)
    dist = {y: c for y, c in sorted(df[kc.COL_YEAR].value_counts().to_dict().items()) if y}
    return df, report, dist


def session_id() -> str:
    """브라우저 세션마다 다른 짧은 식별자 — 여러 사람이 같은 앱을 써도 파일이 섞이지 않게."""
    if "sid" not in st.session_state:
        st.session_state["sid"] = uuid.uuid4().hex[:10]
    return st.session_state["sid"]


def sweep_old_sessions() -> None:
    """오래된 다른 세션 폴더 삭제 — 올린 논문 파일이 서버에 계속 남지 않게."""
    cutoff = time.time() - KEEP_HOURS * 3600
    for kind in ("uploads", "output"):
        root = kc.BASE_DIR / kind
        if not root.is_dir():
            continue
        for child in root.iterdir():
            if not child.is_dir() or child.name == session_id():
                continue
            try:
                if child.stat().st_mtime < cutoff:
                    shutil.rmtree(child, ignore_errors=True)
            except OSError:
                pass


def session_dir(kind: str) -> Path:
    """이 세션 전용 업로드/결과 폴더."""
    path = kc.BASE_DIR / kind / session_id()
    path.mkdir(parents=True, exist_ok=True)
    if not st.session_state.get("swept"):
        st.session_state["swept"] = True
        sweep_old_sessions()
    return path


def save_upload(f) -> Path:
    path = session_dir("uploads") / Path(f.name).name
    data = f.getvalue()
    if not path.exists() or path.stat().st_size != len(data) or path.read_bytes() != data:
        path.write_bytes(data)
    return path


def settings() -> dict:
    """사용자별 단어 설정 (브라우저 세션에만 남고, 파일로 내려받아 다시 쓸 수 있음)."""
    if "settings" not in st.session_state:
        st.session_state["settings"] = kc.default_settings()
    return st.session_state["settings"]


def append_config_lines(fname: str, items: list[str]) -> int:
    """내 설정에 줄을 덧붙인다 (프로그램 폴더의 파일은 건드리지 않음)."""
    cfg = settings()
    text = cfg.get(fname, "")
    existing = {line.strip().lower() for line in text.splitlines()}
    new = [i for i in items if i.strip().lower() not in existing]
    if new:
        cfg[fname] = text.rstrip("\n") + "\n" + "\n".join(new) + "\n"
    return len(new)


def safe_label(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip() or "result"


# ── 분석 실행 ─────────────────────────────────────────────────────────────────
def run_analysis(paths, years, label, sections, weight, top, ratio_pct, min_phrase, angles, doc_types, topics):
    argv = [
        *[str(p) for p in paths], "--label", label, "--top", str(top), "--angles", angles,
        "--doc-types", doc_types, "--topics", topics, "--out", str(session_dir("output")),
        "--weight", weight, "--min-keyword-ratio", str(ratio_pct / 100),
        "--min-phrase-papers", str(min_phrase), "--sections", *[SECTION_OPTIONS[s] for s in sections],
    ]
    if years:
        argv += ["--year", *years]
    args = kc.parse_args(argv)
    if ratio_pct == 0:
        args.min_keyword_papers = 0

    logs = []
    with st.status("분석 중… 처음 분석하는 파일은 3~5분 정도 걸립니다.", expanded=True) as status:
        bar = st.progress(0.0, text="준비 중")
        log_box = st.empty()

        def on_log(msg):
            logs.append(msg)
            log_box.code("\n".join(logs[-10:]), language=None)

        def on_progress(done, total):
            bar.progress(done / total, text=f"문장 분석(품사 태깅) {done:,} / {total:,}")

        st.session_state["last_args"] = args
        kc.set_logger(on_log)
        try:
            res = kc.run_pipeline(args, progress=on_progress, settings=settings())  # 설정도 엑셀에 기록됨
        except SystemExit as e:
            status.update(label=f"오류: {e}", state="error")
            return
        except Exception as e:  # noqa: BLE001
            status.update(label="오류가 발생했습니다", state="error")
            st.exception(e)
            return
        finally:
            kc.set_logger(None)
        bar.progress(1.0, text="완료")
        status.update(label=f"분석 완료 ({res['seconds']}초)", state="complete", expanded=False)
    st.session_state["result"] = res


# ── 결과 표시 ─────────────────────────────────────────────────────────────────
def render_result():
    res = st.session_state.get("result")
    if not res:
        return
    table, dropped = res["table"], res["dropped"]
    outputs = [Path(p) for p in res["outputs"]]

    st.divider()
    st.subheader(f"3. 결과 — {res['label']}")
    m1, m2, m3 = st.columns(3)
    m1.metric("분석한 논문", f"{res['settings']['집계 논문 수']:,}편")  # 중복 제거 후
    m2.metric("키프레이즈", f"{len(table):,}개")
    m3.metric("그중 2단어 이상 구문", f"{int((table['단어수'] >= 2).sum()):,}개")

    pngs = [p for p in outputs if p.suffix == ".png"]
    for tab, p in zip(st.tabs(["☁️ 전체 (단어 + 구문)" if p.stem.endswith("전체") else "☁️ 구문만 (2단어 이상)"
                               for p in pngs]), pngs):
        with tab:
            st.image(str(p), width="stretch")

    st.markdown("**내려받기**")
    cols = st.columns(len(outputs) + 1)
    for col, p in zip(cols, outputs):
        if p.suffix == ".xlsx":
            name, mime = "📥 엑셀 표", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            name = "🖼️ 전체 워드클라우드" if p.stem.endswith("전체") else "🖼️ 구문 워드클라우드"
            mime = "image/png"
        col.download_button(name, data=p.read_bytes(), file_name=p.name, mime=mime, key=f"dl_{p.name}")
    if hasattr(os, "startfile") and cols[-1].button("📂 결과 폴더 열기"):  # 내 컴퓨터에서 실행할 때만
        os.startfile(outputs[0].parent)

    st.markdown("**키프레이즈 표**")
    c1, c2 = st.columns([2, 3])
    query = c1.text_input("검색", placeholder="예: cancer", key="search")
    kind = c2.radio("보기", ["전체", "구문(2단어 이상)", "단어(1단어)"], horizontal=True, key="kind")
    view = table
    if kind.startswith("구문"):
        view = view[view["단어수"] >= 2]
    elif kind.startswith("단어"):
        view = view[view["단어수"] == 1]
    if query:
        hit = view["병합된 표기"].str.contains(query, case=False, regex=False)
        view = view[hit | view["키프레이즈"].str.contains(query, case=False, regex=False)]
    st.dataframe(to_display(view), height=460)

    with st.expander("✂️ 결과에서 빼고 싶은 키프레이즈가 있나요?"):
        picks = st.multiselect("제외할 항목 선택 (상위 300개 중)", table["키프레이즈"].head(300).tolist(), key="picks")
        if st.button("제외 목록에 추가", disabled=not picks):
            n = append_config_lines("exclude_phrases.txt", picks)
            st.success(f"{n}개를 **내 설정**의 제외 목록에 넣었습니다. 위의 **분석 시작**을 다시 누르면 반영됩니다 (수십 초).")

    if res.get("relaxed"):
        st.info(f"**논문이 {res['settings']['집계 논문 수']}편이라 작은 표본용 기준을 적용했습니다.** "
                "‘여러 편에서 반복될 것’을 요구하는 조건(저자키워드 2편 이상, 표 포함 2편 이상, 구문 3편 이상)을 "
                "각각 1·1·2편으로 낮췄습니다. 이 조건을 그대로 두면 10편 규모에서는 거의 전부 걸러집니다. "
                "고급 설정에서 직접 지정하시면 그 값이 우선합니다.", icon="📉")

    render_suggestions()

    merged_rows = table[table["병합된 표기"].str.contains(";", na=False)]
    with st.expander(f"🔗 자동으로 합쳐진 표기 ({len(merged_rows):,}개 키프레이즈)"):
        st.caption("대소문자·하이픈·소유격·단수복수 차이를 하나로 합친 내역입니다(한 항목당 최대 8개 표기 표시). "
                   "엑셀 표의 **병합된 표기** 열에도 같은 내용이 있습니다. "
                   "잘못 합쳐진 게 보이면 옆칸(⚙️ 내 단어 설정)의 ‘항상 포함할 구문’에 원하는 표기를 적어 분리할 수 있습니다.")
        view = merged_rows[["키프레이즈", "병합된 표기", "논문수", "총빈도"]].reset_index(drop=True)
        view.index += 1
        st.dataframe(view, height=360)

    with st.expander("📐 임계값 민감도 — 필터 기준을 바꾸면 결과가 얼마나 달라지나"):
        st.caption("일반어 필터 비율과 자동 구문 기준을 바꿔가며 다시 집계해 비교합니다. "
                   "논문 방법론에 “임계값 5%를 썼고 2~10% 구간에서 상위 결과는 대체로 같았다”처럼 쓸 수 있습니다. "
                   "품사 태깅은 다시 하지 않지만 집계를 여러 번 반복하므로 1~2분 걸립니다.")
        if st.button("민감도 표 만들기", key="run_sensitivity", disabled="last_args" not in st.session_state):
            with st.spinner("임계값을 바꿔가며 다시 집계하는 중…"):
                kc.set_logger(None)
                st.session_state["sensitivity"] = kc.sensitivity_table(st.session_state["last_args"], settings())
        packed = st.session_state.get("sensitivity")
        if packed is not None:
            sens, scored, best = packed
            st.dataframe(sens, hide_index=True, width="stretch")
            st.caption("**상위 20 일치**는 기준 설정의 상위 20개 중 몇 개가 그대로 남았는지, "
                       "**전체 겹침(자카드)**은 키프레이즈 집합이 얼마나 같은지(1이면 동일)를 뜻합니다.")

            st.markdown("**임계값 추천** — 논문을 무작위 절반으로 나눠 **한쪽만 보고** 필터를 적용하고, "
                        "**다른 쪽의 저자키워드**로 채점한 결과입니다 (5회 반복 평균).")
            st.dataframe(scored, hide_index=True, width="stretch")
            pct = int(round(best * 100))
            same = abs(best - st.session_state["last_args"].min_keyword_ratio) < 1e-9
            if same:
                st.success(f"현재 설정({pct}%)이 이 데이터에서 가장 좋습니다.")
            else:
                st.info(f"이 데이터에서 가장 좋게 나온 값은 **{pct} %** 입니다. "
                        f"(현재 {int(round(st.session_state['last_args'].min_keyword_ratio * 100))}%). "
                        "F1 차이가 편차보다 작으면 둘 중 아무거나 써도 무방합니다.")
                if st.button(f"{pct}%로 바꾸기 (고급 설정 반영)", key="apply_ratio"):
                    st.session_state["ratio"] = pct
                    st.rerun()
            st.download_button("📥 민감도 표 내려받기 (CSV)", data=sens.to_csv(index=False).encode("utf-8-sig"),
                               file_name=f"{res['label']}_민감도.csv", mime="text/csv", key="dl_sens")

    with st.expander(f"🔍 일반어 필터로 제외된 항목 보기 ({len(dropped):,}개)"):
        st.caption("초록에는 자주 나오지만 저자키워드로는 거의 쓰이지 않는 일반어입니다. "
                   "살리고 싶은 항목은 아래에서 선택하거나 옆칸(⚙️ 내 단어 설정)의 ‘항상 포함할 구문’에 추가하세요.")
        keep = st.multiselect("항상 포함할 항목으로 추가", dropped["키프레이즈"].head(300).tolist(), key="keep")
        if st.button("항상 포함 목록에 추가", disabled=not keep):
            n = append_config_lines("custom_phrases.txt", keep)
            st.success(f"{n}개를 **내 설정**에 넣었습니다. **분석 시작**을 다시 누르면 반영됩니다.")
        st.dataframe(to_display(dropped.head(500)), height=360)


def to_display(df):
    out = df.drop(columns=["정규화 키"]).reset_index(drop=True)
    out["저자키워드 사용률"] = (out["저자키워드 사용률"] * 100).map(
        lambda v: "" if pd.isna(v) else f"{v:.1f}%")
    out.index += 1
    return out


# ── 탭 1: 분석 ────────────────────────────────────────────────────────────────
def render_run_tab():
    st.subheader("1. 파일 올리기")
    uploaded = st.file_uploader(
        "Scopus · Web of Science · ProQuest에서 내보낸 논문 목록 (.xlsx · .xls · .csv) — 여러 개 한꺼번에 가능",
        type=["xlsx", "xls", "csv"], accept_multiple_files=True)
    with st.expander("또는 컴퓨터에 있는 파일 경로로 불러오기 (여러 개면 한 줄에 하나씩)"):
        typed = st.text_area("파일 경로", key="typed_path", height=80,
                             placeholder="D:\\논문자료\\scopus.csv\nD:\\논문자료\\savedrecs.xls")

    paths = [save_upload(f) for f in uploaded] if uploaded else []
    for line in (typed or "").splitlines():
        line = line.strip().strip('"')
        if not line:
            continue
        if Path(line).is_file():
            paths.append(Path(line))
        else:
            st.error(f"해당 경로에 파일이 없습니다: {line}")
    if not paths:
        st.info("파일을 올리면 통합 결과와 **분석 시작** 버튼이 나타납니다. "
                "여러 DB에서 받은 파일을 함께 올리면 중복 문헌을 제거한 뒤 분석합니다.")
        st.caption(f"올린 파일과 결과는 **내 세션 전용 폴더**에 저장되어 다른 사용자와 섞이지 않고, "
                   f"{KEEP_HOURS}시간이 지나면 서버에서 자동으로 지워집니다.")
        render_result()
        return

    doc_types = "journal" if st.radio(
        "문헌 유형", ["저널 논문만", "전부 포함 (학위논문·회의록 등)"], horizontal=True,
        help="ProQuest에는 학위논문이 섞여 있는 경우가 많습니다.") == "저널 논문만" else "all"

    try:
        merged, report, year_dist = inspect_files(tuple(str(p) for p in paths),
                                                  tuple(p.stat().st_mtime for p in paths), doc_types)
    except SystemExit as e:
        st.error(f"파일을 읽을 수 없습니다: {e}")
        return

    st.markdown("**통합 결과**")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("올린 파일", f"{len(report['files'])}개")
    m2.metric("원본 합계", f"{report['raw_total']:,}편")
    m3.metric("중복 제거", f"{report['dup_total']:,}편",
              help=f"DOI 일치 {report['dup_doi']:,}편 · 제목 일치 {report['dup_title']:,}편")
    m4.metric("분석 대상", f"{len(merged):,}편",
              help="중복 제거 후 문헌 유형 필터까지 적용한 건수")
    st.dataframe(pd.DataFrame(report["files"]), hide_index=True, width="stretch")
    if report["dup_total"]:
        with st.expander(f"🔁 중복으로 제거된 {report['dup_total']:,}건 보기"):
            st.caption("DOI가 같거나, 정규화한 제목이 완전히 같고 발행연도가 1년 이내인 경우 같은 문헌으로 봅니다. "
                       "빠진 레코드의 초록·키워드가 더 충실하면 그 내용으로 채워 넣습니다.")
            st.dataframe(pd.DataFrame(report["removed"]), hide_index=True, height=300)
    if doc_types == "journal" and report["after_merge"] != report["after_type"]:
        st.caption(f"문헌 유형 필터로 {report['after_merge'] - report['after_type']:,}건 제외 "
                   f"(유형 분포: {', '.join(f'{k} {v:,}' for k, v in report['type_counts'].items())})")

    render_pre_suggestions(merged)

    st.subheader("2. 옵션")
    c1, c2 = st.columns(2)
    with c1:
        years = []
        if year_dist:
            years = st.multiselect("발행연도", list(year_dist), default=list(year_dist))
        stem = paths[0].stem if len(paths) == 1 else "통합"
        default_label = stem if not years or len(years) == len(year_dist) else f"{stem}_{'-'.join(years)}"
        label = safe_label(st.text_input("결과 이름 (파일 이름에 사용)", value=default_label))
    with c2:
        sections = st.multiselect("집계할 부분", list(SECTION_OPTIONS), default=list(SECTION_OPTIONS))
        weight = st.radio("워드클라우드 글자 크기 기준", ["count", "papers"], horizontal=True,
                          format_func={"count": "노출빈도 (총 등장 횟수)", "papers": "논문수 (등장한 논문 수)"}.get)
        angles = st.radio("워드클라우드 글자 각도", list(ANGLE_OPTIONS), horizontal=True, format_func=ANGLE_OPTIONS.get)
    with st.expander("고급 설정"):
        top = st.slider("워드클라우드에 표시할 키프레이즈 수", 20, 100, 100, 10,
                help="적게 둘수록 글자가 크고 읽기 쉬워집니다. 발표 슬라이드용이면 20~50개를 권합니다.")
        ratio = st.slider("일반어 필터 강도 — 저자키워드 사용률 하한 (%)", 0, 30, 5, key="ratio",
                          help="`risk`·`approach`처럼 초록에 흔하지만 주제어가 아닌 말을 거르는 기준입니다. "
                               "그 표현이 등장한 논문 중 저자가 키워드로도 단 논문의 비율이 이 값 이상이면 남깁니다. "
                               "0이면 필터를 끄고, 높일수록 깐깐해집니다. "
                               "이 데이터에 맞는 값은 위쪽 ‘📏 일반어 필터 강도 추천’에서 가늠할 수 있습니다.")
        min_phrase = st.number_input("제목·초록에서 구문으로 인정할 최소 논문 수", 2, 20, 3,
                                     help="데이터가 적을 때(예: 몇 백 편)는 2로 낮추세요.")
        topic_mode = st.radio("토픽 모델", ["자동 (응집도로 토픽 수 선택)", "끄기"], horizontal=True,
                              help=f"논문 {kc.TOPIC_MIN_DOCS}편 이상일 때만 실행됩니다. "
                                   "결과는 ‘🧭 토픽’ 탭에서 토픽 수를 바꿔가며 볼 수 있습니다.")

    if st.button("🚀 분석 시작", type="primary", disabled=not sections):
        run_analysis(paths, years if len(years) != len(year_dist) else [], label, sections, weight, top, ratio,
                     int(min_phrase), angles, doc_types, "auto" if topic_mode.startswith("자동") else "off")
    render_result()


# ── 탭 2: 토픽 ────────────────────────────────────────────────────────────────
def render_topic_tab():
    res = st.session_state.get("result")
    topics = (res or {}).get("topics")
    if not topics:
        st.info(f"‘📊 분석’ 탭에서 분석을 실행하면 토픽 결과가 여기에 나옵니다. "
                f"논문이 {kc.TOPIC_MIN_DOCS}편 이상이어야 하고, 고급 설정에서 토픽 모델을 켜 두어야 합니다.")
        return

    n_docs = len(topics["docs"])
    if n_docs < kc.TOPIC_SMALL_DOCS:
        st.warning(f"**표본이 적습니다 — 논문 {n_docs}편.** 토픽 모델은 보통 수백 편 이상에서 안정적이고, "
                   f"{kc.TOPIC_SMALL_DOCS}편 미만에서는 몇 편만 바뀌어도 토픽 구성이 달라집니다. "
                   "응집도로 고른 토픽 수도 믿기 어려우니 **직접 토픽 수를 바꿔 보시고, 결과는 참고용으로**만 쓰세요. "
                   "논문에 싣는다면 표본 수와 함께 한계를 밝히셔야 합니다.", icon="⚠️")
    st.caption("집계된 키프레이즈를 입력으로 논문을 몇 개의 주제로 묶습니다(NMF). "
               "응집도(NPMI)가 가장 높은 토픽 수를 추천하지만, 지표가 연구 목적과 늘 맞지는 않으니 직접 바꿔 보세요.")
    curve = topics["curve"]
    if not curve.empty:
        c1, c2 = st.columns([2, 3])
        c1.metric("추천 토픽 수", f"{topics['k']}개")
        c2.line_chart(curve.set_index("토픽 수"), height=160)

    limit = int(min(min(topics["matrix"].shape), 30))  # 문헌 수·키프레이즈 수를 넘을 수 없음
    k = st.slider("토픽 수", 2, max(3, limit), min(int(topics["k"]), limit), key="topic_k")
    if k != topics["k"]:
        with st.spinner("토픽 다시 계산 중…"):
            weights, components, scores = kc.fit_topics(topics["matrix"], k)
            summary, docs_tab, _ = kc.topic_tables(weights, components, scores, topics["keys"],
                                                   topics["labels"], topics["papers"])
    else:
        summary, docs_tab = topics["summary"], topics["docs"]

    st.markdown("**토픽별 요약**")
    st.dataframe(summary, hide_index=True, width="stretch")
    st.caption("`평균 인용`은 인용 자료가 있는 논문만으로 계산합니다(ProQuest에는 인용수 열이 없음). "
               "토픽 간 비교는 같은 기간·같은 DB 안에서만 의미가 있습니다.")

    st.markdown("**어떤 주제가 뜨고 지는가 — 연도별 점유율**")
    st.caption("편수 그대로 보면 문헌 전체가 늘어난 효과와 섞입니다. 그래서 **그 해 논문 중 차지하는 비율**로 봅니다. "
               "추세 판정은 Mann-Kendall 검정(비모수)이며, 연도가 5년 미만이면 판단하지 않습니다.")
    counts = (docs_tab.assign(연도=pd.to_numeric(docs_tab["연도"], errors="coerce"))
              .dropna(subset=["연도"]).pivot_table(index="연도", columns="토픽", values="제목", aggfunc="count")
              .fillna(0))
    if len(counts) > 1:
        mode = st.radio("보기", ["점유율(%)", "논문 수"], horizontal=True, key="topic_trend_mode",
                        label_visibility="collapsed")
        shown = counts.copy()
        shown.index = shown.index.astype(int).astype(str)   # 연도에 천 단위 쉼표가 붙지 않도록
        if mode == "점유율(%)":
            st.line_chart(100 * shown.div(shown.sum(axis=1).replace(0, 1), axis=0), height=260)
        else:
            st.bar_chart(shown, height=260)
        trends = topics.get("trends") if k == topics["k"] and topics.get("trends") is not None             else kc.topic_trends(docs_tab)
        if trends is not None and not trends.empty:
            shown_trends = trends.drop(columns=[c for c in trends.columns
                                                if trends[c].isna().all() or (trends[c] == "").all()])
            st.dataframe(shown_trends, hide_index=True, width="stretch")
            if (trends["추세"] == "판단 보류 (연도 5년 미만)").all():
                st.caption(f"연도가 {len(counts)}년치뿐이라 추세 검정은 하지 않았습니다 "
                           f"({kc.TREND_MIN_YEARS}년 이상 필요). 위 그래프와 점유율 수치는 그대로 보실 수 있습니다.")
            up = trends[trends["추세"] == "증가"]["토픽"].tolist()
            down = trends[trends["추세"] == "감소"]["토픽"].tolist()
            if up or down:
                st.info(f"**증가 추세**: {', '.join(up) or '없음'}  ·  **감소 추세**: {', '.join(down) or '없음'}  "
                        f"(p < 0.05)")
    else:
        st.caption("연도가 한 해뿐이라 추이를 그리지 않습니다.")

    stale = [col for col in ("인용수", "2순위 토픽") if col not in docs_tab.columns]
    if stale:
        st.info("이 결과는 이전 버전으로 만든 것이라 일부 항목이 없습니다. "
                "‘📊 분석’ 탭에서 **분석 시작**을 다시 누르면 전부 나옵니다.")

    st.markdown("**토픽별 대표 논문**")
    st.caption("리뷰에 인용할 논문을 고를 때 씁니다. `대표도`는 그 논문이 해당 토픽에 얼마나 치우쳐 있는지입니다.")
    pick_t = st.selectbox("토픽", sorted(docs_tab["토픽"].unique()), key="topic_rep")
    mine = docs_tab[docs_tab["토픽"] == pick_t]
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("가장 전형적인 논문 (토픽 비중 순)")
        st.dataframe(mine.nlargest(5, "토픽 비중")[["제목", "연도", "토픽 비중", "DOI"]]
                     .rename(columns={"토픽 비중": "대표도"}), hide_index=True, height=220)
    with c2:
        st.markdown("가장 많이 인용된 논문")
        cited = mine.dropna(subset=["인용수"]) if "인용수" in mine.columns else mine.iloc[0:0]
        if cited.empty:
            st.caption("인용 자료가 있는 논문이 없습니다.")
        else:
            st.dataframe(cited.nlargest(5, "인용수")[["제목", "연도", "인용수", "DOI"]],
                         hide_index=True, height=220)

    pairs, bridges = topics.get("pairs"), topics.get("bridges")
    if pairs is not None and not pairs.empty and k == topics["k"]:
        with st.expander("🔗 두 토픽에 걸친 논문 — 주제 경계에서 나온 연구"):
            st.caption("1순위 토픽 비중의 절반 이상을 2순위 토픽이 차지하는 논문입니다. "
                       "분야가 섞이는 지점을 찾거나, 토픽 수가 적절한지 가늠할 때 씁니다.")
            cc1, cc2 = st.columns([1, 2])
            cc1.dataframe(pairs, hide_index=True, height=260)
            cc2.dataframe(bridges[["제목", "연도", "1순위", "2순위", "2순위 비중"]].head(30),
                          hide_index=True, height=260)

    themes = topics["themes"]
    if themes is not None and not themes.empty:
        st.markdown("**주제 지도 — 중심성(다른 주제와의 연결) × 밀도(주제 내부 응집)**")
        st.scatter_chart(themes, x="중심성", y="밀도", color="사분면", size="키프레이즈 수", height=320)
        st.dataframe(themes, hide_index=True, width="stretch")

    with st.expander(f"📄 문헌별 토픽 배정 ({len(docs_tab):,}편)"):
        pick = st.selectbox("토픽 고르기", ["전체"] + sorted(docs_tab["토픽"].unique()), key="topic_pick")
        view = docs_tab if pick == "전체" else docs_tab[docs_tab["토픽"] == pick]
        st.dataframe(view, hide_index=True, height=360)
    st.caption("같은 내용이 엑셀의 ‘토픽’·‘토픽 추세’·‘토픽 간 연결’·‘경계 논문’·‘문헌별 토픽’·‘주제 지도’ 시트에도 저장됩니다.")


# ── 탭 3: 논문 지도 ───────────────────────────────────────────────────────────
def render_map_tab():

    res = st.session_state.get("result")
    mapping = (res or {}).get("map")
    if not mapping:
        st.info(f"‘📊 분석’ 탭에서 분석을 실행하면 논문 지도가 여기에 나옵니다. "
                f"논문 {kc.MAP_MIN_DOCS}편 이상이어야 합니다.")
        return

    st.caption("초록·제목·키워드의 단어 구성이 비슷한 논문끼리 가깝게 놓고, 가장 비슷한 논문끼리 선으로 이었습니다. "
               "인용 관계가 아니라 **내용 유사도**입니다. 점 색은 토픽, 크기는 인용수(Scopus·WoS 기준)입니다.")
    st.caption("**논문을 누르면** 그 논문이 기준이 되어, 내용이 가장 비슷한 논문 목록(유사도 %)과 연결선이 나타납니다. 목록의 제목을 누르면 다시 그 논문 기준으로 옮겨 가며 따라갈 수 있습니다. **거품을 누르면** 그 영역의 논문 목록이 나옵니다.")
    if len(mapping["nodes"]) < kc.TOPIC_SMALL_DOCS:
        st.warning(f"**논문 {len(mapping['nodes'])}편으로 그린 지도입니다.** 점이 적으면 덩어리가 잘 생기지 않고, "
                   "‘가장 비슷한 논문’의 유사도도 20% 안팎까지 내려갈 수 있습니다(비슷해서가 아니라 "
                   "그보다 비슷한 논문이 없어서입니다). 목록의 유사도 값을 함께 보고 판단하세요.", icon="⚠️")
    st.warning("**해석 주의** — 가까이 있으면 내용이 비슷한 것이 맞지만, **멀리 떨어진 두 덩어리 사이의 거리나 "
               "덩어리의 크기·모양은 의미가 없습니다**(t-SNE 배치의 성질). 어떤 주제가 더 크다/가깝다는 결론은 "
               "토픽 탭의 논문 수로 판단하세요.", icon="⚠️")
    missing = mapping.get("인용수 없는 논문", 0)
    if missing:
        st.caption(f"인용수 자료가 없는 논문 {missing:,}편(주로 ProQuest)은 **중립 크기**로 그렸습니다. "
                   "작은 점이 '인용 0'을 뜻하지 않도록 구분했고, 마우스를 올리면 '인용 정보 없음'으로 표시됩니다.")
    nodes = mapping["nodes"]
    m1, m2, m3 = st.columns(3)
    m1.metric("논문", f"{len(nodes):,}편")
    m2.metric("연결선", f"{len(mapping['edges']):,}개")
    m3.metric("색으로 구분한 토픽", f"{nodes['토픽'].nunique()}개")

    key = f"map_html::{res['label']}::{len(nodes)}"
    if key not in st.session_state:
        names = {}
        if res.get("topics") is not None:
            names = {row["토픽"]: f"{row['토픽']} {row['상위 키프레이즈'].split(',')[0]}"
                     for _, row in res["topics"]["summary"].iterrows()}
        st.session_state[key] = kc.paper_map_html(mapping, res["label"], names)
    html = st.session_state[key]
    # st.iframe은 URL만 받고 inline HTML을 받지 못한다 (지도는 자체 완결 HTML이라 이 방식 유지)
    import streamlit.components.v1 as components
    components.html(html, height=720, scrolling=False)
    st.download_button("📥 지도 HTML 내려받기 (혼자서도 열립니다)", data=html.encode("utf-8"),
                       file_name=f"{res['label']}_논문지도.html", mime="text/html", key="dl_map")
    st.caption("왼쪽 위에서 보기를 바꿉니다 — **거품**(원을 누르면 논문 목록) · **분포**(유사도 좌표 그대로) · "
               "**연결망**(비슷한 논문끼리 선) · **밀도**(몰린 곳 열지도). "
               "거품 보기의 **원 기준**은 **토픽**(색과 1:1 일치)과 **위치**(지도에서 가까운 것끼리) 중 선택합니다. "
               "휠로 확대하면 논문 제목이 나타나고, 범례를 누르면 토픽을 숨기며, 검색창으로 제목을 거릅니다.")


# ── 탭 4: 단어 설정 ───────────────────────────────────────────────────────────
@st.cache_data(show_spinner="필터 강도를 데이터로 가늠하는 중…")
def _pre_threshold(keywords: tuple, titles: tuple, abstracts: tuple, stopwords: str, modifiers: str):
    """분석 전 일반어 필터 강도 추천 (추정치)"""
    frame = pd.DataFrame({kc.COL_KEYWORDS: list(keywords), kc.COL_TITLE: list(titles),
                          kc.COL_ABSTRACT: list(abstracts)})
    return kc.recommend_ratio_quick(frame, settings={**settings(), "stopwords.txt": stopwords,
                                                     "modifiers.txt": modifiers})


@st.cache_data(show_spinner=False)
def _pre_suggestions(keywords: tuple, titles: tuple, abstracts: tuple):
    """올린 파일 내용으로 분석 전 추천 (캐시 — 같은 파일이면 다시 계산하지 않음)"""
    frame = pd.DataFrame({kc.COL_KEYWORDS: list(keywords), kc.COL_TITLE: list(titles),
                          kc.COL_ABSTRACT: list(abstracts)})
    return kc.suggest_before_analysis(frame)


def render_pre_suggestions(merged: pd.DataFrame) -> None:
    """분석을 시작하기 전에 '이건 빼는 게 좋겠다'를 근거와 함께 보여 준다."""
    sug = _pre_suggestions(tuple(merged[kc.COL_KEYWORDS]), tuple(merged[kc.COL_TITLE]),
                           tuple(merged[kc.COL_ABSTRACT]))
    total = len(sug["tags"]) + len(sug["queries"]) + len(sug["merges"])
    if not total:          # 추천이 없어도 아래 '필터 강도 추천'은 그대로 둔다
        st.caption("✨ **분석 전 정리 추천: 없음** — 지울 만한 색인어·검색어나 철자만 다른 표기가 보이지 않습니다. "
                   "논문 수가 적으면 (3편 이상에서 반복되는 것만 보므로) 추천이 비는 것이 정상입니다.")
    else:
        with st.expander(f"✨ 분석 전 정리 추천 {total}건 — 올린 파일의 저자키워드에서 찾았습니다", expanded=total <= 12):
            st.caption("분야마다 정답이 달라 **자동으로 지우지 않습니다.** 근거를 보고 고른 것만 내 설정에 들어가고, "
                       "그 상태로 **분석 시작**을 누르면 반영됩니다.")
            if sug["tags"]:
                st.markdown("**DB가 자동으로 붙인 색인어로 보이는 항목**")
                st.dataframe(pd.DataFrame(sug["tags"]), hide_index=True, height=180)
                picked = st.multiselect("제외할 항목", [t["키워드"] for t in sug["tags"]],
                                        key="pre_tags", label_visibility="collapsed")
                if st.button("제외 목록에 넣기", key="pre_tag_apply", disabled=not picked):
                    st.success(f"{append_config_lines('exclude_phrases.txt', picked)}개를 내 설정에 넣었습니다.")
            if sug["queries"]:
                st.markdown("**검색어 자체로 보이는 항목** — 논문 40% 이상에 달려 있어 변별력이 없습니다")
                st.dataframe(pd.DataFrame(sug["queries"]), hide_index=True, height=140)
                picked = st.multiselect("제외할 항목", [q["키워드"] for q in sug["queries"]],
                                        key="pre_queries", label_visibility="collapsed")
                if st.button("제외 목록에 넣기", key="pre_query_apply", disabled=not picked):
                    st.success(f"{append_config_lines('exclude_phrases.txt', picked)}개를 내 설정에 넣었습니다.")
            if sug["merges"]:
                st.markdown("**철자·어미만 다른 같은 말** — 하나로 합치면 빈도가 합산됩니다")
                opts = {f"{m['대표 표기']}  ←  {m['합칠 표기']}  ({m['합친 논문']}편)": m for m in sug["merges"]}
                picked = st.multiselect("합칠 묶음", list(opts), key="pre_merges", label_visibility="collapsed")
                if st.button("합치기 적용", key="pre_merge_apply", disabled=not picked):
                    lines = [f"{opts[p]['합칠 표기'].replace(', ', ' | ')} => {opts[p]['대표 표기']}" for p in picked]
                    st.success(f"{append_config_lines('synonyms.txt', lines)}개 묶음을 내 설정에 넣었습니다.")
            if sug.get("auto"):
                st.markdown("**🔗 프로그램이 알아서 합치는 표기** — 손대지 않아도 됩니다")
                st.caption("대소문자·하이픈·소유격·단수복수 차이는 자동으로 하나로 셉니다. "
                           "영국식/미국식 철자처럼 자동으로 안 되는 것만 위의 ‘합칠 표기’에 올라옵니다.")
                st.dataframe(pd.DataFrame(sug["auto"]), hide_index=True, height=200)
            st.caption("고른 항목은 왼쪽 옆칸 **⚙️ 내 단어 설정**에서 확인·수정할 수 있고, 파일로 내려받아 다음에 다시 쓸 수 있습니다.")

    with st.expander("📏 일반어 필터 강도 추천 — 이 데이터에는 몇 %가 맞나"):
        st.caption("**일반어 필터**는 `risk`·`approach`처럼 초록에 흔하지만 주제어가 아닌 말을 거르는 기준입니다. "
                   "“그 표현이 등장한 논문 중 **저자가 키워드로도 단 논문**의 비율”이 기준치 이상일 때만 남깁니다. "
                   "아래는 논문을 무작위 절반으로 나눠 한쪽으로 고르고 **다른 쪽의 저자키워드로 채점**한 결과입니다.")
        if st.button("필터 강도 계산", key="pre_threshold_run"):
            cfg = settings()
            st.session_state["pre_threshold"] = _pre_threshold(
                tuple(merged[kc.COL_KEYWORDS]), tuple(merged[kc.COL_TITLE]), tuple(merged[kc.COL_ABSTRACT]),
                cfg.get("stopwords.txt", ""), cfg.get("modifiers.txt", ""))
        got = st.session_state.get("pre_threshold")
        if got is not None:
            scored, best = got
            if scored.empty or best is None:
                st.info("저자키워드가 적어 가늠할 수 없습니다. 기본값(5%)으로 분석해 보세요.")
            else:
                st.dataframe(scored, hide_index=True, width="stretch")
                pct = int(round(best * 100))
                st.info(f"이 데이터에서 가장 좋게 나온 값은 **{pct} %** 입니다. "
                        "품사 태깅 전이라 **추정치**이고, 분석 후 ‘📐 임계값 민감도’에서 정확한 값으로 다시 확인할 수 있습니다. "
                        "F1 차이가 편차보다 작으면 어느 쪽을 써도 무방합니다.")
                if st.button(f"고급 설정을 {pct}%로 맞추기", key="pre_threshold_apply"):
                    st.session_state["ratio"] = pct
                    st.rerun()


def render_suggestions():
    """업로드한 파일 내용으로 뽑은 정리 후보 — 고른 것만 내 설정에 넣는다."""
    res = st.session_state.get("result")
    sug = (res or {}).get("suggest")
    if not sug:
        return
    total = len(sug["merges"]) + len(sug["tags"]) + len(sug["queries"])
    if not total:
        return
    with st.expander(f"✨ 자동 정리 제안 {total}건 — 올린 파일 내용에서 찾은 후보"):
        st.caption("분야마다 정답이 달라 자동으로 지우지 않습니다. 근거를 보고 고른 것만 내 설정에 들어갑니다. "
                   "적용 후 **분석 시작**을 다시 누르면 반영됩니다.")

        if sug["merges"]:
            st.markdown("**합칠 표기** — 어미만 다른 같은 말 (`optimisation` ↔ `optimization`)")
            opts = {f"{m['대표 표기']}  ←  {m['합칠 표기']}  ({m['합친 논문수']:,}편)": m for m in sug["merges"][:40]}
            picked = st.multiselect("합칠 묶음 고르기", list(opts), key="sug_merge", label_visibility="collapsed")
            if st.button("합치기 적용", disabled=not picked, key="apply_merge"):
                lines = [f"{opts[p]['합칠 표기'].replace(', ', ' | ')} => {opts[p]['대표 표기']}" for p in picked]
                n = append_config_lines("synonyms.txt", lines)
                st.success(f"{n}개 묶음을 ‘같은 뜻으로 합치기’에 넣었습니다.")

        if sug["tags"]:
            st.markdown("**색인 태그로 보이는 항목** — 저자키워드 열에만 있고 제목·초록에는 거의 없음")
            st.dataframe(pd.DataFrame(sug["tags"]).head(30), hide_index=True, height=180)
            picked = st.multiselect("제외할 항목 고르기", [t["키프레이즈"] for t in sug["tags"][:50]],
                                    key="sug_tag", label_visibility="collapsed")
            if st.button("제외 적용", disabled=not picked, key="apply_tag"):
                n = append_config_lines("exclude_phrases.txt", picked)
                st.success(f"{n}개를 ‘제외할 키프레이즈’에 넣었습니다.")

        if sug["queries"]:
            st.markdown("**검색어 자체로 보이는 항목** — 대부분의 논문에 나와 변별력이 없음")
            st.dataframe(pd.DataFrame(sug["queries"]).head(10), hide_index=True, height=150)
            picked = st.multiselect("제외할 항목 고르기", [q["키프레이즈"] for q in sug["queries"]],
                                    key="sug_query", label_visibility="collapsed")
            if st.button("제외 적용", disabled=not picked, key="apply_query"):
                n = append_config_lines("exclude_phrases.txt", picked)
                st.success(f"{n}개를 ‘제외할 키프레이즈’에 넣었습니다.")


def render_settings_sidebar():
    """단어 설정은 옆칸(사이드바)으로 — 평소에는 접혀 있고, 필요할 때만 연다."""
    cfg = settings()
    changed = kc.settings_diff(cfg)
    with st.sidebar:
        st.markdown("### ⚙️ 내 단어 설정")
        st.caption("제외·합치기 목록입니다. **내 브라우저에만** 저장되고 창을 닫으면 사라지므로, "
                   "계속 쓰실 설정은 내려받아 두었다가 다음에 올리세요.")
        st.metric("기본값에서 바꾼 줄", f"{len(changed)}줄")
        if changed:
            with st.expander("바꾼 내용 보기"):
                st.dataframe(pd.DataFrame(changed), hide_index=True, height=200)

        st.download_button("📥 내려받기", data=json.dumps(cfg, ensure_ascii=False, indent=1).encode("utf-8"),
                           file_name="키프레이즈_단어설정.json", mime="application/json",
                           key="dl_settings", width="stretch")
        uploaded = st.file_uploader("📤 설정 파일 올리기", type=["json"], key="up_settings")
        if uploaded is not None:
            try:
                loaded = json.loads(uploaded.getvalue().decode("utf-8"))
                st.session_state["settings"] = {k: loaded.get(k, v) for k, v in kc.default_settings().items()}
                st.success("불러왔습니다. ‘분석 시작’을 다시 누르세요.")
            except Exception as e:  # noqa: BLE001
                st.error(f"설정 파일을 읽을 수 없습니다: {e}")
        if st.button("기본값으로 되돌리기", key="reset_settings", width="stretch"):
            st.session_state["settings"] = kc.default_settings()
            st.rerun()

        with st.expander("목록 직접 고치기"):
            st.caption("보통은 분석 화면의 추천으로 충분합니다. `#` 뒤는 설명입니다. "
                       "고친 뒤 **분석 시작**을 다시 누르면 반영됩니다.")
            name = st.selectbox("고칠 목록", list(CONFIG_FILES), key="cfg_pick")
            fname, desc = CONFIG_FILES[name]
            st.markdown(desc)
            new = st.text_area(name, value=cfg.get(fname, ""), height=260,
                               key=f"cfg_{fname}", label_visibility="collapsed")
            if st.button("💾 적용", key=f"save_{fname}", width="stretch"):
                cfg[fname] = new
                st.toast(f"‘{name}’ 적용 완료 — 분석을 다시 실행하세요")


# ── 탭 3: 도움말 ──────────────────────────────────────────────────────────────
HELP = """
#### 사용 순서
1. Scopus · Web of Science · ProQuest에서 논문 목록을 **Excel 또는 CSV로 내보내기** (제목 · 초록 · 저자키워드 포함)
   여러 DB 파일을 한꺼번에 올리면 중복 문헌을 제거한 통합 목록으로 분석합니다
2. **📊 분석** 탭에 파일을 끌어다 놓기 → 연도 확인 → **분석 시작**
3. 워드클라우드·엑셀 표 확인 및 내려받기 (내 컴퓨터에서 실행하면 `output` 폴더에도 저장됩니다)
4. 빼고 싶은 단어는 결과 아래 **✂️**에서, 표기 합치기 등은 왼쪽 옆칸 **⚙️ 내 단어 설정**에서 고친 뒤 다시 분석

집계·필터 기준은 옆의 **처리 방식** 탭에 자세히 정리되어 있습니다.
논문이 10편 이상이면 **🧭 토픽** 탭에서 주제별 묶음·추세·대표 논문을,
20편 이상이면 **🗺️ 논문 지도** 탭에서 내용이 비슷한 논문끼리 묶인 지도를 볼 수 있습니다.
(50편 미만이면 참고용이라는 경고가 함께 표시됩니다.)

#### 엑셀 열 설명
- **총빈도**: 제목+초록+저자키워드에 등장한 총 횟수 (노출빈도)
- **논문수**: 한 번이라도 등장한 논문 수
- **저자키워드 사용률**: 등장한 논문 중 저자키워드로도 쓴 논문의 비율
- **병합된 표기**: 하나로 합쳐진 실제 표기들 (단복수·표기 차이 확인용)

#### 종료
내 컴퓨터에서 실행한 경우, 검은색 실행 창을 닫으면 프로그램이 종료됩니다.
"""

METHOD_HELP = """
결과는 아래 6단계를 거쳐 나옵니다. 숫자는 **기본값** 기준이며, 일부는 ‘2. 옵션 → 고급 설정’에서 바꿀 수 있습니다.

**① 논문 고르기 → ② 텍스트 정리 → ③ 표기 통일 → ④ 구문 사전 만들기 → ⑤ 세기 → ⑥ 거르기·정렬**

##### ① 논문 고르기 (여러 DB 통합)
- **읽는 열**: 열 이름으로 DB를 판단해 제목·초록·저자키워드·DOI·연도만 가져옵니다

  | DB | 제목 | 저자키워드 | DOI |
  |---|---|---|---|
  | Scopus | `Title` | `Author Keywords` | `DOI` |
  | Web of Science | `Article Title` | `Author Keywords` | `DOI` |
  | ProQuest | `Title` | `identifierKeywords` | `digitalObjectIdentifier` |

- **중복 제거**: ① DOI가 같으면 같은 문헌 ② DOI가 없거나 다르면 대소문자·구두점을 지운 제목이 같고 발행연도가 1년 이내일 때 같은 문헌.
  빠진 레코드의 초록·키워드가 더 충실하면 그 내용으로 채웁니다. 제거 내역은 엑셀 **중복 제거 목록** 시트에 남습니다
- **색인어는 쓰지 않습니다**: WoS `Keywords Plus`, ProQuest `subjectTerms`, Scopus `Index Keywords`는 기계·색인자가 붙인 것이라 저자키워드와 섞지 않습니다
- **문헌 유형**: 기본은 저널 논문만. 학위논문·회의록·도서·서평은 제외하며, 화면에서 전부 포함으로 바꿀 수 있습니다
- **연도**: 발행연도를 고르면 해당 연도만 남김
- **빈 행**: 제목과 초록이 모두 비어 있으면 제외. `[No abstract available]`은 초록 없음으로 처리

##### ② 텍스트 정리
주제어가 아닌 문구를 먼저 지웁니다.

| 지우는 것 | 예 |
|---|---|
| HTML 태그 | `<sup>`, `<i>` |
| 저작권 문구 (그 뒤 전부) | `© 2025 Elsevier…`, `Copyright 2024…` |
| 구조화 초록 소제목 | `Background:`, `Methods:`, `Main outcome measures:` |
| 통계 괄호 | `(95% CI 1.2–3.4)`, `(p<0.05)`, `(OR = 2.1)` |
| 인용 괄호 | `(Kim et al 2020; …)` |

**약어 풀기** — 한 논문 안에서 `large language models (LLMs)`처럼 정의된 약어를 찾아,
**그 논문의** 제목·초록·키워드에 나오는 `LLMs`·`LLM`을 `large language models`로 바꿉니다. 다른 논문에는 적용하지 않습니다.

##### ③ 표기 통일
- **대소문자 무시**: `Deep Learning` = `deep learning`
- **하이픈·소유격**: `meta-analysis` = `meta analysis`, `Alzheimer's disease` = `Alzheimer disease`
- **단수형** (영어 명사 사전 lemminflect): `models → model`, `analyses → analysis`, `mice → mouse`, `LLMs → LLM`
- **예외**: 대문자 약어(`AIDS`, `SARS`)와 3글자 이하 단어는 그대로 · 사전에 없는 단어는 단수형이 파일 안에 실제로 나올 때만 변환
- 무엇이 합쳐졌는지는 표의 **병합된 표기** 열에서 확인

##### ④ 구문 사전 만들기
품사 태깅(spaCy) 후, 아래 세 출처를 합쳐 ‘한 덩어리로 볼 구문’ 목록을 만듭니다.

| 출처 | 조건 |
|---|---|
| 저자키워드 | 전체 논문의 저자키워드 중 **2–8단어**짜리 전부 (한 편에만 있어도 포함) |
| 자동 추출 명사구 | 형용사·명사가 연달아 나오는 **2–4단어** 구간 · 끝 단어는 명사 · **3편 이상** 논문에 등장 |
| 항상 포함할 구문 | 옆칸(⚙️ 내 단어 설정)의 ‘항상 포함할 구문’에 적은 항목 |

자동 추출에서는 `novel`, `significant` 같은 수식어로 **시작**하거나 `model`, `approach` 같은 일반 명사로 **끝나는** 구문을 뺍니다.

##### ⑤ 세기
- **제목·초록**: 구문 사전에서 **가장 긴 구문부터** 맞춤 (쉼표·마침표를 넘어서는 매칭하지 않음).
  구문에 걸리지 않은 **명사만** 1단어로 세고, 형용사·동사 단독은 세지 않음
- **저자키워드**: 세미콜론으로 나눈 키워드 하나하나를 그대로 1회씩 셈 (1단어 키워드 포함)
- **합치기·빼기**: ‘같은 뜻으로 합치기’로 대표 표기에 합치고, ‘제외할 키프레이즈’·‘1단어 일반어’,
  영어 기능어(the, of …), 1글자·숫자만 있는 항목은 이때 뺌

> ⚠️ **긴 구문으로 잡힌 부분은 짧은 구문으로 다시 세지 않습니다.**
> 초록의 `non-small cell lung cancer`는 그 구문으로만 1회 세고, `lung cancer`나 `cancer`로는 세지 않습니다.

| 기록하는 값 | 뜻 |
|---|---|
| 제목 · 초록 · 저자키워드 | 각 부분의 **등장 횟수** (한 초록에 5번 나오면 5) |
| 총빈도 | 세 값의 합 (노출빈도) |
| 논문수 | 한 번이라도 등장한 논문 수 |
| 저자키워드 사용률 | 저자키워드로 쓴 논문 수 ÷ 등장한 논문 수 |

##### ⑥ 거르기와 정렬
결과에 남으려면 **세 조건을 모두** 만족해야 합니다.

1. 등장 논문 **2편 이상** (미달이면 표에서 아예 빠짐)
2. 저자키워드로 쓴 논문 **2편 이상**
3. 저자키워드 사용률 **5% 이상** (고급 설정의 ‘일반어 필터 강도’ — 0으로 하면 2·3번 조건을 끔)

- 2·3번에서 탈락한 항목은 결과 아래 **🔍 일반어 필터로 제외된 항목**과 엑셀 ‘제외됨(일반어 필터)’ 시트에서 확인.
  ‘항상 포함할 구문’에 넣은 항목은 조건과 관계없이 남음
- **정렬**: 총빈도 순, 같으면 논문수 순 (글자 크기 기준을 ‘논문수’로 고르면 반대)
- **표시 이름**: 합쳐진 표기 중 가장 많이 쓰인 형태 (예: `COVID-19`)
- **워드클라우드**: 남은 항목 중 상위 100개 (구문 이미지는 2단어 이상만 따로 100개). 고급 설정에서 20~100개 사이로 바꿀 수 있고, 적을수록 글자가 커집니다.
  글자 크기는 총빈도(또는 논문수)를 따르지만 정확히 비례하지는 않으므로 정확한 수치는 표에서 확인.
  상위 3개는 가로로, 나머지는 선택한 각도로 가운데부터 배치

##### ⑦ 토픽 모델 (논문 10편 이상일 때)
- **입력**: ⑥까지 통과한 키프레이즈의 논문별 출현 횟수 (TF-IDF 가중)
- **방법**: NMF로 토픽을 뽑고, 토픽마다 상위 키프레이즈 8개의 NPMI 응집도를 계산합니다.
  토픽 수 후보(5·8·10·12·15·20) 중 응집도가 가장 높은 값을 추천하되, 지표가 연구 목적과 늘 맞지는 않으니
  ‘🧭 토픽’ 탭에서 직접 바꿔 보세요.
- **주제 지도**: 같은 논문에 함께 나온 키프레이즈로 네트워크를 만들고 Louvain으로 묶은 뒤,
  Callon의 중심성(다른 주제와의 연결)·밀도(주제 내부 응집)로 4분면에 배치합니다 — 핵심 / 기반 / 니치 / 주변 주제

##### 예시 한 편으로 따라가 보기
> **초록** — *Large language models (LLMs) are increasingly used. We evaluated LLMs for predicting survival in non-small cell lung cancer patients.*
>
> **저자키워드** — *LLM; Lung cancer; Survival analysis*

| 키프레이즈 | 어떻게 세어지나 |
|---|---|
| large language model | 초록 2 (`LLMs` 약어 풀기) + 키워드 1 (`LLM`도 치환) → 총빈도 3, 논문 1 |
| non-small cell lung cancer | 구문 사전에 있으면 초록 1 (다른 논문의 키워드에 있거나 3편 이상에 반복된 경우) |
| lung cancer | 키워드 1 — 초록 속 lung cancer는 더 긴 구문에 포함돼 따로 세지 않음 |
| survival analysis | 키워드 1 |
| survival · patient | 1단어 명사로 초록 각 1 — 저자키워드로 거의 안 쓰이면 ⑥에서 빠질 수 있음 |

##### 해석할 때 주의할 점
- **저자키워드를 빼면 결과가 거의 사라집니다.** ‘집계할 부분’에서 저자키워드를 해제하면 사용률이 모두 0이 되어
  ‘항상 포함할 구문’만 남습니다. 이때는 고급 설정의 일반어 필터 강도를 **0**으로 바꾸세요.
- **긴 구문이 우선합니다.** 짧은 구문의 초록 빈도가 생각보다 낮다면 더 긴 구문에 흡수됐을 수 있습니다. 필요하면 ‘같은 뜻으로 합치기’를 쓰세요.
- **1단어 형용사는 나오지 않습니다.** `pediatric`처럼 명사가 아닌 단어는 구문의 일부로만 집계됩니다.
- **품사 태깅은 자동이라 가끔 틀립니다.** 꼭 잡아야 하는 표현은 ‘항상 포함할 구문’에 적으세요.
"""


session_id()  # 세션 식별자를 미리 만들어 둔다 (업로드·결과 폴더 분리에 사용)
st.title("☁️ 논문 키프레이즈 워드클라우드")
st.caption("Scopus 논문 목록 파일을 올리면 제목·초록·저자키워드의 키프레이즈 노출빈도를 집계해 워드클라우드와 엑셀 표를 만듭니다.")
render_settings_sidebar()
tab_run, tab_topic, tab_map, tab_help = st.tabs(["📊 분석", "🧭 토픽", "🗺️ 논문 지도", "❓ 도움말"])
with tab_run:
    render_run_tab()
with tab_topic:
    render_topic_tab()
with tab_map:
    render_map_tab()
with tab_help:
    help_use, help_method = st.tabs(["사용법", "처리 방식 (집계·필터 기준)"])
    with help_use:
        st.markdown(HELP)
    with help_method:
        st.markdown(METHOD_HELP)
