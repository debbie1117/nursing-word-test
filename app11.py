import streamlit as st
import random
from gtts import gTTS
import io
import os
import csv
import openpyxl
import base64
import requests
import datetime

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

st.set_page_config(page_title="간호학 의학용어 테스트 & 데이터 분석", page_icon="🩺", layout="wide")

FLASHCARD_APP_URL = "https://streamlit.app"
LOG_FILE_PATH = "quiz_logs.csv"

# ---------------------------------------------------------
# 1. 영구 저장 및 불러오기 (깃허브 연동 + 로컬 파일 백업)
# ---------------------------------------------------------
def get_github_config():
    try:
        token = st.secrets.get("GITHUB_TOKEN", "")
        repo = st.secrets.get("REPO_NAME", "")
        return token, repo
    except Exception:
        return "", ""

def load_all_logs():
    # 1) 깃허브 토큰이 설정되어 있으면 깃허브에서 영구 기록 불러오기
    token, repo = get_github_config()
    if token and repo:
        url = f"https://github.com{repo}/contents/{LOG_FILE_PATH}"
        headers = {"Authorization": f"token {token}"}
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            content_b64 = res.json().get("content", "")
            decoded = base64.b64decode(content_b64).decode("utf-8-sig")
            reader = csv.DictReader(io.StringIO(decoded))
            return [
                {
                    "날짜": r.get("날짜", ""),
                    "파트": r.get("파트", ""),
                    "영단어": r.get("영단어", ""),
                    "한글뜻": r.get("한글뜻", ""),
                    "정답여부": int(r.get("정답여부", 0))
                }
                for r in reader
            ]

    # 2) 로컬 CSV 파일이 있으면 불러오기
    if os.path.exists(LOG_FILE_PATH):
        with open(LOG_FILE_PATH, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            return [
                {
                    "날짜": r.get("날짜", ""),
                    "파트": r.get("파트", ""),
                    "영단어": r.get("영단어", ""),
                    "한글뜻": r.get("한글뜻", ""),
                    "정답여부": int(r.get("정답여부", 0))
                }
                for r in reader
            ]
    return []

def save_all_logs(logs):
    # 로컬 CSV 파일 저장
    fieldnames = ["날짜", "파트", "영단어", "한글뜻", "정답여부"]
    with open(LOG_FILE_PATH, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in logs:
            writer.writerow(row)

    # 깃허브 영구 저장 (토큰 설정 시)
    token, repo = get_github_config()
    if token and repo:
        url = f"https://github.com{repo}/contents/{LOG_FILE_PATH}"
        headers = {"Authorization": f"token {token}"}
        sha = None
        res_get = requests.get(url, headers=headers)
        if res_get.status_code == 200:
            sha = res_get.json().get("sha")

        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=fieldnames)
        writer.writeheader()
        for row in logs:
            writer.writerow(row)

        encoded_content = base64.b64encode(output.getvalue().encode("utf-8-sig")).decode("utf-8")
        payload = {"message": "Update quiz logs", "content": encoded_content}
        if sha:
            payload["sha"] = sha
        requests.put(url, headers=headers, json=payload)

# ---------------------------------------------------------
# 2. 단어장 엑셀/CSV 로드
# ---------------------------------------------------------
def load_vocabulary():
    target_file = None
    for file in os.listdir("."):
        if (file.endswith(".xlsx") or file.endswith(".csv")) and not file.startswith("~$") and file != LOG_FILE_PATH:
            target_file = file
            break

    if not target_file:
        return None, None

    rows = []
    if target_file.endswith(".xlsx"):
        wb = openpyxl.load_workbook(target_file, data_only=True)
        if wb.active is not None:
            rows = [list(r) for r in wb.active.iter_rows(values_only=True)]
    else:
        for enc in ["utf-8-sig", "cp949", "utf-8", "euc-kr"]:
            try:
                with open(target_file, mode="r", encoding=enc) as f:
                    rows = list(csv.reader(f))
                break
            except UnicodeDecodeError:
                continue

    records = []
    for row in rows:
        if not row or len(row) < 3:
            continue
        it = iter(row)
        no_raw = next(it, None)
        eng_raw = next(it, None)
        kor_raw = next(it, None)
        part_raw = next(it, None)

        no_val = str(no_raw).strip() if no_raw is not None else ""
        if no_val.isdigit():
            english = str(eng_raw).strip() if eng_raw is not None else ""
            korean = str(kor_raw).strip() if kor_raw is not None else ""
            part = str(part_raw).strip() if part_raw is not None and str(part_raw).strip() else "기본 범위"

            if english == "соссух":
                english = "coccyx"
            elif english == "costal nec":
                english = "costal neck"

            if korean and english:
                records.append({"korean": korean, "english": english, "part": part})

    word_dict = {}
    for item in records:
        p = item.get("part", "기본 범위")
        if p not in word_dict:
            word_dict[p] = []
        word_dict[p].append(item)

    return word_dict, records

# ---------------------------------------------------------
# 3. 시험 초기화 및 음성 생성
# ---------------------------------------------------------
def init_quiz(word_data, selected_range, custom_words=None):
    if custom_words is not None:
        words = list(custom_words)
    elif selected_range == "전체 범위":
        words = [item for page_words in word_data.values() for item in page_words]
    else:
        words = list(word_data.get(selected_range, []))

    random.shuffle(words)
    st.session_state.quiz_list = words
    st.session_state.current_index = 0
    st.session_state.score = 0
    st.session_state.submitted = False
    st.session_state.user_answer = ""
    st.session_state.wrong_answers = []
    st.session_state.current_range = selected_range
    if "scored_this_q" in st.session_state:
        del st.session_state.scored_this_q

if "attempt_logs" not in st.session_state:
    st.session_state.attempt_logs = load_all_logs()

@st.cache_data(show_spinner=False)
def get_audio_bytes(word):
    clean_word = word.replace("~", " to ").replace("(", " ").replace(")", " ")
    tts = gTTS(text=clean_word, lang='en')
    fp = io.BytesIO()
    tts.write_to_fp(fp)
    fp.seek(0)
    return fp.read()

# ---------------------------------------------------------
# 4. 메인 화면 UI (시험 + 실시간 그래프 한 화면 배치)
# ---------------------------------------------------------
top_col1, top_col2 = st.columns([2.5, 1])
with top_col1:
    st.title("🩺 간호학 의학용어 받아쓰기 & 학습 분석")
with top_col2:
    st.write("")
    st.link_button("🗂️ 플래시카드 공부하러 가기", FLASHCARD_APP_URL, use_container_width=True)

word_data, all_records = load_vocabulary()

if not word_data or not all_records:
    st.error("⚠️ 단어장 파일(`.xlsx` 또는 `.csv`)을 찾을 수 없습니다.")
    st.stop()

# 화면을 좌우로 분할: 왼쪽은 받아쓰기 시험, 오른쪽은 항상 보이는 데이터 그래프
col_exam, col_graph = st.columns([1.2, 1])

# =========================================================
# [왼쪽] 받아쓰기 시험 영역
# =========================================================
with col_exam:
    st.subheader("✍️ 받아쓰기 시험")
    range_options = ["전체 범위"] + list(word_data.keys())
    selected_range = st.selectbox("📚 시험 볼 파트를 선택하세요:", range_options)

    if "quiz_list" not in st.session_state or st.session_state.get("current_range") != selected_range:
        init_quiz(word_data, selected_range)

    total_questions = len(st.session_state.quiz_list)
    current_idx = st.session_state.current_index

    if current_idx >= total_questions:
        st.success("🎉 선택한 범위의 단어 테스트가 완료되었습니다!")
        accuracy = round((st.session_state.score / total_questions) * 100, 1)
        st.subheader(f"최종 점수: {st.session_state.score} / {total_questions}점 (정답률 {accuracy}%)")

        if st.session_state.wrong_answers:
            st.write("### ❌ 오답 노트")
            md_table = "| 파트 | 한국어 뜻 | 입력한 답 | 정답 |\n| :--- | :--- | :--- | :--- |\n"
            for item in st.session_state.wrong_answers:
                md_table += f"| {item.get('파트')} | {item.get('뜻')} | `{item.get('입력한 답')}` | **{item.get('정답')}** |\n"
            st.markdown(md_table)

            c1, c2 = st.columns(2)
            with c1:
                if st.button("🔁 틀린 문제만 재시험", type="primary", use_container_width=True):
                    retry_list = [{"korean": w.get("뜻"), "english": w.get("정답"), "part": w.get("파트")} for w in st.session_state.wrong_answers]
                    init_quiz(word_data, selected_range, custom_words=retry_list)
                    st.rerun()
            with c2:
                if st.button("🔄 전체 다시 테스트", use_container_width=True):
                    init_quiz(word_data, selected_range)
                    st.rerun()
        else:
            st.balloons()
            if st.button("같은 범위 다시 테스트하기"):
                init_quiz(word_data, selected_range)
                st.rerun()
    else:
        current_q = st.session_state.quiz_list[current_idx]
        st.progress(current_idx / total_questions)
        st.caption(f"진행: {current_idx + 1} / {total_questions} 문제 | 파트: {current_q.get('part')} | 맞힌 개수: {st.session_state.score}개")

        st.markdown(f"### 뜻: **{current_q.get('korean')}**")
        st.audio(get_audio_bytes(current_q.get('english', '')), format="audio/mp3")

        with st.form(key=f"quiz_form_{selected_range}_{current_idx}"):
            user_input = st.text_input("들리는 영단어를 입력하세요 (대소문자 무관):", disabled=st.session_state.submitted)
            submit_btn = st.form_submit_button("정답 확인", disabled=st.session_state.submitted)

            if submit_btn:
                st.session_state.submitted = True
