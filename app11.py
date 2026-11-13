import streamlit as st
import random
from gtts import gTTS
import io
import os
import csv
import openpyxl

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

st.set_page_config(page_title="간호학 의학용어 테스트 & 데이터 분석", page_icon="🩺", layout="wide")

FLASHCARD_APP_URL = "https://streamlit.app"

# ---------------------------------------------------------
# 1. 엑셀/CSV 단어장 로드
# ---------------------------------------------------------
def load_vocabulary():
    target_file = None
    for file in os.listdir("."):
        if (file.endswith(".xlsx") or file.endswith(".csv")) and not file.startswith("~$"):
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
# 2. 시험 초기화 및 누적 통계 관리
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
    st.session_state.attempt_logs = []

# ---------------------------------------------------------
# 3. 음성 생성 함수
# ---------------------------------------------------------
@st.cache_data(show_spinner=False)
def get_audio_bytes(word):
    clean_word = word.replace("~", " to ").replace("(", " ").replace(")", " ")
    tts = gTTS(text=clean_word, lang='en')
    fp = io.BytesIO()
    tts.write_to_fp(fp)
    fp.seek(0)
    return fp.read()

# ---------------------------------------------------------
# 4. 메인 화면 UI
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

page_mode = st.radio(
    "화면 선택",
    ["✍️ 받아쓰기 시험", "📊 나의 학습 데이터 분석"],
    horizontal=True,
    label_visibility="collapsed"
)
st.divider()

# =========================================================
# [화면 1] 받아쓰기 시험 화면
# =========================================================
if page_mode == "✍️ 받아쓰기 시험":
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
                if st.button("🔁 틀린 문제만 재시험 보기", type="primary", use_container_width=True):
                    retry_list = [{"korean": w.get("뜻"), "english": w.get("정답"), "part": w.get("파트")} for w in st.session_state.wrong_answers]
                    init_quiz(word_data, selected_range, custom_words=retry_list)
                    st.rerun()
            with c2:
                if st.button("🔄 전체 범위 다시 테스트하기", use_container_width=True):
                    init_quiz(word_data, selected_range)
                    st.rerun()
        else:
            st.balloons()
            st.write("완벽합니다! 틀린 문제가 하나도 없습니다.")
            if st.button("같은 범위 다시 테스트하기"):
                init_quiz(word_data, selected_range)
                st.rerun()

    else:
        current_q = st.session_state.quiz_list[current_idx]

        st.progress(current_idx / total_questions)
        st.caption(f"진행 상황: {current_idx + 1} / {total_questions} 문제 | 파트: {current_q.get('part')} | 맞힌 개수: {st.session_state.score}개")

        st.markdown(f"### 뜻: **{current_q.get('korean')}**")

        audio_data = get_audio_bytes(current_q.get('english', ''))
        st.audio(audio_data, format="audio/mp3")

        with st.form(key=f"quiz_form_{selected_range}_{current_idx}"):
            user_input = st.text_input("들리는 영단어를 입력하세요 (대소문자 무관):", disabled=st.session_state.submitted)
            submit_btn = st.form_submit_button("정답 확인", disabled=st.session_state.submitted)

            if submit_btn:
                st.session_state.submitted = True
                st.session_state.user_answer = user_input.strip()
                st.rerun()

        if st.session_state.submitted:
            correct_ans = current_q.get('english', '').strip().lower()
            user_ans = st.session_state.user_answer.lower()
            is_correct = (user_ans == correct_ans)

            if is_correct:
                st.success(f"⭕ 정답입니다! (**{current_q.get('english')}**)")
                if "scored_this_q" not in st.session_state:
                    st.session_state.score += 1
                    st.session_state.scored_this_q = True
                    st.session_state.attempt_logs.append({
                        "파트": current_q.get("part"),
                        "영단어": current_q.get("english"),
                        "한글뜻": current_q.get("korean"),
                        "정답여부": 1
                    })
            else:
                st.error(f"❌ 오답입니다. 입력한 답: `{st.session_state.user_answer}` / 정답: **{current_q.get('english')}**")
                if "scored_this_q" not in st.session_state:
                    st.session_state.wrong_answers.append({
                        "파트": current_q.get("part"),
                        "뜻": current_q.get("korean"),
                        "입력한 답": st.session_state.user_answer,
                        "정답": current_q.get('english')
                    })
                    st.session_state.scored_this_q = True
                    st.session_state.attempt_logs.append({
                        "파트": current_q.get("part"),
                        "영단어": current_q.get("english"),
                        "한글뜻": current_q.get("korean"),
                        "정답여부": 0
                    })

            if st.button("다음 문제 ➡️"):
                st.session_state.current_index += 1
                st.session_state.submitted = False
                st.session_state.user_answer = ""
                if "scored_this_q" in st.session_state:
                    del st.session_state.scored_this_q
                st.rerun()

# =========================================================
# [화면 2] 데이터 분석 및 시각화 화면
# =========================================================
else:
    st.subheader("📊 의학용어 학습 데이터 분석 대시보드")

    # 1. 전체 단어장 분포 시각화
    st.markdown("#### 1️⃣ 파트별 수록 영단어 분포")
    part_count_dict = {p: len(items) for p, items in word_data.items()}
    st.bar_chart(part_count_dict)

    st.divider()

    # 2. 실시간 시험 결과 분석
    st.markdown("#### 2️⃣ 실시간 시험 결과 분석 (정답률 및 취약 용어)")

    if len(st.session_state.attempt_logs) == 0:
        st.info("💡 아직 풀이한 시험 데이터가 없습니다. **[✍️ 받아쓰기 시험]**에서 문제를 풀고 정답을 확인하면 그래프가 자동으로 생성됩니다.")
    else:
        logs = st.session_state.attempt_logs
        total_solved = len(logs)
        total_correct = sum(item.get("정답여부", 0) for item in logs)
