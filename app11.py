import streamlit as st
import random
from gtts import gTTS
import io
import os
import csv
import openpyxl

# 페이지 기본 설정
st.set_page_config(page_title="간호학 해부학 영단어 테스트", page_icon="🩺", layout="centered")


# ---------------------------------------------------------
# 1. 엑셀(.xlsx) 및 CSV(.csv) 파일 자동 탐색 및 불러오기
# ---------------------------------------------------------
def load_vocabulary():
    target_file = None
    # 현재 폴더에 있는 .xlsx 또는 .csv 파일 자동 탐색
    for file in os.listdir("."):
        if (file.endswith(".xlsx") or file.endswith(".csv")) and not file.startswith("~$"):
            target_file = file
            break

    if not target_file:
        return None, None

    rows = []
    # 1) 엑셀 파일(.xlsx)인 경우
    if target_file.endswith(".xlsx"):
        wb = openpyxl.load_workbook(target_file, data_only=True)
        sheet = wb.active
        for r in sheet.iter_rows(values_only=True):
            rows.append(list(r))

    # 2) CSV 파일(.csv)인 경우 (한글 인코딩 자동 처리)
    elif target_file.endswith(".csv"):
        for enc in ["utf-8-sig", "cp949", "utf-8", "euc-kr"]:
            try:
                with open(target_file, mode="r", encoding=enc) as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                break
            except UnicodeDecodeError:
                continue

    word_dict = {}
    for row in rows:
        if not row or len(row) < 3:
            continue

        no_val = str(row[0]).strip() if row[0] is not None else ""
        # A열(번호)이 숫자인 데이터 행만 추출 (1~3행 제목줄 자동 제외)
        if not no_val.isdigit():
            continue

        english = str(row[1]).strip() if row[1] is not None else ""  # B열: 영어 용어
        korean = str(row[2]).strip() if row[2] is not None else ""  # C열: 한글 용어
        page = str(row[3]).strip() if len(row) > 3 and row[3] is not None else "기본 범위"  # D열: 페이지

        # 원문 오타 보정 (발음 및 채점 정확도 향상)
        if english == "соссух":
            english = "coccyx"
        elif english == "costal nec":
            english = "costal neck"

        if korean and english:
            if page not in word_dict:
                word_dict[page] = []
            word_dict[page].append({
                "korean": korean,
                "english": english
            })

    return word_dict, target_file


# ---------------------------------------------------------
# 2. 시험 범위 선택 및 세션 초기화 함수
# ---------------------------------------------------------
def init_quiz(word_data, selected_range):
    if selected_range == "전체 범위":
        words = [item for page_words in word_data.values() for item in page_words]
    else:
        words = list(word_data[selected_range])

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


# ---------------------------------------------------------
# 3. 음성 생성 함수 (캐싱 적용)
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
# 4. 메인 화면 UI 구성
# ---------------------------------------------------------
st.title("🩺 간호학 해부학 영단어 받아쓰기")

word_data, loaded_filename = load_vocabulary()

if word_data is None or len(word_data) == 0:
    st.error("⚠️ 단어장 파일(`.xlsx` 또는 `.csv`)을 찾을 수 없거나 데이터를 읽지 못했습니다.")
    st.write(f"**현재 파이참 실행 폴더 위치:** `{os.getcwd()}`")
    st.write("**이 폴더 안에 있는 파일 목록:**", os.listdir("."))
    st.stop()

st.caption(f"연결된 파일: `{loaded_filename}`")

# D열(P.1 ~ P.8) 기준으로 시험 범위 선택
range_options = ["전체 범위"] + list(word_data.keys())
selected_range = st.selectbox("📚 시험 볼 페이지를 선택하세요:", range_options)

if "quiz_list" not in st.session_state or st.session_state.get("current_range") != selected_range:
    init_quiz(word_data, selected_range)

st.divider()

total_questions = len(st.session_state.quiz_list)
current_idx = st.session_state.current_index

if current_idx >= total_questions:
    st.success("🎉 선택한 범위의 단어 테스트가 완료되었습니다!")
    st.subheader(f"최종 점수: {st.session_state.score} / {total_questions}점")

    if st.session_state.wrong_answers:
        st.write("### ❌ 오답 노트")
        st.table(st.session_state.wrong_answers)
    else:
        st.balloons()
        st.write("완벽합니다! 틀린 문제가 하나도 없습니다.")

    if st.button("같은 범위 다시 테스트하기"):
        init_quiz(word_data, selected_range)
        st.rerun()

else:
    current_q = st.session_state.quiz_list[current_idx]

    st.progress((current_idx) / total_questions)
    st.caption(f"진행 상황: {current_idx + 1} / {total_questions} 문제 (현재 맞힌 개수: {st.session_state.score}개)")

    st.markdown(f"### 뜻: **{current_q['korean']}**")

    audio_data = get_audio_bytes(current_q['english'])
    st.audio(audio_data, format="audio/mp3")

    with st.form(key=f"quiz_form_{selected_range}_{current_idx}"):
        user_input = st.text_input("들리는 영단어를 입력하세요 (대소문자 무관):", disabled=st.session_state.submitted)
        submit_btn = st.form_submit_button("정답 확인", disabled=st.session_state.submitted)

        if submit_btn:
            st.session_state.submitted = True
            st.session_state.user_answer = user_input.strip()
            st.rerun()

    if st.session_state.submitted:
        correct_ans = current_q['english'].strip().lower()
        user_ans = st.session_state.user_answer.lower()

        if user_ans == correct_ans:
            st.success(f"⭕ 정답입니다! (**{current_q['english']}**)")
            if "scored_this_q" not in st.session_state:
                st.session_state.score += 1
                st.session_state.scored_this_q = True
        else:
            st.error(f"❌ 오답입니다. 입력한 답: `{st.session_state.user_answer}` / 정답: **{current_q['english']}**")
            if "scored_this_q" not in st.session_state:
                st.session_state.wrong_answers.append({
                    "뜻": current_q['korean'],
                    "입력한 답": st.session_state.user_answer,
                    "정답": current_q['english']
                })
                st.session_state.scored_this_q = True

        if st.button("다음 문제 ➡️"):
            st.session_state.current_index += 1
            st.session_state.submitted = False
            st.session_state.user_answer = ""
            if "scored_this_q" in st.session_state:
                del st.session_state.scored_this_q
            st.rerun()
