import streamlit as st
import random
from gtts import gTTS
import io
import os
import csv
import openpyxl
import heapq
from fpdf import FPDF

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

st.set_page_config(page_title="간호학 의학용어 테스트 & 데이터 분석", page_icon="🩺", layout="wide")

FLASHCARD_APP_URL = "https://nursing-word-test-xcixlbea3grxdnxwxluhgq.streamlit.app/"

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
# 3. 음성 생성 및 오답 PDF 생성 함수
# ---------------------------------------------------------
@st.cache_data(show_spinner=False)
def get_audio_bytes(word):
    clean_word = word.replace("~", " to ").replace("(", " ").replace(")", " ")
    tts = gTTS(text=clean_word, lang='en')
    fp = io.BytesIO()
    tts.write_to_fp(fp)
    fp.seek(0)
    return fp.read()

def create_wrong_answers_pdf(wrong_list, range_name):
    pdf = FPDF()
    pdf.add_page()

    font_candidates = [
        "NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf"
    ]
    font_path = next((p for p in font_candidates if os.path.exists(p)), None)

    if font_path:
        pdf.add_font("CustomKorean", "", font_path)
        pdf.set_font("CustomKorean", size=16)
    else:
        pdf.set_font("Helvetica", size=16)

    pdf.cell(0, 12, text=f"간호학 의학용어 오답 노트 ({range_name})", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)

    if font_path:
        pdf.set_font("CustomKorean", size=11)
    else:
        pdf.set_font("Helvetica", size=11)

    pdf.set_fill_color(230, 236, 245)
    pdf.cell(45, 10, text="파트", border=1, fill=True, align="C")
    pdf.cell(55, 10, text="정답 (영단어)", border=1, fill=True, align="C")
    pdf.cell(55, 10, text="한국어 뜻", border=1, fill=True, align="C")
    pdf.cell(35, 10, text="내가 쓴 답", border=1, fill=True, align="C", new_x="LMARGIN", new_y="NEXT")

    for item in wrong_list:
        pdf.cell(45, 10, text=str(item.get("파트", "")), border=1)
        pdf.cell(55, 10, text=str(item.get("정답", "")), border=1)
        pdf.cell(55, 10, text=str(item.get("뜻", "")), border=1)
        pdf.cell(35, 10, text=str(item.get("입력한 답", "")), border=1, new_x="LMARGIN", new_y="NEXT")

    return bytes(pdf.output())

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

            pdf_bytes = create_wrong_answers_pdf(st.session_state.wrong_answers, selected_range)

            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("🔁 틀린 문제만 재시험 보기", type="primary", use_container_width=True):
                    retry_list = [{"korean": w.get("뜻"), "english": w.get("정답"), "part": w.get("파트")} for w in st.session_state.wrong_answers]
                    init_quiz(word_data, selected_range, custom_words=retry_list)
                    st.rerun()
            with c2:
                st.download_button(
                    label="📄 오답 노트 PDF 다운로드",
                    data=pdf_bytes,
                    file_name=f"의학용어_오답노트_{selected_range}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with c3:
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

                if "scored_this_q" not in st.session_state:
                    correct_ans = current_q.get('english', '').strip().lower()
                    user_ans = st.session_state.user_answer.lower()
                    is_correct = (user_ans == correct_ans)

                    if is_correct:
                        st.session_state.score += 1
                    else:
                        st.session_state.wrong_answers.append({
                            "파트": current_q.get("part"),
                            "뜻": current_q.get("korean"),
                            "입력한 답": st.session_state.user_answer,
                            "정답": current_q.get('english')
                        })

                    st.session_state.attempt_logs.append({
                        "파트": current_q.get("part"),
                        "영단어": current_q.get("english"),
                        "한글뜻": current_q.get("korean"),
                        "정답여부": is_correct,
                        "입력답안": st.session_state.user_answer
                    })
                    st.session_state.scored_this_q = True
                st.rerun()

        if st.session_state.submitted:
            correct_word = current_q.get("english", "")
            if st.session_state.user_answer.lower() == correct_word.strip().lower():
                st.success(f"⭕ 정답입니다! (**{correct_word}**)")
            else:
                st.error(f"❌ 오답입니다. 입력한 답: `{st.session_state.user_answer}` / 정답: **{correct_word}**")

            if st.button("다음 문제 ➡️", type="primary"):
                st.session_state.current_index += 1
                st.session_state.submitted = False
                st.session_state.user_answer = ""
                if "scored_this_q" in st.session_state:
                    del st.session_state.scored_this_q
                st.rerun()

# =========================================================
# [화면 2] 나의 학습 데이터 분석 화면
# =========================================================
elif page_mode == "📊 나의 학습 데이터 분석":
    st.subheader("📊 누적 학습 통계 및 취약 단어 분석")
    logs = st.session_state.attempt_logs

    if not logs:
        st.info("아직 풀이한 문제 기록이 없습니다. 먼저 받아쓰기 시험을 진행해 주세요!")
    else:
        total_tried = len(logs)
        total_correct = sum(1 for log in logs if log["정답여부"])
        overall_acc = round((total_correct / total_tried) * 100, 1)

        m1, m2, m3 = st.columns(3)
        m1.metric("총 풀이 문제 수", f"{total_tried}문제")
        m2.metric("누적 정답 수", f"{total_correct}문제")
        m3.metric("전체 평균 정답률", f"{overall_acc}%")

        st.divider()

        wrong_counts = {}
        for log in logs:
            if not log["정답여부"]:
                key = (log["영단어"], log["한글뜻"], log["파트"])
                wrong_counts[key] = wrong_counts.get(key, 0) + 1

        st.write("### 🚨 자주 틀리는 취약 단어 TOP 5")
        if wrong_counts:
            top_wrong = heapq.nlargest(5, wrong_counts.items(), key=lambda x: x[1])
            for (eng, kor, part), cnt in top_wrong:
                st.write(f"- **{eng}** ({kor}) | 파트: {part} — **오답 {cnt}회**")
        else:
            st.success("틀린 단어가 없습니다! 훌륭합니다.")

        if PANDAS_AVAILABLE:
            st.write("### 📋 전체 풀이 로그")
            df = pd.DataFrame(logs)
            st.dataframe(df, use_container_width=True)
