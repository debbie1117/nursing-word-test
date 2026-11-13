import streamlit as st
import random
from gtts import gTTS
import io
import os
import csv
import openpyxl

# 페이지 기본 설정
st.set_page_config(page_title="간호학 의학용어 플래시카드", page_icon="🗂️", layout="centered")

# 🔗 기존 받아쓰기 시험 웹사이트 주소
DICTATION_APP_URL = "https://nursing-word-test-58ddvddugjimbyvjhdgyin.streamlit.app/"

# ---------------------------------------------------------
# 1. 엑셀(.xlsx) 및 CSV(.csv) 파일 불러오기
# ---------------------------------------------------------
def load_vocabulary():
    target_file = None
    for file in os.listdir("."):
        if (file.endswith(".xlsx") or file.endswith(".csv")) and not file.startswith("~$"):
            target_file = file
            break

    if not target_file:
        return [], {}

    rows = []
    if target_file.endswith(".xlsx"):
        wb = openpyxl.load_workbook(target_file, data_only=True)
        sheet = wb.active
        for r in sheet.iter_rows(values_only=True):
            rows.append(list(r))
    elif target_file.endswith(".csv"):
        for enc in ["utf-8-sig", "cp949", "utf-8", "euc-kr"]:
            try:
                with open(target_file, mode="r", encoding=enc) as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                break
            except UnicodeDecodeError:
                continue

    all_words = []
    word_dict = {}

    for row in rows:
        if not row or len(row) < 3:
            continue

        no_val = str(row[0]).strip() if row[0] is not None else ""
        if not no_val.isdigit():
            continue

        english = str(row[1]).strip() if row[1] is not None else ""
        korean = str(row[2]).strip() if row[2] is not None else ""
        part = str(row[3]).strip() if len(row) > 3 and row[3] is not None else "기본 범위"

        if english == "соссух":
            english = "coccyx"
        elif english == "costal nec":
            english = "costal neck"

        if korean and english:
            item = {
                "no": int(no_val),
                "english": english,
                "korean": korean,
                "part": part
            }
            all_words.append(item)
            if part not in word_dict:
                word_dict[part] = []
            word_dict[part].append(item)

    return all_words, word_dict

# ---------------------------------------------------------
# 2. 음성 생성 함수 (캐싱 적용)
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
# 3. 메인 화면 UI
# ---------------------------------------------------------
all_words, word_dict = load_vocabulary()

if "unknown_words" not in st.session_state:
    st.session_state.unknown_words = set()

# 상단 제목 및 받아쓰기 시험 페이지 연결 버튼
top_col1, top_col2 = st.columns([2.2, 1])
with top_col1:
    st.title("🗂️ 의학용어 플래시카드")
with top_col2:
    st.write("")
    st.link_button("✍️ 받아쓰기 시험보러 가기", DICTATION_APP_URL, use_container_width=True)

if not all_words:
    st.error("⚠️ 단어장 엑셀/CSV 파일을 찾을 수 없습니다.")
    st.stop()

# 파트 선택 및 앞면 기준 설정
c1, c2 = st.columns([1.8, 1.2])
with c1:
    fc_part = st.selectbox("📚 학습할 파트 선택:", ["전체 범위", "❌ 모르는 단어만 반복 학습"] + list(word_dict.keys()))
with c2:
    fc_direction = st.radio("카드 앞면 표시:", ["한글 뜻 → 영어", "영어 → 한글 뜻"], horizontal=True)

if fc_part == "전체 범위":
    cards = list(all_words)
elif fc_part == "❌ 모르는 단어만 반복 학습":
    cards = [w for w in all_words if w["english"] in st.session_state.unknown_words]
else:
    cards = list(word_dict[fc_part])

if not cards:
    st.success("🎉 현재 '모르는 단어'로 체크된 항목이 없습니다!")
    st.stop()

state_key = f"{fc_part}_{fc_direction}"
if st.session_state.get("fc_state_key") != state_key:
    st.session_state.fc_state_key = state_key
    st.session_state.fc_cards = cards
    st.session_state.fc_idx = 0
    st.session_state.fc_flipped = False

active_cards = st.session_state.fc_cards
idx = min(st.session_state.fc_idx, len(active_cards) - 1)
curr = active_cards[idx]

st.progress((idx + 1) / len(active_cards))
st.caption(f"진행률: {idx + 1} / {len(active_cards)} 장 | 복습 대상(모르는 단어): {len(st.session_state.unknown_words)}개")

# 플래시카드 본문 (빈 공간 제거 + 중앙 정렬)
with st.container(border=True):
    front_label = "🇰🇷 한국어 뜻" if fc_direction == "한글 뜻 → 영어" else "🇺🇸 영어 용어"
    front_text = curr["korean"] if fc_direction == "한글 뜻 → 영어" else curr["english"]
    back_text = curr["english"] if fc_direction == "한글 뜻 → 영어" else curr["korean"]

    # 카드 상단 파트명 및 단어 중앙 배치
    st.markdown(
        f"""
        <div style="text-align: center; padding: 10px 0;">
            <span style="color: #9e9e9e; font-size: 0.9rem;">📂 {curr['part']}</span>
            <h4 style="margin-top: 12px; margin-bottom: 8px;">{front_label}</h4>
            <h1 style="margin-top: 0px; margin-bottom: 16px;">{front_text}</h1>
        </div>
        """,
        unsafe_allow_html=True
    )

    # 오디오 플레이어 중앙 배치
    _, audio_col, _ = st.columns([1, 2, 1])
    with audio_col:
        st.audio(get_audio_bytes(curr["english"]), format="audio/mp3")

    st.divider()

    if not st.session_state.fc_flipped:
        if st.button("👀 클릭해서 정답 확인", type="primary", use_container_width=True):
            st.session_state.fc_flipped = True
            st.rerun()
    else:
        st.markdown(
            f"""
            <div style="background-color: rgba(46, 125, 50, 0.25); border-radius: 8px; padding: 16px; text-align: center; margin-bottom: 16px;">
                <h2 style="margin: 0; color: #81c784;">{back_text}</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

        chk1, chk2 = st.columns(2)
        with chk1:
            if st.button("⭕ 암기 완료", use_container_width=True):
                if curr["english"] in st.session_state.unknown_words:
                    st.session_state.unknown_words.remove(curr["english"])
                if idx < len(active_cards) - 1:
                    st.session_state.fc_idx += 1
                st.session_state.fc_flipped = False
                st.rerun()
        with chk2:
            if st.button("❌ 모르는 단어 추가", use_container_width=True):
                st.session_state.unknown_words.add(curr["english"])
                if idx < len(active_cards) - 1:
                    st.session_state.fc_idx += 1
                st.session_state.fc_flipped = False
                st.rerun()

# 이전 / 섞기 / 다음 버튼
nav1, nav2, nav3 = st.columns(3)
with nav1:
    if st.button("⬅️ 이전 카드", disabled=(idx == 0), use_container_width=True):
        st.session_state.fc_idx -= 1
        st.session_state.fc_flipped = False
        st.rerun()
with nav2:
    if st.button("🔀 카드 섞기", use_container_width=True):
        random.shuffle(st.session_state.fc_cards)
        st.session_state.fc_idx = 0
        st.session_state.fc_flipped = False
        st.rerun()
with nav3:
    if st.button("다음 카드 ➡️", disabled=(idx >= len(active_cards) - 1), use_container_width=True):
        st.session_state.fc_idx += 1
        st.session_state.fc_flipped = False
        st.rerun()
