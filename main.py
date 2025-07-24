
import streamlit as st
from datetime import date
import base64

st.set_page_config(layout="centered")

# 로고 Base64 변환
with open("intro_logo.png", "rb") as img_file:
    encoded = base64.b64encode(img_file.read()).decode()

# 중앙 로고 출력
st.markdown(f"""
<div style='text-align: center;'>
    <img src='data:image/png;base64,{encoded}' width='250'>
</div>
""", unsafe_allow_html=True)


# 오늘 날짜 기준
today = date.today()

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("조회 시작일", value=today, key="start_date")
with col2:
    end_date = st.date_input("조회 종료일", value=today, key="end_date")

if start_date > end_date:
    st.error("시작 날짜는 끝 날짜보다 이전이어야 합니다.")
else:
    st.success(f"선택된 기간: {start_date} ~ {end_date}")



user_id = st.text_input(label="고객번호",label_visibility="collapsed", placeholder="👤 고객번호 혹은 한전ON ID", key="user_id_style")
user_pw = st.text_input(label="비밀번호",label_visibility="collapsed", placeholder="🔒 비밀번호", type="password", key="user_pw_style")


# CSS 정의
st.markdown("""
<style>
div.stButton > button:first-child {
    background-color: #09A3E0;
    color: white;
    border: none;
    border-radius: 25px;
    padding: 12px;
    font-size: 15px;
    font-weight: bold;
    width: 100%;
    cursor: pointer;
}
div.stButton > button:first-child:hover {
    background-color: #078FC2;
}
</style>
""", unsafe_allow_html=True)

search_button_clicked = st.button("조회")

if search_button_clicked:
    try:
        if not user_id or not user_pw:
            st.warning("고객번호(ID)와 비밀번호를 모두 입력해주세요.")
        else:
            from kepco import login, download_data, merge_files

            with st.spinner("🔐 로그인 중..."):
                login_success = login(user_id, user_pw)

            if login_success:
                with st.spinner("📥 데이터 다운로드 중..."):
                    download_data(start_date, end_date)

                with st.spinner("🗃️ 파일 병합 중..."):
                    merge_files()

                st.success("✅ 데이터 다운로드 및 병합이 완료되었습니다.")
            else:
                st.error("❌ ID 또는 비밀번호가 올바르지 않습니다. 다시 확인해주세요.")

    except PermissionError:
        st.error("⚠️ 엑셀 파일이 열려 있어 작업을 완료할 수 없습니다. 열려 있는 파일을 모두 닫고 다시 시도해 주세요.")
    except ValueError as e:
        st.error(str(e))


