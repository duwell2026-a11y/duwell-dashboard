import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import re
import time
import requests
import base64
import json
import os
from PIL import Image
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 
from streamlit_calendar import calendar
import google.generativeai as genai 

# --------------------------------------------------------------------------
# 🚨 [스마트 설정 구역] - 웹/로컬 자동 감지 (수정 금지)
# --------------------------------------------------------------------------

# 페이지 설정 (가장 먼저 실행되어야 함)
st.set_page_config(page_title="DUWELL 통합 관제센터", layout="wide", page_icon="🍷")

# 1. 내 컴퓨터(로컬)에 키 파일이 있는지 확인
local_key_path = r"D:\비서\google_key.json"
is_local = os.path.exists(local_key_path)

SHEET_ID = ""
GOOGLE_API_KEY = ""
SENDER_EMAIL = ""
SENDER_PASSWORD = ""
GOOGLE_CREDENTIALS = None

try:
    if is_local:
        # 🏠 [내 컴퓨터 모드] - D드라이브 파일 사용
        print("💻 내 컴퓨터(로컬) 환경에서 실행 중입니다.")
        
        # 사장님 원래 설정값 (로컬용)
        SHEET_ID = "1xqcbuzRzzp4i_Qsy4CKRjIIvGOTthT88bXxxY5RjEjQ"
        GOOGLE_API_KEY = "AIzaSyBBReb6mUNBeIGa2n-GJEt-lUphanHq3jg"
        SENDER_EMAIL = "duwell2026@gmail.com"
        SENDER_PASSWORD = "mvxo jzki djzg iwor"
        
        # 로컬 파일에서 인증 정보 로드
        with open(local_key_path, "r", encoding="utf-8") as f:
            GOOGLE_CREDENTIALS = json.load(f)

    else:
        # ☁️ [웹 배포 모드] - Streamlit Secrets 사용
        # Streamlit Cloud에 올리면 자동으로 이 부분이 실행됩니다.
        
        SHEET_ID = st.secrets["SHEET_ID"]
        GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
        SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
        SENDER_PASSWORD = st.secrets["SENDER_PASSWORD"]
        
        # Secrets에 저장된 JSON 문자열을 딕셔너리로 변환
        if "GOOGLE_JSON_KEY" in st.secrets:
            GOOGLE_CREDENTIALS = json.loads(st.secrets["GOOGLE_JSON_KEY"])
        else:
            # 예비용 (혹시 json 문자열 방식이 아닐 경우)
            GOOGLE_CREDENTIALS = st.secrets["google_credentials"]

    # AI 설정 초기화
    if GOOGLE_API_KEY:
        genai.configure(api_key=GOOGLE_API_KEY)

except Exception as e:
    st.error(f"❌ 설정 로드 실패: {e}")
    st.stop()

# --------------------------------------------------------------------------
# 🛠️ 함수 모음
# --------------------------------------------------------------------------

def get_best_model():
    """구글 서버에 직접 물어봐서 사용 가능한 모델을 찾아냅니다."""
    try:
        available_models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                available_models.append(m.name)
        
        for m in available_models:
            if 'flash' in m.lower(): return m
        for m in available_models:
            if 'pro' in m.lower() and 'vision' not in m.lower(): return m

        return available_models[0] if available_models else "models/gemini-pro"
        
    except Exception as e:
        return "gemini-pro"

def ask_ai(prompt, images=None):
    if not GOOGLE_API_KEY: return "🚫 API 키가 설정되지 않았습니다."
    try:
        model_name = get_best_model()
        model = genai.GenerativeModel(model_name)
        
        content = [prompt]
        if images:
            if isinstance(images, list):
                for img in images: content.append(Image.open(img))
            else:
                content.append(Image.open(images))
        
        response = model.generate_content(content)
        return response.text
    except Exception as e:
        return f"🚨 AI 오류 ({get_best_model()}): {str(e)}"

def get_client():
    """구글 시트 인증 함수 (수정됨)"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        # 파일 경로가 아니라, 위에서 로드한 딕셔너리(GOOGLE_CREDENTIALS)를 직접 사용
        if GOOGLE_CREDENTIALS:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(GOOGLE_CREDENTIALS, scope)
            return gspread.authorize(creds)
        else:
            return None
    except Exception as e:
        st.error(f"구글 시트 인증 실패: {e}")
        return None

def clean_date_str(date_val):
    s = str(date_val).strip()
    if not s: return None
    nums = re.findall(r'\d+', s)
    if len(nums) >= 3:
        y, m, d = nums[0], nums[1], nums[2]
        if len(y) == 2: y = "20" + y
        return f"{y}-{int(m):02d}-{int(d):02d}"
    return s

def load_data(sheet_name):
    client = get_client()
    if not client: return pd.DataFrame(), None
    
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        
        if df.empty: return df, sheet

        df.columns = [c.strip() for c in df.columns]
        for col in ['날짜', '시작일', '종료일', '주문일시', '주문일']:
            if col in df.columns:
                df[col] = df[col].apply(clean_date_str)

        rename_map = {
            '주문일시': '날짜', '주문일': '날짜', '일자': '날짜',
            '금액': '결제금액', '예상견적': '결제금액',
            '성함': '구매자명', '고객명': '구매자명', '이름': '구매자명',
            '상품': '상품명', '품목': '상품명',
            '디자인파일': '디자인파일', '첨부파일': '디자인파일',
            '상태': '상태', '진행상태': '상태'
        }
        df.rename(columns=rename_map, inplace=True)
        df = df.loc[:, ~df.columns.duplicated()]

        if '주문처' not in df.columns: df['주문처'] = '🏠 자사몰'
        if '상태' not in df.columns: df['상태'] = '신규'
        
        return df, sheet
    except Exception as e:
        return pd.DataFrame(), None

def update_status_in_sheet(sheet, row_data, new_status="완료"):
    try:
        records = sheet.get_all_records()
        target_row_idx = -1
        for idx, record in enumerate(records):
            r_name = record.get('성함') or record.get('구매자명') or record.get('이름')
            r_item = record.get('상품') or record.get('상품명')
            if str(r_name) == str(row_data.get('구매자명')) and str(r_item) == str(row_data.get('상품명')):
                target_row_idx = idx + 2 
                break
        
        if target_row_idx != -1:
            header = sheet.row_values(1)
            col_idx = -1
            for i, h in enumerate(header):
                if h.strip() in ['상태', '진행상태']:
                    col_idx = i + 1
                    break
            if col_idx != -1:
                sheet.update_cell(target_row_idx, col_idx, new_status)
                return True, "✅ 상태 업데이트 성공!"
            else:
                return False, "❌ '상태' 컬럼 없음"
        else:
            return False, "❌ 주문 찾기 실패"
    except Exception as e:
        return False, f"❌ 오류: {str(e)}"

def get_drive_id(url):
    if not url: return None
    url = str(url)
    patterns = [r'id=([-a-zA-Z0-9_]+)', r'/file/d/([-a-zA-Z0-9_]+)', r'open\?id=([-a-zA-Z0-9_]+)']
    for p in patterns:
        match = re.search(p, url)
        if match: return match.group(1)
    return None

def send_email_with_attach(to, subject, body, attachment_file=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        if attachment_file:
            part = MIMEApplication(attachment_file.read(), Name=attachment_file.name)
            part['Content-Disposition'] = f'attachment; filename="{attachment_file.name}"'
            msg.attach(part)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(SENDER_EMAIL, SENDER_PASSWORD)
            s.send_message(msg)
        return True, "✅ 전송 성공"
    except Exception as e:
        return False, f"❌ 전송 실패: {str(e)}"

def process_audio(uploaded_file):
    try:
        if not GOOGLE_API_KEY: return "API 키 없음"
        with open("temp_audio_file.mp3", "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        myfile = genai.upload_file("temp_audio_file.mp3")
        model = genai.GenerativeModel(get_best_model()) 
        result = model.generate_content(["이 음성 파일 내용을 요약하고, 일정(날짜,시간,내용)이 있다면 추출해줘.", myfile])
        return result.text
    except Exception as e:
        return f"오류: {str(e)}"

# --- 🏠 메인 UI 로직 ---

with st.sidebar:
    st.markdown("<h1 style='color:#800020;'>🍷 DUWELL</h1>", unsafe_allow_html=True)
    if st.button("🔄 데이터 새로고침", type="primary"):
        st.rerun()
    menu = st.radio("메뉴 이동", ["🏠 통합 모니터링", "🏭 공장 발주", "📢 마케팅 센터", "🎨 디자인 시안실", "📅 일정 관리", "📋 주문 장부", "🛠️ 옵션 관리"])

st.markdown(f"<h2 style='color:#333;'>{menu}</h2>", unsafe_allow_html=True)
st.divider()

# 데이터 로드
df_duwell, sheet_main = load_data("시트1") 
df_all = df_duwell.copy()
if not df_all.empty and '날짜' in df_all.columns:
    df_all['날짜'] = pd.to_datetime(df_all['날짜'], errors='coerce')
    df_all = df_all.sort_values(by='날짜', ascending=False)
    df_all['날짜'] = df_all['날짜'].dt.strftime('%Y-%m-%d')

# === [1] 🏠 통합 모니터링 ===
if menu == "🏠 통합 모니터링":
    today = datetime.now().strftime("%Y-%m-%d")
    c1, c2, c3 = st.columns(3)
    if not df_all.empty:
        df_all['금액_숫자'] = pd.to_numeric(df_all['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
        today_orders = df_all[df_all['날짜'] == today]
        today_sales = today_orders['금액_숫자'].sum()
        total_sales = df_all['금액_숫자'].sum()
        c1.metric("📦 오늘 주문건수", f"{len(today_orders)}건")
        c2.metric("💰 오늘 매출", f"{today_sales:,.0f}원")
        c3.metric("🏆 총 누적 매출", f"{total_sales:,.0f}원")
    else:
        c1.metric("📦 오늘 주문건수", "0건")
        c2.metric("💰 오늘 매출", "0원")
        c3.metric("🏆 총 누적 매출", "0원")

    st.markdown("---")
    if st.button("🚀 AI 일일 경영 브리핑 생성"):
        with st.spinner("AI가 데이터를 분석 중입니다..."):
            if not df_all.empty:
                sales_summary = f"오늘 날짜: {today}. 오늘 주문 {len(today_orders)}건, 매출 {today_sales:,.0f}원."
                prompt = f"{sales_summary} 사장님께 하루를 시작하는 활기차고 격식있는 브리핑 멘트를 작성해줘."
                st.success(ask_ai(prompt))
            else:
                st.warning("데이터가 없습니다.")

    col_l, col_r = st.columns([1, 1])
    with col_l:
        st.subheader("📅 오늘의 일정")
        df_sch, _ = load_data("일정관리")
        if not df_sch.empty:
            today_sch = df_sch[df_sch['시작일'] == today]
            if not today_sch.empty:
                for _, r in today_sch.iterrows(): st.info(f"⏰ {r.get('시간','')} | {r.get('일정명','')}")
            else: st.write("일정 없음")
        else: st.write("일정 로드 실패")

    with col_r:
        st.subheader("📦 최근 주문 (5건)")
        if not df_all.empty:
            possible_cols = ['날짜', '구매자명', '상품명', '상태']
            cols = [c for c in possible_cols if c in df_all.columns]
            st.dataframe(df_all[cols].head(5), hide_index=True, use_container_width=True)
        else: st.info("주문 없음")

# === [2] 🏭 공장 발주 ===
elif menu == "🏭 공장 발주":
    if 'mail_body' not in st.session_state: st.session_state['mail_body'] = ""
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("📝 발주 내용 입력")
        with st.form("order_form"):
            factory_name = st.text_input("공장명")
            factory_email = st.text_input("공장 이메일")
            items = st.text_area("발주 품목 및 내용")
            uploaded_file = st.file_uploader("발주서 파일(엑셀/PDF)", type=['xlsx', 'xls', 'pdf'])
            if st.form_submit_button("🤖 AI 메일 초안 작성"):
                if not items: st.warning("내용을 입력하세요.")
                else:
                    prompt = f"수신: {factory_name}. 내용: {items}. 정중하고 명확한 발주 메일 본문 작성해줘."
                    st.session_state['mail_body'] = ask_ai(prompt)
                    st.rerun()

    with c2:
        st.subheader("📧 메일 전송")
        final_body = st.text_area("메일 본문", value=st.session_state['mail_body'], height=300)
        if st.button("🚀 이메일 전송하기"):
            if not factory_email: st.error("이메일을 입력하세요.")
            else:
                ok, msg = send_email_with_attach(factory_email, f"[발주] (주)DUWELL {factory_name} 발주서 건", final_body, uploaded_file)
                if ok: st.success(msg)
                else: st.error(msg)

# === [3] 📢 마케팅 센터 ===
elif menu == "📢 마케팅 센터":
    st.info("💡 AI 마케팅/기획 올인원")
    t1, t2, t3, t4, t5 = st.tabs(["✍️ 카피라이팅", "💡 네이밍", "📅 프로모션", "🆘 CS/후기", "💎 VIP 분석"])
    
    with t1:
        st.subheader("✍️ SNS 홍보 문구 작성")
        col1, col2 = st.columns(2)
        with col1:
            product = st.text_input("상품명")
            target = st.text_input("타겟 고객 (예: 20대 여성)")
        with col2:
            channel = st.selectbox("업로드 채널", ["인스타그램", "블로그", "스마트스토어 상세페이지"])
            tone = st.selectbox("말투", ["감성적인", "전문적인", "유머러스한"])
        if st.button("✨ 문구 생성"):
            st.write(ask_ai(f"상품: {product}, 타겟: {target}, 채널: {channel}, 말투: {tone}. 마케팅 문구 작성해줘."))

    with t2:
        st.subheader("💡 브랜드/상품 네이밍")
        desc = st.text_area("제품 특징/컨셉")
        if st.button("이름 추천받기"):
            st.write(ask_ai(f"제품 특징: {desc}. 기억에 남는 브랜드 네임 5개 추천해주고 이유도 설명해줘."))

    with t3:
        st.subheader("📅 프로모션 기획")
        goal = st.text_input("행사 목표 (예: 재고 소진)")
        if st.button("기획안 받기"):
            st.write(ask_ai(f"목표: {goal}. 실행 가능한 프로모션 아이디어와 기획안 3가지 제안해줘."))
            
    with t4:
        st.subheader("🆘 고객 후기/문의 분석")
        review_txt = st.text_area("고객의 글 붙여넣기")
        if st.button("답변 생성"):
            st.write(ask_ai(f"이 글을 분석하고 정중한 답변 작성해줘: {review_txt}"))
            
    with t5:
        st.subheader("💎 VIP 고객 분석")
        if not df_all.empty:
            df_vip = df_all.copy()
            df_vip['금액_숫자'] = pd.to_numeric(df_vip['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
            if '구매자명' in df_vip.columns:
                vip_group = df_vip.groupby('구매자명')['금액_숫자'].sum().sort_values(ascending=False).head(10)
                st.dataframe(vip_group, use_container_width=True)
            else: st.warning("구매자명 컬럼 없음")

# === [4] 🎨 디자인 시안실 ===
elif menu == "🎨 디자인 시안실":
    st.subheader("🎨 시안 작업 관리")
    if df_duwell.empty: st.warning("데이터가 없습니다.")
    else:
        tab_wait, tab_done = st.tabs(["🔥 작업 대기중", "✅ 작업 완료"])
        with tab_wait:
            df_wait = df_duwell[df_duwell['상태'] != '완료']
            if df_wait.empty: st.info("대기 중인 작업 없음")
            else:
                for i, r in df_wait.iterrows():
                    with st.expander(f"📌 {r.get('구매자명')} - {r.get('상품명')}"):
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            link = str(r.get('디자인파일', ''))
                            drive_id = get_drive_id(link)
                            if drive_id: st.image(f"https://drive.google.com/thumbnail?id={drive_id}&sz=w400", caption="미리보기")
                            elif link.startswith('http'): st.image(link)
                            else: st.text("이미지 없음")
                        with c2:
                            st.write(f"**요청:** {r.get('요청사항', '-')}")
                            st.write(f"**날짜:** {r.get('날짜', '-')}")
                            st.write(f"**파일:** {link}")
                            if st.button("✅ 완료 처리", key=f"btn_{i}"):
                                success, msg = update_status_in_sheet(sheet_main, r, "완료")
                                if success:
                                    st.success(msg)
                                    time.sleep(1)
                                    st.rerun()
                                else: st.error(msg)
        with tab_done:
            df_done = df_duwell[df_duwell['상태'] == '완료']
            st.dataframe(df_done, use_container_width=True)

# === [5] 📅 일정 관리 ===
elif menu == "📅 일정 관리":
    st.subheader("📅 일정 캘린더")
    df_sch, sheet_sch = load_data("일정관리")
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("#### ➕ 일정 추가")
        with st.form("add_schedule"):
            d_date = st.date_input("날짜")
            d_time = st.time_input("시간")
            d_title = st.text_input("일정명")
            d_desc = st.text_area("상세내용")
            if st.form_submit_button("저장"):
                if sheet_sch:
                    sheet_sch.append_row([str(d_date), str(d_date), str(d_time), d_title, d_desc])
                    st.success("저장됨")
                    st.rerun()
        st.markdown("#### 🎙️ 음성 일정 추가")
        audio_file = st.file_uploader("음성 파일", type=['mp3', 'wav', 'm4a'])
        if audio_file and st.button("음성 분석"):
            st.info(process_audio(audio_file))
    with col2:
        if not df_sch.empty:
            events = []
            for _, row in df_sch.iterrows():
                events.append({
                    "title": f"{row.get('일정명')}",
                    "start": str(row.get('시작일')),
                    "backgroundColor": "#FF4B4B" if "미팅" in str(row.get('일정명')) else "#3788d8"
                })
            calendar(events=events, options={"headerToolbar": {"left": "prev,next today", "center": "title", "right": "dayGridMonth,listMonth"}})
        else: st.write("일정 없음")

# === [6] 📋 주문 장부 ===
elif menu == "📋 주문 장부":
    st.subheader("📋 전체 주문 장부")
    if not df_all.empty:
        st.dataframe(df_all, use_container_width=True)
        csv = df_all.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 엑셀(CSV)로 다운로드", csv, "order_list.csv", "text/csv")
    else: st.info("데이터가 없습니다.")

# === [7] 🛠️ 옵션 관리 ===
elif menu == "🛠️ 옵션 관리":
    st.subheader("🛠️ 옵션 관리")
    df_opt, sheet_opt = load_data("옵션관리")
    if not df_opt.empty:
        edited_df = st.data_editor(df_opt, num_rows="dynamic", use_container_width=True)
        if st.button("💾 저장"):
            if sheet_opt:
                sheet_opt.clear()
                sheet_opt.update([edited_df.columns.values.tolist()] + edited_df.values.tolist())
                st.success("저장됨!")
    else: st.info("'옵션관리' 시트가 없습니다.")