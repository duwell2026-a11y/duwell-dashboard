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
import io

# --------------------------------------------------------------------------
# 1. 페이지 및 디자인 설정 (네이버 스마트스토어 테마 + 레이아웃 고정)
# --------------------------------------------------------------------------

st.set_page_config(page_title="DUWELL 판매자센터", layout="wide", page_icon="🛍️")

# 🎨 [디자인 커스텀] 사장님의 기존 스타일 유지 및 그래프/숫자판 크기 고정
st.markdown("""
    <style>
        /* 1. 폰트 및 기본 배경 */
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        html, body, [class*="css"] {
            font-family: 'Pretendard', sans-serif;
        }
        .stApp {
            background-color: #F5F6F8; /* 연한 회색 배경 */
        }
        
        /* 2. 사이드바 (네이버 스타일 다크 그레이) */
        [data-testid="stSidebar"] {
            background-color: #30343B;
        }
        [data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        
        /* 3. 헤더 숨김 */
        header[data-testid="stHeader"] {
            background-color: transparent;
        }

        /* 4. 숫자판(Metric) 디자인 및 높이 고정 */
        [data-testid="stMetric"] {
            background-color: #FFFFFF;
            padding: 20px;
            border-radius: 0px; 
            border: 1px solid #DEE2E6;
            box-shadow: 0 1px 2px rgba(0,0,0,0.03);
            min-height: 130px;
            max-height: 130px;
        }
        [data-testid="stMetricLabel"] {
            font-size: 14px !important;
            color: #767676 !important;
        }
        [data-testid="stMetricValue"] {
            font-size: 28px !important;
            color: #03C75A !important; /* 네이버 그린 */
            font-weight: 700;
        }

        /* 5. 버튼 디자인 (네이버 그린) */
        div.stButton > button {
            background-color: #03C75A;
            color: white;
            border-radius: 2px;
            border: 1px solid #02b351;
            padding: 10px 20px;
            font-weight: 600;
            font-size: 14px;
            width: 100%;
        }
        div.stButton > button:hover {
            background-color: #00b34e;
            color: white;
            border-color: #00b34e;
        }
        
        /* 6. 탭(Tab) 디자인 */
        .stTabs [data-baseweb="tab-list"] {
            gap: 24px;
            background-color: white;
            padding: 0 10px;
            border-bottom: 1px solid #ddd;
        }
        .stTabs [aria-selected="true"] {
            color: #03C75A !important;
            border-bottom: 3px solid #03C75A !important;
            font-weight: bold;
        }
        
        /* 7. 데이터프레임 (표) 스타일 */
        [data-testid="stDataFrame"] {
            background-color: white;
            border: 1px solid #DEE2E6;
        }

        /* 8. 그래프 영역 고정 */
        .chart-container {
            height: 450px;
            background-color: #FFFFFF;
            padding: 20px;
            border: 1px solid #DEE2E6;
            margin-bottom: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# 2. 키 파일 및 권한 설정
# --------------------------------------------------------------------------

local_key_path = r"D:\비서\google_key.json"
is_local = os.path.exists(local_key_path)

SHEET_ID = ""
GOOGLE_API_KEY = ""
SENDER_EMAIL = ""
SENDER_PASSWORD = ""
GOOGLE_CREDENTIALS = None

try:
    if is_local:
        # 🏠 [내 컴퓨터 모드]
        SHEET_ID = "1xqcbuzRzzp4i_Qsy4CKRjIIvGOTthT88bXxxY5RjEjQ"
        GOOGLE_API_KEY = "AIzaSyBBReb6mUNBeIGa2n-GJEt-lUphanHq3jg"
        SENDER_EMAIL = "duwell2026@gmail.com"
        SENDER_PASSWORD = "mvxo jzki djzg iwor"
        with open(local_key_path, "r", encoding="utf-8") as f:
            GOOGLE_CREDENTIALS = json.load(f)
    else:
        # ☁️ [GitHub/배포 모드]
        SHEET_ID = st.secrets["SHEET_ID"]
        GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
        SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
        SENDER_PASSWORD = st.secrets["SENDER_PASSWORD"]
        if "GOOGLE_JSON_KEY" in st.secrets:
            GOOGLE_CREDENTIALS = json.loads(st.secrets["GOOGLE_JSON_KEY"])
        else:
            GOOGLE_CREDENTIALS = st.secrets["google_credentials"]

    if GOOGLE_API_KEY:
        genai.configure(api_key=GOOGLE_API_KEY)

except Exception as e:
    st.error(f"❌ 설정 로드 실패: {e}")
    st.stop()

# --------------------------------------------------------------------------
# 🛠️ 함수 모음
# --------------------------------------------------------------------------

def get_best_model():
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
    except Exception:
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
        return f"🚨 AI 오류: {str(e)}"

def get_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if GOOGLE_CREDENTIALS:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(GOOGLE_CREDENTIALS, scope)
            return gspread.authorize(creds)
        return None
    except Exception as e:
        st.error(f"구글 시트 인증 실패: {e}")
        return None

def clean_date_str(date_val):
    s = str(date_val).strip()
    if not s or s == 'None': return None
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
    except Exception:
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

# ✨ [신규] 재고 부족 알림 함수 (사장님/사모님 동시 알림용)
def check_stock_and_alert(df_stock):
    df_stock['현재재고'] = pd.to_numeric(df_stock['현재재고'], errors='coerce').fillna(0)
    df_stock['안전재고'] = pd.to_numeric(df_stock['안전재고'], errors='coerce').fillna(0)
    low_items = df_stock[df_stock['현재재고'] <= df_stock['안전재고']]
    if not low_items.empty:
        msg = "🚨 [DUWELL 재고 부족 알림]\n\n다음 상품의 재고가 안전 수준 이하입니다:\n\n"
        for _, row in low_items.iterrows():
            msg += f"- {row['상품명']}: 현재 {int(row['현재재고'])}개 (안전재고: {int(row['안전재고'])}개)\n"
        msg += "\n빠른 확인 및 발주 부탁드립니다. 🍷"
        # 사장님 메일(SENDER_EMAIL)에 발송. 사모님 메일 추가 시 콤마로 연결 가능
        send_email_with_attach(SENDER_EMAIL, "[DUWELL] 🚨 긴급: 재고 부족 알림", msg)
        return True
    return False

# --------------------------------------------------------------------------
# 🏠 메인 UI 로직
# --------------------------------------------------------------------------

with st.sidebar:
    st.markdown("<h1 style='color:#800020;'>🍷 DUWELL</h1>", unsafe_allow_html=True)
    if st.button("🔄 데이터 새로고침", type="primary"):
        st.rerun()
    menu = st.radio("메뉴 이동", [
        "🏠 통합 모니터링", "📦 주문 일괄 등록", "💎 고객 CRM 센터", 
        "🛠️ 재고 관리", "🏭 공장 발주", "📢 마케팅 센터", 
        "🎨 디자인 시안실", "📅 일정 관리", "📋 주문 장부", "🛠️ 옵션 관리"
    ])

st.markdown(f"<h2 style='color:#333;'>{menu}</h2>", unsafe_allow_html=True)
st.divider()

# 데이터 로드
df_duwell, sheet_main = load_data("시트1") 
df_all = df_duwell.copy()
if not df_all.empty and '날짜' in df_all.columns:
    df_all['날짜'] = pd.to_datetime(df_all['날짜'], errors='coerce')
    df_all = df_all.sort_values(by='날짜', ascending=False)
    df_all['날짜_str'] = df_all['날짜'].dt.strftime('%Y-%m-%d')

# === [1] 🏠 통합 모니터링 ===
if menu == "🏠 통합 모니터링":
    today = datetime.now().strftime("%Y-%m-%d")
    c1, c2, c3 = st.columns(3)
    if not df_all.empty:
        df_all['금액_숫자'] = pd.to_numeric(df_all['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
        today_orders = df_all[df_all['날짜_str'] == today]
        today_sales = today_orders['금액_숫자'].sum()
        total_sales = df_all['금액_숫자'].sum()
        c1.metric("📦 오늘 주문건수", f"{len(today_orders)}건")
        c2.metric("💰 오늘 매출", f"{today_sales:,.0f}원")
        c3.metric("🏆 총 누적 매출", f"{total_sales:,.0f}원")
    else:
        c1.metric("📦 오늘 주문건수", "0건"); c2.metric("💰 오늘 매출", "0원"); c3.metric("🏆 총 누적 매출", "0원")

    st.markdown("---")
    if st.button("🚀 AI 일일 경영 브리핑 생성"):
        with st.spinner("AI 분석 중..."):
            if not df_all.empty:
                sales_summary = f"오늘 날짜: {today}. 오늘 주문 {len(today_orders)}건, 매출 {today_sales:,.0f}원."
                prompt = f"{sales_summary} 사장님께 하루를 시작하는 활기차고 격식있는 브리핑 멘트를 작성해줘."
                st.success(ask_ai(prompt))
            else: st.warning("데이터가 없습니다.")

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
            possible_cols = ['날짜_str', '구매자명', '상품명', '상태']
            cols = [c for c in possible_cols if c in df_all.columns]
            st.dataframe(df_all[cols].head(5), hide_index=True, use_container_width=True)
        else: st.info("주문 없음")

# === [2] 📦 주문 일괄 등록 (지능형 재고 차감 탑재) ===
elif menu == "📦 주문 일괄 등록":
    st.info("💡 마켓별로 상품명이 달라도 '매핑명' 키워드를 분석하여 재고를 자동 차감합니다.")
    uploaded_file = st.file_uploader("네이버 주문 엑셀 파일 업로드 (.xlsx)", type=['xlsx'])
    
    if uploaded_file:
        try:
            # 1. 엑셀 읽기
            df_new = pd.read_excel(uploaded_file, header=1)
            target_cols = {
                '상품주문번호': '주문번호', '주문일시': '날짜', '수취인명': '구매자명',
                '수취인연락처1': '연락처', '배송지': '주소', '상품명': '상품명',
                '수량': '수량', '총 주문금액': '결제금액', '배송메세지': '요청사항'
            }
            valid_cols = {k: v for k, v in target_cols.items() if k in df_new.columns}
            df_upload = df_new[list(valid_cols.keys())].rename(columns=valid_cols)
            
            st.write("🔽 업로드될 데이터 미리보기")
            st.dataframe(df_upload.head(3))
            
            # --- 💾 저장 및 지능형 차감 버튼 ---
            if st.button("💾 구글 시트 저장 및 지능형 재고 차감"):
                if sheet_main:
                    try:
                        # (1) 주문 데이터 저장
                        rows_to_add = []
                        for _, row in df_upload.iterrows():
                            rows_to_add.append([
                                str(row.get('날짜', '')), str(row.get('구매자명', '')), str(row.get('연락처', '')),
                                str(row.get('주소', '')), str(row.get('상품명', '')), str(row.get('수량', '1')),
                                str(row.get('결제금액', '0')), "", "", str(row.get('요청사항', '')), "", "신규(스마트스토어)"
                            ])
                        sheet_main.append_rows(rows_to_add)

                        # (2) ✨ 지능형 재고 차감 로직 (매핑명 분석)
                        try:
                            df_opt, _ = load_data("옵션관리")
                            df_stock, sheet_stock = load_data("재고관리")
                            
                            if not df_stock.empty and not df_opt.empty:
                                for _, order in df_upload.iterrows():
                                    market_p_name = str(order.get('상품명', '')) # 주문서의 긴 이름
                                    order_qty = int(order.get('수량', 1))
                                    
                                    target_std_name = None
                                    # 옵션관리 시트의 매핑 키워드 확인
                                    for _, opt in df_opt.iterrows():
                                        # 콤마로 구분된 키워드 리스트화 (예: "와플, 엠보싱" -> ["와플", "엠보싱"])
                                        keywords = [k.strip() for k in str(opt.get('매핑명', '')).split(',') if k.strip()]
                                        
                                        # 주문서 이름에 키워드가 하나라도 포함되어 있다면 매칭 성공!
                                        if any(kw in market_p_name for kw in keywords):
                                            target_std_name = opt.get('상품명') # 기준 상품명 추출
                                            break
                                    
                                    # 매칭된 상품의 재고 차감 실행
                                    if target_std_name:
                                        stock_records = sheet_stock.get_all_records()
                                        for idx, s_item in enumerate(stock_records):
                                            if str(s_item.get('상품명')).strip() == str(target_std_name).strip():
                                                current_qty = int(s_item.get('현재재고', 0))
                                                # B열(2열) 업데이트
                                                sheet_stock.update_cell(idx + 2, 2, current_qty - order_qty)
                                                break
                            
                            # 재고 부족 알림 체크
                            updated_stock, _ = load_data("재고관리")
                            check_stock_and_alert(updated_stock)
                            
                            st.success(f"✅ 총 {len(rows_to_add)}건 저장 및 지능형 재고 차감 완료!")
                            time.sleep(2)
                            st.rerun()

                        except Exception as stock_err:
                            st.warning(f"⚠️ 주문은 저장되었으나 재고 차감 중 오류 발생: {stock_err}")

                    except Exception as e:
                        st.error(f"❌ 데이터 저장 중 오류 발생: {e}")
                else:
                    st.error("구글 시트 연결 실패")
        
        except Exception as e:
            st.error(f"⚠️ 엑셀 파일 읽기 오류: {e}")

# === [3] 🏭 공장 발주 ===
elif menu == "🏭 공장 발주":
    if 'mail_body' not in st.session_state: st.session_state['mail_body'] = ""
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("📝 발주 내용 입력")
        with st.form("order_form"):
            factory_name = st.text_input("공장명"); factory_email = st.text_input("공장 이메일")
            items = st.text_area("발주 품목 및 내용"); uploaded_file = st.file_uploader("발주서 파일", type=['xlsx', 'xls', 'pdf'])
            if st.form_submit_button("🤖 AI 메일 초안 작성"):
                if not items: st.warning("내용을 입력하세요.")
                else:
                    prompt = f"수신: {factory_name}. 내용: {items}. 정중한 발주 메일 작성해줘."
                    st.session_state['mail_body'] = ask_ai(prompt); st.rerun()
    with c2:
        st.subheader("📧 메일 전송")
        final_body = st.text_area("메일 본문", value=st.session_state['mail_body'], height=300)
        if st.button("🚀 이메일 전송하기"):
            if not factory_email: st.error("이메일을 입력하세요.")
            else:
                ok, msg = send_email_with_attach(factory_email, f"[발주] (주)DUWELL 발주서 건", final_body, uploaded_file)
                if ok: st.success(msg)
                else: st.error(msg)

# === [4] 📢 마케팅 센터 ===
elif menu == "📢 마케팅 센터":
    st.info("💡 AI 마케팅 올인원")
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["📂 리뷰 엑셀 일괄 답글", "💬 리뷰 건별 답글", "✍️ 카피라이팅", "💡 네이밍", "📅 프로모션", "🆘 CS/후기", "💎 VIP 분석"])
    with t1:
        uploaded_review = st.file_uploader("리뷰 엑셀 파일 (.xlsx)", type=['xlsx'], key="review_xls")
        if uploaded_review:
            try:
                df_rev = pd.read_excel(uploaded_review, header=1)
                content_col = next((c for c in df_rev.columns if '리뷰' in c or '내용' in c), None)
                score_col = next((c for c in df_rev.columns if '평점' in c or '점수' in c), None)
                if content_col and score_col:
                    if st.button("🤖 AI 답글 일괄 생성 시작"):
                        with st.spinner("생성 중..."):
                            ai_replies = [ask_ai(f"리뷰:{str(row[content_col])}. 평점:{row[score_col]}. 감사 답글 작성.") for _, row in df_rev.iterrows()]
                            df_rev['AI_자동답글'] = ai_replies; st.success("🎉 생성 완료!")
                            buffer = io.BytesIO()
                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer: df_rev.to_excel(writer, index=False)
                            st.download_button("📥 다운로드", data=buffer.getvalue(), file_name="리뷰답글완료.xlsx")
            except Exception as e: st.error(f"오류: {e}")
    with t2:
        rv_text = st.text_area("리뷰 내용")
        if st.button("🤖 답글 추천"): st.write(ask_ai(f"리뷰: {rv_text}. 답글 추천해줘."))
    with t3:
        p_name = st.text_input("상품명")
        if st.button("✨ 문구 생성"): st.write(ask_ai(f"상품:{p_name}. SNS 홍보 문구 작성."))
    with t4:
        n_desc = st.text_area("브랜드 특징")
        if st.button("이름 추천"): st.write(ask_ai(f"특징: {n_desc}. 네이밍 제안."))
    with t5:
        pr_goal = st.text_input("프로모션 목표")
        if st.button("기획안 생성"): st.write(ask_ai(f"목표: {pr_goal}. 프로모션 기획안 작성."))
    with t6:
        cs_txt = st.text_area("고객 문의")
        if st.button("답변 생성"): st.write(ask_ai(f"문의: {cs_txt}. 정중한 CS 답변 작성."))
    with t7:
        if not df_all.empty:
            df_vip = df_all.copy()
            df_vip['amt'] = pd.to_numeric(df_vip['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
            st.dataframe(df_vip.groupby('구매자명')['amt'].sum().sort_values(ascending=False).head(10))

# === [5] 🎨 디자인 시안실 ===
elif menu == "🎨 디자인 시안실":
    st.subheader("🎨 시안 작업 관리")
    if df_duwell.empty: st.warning("데이터가 없습니다.")
    else:
        tab_wait, tab_done = st.tabs(["🔥 작업 대기중", "✅ 작업 완료"])
        with tab_wait:
            df_wait = df_duwell[df_duwell['상태'] != '완료']
            for i, r in df_wait.iterrows():
                with st.expander(f"📌 {r.get('구매자명')} - {r.get('상품명')}"):
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        link = str(r.get('디자인파일', ''))
                        drive_id = get_drive_id(link)
                        if drive_id: st.image(f"https://drive.google.com/thumbnail?id={drive_id}&sz=w400")
                        else: st.text("이미지 없음")
                    with c2:
                        st.write(f"요청: {r.get('요청사항', '-')}")
                        if st.button("✅ 완료 처리", key=f"btn_{i}"):
                            success, msg = update_status_in_sheet(sheet_main, r, "완료")
                            if success: st.success(msg); time.sleep(1); st.rerun()
        with tab_done: st.dataframe(df_duwell[df_duwell['상태'] == '완료'])

# === [6] 📅 일정 관리 ===
elif menu == "📅 일정 관리":
    st.subheader("📅 일정 캘린더")
    df_sch, sheet_sch = load_data("일정관리")
    col1, col2 = st.columns([1, 2])
    with col1:
        with st.form("add_schedule"):
            d_date = st.date_input("날짜"); d_time = st.time_input("시간"); d_title = st.text_input("일정명"); d_desc = st.text_area("상세내용")
            if st.form_submit_button("저장"):
                if sheet_sch: sheet_sch.append_row([str(d_date), str(d_date), str(d_time), d_title, d_desc]); st.success("저장됨"); st.rerun()
        audio_file = st.file_uploader("음성 일정 추가", type=['mp3', 'wav', 'm4a'])
        if audio_file and st.button("음성 분석"): st.info(process_audio(audio_file))
    with col2:
        if not df_sch.empty:
            events = [{"title": str(r.get('일정명')), "start": str(r.get('시작일'))} for _, r in df_sch.iterrows()]
            calendar(events=events)

# === [7] 📋 주문 장부 ===
elif menu == "📋 주문 장부":
    st.subheader("📋 전체 주문 장부")
    if not df_all.empty:
        st.dataframe(df_all, use_container_width=True)
        csv = df_all.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 다운로드", csv, "order_list.csv", "text/csv")

# === [8] 🛠️ 옵션 관리 (매핑 컬럼 활성화 버전) ===
elif menu == "🛠️ 옵션 관리":
    st.subheader("🛠️ 옵션 및 통합 상품명 관리")
    
    df_opt, sheet_opt = load_data("옵션관리")
    
    # 만약 시트에 '매핑명' 열이 없다면 안내 메시지 출력
    if '매핑명' not in df_opt.columns:
        st.error("⚠️ 구글 시트 '옵션관리' 탭 맨 오른쪽에 '매핑명' 컬럼을 추가해 주세요!")
    
    if not df_opt.empty:
        # 표에서 직접 '매핑명'을 입력할 수 있도록 설정
        edited_df = st.data_editor(
            df_opt, 
            num_rows="dynamic", 
            use_container_width=True,
            key="opt_map_editor"
        )
        
        if st.button("💾 설정 및 매핑명 저장"):
            try:
                sheet_opt.clear()
                # 수정된 데이터를 헤더와 함께 시트에 다시 덮어씁니다.
                sheet_opt.update([edited_df.columns.values.tolist()] + edited_df.values.tolist())
                st.success("✅ '매핑명'을 포함한 모든 설정이 저장되었습니다!")
                st.rerun()
            except Exception as e:
                st.error(f"저장 오류: {e}")

    # 하단 가이드 (이미지 2번처럼 예시를 보여줌)
    with st.expander("💡 매핑명 입력 방법 (예시)"):
        st.write("매핑명 칸에 여러 이름을 넣을 때는 콤마(,)로 구분해 주세요.")
        st.code("예: 호텔 타월, 솔리드 타월, 기본 수건")

# === [9] 💎 고객 CRM 센터 (히스토리 누적 포함) ===
elif menu == "💎 고객 CRM 센터":
    st.subheader("💎 고객 통합 프로필 및 상담 관리")
    if df_all.empty: st.warning("데이터가 없습니다.")
    else:
        df_crm = df_all.copy()
        df_crm['amt'] = pd.to_numeric(df_crm['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
        try:
            cust_profile = df_crm.groupby('구매자명').agg({'날짜': ['max', 'count'], 'amt': 'sum'}).reset_index()
            cust_profile.columns = ['고객명', '최근구매일', '구매횟수', '누적금액']
            def analyze_cx(row):
                grade = "💎 VIP" if row['누적금액'] >= 500000 else "🥈 일반"
                days = (datetime.now() - row['최근구매일']).days if pd.notnull(row['최근구매일']) else 0
                status = "🔔 교체주기" if 150 <= days <= 210 else "✅ 정상"
                return pd.Series([grade, status, days])
            cust_profile[['등급', '상태', '경과일']] = cust_profile.apply(analyze_cx, axis=1)

            t1, t2 = st.tabs(["👤 통합 프로필", "🎯 스마트 타겟팅"])
            with t1:
                search_nm = st.text_input("고객명 검색", "")
                f_df = cust_profile[cust_profile['고객명'].str.contains(search_nm, na=False)].copy()
                event = st.dataframe(f_df, on_select="rerun", selection_mode="single-row", use_container_width=True, hide_index=True)
                selected = event.selection.rows
                if selected:
                    sel = f_df.iloc[selected[0]]; st.divider()
                    c_a, c_b = st.columns(2)
                    with c_a:
                        st.markdown(f"### 👤 {sel['고객명']} 프로필")
                        st.markdown("#### 📜 상담 히스토리")
                        current_history = ""
                        try:
                            client = get_client(); target_sh = client.open("주문데이터").worksheet("시트1")
                            h = target_sh.row_values(1)
                            if '비고' in h:
                                cell = target_sh.find(sel['고객명'])
                                if cell:
                                    current_history = target_sh.cell(cell.row, h.index('비고')+1).value
                                    st.text_area("기록", value=current_history or "내용 없음", height=150, disabled=True)
                        except Exception as e: st.caption(f"로드 오류: {e}")
                        memo_in = st.text_area("📝 신규 상담", key=f"memo_{sel['고객명']}")
                        if st.button("💾 누적 저장"):
                            try:
                                now = datetime.now().strftime('%Y-%m-%d %H:%M')
                                final = f"{current_history}\n[{now}] {memo_in}" if current_history else f"[{now}] {memo_in}"
                                target_sh.update_cell(cell.row, h.index('비고')+1, final)
                                st.success("저장됨"); st.rerun()
                            except Exception as e: st.error(f"오류: {e}")
                    with c_b:
                        st.markdown("### 🤖 AI 마케팅"); p_msg = f"{sel['고객명']}님을 위한 와인색 감성 메시지 작성해줘."
                        if st.button("✨ 문구 생성"): st.write(ask_ai(p_msg))
            with t2:
                risk_df = cust_profile[cust_profile['상태']=='🔔 교체주기']
                st.success(f"📍 재구매 알림 대상 ({len(risk_df)}명)"); st.dataframe(risk_df, hide_index=True)
        except Exception as e: st.error(f"오류: {e}")

# === [10] 🛠️ 재고 관리 (그래프 고정 및 배경 버그 수정) ===
elif menu == "🛠️ 재고 관리":
    st.subheader("📦 DUWELL 실시간 재고 모니터링")
    df_stock, sheet_stock = load_data("재고관리")
    
    if not df_stock.empty:
        # 1. 상단 요약 지표 (높이 고정)
        df_stock['현재재고'] = pd.to_numeric(df_stock['현재재고'], errors='coerce').fillna(0)
        df_stock['안전재고'] = pd.to_numeric(df_stock['안전재고'], errors='coerce').fillna(0)
        
        cols = st.columns(4)
        for i, (idx, row) in enumerate(df_stock.iterrows()):
            if i < 4:
                is_low = row['현재재고'] <= row['안전재고']
                cols[i].metric(
                    label=row['상품명'], 
                    value=f"{int(row['현재재고'])}개", 
                    delta="-재고부족" if is_low else "정상", 
                    delta_color="inverse" if is_low else "normal"
                )
        
        st.divider()

        # 2. 📊 재고 시각화 (높이를 400px로 완전 고정하여 '움직임' 방지)
        st.markdown("#### 📊 재고 현황 (현재고 vs 안전재고)")
        
        # 데이터를 시각화용으로 정리
        chart_data = df_stock.set_index('상품명')[['현재재고', '안전재고']]
        
        # st.bar_chart에 height 설정을 직접 주어 높이를 고정합니다.
        st.bar_chart(chart_data, height=400, use_container_width=True)

        st.divider()

        # 3. 입고 처리 및 명세
        c_in, c_list = st.columns([1, 2])
        with c_in:
            with st.form("in_form"):
                st.markdown("##### ➕ 상품 입고")
                target_p = st.selectbox("품목 선택", df_stock['상품명'].tolist())
                qty = st.number_input("입고 수량", min_value=1)
                if st.form_submit_button("입고 완료"):
                    try:
                        cell = sheet_stock.find(target_p)
                        curr = int(sheet_stock.cell(cell.row, 2).value)
                        sheet_stock.update_cell(cell.row, 2, curr + qty)
                        st.success("반영되었습니다.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")
        with c_list:
            st.markdown("##### 📋 상세 재고 데이터")
            st.dataframe(df_stock, use_container_width=True, hide_index=True)
    else:
        st.info("'재고관리' 시트에 데이터를 입력해주세요.")
