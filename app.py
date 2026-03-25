import textwrap
import streamlit as st
import altair as alt
import pandas as pd
import gspread
from datetime import datetime, timedelta
import re
import time
import json
import os
from PIL import Image
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 
from streamlit_calendar import calendar
import google.generativeai as genai 
import streamlit.components.v1 as components 
import pdfkit
import platform
import io

# --------------------------------------------------------------------------
# 0. 페이지 설정 및 디자인 (CSS)
# --------------------------------------------------------------------------
st.set_page_config(page_title="DUWELL 스마트 ERP", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        /* 폰트 및 배경 */
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        html, body, [class*="css"] { font-family: 'Pretendard', sans-serif; }
        .stApp { background-color: #F8F9FA; }

        /* 사이드바 메뉴 스타일링 - 동그라미/빨간선 제거 */
        [data-testid="stSidebar"] { background-color: #FFFFFF !important; border-right: 1px solid #E9ECEF; }
        [data-testid="stSidebarNav"] { display: none; } /* 기본 네비게이션 숨김 */
        
        /* 라디오 버튼 본체(동그라미) 숨기기 */
        [data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child { display: none !important; }
        
        /* 선택 시 빨간 테두리/그림자 제거 */
        [data-testid="stSidebar"] div[role="radiogroup"] > label:focus-within { box-shadow: none !important; outline: none !important; }

        /* 메뉴 버튼 디자인 */
        [data-testid="stSidebar"] div[role="radiogroup"] > label {
            padding: 12px 20px !important;
            margin-bottom: 8px !important;
            border-radius: 10px !important;
            border: 1px solid transparent !important;
            transition: all 0.3s ease;
            cursor: pointer !important;
            background-color: transparent !important;
        }
        
        /* 호버 및 선택 상태 */
        [data-testid="stSidebar"] div[role="radiogroup"] > label:hover { background-color: #F1F3F5 !important; }
        [data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] {
            background-color: #2B3A55 !important; /* DUWELL 네이비 */
            box-shadow: 0 4px 12px rgba(43,58,85,0.15) !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"] p {
            color: #FFFFFF !important; font-weight: 700 !important;
        }

        /* 메트릭 카드 디자인 */
        [data-testid="metric-container"] {
            background-color: #FFFFFF; border-radius: 12px; padding: 20px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04); border: 1px solid #E9ECEF;
        }
        @page { 
            size: A4 landscape; 
            margin: 10mm; 
        }
    </style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# 1. 보안 및 로그인
# --------------------------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.markdown("<h2 style='text-align: center; margin-top: 100px;'>DUWELL 스마트 센터</h2>", unsafe_allow_html=True)
    with st.form("login_form"):
        user_choice = st.selectbox("접속자 선택", ["고은정 (대표)", "두재훈 (팀장)"])
        pwd = st.text_input("비밀번호 (PIN)", type="password")
        if st.form_submit_button("입장하기", use_container_width=True):
            if pwd == "1121":
                st.session_state['logged_in'] = True
                st.session_state['current_user'] = user_choice
                st.rerun()
            else: st.error("🚨 비밀번호가 틀렸습니다.")
    st.stop()

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
        SHEET_ID = "1xqcbuzRzzp4i_Qsy4CKRjIIvGOTthT88bXxxY5RjEjQ"
        SENDER_EMAIL = "duwell2026@gmail.com"
        GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
        SENDER_PASSWORD = st.secrets["SENDER_PASSWORD"]
        with open(local_key_path, "r", encoding="utf-8") as f:
            GOOGLE_CREDENTIALS = json.load(f)
    else:
        SHEET_ID = st.secrets["SHEET_ID"]
        SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
        GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
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
# 3. 핵심 함수 모음 
# --------------------------------------------------------------------------
def render_page_header(title, subtitle):
    st.markdown(f"""
        <div style="background: linear-gradient(135deg, #2B3A55 0%, #1A2235 100%); 
                    padding: 25px 30px; border-radius: 12px; color: white; margin-bottom: 25px;
                    box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-left: 6px solid #800020;">
            <h2 style="margin:0; font-size:1.6rem; font-weight:800; color:white; letter-spacing:-0.5px;">{title}</h2>
            <p style="margin:8px 0 0 0; font-size:0.95rem; opacity:0.85; color:white;">{subtitle}</p>
        </div>
    """, unsafe_allow_html=True)

@st.cache_resource(ttl=3600)
def get_client():
    try:
        creds_dict = dict(GOOGLE_CREDENTIALS)
        return gspread.service_account_from_dict(creds_dict)
    except Exception as e:
        st.error(f"인증 에러: {e}")
        return None

@st.cache_data(ttl=300) 
def fetch_raw_data(sheet_name):
    client = get_client()
    if not client: return []
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        return sheet.get_all_records()
    except Exception:
        return []

def clean_date_str(date_val):
    s = str(date_val).strip()
    if not s or s == 'None': return None
    nums = re.findall(r'\d+', s)
    if len(nums) >= 3:
        y, m, d = nums[0], nums[1], nums[2]
        if len(y) == 2: y = "20" + y
        return f"{y}-{int(m):02d}-{int(d):02d}"
    return s

def get_color_name(raw_color):
    raw_color = str(raw_color).strip()
    color_lower = raw_color.lower()
    
    if len(color_lower) == 6 and not color_lower.startswith('#'):
        color_lower = '#' + color_lower
        
    color_map = {
        "#ffffff": "화이트", "#fae7a3": "베이지", "#624138": "브라운",
        "#d9d9d9": "라이트 그레이", "#838f95": "다크 그레이", "#333131": "딥 그레이",
        "#9ed3ef": "라이트 블루", "#14529e": "로얄 블루", "#fce5cd": "라이트 핑크",
        "#e6b8af": "핑크", "#f6a147": "라이트 오렌지", "#ed6d03": "오렌지",
        "#660000": "와인", "#0c6b59": "그린", "#ffff00": "옐로우"
    }
    return color_map.get(color_lower, raw_color)

def load_data(sheet_name):
    raw_data = fetch_raw_data(sheet_name)
    df = pd.DataFrame(raw_data)
    
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name) if client else None
    
    if df.empty: return df, sheet
    
    df.columns = [str(c).strip() for c in df.columns]
    for col in ['날짜', '시작일', '종료일', '주문일시', '주문일']:
        if col in df.columns: df[col] = df[col].apply(clean_date_str)
    
    rename_map = {
        '주문일시': '날짜', '주문일': '날짜', '일자': '날짜',
        '금액': '결제금액', '총 주문금액': '결제금액', '예상견적': '결제금액',
        '성함': '구매자명', '고객명': '구매자명', '이름': '구매자명', '수취인명': '구매자명',
        '연락처': '연락처', '수취인연락처1': '연락처', '전화번호': '연락처',
        '주소': '주소', '배송지': '주소',
        '상품': '상품명', '품목': '상품명', '제품명': '상품명', 
        '디자인파일': '디자인파일', '첨부파일': '디자인파일', '시안': '디자인파일',
        '상태': '상태', '진행상태': '상태',
        '배송메세지': '요청사항', '비고': '요청사항', '메모': '요청사항',
        '포장옵션': '케이스', '컬러': '컬러', '옵션정보': '컬러',
        '택배비': '택배비', '희망수령일': '작업유형'
    }
    
    df.rename(columns=rename_map, inplace=True)
    df = df.loc[:, ~df.columns.duplicated()]
    
    if '컬러' in df.columns:
        df['컬러'] = df['컬러'].apply(get_color_name)

    required_cols = ['날짜', '구매자명', '연락처', '주소', '상품명', '수량', '결제금액', '요청사항', '디자인파일', '상태', '케이스', '작업유형', '택배비', '컬러']
    for col in required_cols:
        if col not in df.columns: df[col] = "" 
    
    if '주문처' not in df.columns: df['주문처'] = '🏠 자사몰'
    
    return df, sheet

def update_status_in_sheet(sheet, row_data, new_status="발주완료"):
    try:
        records = sheet.get_all_records()
        header = sheet.row_values(1)
        
        col_idx = -1
        status_names = ['상태', '진행상태', '배송상태', '주문상태']
        for i, h in enumerate(header):
            if any(name in h.strip() for name in status_names):
                col_idx = i + 1
                break
        if col_idx == -1: return False, "❌ 시트에서 '진행상태' 열을 찾을 수 없습니다."

        target_row_idx = -1
        t_name = str(row_data.get('구매자명', '')).strip()
        t_item = str(row_data.get('상품명', '')).strip()

        for idx, record in enumerate(records):
            r_name = str(record.get('구매자명') or record.get('성함') or record.get('이름') or record.get('수취인명') or '').strip()
            r_item = str(record.get('상품명') or record.get('상품') or record.get('제품명') or '').strip()
            if r_name == t_name and (t_item in r_item or r_item in t_item):
                target_row_idx = idx + 2
                break
                
        if target_row_idx != -1:
            sheet.update_cell(target_row_idx, col_idx, new_status)
            return True, f"✅ {target_row_idx}행 업데이트 성공"
        else:
            return False, f"❌ '{t_name}' 고객의 주문을 시트에서 매칭하지 못했습니다."
    except Exception as e: return False, f"❌ 시스템 오류: {str(e)}"

def get_drive_id(url):
    if not url or url == "-" or "이미지없음" in url: return None
    match = re.search(r"(?:id=|\/d\/)([\w-]{25,50})", str(url))
    return match.group(1) if match else None

def ask_ai(prompt):
    if not GOOGLE_API_KEY: return "API Key Missing"
    available_models = [] 
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                available_models.append(m.name)
        if not available_models: return "🚨 권한 에러: 구글 API 키로 사용할 수 있는 모델이 하나도 없습니다."
            
        target_model = available_models[-1]
        for m in available_models:
            if 'flash' in m.lower() or 'pro' in m.lower():
                target_model = m
                break
                
        clean_model_name = target_model.replace("models/", "")
        model = genai.GenerativeModel(clean_model_name)
        response = model.generate_content(prompt)
        return response.text
    except Exception as e: 
        return f"🚨 진짜 에러 원인: {str(e)}\n\n💡 [참고] 내 API로 사용 가능한 모델 목록: {available_models}"

def add_log(action_type, details):
    try:
        client = get_client()
        if client:
            log_sheet = client.open_by_key(SHEET_ID).worksheet("작업로그")
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_sheet.append_row([now_str, action_type, details])
    except Exception: pass

def send_email_with_attach(to, subject, body, attachment_file=None, filename="attachment.xlsx", multiple_attachments=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        if attachment_file:
            attachment_file.seek(0)
            part = MIMEApplication(attachment_file.read(), Name=filename)
            part['Content-Disposition'] = f'attachment; filename="{filename}"'
            msg.attach(part)
            
        if multiple_attachments:
            for att in multiple_attachments:
                att['file'].seek(0)
                part = MIMEApplication(att['file'].read(), Name=att['filename'])
                part['Content-Disposition'] = f'attachment; filename="{att["filename"]}"'
                msg.attach(part)
                
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(SENDER_EMAIL, SENDER_PASSWORD)
            s.send_message(msg)
        return True, "전송 성공"
    except Exception as e: return False, str(e)

def deduct_stock_smart(product_name, qty, df_opt, sheet_stock):
    try:
        if df_opt.empty or not sheet_stock: return False, "⚠️ 옵션 설정 또는 재고 시트 로드 실패"
        product_name = str(product_name).strip()
        target_std_name = None
        
        match_candidates = []
        for _, opt in df_opt.iterrows():
            std_name = str(opt.get('상품명', '')).strip()
            mapping_str = str(opt.get('매핑명', '')).strip()
            keywords = [k.strip() for k in mapping_str.split(',') if k.strip()]
            for kw in keywords:
                if kw in product_name:
                    match_candidates.append((len(kw), std_name))
        
        if match_candidates:
            match_candidates.sort(key=lambda x: x[0], reverse=True)
            target_std_name = match_candidates[0][1]
        else: target_std_name = product_name

        stock_records = sheet_stock.get_all_records()
        for idx, s_item in enumerate(stock_records):
            if str(s_item.get('상품명')).strip() == target_std_name:
                try: current_qty = int(pd.to_numeric(s_item.get('현재재고', 0), errors='coerce'))
                except: current_qty = 0
                new_qty = current_qty - int(qty)
                sheet_stock.update_cell(idx + 2, 2, new_qty)
                return True, f"✅ '{target_std_name}' {qty}개 차감 완료 (잔여: {new_qty})"
        return False, f"⚠️ '{target_std_name}' 상품을 재고 목록에서 찾을 수 없습니다."
    except Exception as e: return False, f"❌ 재고 차감 중 에러 발생: {str(e)}"

def add_stock_smart(product_name, qty, df_opt, sheet_stock):
    try:
        if df_opt.empty or not sheet_stock: return False, "⚠️ 옵션 설정 또는 재고 시트 로드 실패"
        product_name = str(product_name).strip()
        target_std_name = None
        
        match_candidates = []
        for _, opt in df_opt.iterrows():
            std_name = str(opt.get('상품명', '')).strip()
            mapping_str = str(opt.get('매핑명', '')).strip()
            keywords = [k.strip() for k in mapping_str.split(',') if k.strip()]
            for kw in keywords:
                if kw in product_name: match_candidates.append((len(kw), std_name))
        
        if match_candidates:
            match_candidates.sort(key=lambda x: x[0], reverse=True)
            target_std_name = match_candidates[0][1]
        else: target_std_name = product_name

        stock_records = sheet_stock.get_all_records()
        for idx, s_item in enumerate(stock_records):
            if str(s_item.get('상품명')).strip() == target_std_name:
                try: current_qty = int(pd.to_numeric(s_item.get('현재재고', 0), errors='coerce'))
                except: current_qty = 0
                new_qty = current_qty + int(qty)
                sheet_stock.update_cell(idx + 2, 2, new_qty)
                return True, f"✅ '{target_std_name}' {qty}개 복구 완료 (잔여: {new_qty})"
        return False, f"⚠️ '{target_std_name}' 상품을 재고 목록에서 찾을 수 없습니다."
    except Exception as e: return False, f"❌ 재고 복구 중 에러 발생: {str(e)}"

def check_stock_and_alert(df_stock):
    df_stock['현재재고'] = pd.to_numeric(df_stock['현재재고'], errors='coerce').fillna(0)
    df_stock['안전재고'] = pd.to_numeric(df_stock['안전재고'], errors='coerce').fillna(0)
    return df_stock[df_stock['현재재고'] <= df_stock['안전재고']]

def process_audio(uploaded_file):
    try:
        if not GOOGLE_API_KEY: return "API 키 없음"
        with open("temp_audio_file.mp3", "wb") as f:
            f.write(uploaded_file.getbuffer())
        myfile = genai.upload_file("temp_audio_file.mp3")
        model = genai.GenerativeModel("gemini-1.5-flash")
        result = model.generate_content(["이 음성 파일 내용을 요약하고, 일정(날짜,시간,내용)이 있다면 추출해줘.", myfile])
        return result.text
    except Exception as e: return f"오류: {str(e)}"

def update_tracking_in_sheet(sheet, row_data, tracking_num, new_status="배송중"):
    try:
        records = sheet.get_all_records()
        header = sheet.row_values(1)
        
        t_idx = -1
        s_idx = -1
        for i, h in enumerate(header):
            if '송장번호' in h.strip(): t_idx = i + 1
            if h.strip() in ['상태', '진행상태']: s_idx = i + 1
        if t_idx == -1: return False, "❌ 시트에 '송장번호' 열이 없습니다."

        target_row_idx = -1
        t_name = str(row_data.get('구매자명', '')).strip()
        t_item = str(row_data.get('상품명', '')).strip()

        for idx, record in enumerate(records):
            r_name = str(record.get('구매자명') or record.get('성함') or '').strip()
            r_item = str(record.get('상품명') or record.get('상품') or '').strip()
            if r_name == t_name and (t_item in r_item or r_item in t_item):
                target_row_idx = idx + 2
                break
                
        if target_row_idx != -1:
            sheet.update_cell(target_row_idx, t_idx, tracking_num)
            if s_idx != -1: sheet.update_cell(target_row_idx, s_idx, new_status)
            return True, "성공"
        return False, "미매칭"
    except Exception as e: return False, str(e)

def bulk_update_tracking_excel(sheet, df_up):
    try:
        all_data = sheet.get_all_values()
        if not all_data: return False, "❌ 시트에 데이터가 없습니다."
        header = all_data[0]
        
        try:
            name_idx = header.index('구매자명')
            status_idx = next(i for i, h in enumerate(header) if h in ['상태', '진행상태'])
            track_idx = header.index('송장번호')
        except ValueError: return False, "❌ 시트 1행(헤더)에 '구매자명', '상태', '송장번호' 컬럼 확인 필요."

        if '구매자명' not in df_up.columns or '송장번호' not in df_up.columns:
            return False, "❌ 업로드한 엑셀에 '구매자명'과 '송장번호' 열이 필수입니다."

        df_up = df_up.dropna(subset=['구매자명', '송장번호'])
        success_count = 0
        fail_count = 0

        for _, row in df_up.iterrows():
            u_name = str(row['구매자명']).strip()
            u_track = str(row['송장번호']).replace('.0', '').strip()
            if not u_name or not u_track or u_track.lower() == 'nan': continue

            matched = False
            for i in range(len(all_data)-1, 0, -1):
                s_row = all_data[i]
                if len(s_row) <= name_idx: continue 
                if str(s_row[name_idx]).strip() == u_name:
                    sheet.update_cell(i + 1, track_idx + 1, f"'{u_track}")
                    sheet.update_cell(i + 1, status_idx + 1, "배송중")
                    success_count += 1
                    matched = True
                    break
            if not matched: fail_count += 1
        
        result_msg = f"✅ 총 {success_count}건 배송 처리 완료!"
        if fail_count > 0: result_msg += f" (⚠️ {fail_count}건은 명단에 없어 실패했습니다.)"
        return True, result_msg
    except Exception as e: return False, f"❌ 시스템 오류 발생: {str(e)}"

# --------------------------------------------------------------------------
# 4. 사이드바 구성
# --------------------------------------------------------------------------
with st.sidebar:
    st.markdown("""
        <div style='text-align: center; padding: 20px 0;'>
            <h1 style='color: #2B3A55; margin: 0; font-weight: 900; font-size: 2rem;'>DUWELL</h1>
            <p style='color: #800020; font-size: 0.8rem; letter-spacing: 2px; margin: 5px 0 20px 0;'>SMART ERP SYSTEM</p>
        </div>
    """, unsafe_allow_html=True)

    menu = st.radio(
        "Menu Selection",
        ["통합 모니터링", "주문/생산 관리", "신제품 개발실", "마케팅 & CRM", "재고 관리", "옵션 관리", "일정 관리", "마진/정산 분석", "AI 비즈니스 센터"],
        label_visibility="collapsed"
    )

    st.write("---")
    st.caption(f"👤 접속자: {st.session_state.get('current_user', '담당자')}")

# --------------------------------------------------------------------------
# 5. 데이터 불러오기 (메인)
# --------------------------------------------------------------------------
df_duwell, sheet_main = load_data("시트1") 
df_all = df_duwell.copy()

if not df_all.empty and '날짜' in df_all.columns:
    df_all['날짜'] = pd.to_datetime(df_all['날짜'], errors='coerce')
    df_all = df_all.sort_values(by='날짜', ascending=False)
    df_all['날짜_str'] = df_all['날짜'].dt.strftime('%Y-%m-%d')
else:
    if not df_all.empty: df_all['날짜_str'] = ""

# ==========================================================================
# 🚀 6. 메뉴별 로직 시작 (완벽한 들여쓰기 정렬)
# ==========================================================================

# === 통합 모니터링 ===
if menu == "통합 모니터링":
    render_page_header("사업 현황 대시보드", "실시간 매출, 재고, 일정 데이터를 한눈에 파악하세요.")
    
    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")
    
    if not df_all.empty:
        df_all['기존금액'] = pd.to_numeric(df_all.get('결제금액', 0).astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
        df_all['수량_숫자'] = pd.to_numeric(df_all.get('수량', 1), errors='coerce').fillna(1)
        
        df_opt, _ = load_data("옵션관리")
        def get_item_price(item_name):
            if df_opt.empty or '가격' not in df_opt.columns: return 0
            item_clean = str(item_name).replace(" ", "").lower()
            for _, opt in df_opt.iterrows():
                std_name = str(opt.get('상품명', ''))
                mapping_str = str(opt.get('매핑명', ''))
                keywords = [k.strip() for k in mapping_str.split(',') if k.strip()] + [std_name]
                for kw in keywords:
                    if not kw: continue
                    if kw.replace(" ", "").lower() in item_clean:
                        return pd.to_numeric(str(opt.get('가격', 0)).replace(',', '').replace('원', ''), errors='coerce')
            return 0
            
        df_all['계산된단가'] = df_all['상품명'].apply(get_item_price)
        df_all['예상금액'] = df_all['계산된단가'] * df_all['수량_숫자']
        df_all['금액_숫자'] = df_all.apply(lambda x: x['기존금액'] if x['기존금액'] > 0 else x['예상금액'], axis=1)

        df_all['주'] = df_all['날짜'].dt.strftime('%Y년 %U주차')
        df_all['월'] = df_all['날짜'].dt.strftime('%Y-%m')

        today_orders = df_all[df_all['날짜_str'] == today_str]
        this_month_orders = df_all[df_all['월'] == today.strftime('%Y-%m')]
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.metric("오늘 주문", f"{len(today_orders)}건", delta=f"{len(today_orders)}건 New")
        with c2: st.metric("오늘 매출", f"{today_orders['금액_숫자'].sum():,.0f}원")
        with c3: st.metric("이달의 매출", f"{this_month_orders['금액_숫자'].sum():,.0f}원", delta=f"{len(this_month_orders)}건 주문")
        with c4: st.metric("누적 매출", f"{df_all['금액_숫자'].sum():,.0f}원")

        st.divider()

        st.markdown("#### AI 비즈니스 인사이트")
        with st.expander("AI에게 현재 매출/재고 상황 분석 맡기기", expanded=False):
            if st.button("맞춤형 AI 리포트 생성", key="ai_report_btn", type="primary"):
                with st.spinner("데이터를 분석하고 인사이트를 작성 중입니다..."):
                    top_items_dict = df_all['상품명'].value_counts().head(3).to_dict()
                    top_items_str = ", ".join([f"{k}({v}건)" for k, v in top_items_dict.items()])
                    
                    df_stock_temp, _ = load_data("재고관리")
                    low_stock_str = "없음 (모두 정상)"
                    if not df_stock_temp.empty:
                        df_stock_temp['현재재고'] = pd.to_numeric(df_stock_temp['현재재고'], errors='coerce').fillna(0)
                        df_stock_temp['안전재고'] = pd.to_numeric(df_stock_temp['안전재고'], errors='coerce').fillna(0)
                        low_stock_items = df_stock_temp[df_stock_temp['현재재고'] <= df_stock_temp['안전재고']]['상품명'].tolist()
                        if low_stock_items: low_stock_str = ", ".join(low_stock_items)

                    prompt = f"오늘 매출: {today_orders['금액_숫자'].sum():,.0f}원, 이달 매출: {this_month_orders['금액_숫자'].sum():,.0f}원, 베스트셀러: {top_items_str}, 재고부족: {low_stock_str}. 이를 바탕으로 현 상황 긍정 요약, 당면 과제, 매출 액션플랜 2가지를 짧고 명확하게 작성해줘."
                    st.success("AI 리포트 생성이 완료되었습니다.")
                    st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px; border:1px solid #E9ECEF;'>{ask_ai(prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

        st.divider()

        st.markdown("#### 기간별 매출 추이")
        tab_d, tab_w, tab_m = st.tabs(["일별 매출", "주간 매출", "월간 매출"])

        with tab_d:
            df_trend = df_all.groupby('날짜_str')['금액_숫자'].sum().reset_index().sort_values('날짜_str').tail(15)
            line_chart = alt.Chart(df_trend).mark_line(point=True, color='#800020').encode(
                x=alt.X('날짜_str:N', title='날짜', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('금액_숫자:Q', title='매출액'), tooltip=['날짜_str', '금액_숫자']
            ).properties(height=300)
            st.altair_chart(line_chart, use_container_width=True)

        with tab_w:
            df_week = df_all.groupby('주').agg(매출액=('금액_숫자', 'sum'), 주문건수=('구매자명', 'count')).reset_index().sort_values('주').tail(10)
            bar_week = alt.Chart(df_week).mark_bar(color='#30343B', opacity=0.9).encode(
                x=alt.X('주:N', title='주차', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('매출액:Q', title='매출액'), tooltip=['주', '매출액', '주문건수']
            ).properties(height=300)
            st.altair_chart(bar_week, use_container_width=True)

        with tab_m:
            df_month = df_all.groupby('월').agg(매출액=('금액_숫자', 'sum'), 주문건수=('구매자명', 'count')).reset_index().sort_values('월').tail(12)
            bar_month = alt.Chart(df_month).mark_bar(color='#2ca02c', opacity=0.8).encode(
                x=alt.X('월:N', title='월별', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('매출액:Q', title='매출액'), tooltip=['월', '매출액', '주문건수']
            ).properties(height=300)
            st.altair_chart(bar_month, use_container_width=True)

        st.divider()

        col_stock, col_sch = st.columns([1, 1])
        with col_stock:
            st.markdown("#### 재고 경고 (안전재고 미달)")
            df_stock_alert, _ = load_data("재고관리")
            if not df_stock_alert.empty:
                df_stock_alert['현재재고'] = pd.to_numeric(df_stock_alert['현재재고'], errors='coerce').fillna(0)
                df_stock_alert['안전재고'] = pd.to_numeric(df_stock_alert['안전재고'], errors='coerce').fillna(0)
                low_stock = df_stock_alert[df_stock_alert['현재재고'] <= df_stock_alert['안전재고']]
                if not low_stock.empty:
                    for _, s in low_stock.iterrows(): st.error(f"**{s['상품명']}**: 현재 {int(s['현재재고'])}개")
                else: st.success("모든 상품 재고 정상")
            else: st.write("재고 데이터 없음")

        with col_sch:
            st.markdown("#### 다가오는 일정")
            df_sch, _ = load_data("일정관리")
            t_today, t_week, t_month = st.tabs(["오늘", "이번 주", "이번 달"])
            if not df_sch.empty:
                df_sch['시작일_dt'] = pd.to_datetime(df_sch['시작일'], errors='coerce')
                with t_today:
                    today_sch = df_sch[df_sch['시작일'] == today_str]
                    if not today_sch.empty:
                        for _, r in today_sch.iterrows(): st.info(f"**{r.get('시간','')}** | {r.get('일정명','')}")
                    else: st.write("오늘 예정된 일정이 없습니다.")
                with t_week:
                    week_end = today + timedelta(days=7)
                    week_sch = df_sch[(df_sch['시작일_dt'] >= today) & (df_sch['시작일_dt'] <= week_end)].sort_values('시작일_dt')
                    if not week_sch.empty:
                        for _, r in week_sch.iterrows(): st.success(f"[{r['시작일']}] {r.get('시간','')} | {r.get('일정명','')}")
                    else: st.write("이번 주 예정된 일정이 없습니다.")
                with t_month:
                    this_month_str = today.strftime('%Y-%m')
                    month_sch = df_sch[df_sch['시작일'].astype(str).str.startswith(this_month_str, na=False)].sort_values('시작일_dt')
                    if not month_sch.empty:
                        for _, r in month_sch.iterrows(): st.warning(f"[{r['시작일']}] {r.get('일정명','')}")
                    else: st.write("이번 달 일정이 없습니다.")

        st.divider()
        st.markdown("#### 최근 주문 리포트 (최신 5건)")
        possible_cols = ['날짜_str', '구매자명', '상품명', '수량', '결제금액', '상태']
        cols = [c for c in possible_cols if c in df_all.columns]
        st.dataframe(df_all[cols].head(5), hide_index=True, use_container_width=True)

    else:
        st.warning("아직 등록된 주문 데이터가 없습니다.")

# === 주문/생산 관리 ===
elif menu == "주문/생산 관리":
    render_page_header("주문부터 생산까지 One-Stop", "주문 등록, 시안 확인, 공장 발주, 지시서 출력, 송장 입력을 순서대로 처리하세요.")
    
    op_tab1, op_tab2, op_tab3, op_tab4 = st.tabs([
        "1. 주문 등록", "2. 시안 & 장부", "3. 발주 및 지시서 생성", "4. 송장 등록"
    ])

    with op_tab1:
        st.markdown("#### 신규 주문 등록 (재고 자동 차감)")
        sub1, sub2 = st.tabs(["엑셀 일괄 업로드", "건별 수동 등록"])
        
        with sub1:
            uploaded_file = st.file_uploader("네이버/자사몰 엑셀 업로드", type=['xlsx'], key="order_up")
            if uploaded_file and st.button("저장 및 재고 차감", type="primary"):
                try:
                    df_new = pd.read_excel(uploaded_file, header=1)
                    df_opt, _ = load_data("옵션관리")
                    _, sheet_stock = load_data("재고관리")
                    df_new = df_new.dropna(subset=['수취인명', '상품명'])
                    rows_add = []
                    
                    for _, row in df_new.iterrows():
                        p_name = str(row.get('상품명','')).strip()
                        try: qty = int(pd.to_numeric(row.get('수량', 1), errors='coerce'))
                        except: qty = 1
                        try: price = int(pd.to_numeric(str(row.get('총 주문금액', '0')).replace(',', '').replace('원', '').strip(), errors='coerce'))
                        except: price = 0
                        try: ship_fee = int(pd.to_numeric(str(row.get('배송비', '0')).replace(',', '').replace('원', '').strip(), errors='coerce'))
                        except: ship_fee = 0

                        final_color = get_color_name(str(row.get('컬러', row.get('옵션정보', ''))))

                        rows_add.append([
                            str(row.get('주문일시','')), str(row.get('수취인명','')), str(row.get('수취인연락처1','')), str(row.get('배송지','')), 
                            p_name, final_color, 
                            "일반(무지)", 
                            "", str(qty), str(price), 
                            "기본(폴리백)", 
                            "신규", "엑셀일괄", str(row.get('배송메세지','')), "", str(ship_fee)
                        ])
                        
                    if sheet_main:
                        sheet_main.append_rows(rows_add, value_input_option='USER_ENTERED', table_range='A1')
                        st.success(f"{len(rows_add)}건 처리 완료"); st.cache_data.clear(); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"오류: {e}")

        with sub2:
            with st.form("manual"):
                st.markdown("##### 📝 수동 주문 입력")
                col1, col2 = st.columns(2)
                with col1:
                    m_date = st.date_input("날짜", datetime.now())
                    m_market = st.selectbox("판매 채널", ["개인판매(수수료 0%)", "스마트스토어", "쿠팡", "자사몰", "기타"])
                    m_name = st.text_input("구매자명")
                    m_phone = st.text_input("연락처")
                    m_addr = st.text_input("주소")
                    m_work = st.selectbox("작업 유형", ["일반(무지)", "자수", "전사", "나염", "기타"]) 
                    
                with col2:
                    m_prod = st.text_input("상품명 (옵션매핑명)")
                    m_color = st.text_input("컬러 (예: 웜그레이, 또는 #333131)") 
                    m_qty = st.number_input("수량", min_value=1, value=1)
                    m_price = st.number_input("결제금액 (제품가만)", value=0, step=1000)
                    m_shipping = st.number_input("택배비 (무배면 0, 유배면 3000 등)", value=0, step=500)
                    m_case = st.selectbox("포장(케이스)", ["기본(폴리백)", "1매입 박스", "2매입 박스", "3매입 박스", "띠지 포장", "기타"])
                
                m_file = st.text_input("디자인링크 (자수/전사 도안이 있는 경우)")
                m_req = st.text_area("요청사항 (자수 문구 등 상세기재)")
                
                if st.form_submit_button("등록 및 재고차감", type="primary"):
                    if sheet_main:
                        final_m_color = get_color_name(m_color)
                        new_row_data = [
                            str(m_date), m_name, m_phone, m_addr, m_prod, final_m_color, 
                            m_work, 
                            m_file, str(m_qty), str(m_price), 
                            m_case, 
                            "신규", m_market, m_req, "", str(m_shipping)
                        ]
                        sheet_main.insert_row(new_row_data, index=2, value_input_option='USER_ENTERED')
                        df_opt, _ = load_data("옵션관리")
                        _, sheet_stock = load_data("재고관리")
                        ok, msg = deduct_stock_smart(m_prod, m_qty, df_opt, sheet_stock)
                        st.success(msg); st.cache_data.clear(); time.sleep(1); st.rerun()

    with op_tab2:
        st.markdown("#### 디자인 시안 확인 및 전체 장부")
        c_tab1, c_tab2 = st.tabs([" 디자인 작업 대기중 (시안실)", "전체 주문 장부"])
        with c_tab1:
            if df_all.empty: st.warning("데이터가 없습니다.")
            else:
                df_wait = df_all[df_all['상태'] != '완료']
                for i, r in df_wait.iterrows():
                    with st.expander(f"📌 {r.get('구매자명')} - {r.get('상품명')}"):
                        wc1, wc2 = st.columns([1, 2])
                        with wc1:
                            link = str(r.get('디자인파일', ''))
                            drive_id = get_drive_id(link)
                            if drive_id: st.image(f"https://drive.google.com/thumbnail?id={drive_id}&sz=w400")
                            else: st.text("이미지 없음")
                        with wc2:
                            st.write(f"요청사항: {r.get('요청사항', '-')}")
                            if st.button("시안 확정 (완료 처리)", key=f"btn_sian_{i}"):
                                success, msg = update_status_in_sheet(sheet_main, r, "완료")
                                if success: st.success(msg); st.cache_data.clear(); time.sleep(1); st.rerun()
        with c_tab2:
            st.markdown("#### 전체 주문 장부 및 취소/반품 처리")
            st.dataframe(df_all, use_container_width=True)
            csv = df_all.to_csv(index=False).encode('utf-8-sig')
            st.download_button("장부 엑셀 다운로드", csv, "order_list.csv", "text/csv")
            st.divider()
            st.markdown("##### 주문 취소 및 반품 (재고 자동 복구)")
            with st.form("cancel_return_form"):
                cr_col1, cr_col2 = st.columns(2)
                with cr_col1: search_buyer = st.text_input("취소/반품 처리할 구매자명 입력")
                with cr_col2: cr_status = st.selectbox("변경할 상태", ["취소", "반품", "교환"])
                if st.form_submit_button("상태 변경 및 재고 복구", type="primary"):
                    if not search_buyer: st.warning("구매자명을 입력해주세요.")
                    else:
                        target_orders = df_all[df_all['구매자명'].astype(str) == search_buyer]
                        if target_orders.empty: st.error("해당 구매자의 주문을 찾을 수 없습니다.")
                        else:
                            target_row = target_orders.iloc[0] 
                            success, msg = update_status_in_sheet(sheet_main, target_row, cr_status)
                            if success:
                                if cr_status in ["취소", "반품"]:
                                    df_opt, _ = load_data("옵션관리")
                                    _, sheet_stock = load_data("재고관리")
                                    try: qty_to_restore = int(pd.to_numeric(target_row.get('수량', 1), errors='coerce'))
                                    except: qty_to_restore = 1
                                    ok, stock_msg = add_stock_smart(target_row.get('상품명', ''), qty_to_restore, df_opt, sheet_stock)
                                    st.success(f"{msg} / {stock_msg}")
                                else: st.success(f"{msg} (교환은 재고가 복구되지 않습니다)")
                                st.cache_data.clear(); time.sleep(2); st.rerun()
                            else: st.error(msg)

    with op_tab3:
        st.markdown("#### 통합 발주 및 지시서 생성")
        if not df_all.empty:
            pending_orders = df_all[~df_all['상태'].isin(['발주완료', '배송중', '배송완료', '취소', '반품'])].copy()
            if pending_orders.empty: 
                st.success("현재 새로 공장에 넘길 발주 대기 건이 없습니다.")
            else:
                if "선택" not in pending_orders.columns: pending_orders.insert(0, "선택", False)
                if '작업유형' not in pending_orders.columns: pending_orders['작업유형'] = '일반(무지)'
                if '케이스' not in pending_orders.columns: pending_orders['케이스'] = '기본(폴리백)'

                st.markdown("##### 1️⃣ 발주 대상 선택")
                
                btn_col1, btn_col2, btn_col3 = st.columns([1.5, 1.5, 7])
                with btn_col1:
                    if st.button("✅ 전체 선택", use_container_width=True):
                        st.session_state['select_all_toggle'] = True
                        if 'order_editor' in st.session_state: del st.session_state['order_editor']
                        st.rerun()
                with btn_col2:
                    if st.button("🔲 전체 해제", use_container_width=True):
                        st.session_state['select_all_toggle'] = False
                        if 'order_editor' in st.session_state: del st.session_state['order_editor']
                        st.rerun()

                if 'select_all_toggle' in st.session_state:
                    pending_orders['선택'] = st.session_state['select_all_toggle']
                    del st.session_state['select_all_toggle']

                edited_orders = st.data_editor(
                    pending_orders, column_config={"선택": st.column_config.CheckboxColumn(required=True)},
                    column_order=['선택', '날짜_str', '구매자명', '상품명', '컬러', '작업유형', '케이스', '수량', '요청사항'], 
                    hide_index=True, use_container_width=True, key='order_editor'
                )
                selected_data = edited_orders[edited_orders['선택'] == True]
                
                if not selected_data.empty:
                    st.divider()
                    if st.button("발주 확정 및 파일 생성", type="primary", use_container_width=True):
                        for _, row in selected_data.iterrows(): update_status_in_sheet(sheet_main, row, "발주완료")
                        
                        excel_out = io.BytesIO()
                        with pd.ExcelWriter(excel_out, engine='xlsxwriter') as writer:
                            selected_data[['구매자명', '연락처', '주소', '상품명', '컬러', '작업유형', '케이스', '수량', '요청사항']].to_excel(writer, index=False)
                        
                        all_html = "<html><body style='font-family:sans-serif;'>"
                        for _, row in selected_data.iterrows():
                            b_name = str(row['구매자명']).strip()
                            color_val = str(row.get('컬러', '-')).strip()
                            work_val = str(row.get('작업유형', '-')).strip()
                            case_val = str(row.get('케이스', '-')).strip()
                            
                            drive_id = get_drive_id(str(row.get('디자인파일', '')))
                            img_tag = f"<img src='https://drive.google.com/thumbnail?id={drive_id}&sz=w500' style='max-width:100%;'>" if drive_id else "<p>[이미지 없음]</p>"
                            
                            all_html += f"<div style='page-break-after:always; border:2px solid #2B3A55; padding:20px; margin-bottom:40px; border-radius:10px;'><h2 style='color:#2B3A55;'>제작 지시서 ({b_name}님)</h2><table border='1' style='width:100%; border-collapse:collapse; margin-bottom:20px; text-align:center;'><tr style='background:#f4f4f4;'><th>상품명</th><th>컬러</th><th>작업유형</th><th>포장(케이스)</th><th>수량</th><th>요청사항(자수)</th></tr><tr><td>{row['상품명']}</td><td><b>{color_val}</b></td><td style='color:#D32F2F;'><b>{work_val}</b></td><td><b>{case_val}</b></td><td>{row['수량']}</td><td>{row['요청사항']}</td></tr></table><div style='text-align:center;'><p><b>[확정 디자인 시안]</b></p>{img_tag}</div></div>"
                        all_html += "</body></html>"
                        
                        st.session_state['ready_excel'] = excel_out.getvalue()
                        st.session_state['ready_html'] = all_html.encode('utf-8-sig')
                        st.success("발주 처리가 완료되었습니다.")
                        st.cache_data.clear()
                    
                    if 'ready_excel' in st.session_state:
                        c1, c2 = st.columns(2)
                        with c1: st.download_button("발주 리스트 다운로드 (Excel)", st.session_state['ready_excel'], f"발주리스트_{datetime.now().strftime('%m%d')}.xlsx")
                        with c2: st.download_button("상세 작업지시서 다운로드 (HTML)", st.session_state['ready_html'], f"작업지시서_{datetime.now().strftime('%m%d')}.html", "text/html")

    with op_tab4:
        st.markdown("#### 배송 정보(송장) 업데이트")
        if not df_all.empty:
            completed_orders = df_all[df_all['상태'] == '발주완료'].copy()
            sc1, sc2 = st.columns(2)
            with sc1:
                st.markdown("##### 수동 입력 (건별)")
                if completed_orders.empty: st.write("발주 완료된 대기 내역이 없습니다.")
                else:
                    if '송장번호' not in completed_orders.columns: completed_orders['송장번호'] = ""
                    tr_edited = st.data_editor(completed_orders, column_order=['구매자명', '상품명', '송장번호'], hide_index=True)
                    if st.button("수동 송장 저장"):
                        cnt = 0
                        for _, row in tr_edited.iterrows():
                            t_num = str(row.get('송장번호', '')).strip()
                            if t_num and t_num != "nan":
                                ok, _ = update_tracking_in_sheet(sheet_main, row, t_num)
                                if ok: cnt += 1
                        st.success(f"{cnt}건 저장 완료!"); st.cache_data.clear(); time.sleep(1); st.rerun()
            with sc2:
                st.markdown("##### 엑셀 일괄 등록")
                up_f = st.file_uploader("공장 송장 엑셀 파일 (.xlsx)", type=['xlsx'])
                if up_f and st.button("엑셀 데이터 시트 반영"):
                    df_up = pd.read_excel(up_f)
                    ok, msg = bulk_update_tracking_excel(sheet_main, df_up)
                    if ok: st.success(msg); st.cache_data.clear(); time.sleep(1); st.rerun()
                    else: st.error(msg)

# === 신제품 개발실 ===
elif menu == "신제품 개발실":
    render_page_header("신제품 샘플 개발실", "생산 지시서(Tech Pack) 작성부터 샘플 검수(Check-list)까지 원스톱 관리")
    tab_techpack, tab_checklist = st.tabs(["1. 생산 지시서 (Tech Pack)", "2. 샘플 검수 (Check-list)"])

    with tab_techpack:
        with st.form("new_product_dev_form"):
            st.markdown("#### 신제품 생산 스펙 작성")
            c1, c2 = st.columns([1, 1])
            with c1: 
                dev_factory = st.text_input("공장명", placeholder="예: A방직")
                dev_prod_name = st.text_input("상품명 (ITEM)", placeholder="예: 프리미엄 와플 타월")
            with c2: 
                dev_color_qty = st.text_area("초도 발주 수량 (컬러 / 수량)", placeholder="예시) 기호(/)로 구분하여 작성\n웜그레이 / 500장\n네이비 / 500장\n차콜 / 300장", height=110)

            st.markdown("##### 세부 사양 (SPEC)")
            s1, s2, s3, s4 = st.columns(4)
            with s1: dev_size = st.text_input("사이즈", placeholder="예: 40 x 80 cm")
            with s2: dev_weight = st.text_input("중량", placeholder="예: 200g")
            with s3: dev_yarn = st.text_input("사종 (소재)", placeholder="예: 최고급 코마사 40수")
            with s4: dev_dyeing = st.radio("염색 방식", ["선염", "후염", "해당없음"], horizontal=True)

            p1, p2, p3 = st.columns(3)
            with p1: dev_border = st.text_input("보더 디자인", placeholder="예: 양끝 3선 피카소 보더")
            with p2: dev_pkg = st.text_input("포장 방법", placeholder="예: 개별 띠지 + OPP 폴리백")
            with p3: dev_label_pos = st.text_input("라벨/택 위치", placeholder="예: 우측 하단 1cm 띄우고 봉제")

            st.markdown("##### 디자인 상세 및 참고 이미지")
            i1, i2 = st.columns(2)
            with i1:
                dev_design_detail = st.text_area("디자인 (선염 및 보더 등 특이사항)", placeholder="예: 선염 3컬러 교차 배열, 보더 부분 자가드 포인트", height=120)
                dev_extra = st.text_area("작업 시 주의사항 (*중요*)", placeholder="- 봉사 간격 잘게 치기\n- 잔실 없도록 깔끔하게 처리\n- 세탁 라벨은 뒷면에 겹쳐서 봉제", height=120)
            with i2:
                dev_ref_img = st.file_uploader("참고 이미지 (디자인 시안/도식화)", type=['png', 'jpg', 'jpeg'])
                dev_label_img = st.file_uploader("라벨/택 이미지", type=['png', 'jpg', 'jpeg'])

            st.divider()
            submit_dev = st.form_submit_button("작업지시서 문서 생성 및 구글 시트 저장", type="primary")

        if submit_dev:
            if not dev_prod_name or not dev_factory:
                st.warning("공장명과 상품명은 필수 입력 항목입니다.")
            else:
                with st.spinner("지시서를 생성 중입니다... 잠시만 기다려주세요."):
                    import base64
                    import io
                    today_str = datetime.now().strftime("%Y-%m-%d")
                    color_qty_html = ""
                    lines = [line.strip() for line in dev_color_qty.split('\n') if line.strip()]
                    if not lines:
                        color_qty_html = "<tr><td style='border-left:none; border-bottom:none; border-right:2px solid #222; padding:4px;'>-</td><td style='border-right:none; border-bottom:none; padding:4px;'>-</td></tr>"
                    else:
                        for i, line in enumerate(lines):
                            if '/' in line: parts = line.split('/')
                            elif '-' in line: parts = line.split('-')
                            else: parts = [line, ""]
                            color = parts[0].strip()
                            qty = parts[1].strip() if len(parts) > 1 else ""
                            b_bottom = "border-bottom:none;" if i == len(lines) - 1 else "border-bottom:2px solid #222;"
                            color_qty_html += f"<tr><td style='border-left:none; {b_bottom} border-right:2px solid #222; font-weight:bold; padding:4px;'>{color}</td><td style='border-right:none; {b_bottom} padding:4px;'>{qty}</td></tr>"

                    try:
                        client = get_client()
                        if client:
                            sheet_dev = client.open_by_key(SHEET_ID).worksheet("신제품개발")
                            flat_color_qty = dev_color_qty.replace('\n', ', ')
                            row_data = [
                                today_str, dev_factory, dev_prod_name, flat_color_qty,
                                dev_size, dev_weight, dev_yarn, dev_dyeing, dev_border, 
                                dev_design_detail, dev_extra, dev_pkg, dev_label_pos
                            ]
                            sheet_dev.append_row(row_data)
                    except Exception as e:
                        st.error(f"⚠️ 구글 시트 저장 실패: {e}")

                    def get_image_base64(uploaded_file):
                        if uploaded_file is not None:
                            bytes_data = uploaded_file.getvalue()
                            b64 = base64.b64encode(bytes_data).decode()
                            return f"data:{uploaded_file.type};base64,{b64}"
                        return ""
                    
                    b64_ref = get_image_base64(dev_ref_img)
                    b64_label = get_image_base64(dev_label_img)
                    ref_html = f"<img src='{b64_ref}' style='width:100%; max-height:420px; object-fit:contain;'>" if b64_ref else "<span style='color:#999;'>이미지 없음</span>"
                    label_html = f"<img src='{b64_label}' style='width:100%; max-height:150px; object-fit:contain;'>" if b64_label else "<span style='color:#999;'>라벨 없음</span>"
                    extra_html = dev_extra.replace('\n', '<br>')
                    design_html = dev_design_detail.replace('\n', '<br>')

                    html_content = f"""
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <style>
                            @page {{ margin: 5mm; }}
                            * {{ box-sizing: border-box; }}
                            body {{ font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; padding: 0; margin: 0; color: #111; font-size: 18pt; }}
                            table {{ width: 100%; border-collapse: collapse; margin-bottom: 8px; table-layout: fixed; }}
                            th, td {{ border: 2px solid #222; padding: 6px 8px; text-align: center; vertical-align: middle; font-size: 18pt; line-height: 1.4; word-break: keep-all; }}
                            th {{ background-color: #F1F5F9; font-weight: bold; color: #111; }}
                            .left-align {{ text-align: left; vertical-align: top; padding: 10px; }}
                        </style>
                    </head>
                    <body>
                        <div style="text-align: center; margin-bottom: 15px;">
                            <div style="font-size: 45pt; font-weight: 900; letter-spacing: 25px; margin-bottom: 15px; padding-left: 25px;">작업지시서</div>
                            <table style="width: 100%; border: none; margin: 0; border-top: 3px solid #111; border-bottom: 3px solid #111;">
                                <tr>
                                    <td style="border: none; padding: 10px 5px; text-align: center; font-size: 15pt;"><strong>작성일자 :</strong> {today_str}</td>
                                    <td style="border: none; padding: 10px 5px; text-align: center; font-size: 15pt;"><strong>발주처 :</strong> <span style="font-weight:900; color:#2B3A55;">DUWELL</span></td>
                                    <td style="border: none; padding: 10px 5px; text-align: center; font-size: 15pt;"><strong>업체(공장) :</strong> {dev_factory}</td>
                                    <td style="border: none; padding: 10px 5px; text-align: center; font-size: 18pt; font-weight: 900;"><strong>ITEM :</strong> {dev_prod_name}</td>
                                </tr>
                            </table>
                        </div>
                        <table>
                            <colgroup>
                                <col style="width: 45%;"> <col style="width: 20%;"> <col style="width: 35%;"> </colgroup>
                            <tr>
                                <th style="font-size:20pt; letter-spacing:2px;">DESIGN & LABEL / 디자인 및 라벨</th>
                                <th colspan="2" style="font-size:20pt; letter-spacing:2px;">PRODUCTION SPECS / 생산 사양</th>
                            </tr>
                            <tr>
                                <td rowspan="7" style="text-align:center; vertical-align:middle; padding:5px;">
                                    {ref_html}
                                </td>
                                <th>염색방식</th>
                                <td style="font-weight:bold; color:#2B3A55;">{dev_dyeing}</td>
                            </tr>
                            <tr><th>사이즈</th><td>{dev_size}</td></tr>
                            <tr><th>중량</th><td style="color:#D32F2F; font-weight:bold;">{dev_weight}</td></tr>
                            <tr><th>소재(사종)</th><td>{dev_yarn}</td></tr>
                            <tr><th>보더디자인</th><td>{dev_border}</td></tr>
                            <tr><th>포장방법(PKG)</th><td>{dev_pkg}</td></tr>
                            <tr>
                                <th>초도발주수량</th>
                                <td style="padding: 0; vertical-align: top;">
                                    <table style="width:100%; height:100%; margin:0; border-collapse:collapse; border-style:hidden;">
                                        <tr>
                                            <th style="width:50%; border-top:none; border-left:none; border-bottom:2px solid #222; border-right:2px solid #222; background-color:#F8F9FA;">컬러</th>
                                            <th style="width:50%; border-top:none; border-right:none; border-bottom:2px solid #222; background-color:#F8F9FA;">수량</th>
                                        </tr>
                                        {color_qty_html}
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td rowspan="4" style="text-align:center; vertical-align:middle; padding:10px;">
                                    <div style="font-size:14pt; margin-bottom:8px;"><strong>[라벨 부착 위치]</strong><br>{dev_label_pos}</div>
                                    {label_html}
                                </td>
                                <th colspan="2">* 디자인 상세 (선염/보더 등) *</th>
                            </tr>
                            <tr><td colspan="2" class="left-align" style="height: 100px;">{design_html}</td></tr>
                            <tr><th colspan="2">* 작업 시 주의사항 *</th></tr>
                            <tr><td colspan="2" class="left-align" style="height: 100px;">{extra_html}</td></tr>
                        </table>
                    </body>
                    </html>
                    """
                    st.session_state['dev_html_content'] = html_content
                    st.session_state['dev_html_name'] = f"작업지시서_{dev_prod_name}.html"
                    st.success("✅ 지시서 생성이 완료되었습니다! 아래에서 다운로드 버튼을 눌러주세요.")
                    time.sleep(1) 
                    st.rerun()

        if 'dev_html_content' in st.session_state:
            st.divider()
            st.markdown("##### 🖨️ 생성된 문서 다운로드")
            st.info("💡 우측의 HTML 문서를 다운로드하여 브라우저에서 열고 'Ctrl+P(인쇄) -> PDF로 저장'을 이용해 주세요.")
            st.download_button(
                label="✅ 작업지시서 다운로드 (HTML)", 
                data=st.session_state['dev_html_content'].encode('utf-8-sig'), 
                file_name=st.session_state['dev_html_name'], 
                mime="text/html", 
                use_container_width=True
            )

    with tab_checklist:
        st.markdown("#### 🔎 입고 샘플 품질 검수 리스트")
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            chk_prod_name = st.text_input("제품명 (Product Name)", value="와플 타월")
            chk_reviewer = st.text_input("검수자", value=st.session_state.get('current_user', '담당자'))
        with cc2:
            chk_color_size = st.text_input("컬러 및 옵션 (Color/Size)", value="진행전 / 40*90")
            chk_round = st.selectbox("차수", ["1차", "2차", "3차", "최종(Pre-Pro)"])
        with cc3:
            chk_date = st.date_input("작성일", datetime.now())
            chk_result = st.selectbox("최종 판정 (Result)", ["대기중", "합격 (PASS)", "수정 후 재샘플 (RE-WORK)", "드랍 (DROP)"])

        st.divider()
        st.markdown("##### 세부 검수 항목 (Check-list)")
        
        default_chk_data = [
            {"항목": "사이즈", "세부 검수 기준": "작업 지시서 상의 규격과 일치하는가?", "판정": "대기", "비고": ""},
            {"항목": "중량", "세부 검수 기준": "요구한 평량(g) 기준 오차 범위 내에 있는가?", "판정": "대기", "비고": ""},
            {"항목": "제품디자인", "세부 검수 기준": "와플 조직감, 패턴, 직조 형태가 기획안과 동일한가?", "판정": "대기", "비고": ""},
            {"항목": "색상", "세부 검수 기준": "지정된 컬러 발색이 정확하고 탕차이(이색)가 없는가?", "판정": "진행전", "비고": ""},
            {"항목": "봉제", "세부 검수 기준": "테두리 마감(헤밍 등) 일정하고 바른가?", "판정": "대기", "비고": ""},
            {"항목": "올풀림", "세부 검수 기준": "테두리 마감 부위 및 표면에 올이 나간 곳이 없는가?", "판정": "대기", "비고": ""},
            {"항목": "보풀", "세부 검수 기준": "표면 잔털이 심하거나 마찰 시 보풀이 생기지 않는가?", "판정": "대기", "비고": ""},
            {"항목": "라벨", "세부 검수 기준": "케어/브랜드 라벨의 부착 위치와 봉제가 반듯한가?", "판정": "진행전", "비고": ""},
            {"항목": "기타", "세부 검수 기준": "오염, 잡사, 포장 상태 등에 문제가 없는가?", "판정": "진행전", "비고": ""}
        ]
        df_chk = pd.DataFrame(default_chk_data)
        edited_chk = st.data_editor(
            df_chk, 
            column_config={
                "판정": st.column_config.SelectboxColumn("판정 (O/X)", options=["O", "X", "진행전", "대기"], required=True),
                "비고": st.column_config.TextColumn("비고 (불량 사유 등)")
            },
            hide_index=True, use_container_width=True
        )

        st.divider()
        st.markdown("##### 공장 커뮤니케이션 요약")
        com1, com2 = st.columns(2)
        with com1: chk_inquiry = st.text_area("본사 문의 사항 및 개선 요청", placeholder="예: 1. 세탁 후 올나감 현상 원인 파악\n2. 테두리 단봉 -> 삼봉 변경 가능 여부", height=100)
        with com2: chk_feedback = st.text_area("공장 피드백 (답변)", placeholder="예: 삼봉 변경 시 단가 건당 50원 인상됨.", height=100)

        st.markdown("##### 📸 문제점 참고 이미지 (비교용)")
        img1, img2 = st.columns(2)
        with img1: chk_img1 = st.file_uploader("참고 이미지 1 (불량 부위 / 변경 전)", type=['png', 'jpg', 'jpeg'], key="chk1")
        with img2: chk_img2 = st.file_uploader("참고 이미지 2 (정상 부위 / 변경 후)", type=['png', 'jpg', 'jpeg'], key="chk2")

        submit_check = st.button("검수 리스트 문서 생성 및 구글 시트 저장", type="primary", use_container_width=True)

        if submit_check:
            with st.spinner("검수 리스트를 처리 중입니다..."):
                import base64
                import io
                try:
                    client = get_client()
                    if client:
                        sheet_chk = client.open_by_key(SHEET_ID).worksheet("샘플검수")
                        row_data = [
                            str(chk_date), chk_prod_name, chk_round, chk_color_size, 
                            chk_reviewer, chk_result, chk_inquiry.replace('\n', ' / '), chk_feedback.replace('\n', ' / ')
                        ]
                        sheet_chk.append_row(row_data)
                except Exception: pass

                chk_tbody = ""
                for idx, row in edited_chk.iterrows():
                    res_color = "#D32F2F" if row['판정'] == 'X' else ("#1976D2" if row['판정'] == 'O' else "#555")
                    chk_tbody += f"<tr><td style='text-align:center;'>{idx+1}</td><td style='text-align:center; font-weight:bold;'>{row['항목']}</td><td style='text-align:left;'>{row['세부 검수 기준']}</td><td style='text-align:center; color:{res_color}; font-weight:bold;'>{row['판정']}</td><td style='text-align:left;'>{row['비고']}</td></tr>"

                def get_img_b64(f):
                    if f: return f"data:{f.type};base64,{base64.b64encode(f.getvalue()).decode()}"
                    return ""
                
                b64_img1 = get_img_b64(chk_img1)
                b64_img2 = get_img_b64(chk_img2)

                img1_html = f"<img src='{b64_img1}' style='width:100%; max-height:280px; object-fit:contain;'>" if b64_img1 else "이미지 없음"
                img2_html = f"<img src='{b64_img2}' style='width:100%; max-height:280px; object-fit:contain;'>" if b64_img2 else "이미지 없음"

                chk_html_content = f"""
                <html>
                <head>
                    <meta charset="utf-8">
                    <style>
                        @page {{ margin: 5mm; }}
                        body {{ font-family: 'Malgun Gothic', sans-serif; padding: 0; margin: 0; color: #111; font-size: 18pt; }}
                        .title {{ text-align: center; font-size: 38pt; font-weight: 900; letter-spacing: 5px; margin-bottom: 10px; }}
                        table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; table-layout: fixed; }}
                        th, td {{ border: 2px solid #222; padding: 6px 8px; vertical-align: middle; font-size: 18pt; line-height: 1.4; word-break: keep-all; }}
                        th {{ background-color: #F1F5F9; font-weight: bold; text-align: center; color: #111; }}
                    </style>
                </head>
                <body>
                    <div class="title">샘플 체크 리스트 ({chk_round})</div>
                    <table>
                        <tr>
                            <th style="width:15%;">제품명</th><td style="width:35%; font-weight:bold; font-size:20pt;">{chk_prod_name}</td>
                            <th style="width:15%;">작성일자</th><td style="width:35%;">{str(chk_date)}</td>
                        </tr>
                        <tr>
                            <th>컬러 및 옵션</th><td style="font-size:20pt;">{chk_color_size}</td>
                            <th>검수자</th><td>{chk_reviewer}</td>
                        </tr>
                        <tr>
                            <th>최종 판정</th><td colspan="3" style="font-weight:900; font-size:24pt; color:#EE0979; text-align:center;">{chk_result}</td>
                        </tr>
                    </table>
                    <table>
                        <tr><th style="width:5%;">No.</th><th style="width:15%;">검수 항목</th><th style="width:40%;">세부 검수 기준</th><th style="width:10%;">판정</th><th style="width:30%;">비고</th></tr>
                        {chk_tbody}
                    </table>
                    <table>
                        <tr><th style="width:50%;">본사 문의 및 개선 요청 사항</th><th style="width:50%;">공장 피드백</th></tr>
                        <tr><td style="height:80px; vertical-align:top; padding:10px;">{chk_inquiry.replace(chr(10), '<br>')}</td><td style="height:80px; vertical-align:top; padding:10px;">{chk_feedback.replace(chr(10), '<br>')}</td></tr>
                    </table>
                    <table>
                        <tr><th style="width:50%;">참고 이미지 1</th><th style="width:50%;">참고 이미지 2</th></tr>
                        <tr><td style="height:300px; text-align:center; vertical-align:middle; padding:5px;">{img1_html}</td><td style="height:300px; text-align:center; vertical-align:middle; padding:5px;">{img2_html}</td></tr>
                    </table>
                </body>
                </html>
                """
                st.session_state['chk_html_content'] = chk_html_content
                st.session_state['chk_html_name'] = f"샘플검수서_{chk_prod_name}_{chk_round}.html"
                st.success("✅ 검수 리스트 생성이 완료되었습니다! 아래에서 다운로드 버튼을 눌러주세요.")
                time.sleep(1); st.rerun()

        if 'chk_html_content' in st.session_state:
            st.divider()
            st.markdown("##### 🖨️ 생성된 문서 다운로드")
            st.info("💡 우측의 HTML 문서를 다운로드하여 브라우저에서 열고 'Ctrl+P(인쇄) -> PDF로 저장'을 이용해 주세요.")
            st.download_button(
                label="✅ 검수 리스트 다운로드 (HTML)", 
                data=st.session_state['chk_html_content'].encode('utf-8-sig'), 
                file_name=st.session_state['chk_html_name'], 
                mime="text/html", 
                use_container_width=True
            )

# === 마진/정산 분석 ===
elif menu == "마진/정산 분석":
    render_page_header("마진/정산 분석", "실시간 마진 및 정산 분석기")
    with st.expander("정산 기준 설정 (수수료율 입력)", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fee_smart = st.number_input("스마트스토어 수수료 (%)", value=5.5)
            fee_own = st.number_input("자사몰(PG) 수수료 (%)", value=3.0)
        with col_f2:
            fee_coupang = st.number_input("쿠팡 수수료 (%)", value=10.8)
            fee_personal = st.number_input("개인판매/B2B 수수료 (%)", value=0.0) 
        with col_f3:
            fee_etc = st.number_input("기타 마켓 수수료 (%)", value=5.0)

    if not df_all.empty:
        df_cost, _ = load_data("옵션관리") 
        df_calc = df_all[~df_all['상태'].isin(['취소', '반품', '교환'])].copy()

        if df_calc.empty:
            st.warning("현재 정산 가능한 유효 판매 데이터가 없습니다.")
        else:
            df_calc['수량'] = pd.to_numeric(df_calc['수량'], errors='coerce').fillna(1)
            if '주문처' not in df_calc.columns: df_calc['주문처'] = ''

            def calculate_profit(row):
                market = str(row.get('주문처', '')).strip()
                qty = row['수량']
                item_name = str(row['상품명']).strip()
                item_clean = item_name.replace(" ", "").lower()
                
                raw_paid = str(row.get('결제금액', '0')).replace(',', '').replace('원', '').strip()
                actual_paid = pd.to_numeric(raw_paid, errors='coerce')
                if pd.isna(actual_paid): actual_paid = 0
                
                raw_ship = str(row.get('택배비', '0')).replace(',', '').replace('원', '').strip()
                actual_ship_cost = pd.to_numeric(raw_ship, errors='coerce')
                if pd.isna(actual_ship_cost): actual_ship_cost = 0
                
                if market == "" or market == "nan": fee_rate = 0.0
                elif '스마트스토어' in market: fee_rate = fee_smart / 100
                elif '쿠팡' in market: fee_rate = fee_coupang / 100
                elif '자사몰' in market: fee_rate = fee_own / 100
                elif '개인' in market or 'B2B' in market: fee_rate = fee_personal / 100
                else: fee_rate = fee_etc / 100
                
                unit_cost = 0
                unit_price = 0
                
                if not df_cost.empty:
                    for _, opt in df_cost.iterrows():
                        std_name = str(opt.get('상품명', '')).strip()
                        mapping_str = str(opt.get('매핑명', '')).strip()
                        keywords = [k.strip() for k in mapping_str.split(',') if k.strip()] + [std_name]
                        for kw in keywords:
                            if not kw: continue
                            if kw.replace(" ", "").lower() in item_clean:
                                raw_price = str(opt.get('가격', 0)).replace(',', '').replace('원', '').replace('₩', '').replace('\\', '').strip()
                                raw_cost = str(opt.get('원가', 0)).replace(',', '').replace('원', '').replace('₩', '').replace('\\', '').strip()
                                unit_price = pd.to_numeric(raw_price, errors='coerce')
                                unit_cost = pd.to_numeric(raw_cost, errors='coerce')
                                if pd.isna(unit_price): unit_price = 0
                                if pd.isna(unit_cost): unit_cost = 0
                                break 
                
                total_revenue = actual_paid if actual_paid > 0 else (unit_price * qty)
                total_cost = unit_cost * qty
                commission_fee = total_revenue * fee_rate
                net_profit = total_revenue - total_cost - commission_fee - actual_ship_cost
                margin_rate = (net_profit / total_revenue * 100) if total_revenue > 0 else 0
                return pd.Series([total_revenue, commission_fee, unit_cost, total_cost, actual_ship_cost, net_profit, margin_rate])

            df_calc[['예상결제금액', '마켓수수료', '매입단가(1개)', '총매입원가', '적용택배비', '예상순이익', '마진율(%)']] = df_calc.apply(calculate_profit, axis=1)

            tab_sum, tab_month, tab_cal, tab_detail = st.tabs(["전체 요약", "월별 정산 내역", "일별 매출 캘린더", "주문별 상세 내역"])

            with tab_sum:
                st.markdown("### 누적 정산 리포트")
                total_sales = df_calc['예상결제금액'].sum()
                total_profit = df_calc['예상순이익'].sum()
                avg_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
                
                c1, c2, c3 = st.columns(3)
                c1.metric("누적 총 매출액", f"{total_sales:,.0f} 원")
                c2.metric("누적 총 순이익", f"{total_profit:,.0f} 원")
                c3.metric("평균 마진율", f"{avg_margin:.1f} %")

            with tab_month:
                df_calc['월'] = df_calc['날짜'].dt.strftime('%Y-%m')
                monthly_profit = df_calc.groupby('월').agg(
                    주문건수=('구매자명', 'count'), 매출액=('예상결제금액', 'sum'), 마켓수수료=('마켓수수료', 'sum'),
                    총매입원가=('총매입원가', 'sum'), 택배비총합=('적용택배비', 'sum'), 순이익=('예상순이익', 'sum')
                ).reset_index().sort_values('월', ascending=False)
                monthly_profit['평균마진율(%)'] = (monthly_profit['순이익'] / monthly_profit['매출액'] * 100).fillna(0)
                
                styled_monthly = monthly_profit.style.format({
                    '매출액': '{:,.0f}', '마켓수수료': '{:,.0f}', '총매입원가': '{:,.0f}', '택배비총합': '{:,.0f}',
                    '순이익': '{:,.0f}', '평균마진율(%)': '{:.1f}%'
                })
                try: styled_monthly = styled_monthly.background_gradient(subset=['평균마진율(%)'], cmap='RdYlGn')
                except: pass
                st.dataframe(styled_monthly, use_container_width=True, hide_index=True)

            with tab_cal:
                st.markdown("### 캘린더 뷰 (일별 매출 & 순이익)")
                cal_options = {
                    "headerToolbar": {
                        "left": "prev,next today",
                        "center": "title",
                        "right": "dayGridMonth"
                    },
                    "initialView": "dayGridMonth", 
                    "height": 650, 
                }
                
                valid_dates = df_calc.dropna(subset=['날짜_str']).copy()
                valid_dates['날짜_str'] = valid_dates['날짜_str'].astype(str).str.strip()
                
                valid_dates = valid_dates[
                    (valid_dates['날짜_str'] != '') & 
                    (valid_dates['날짜_str'].str.lower() != 'nan') & 
                    (valid_dates['날짜_str'].str.lower() != 'nat')
                ]
                
                events = []
                if not valid_dates.empty:
                    daily_sales = valid_dates.groupby('날짜_str').agg(매출액=('예상결제금액', 'sum'), 순이익=('예상순이익', 'sum')).reset_index()
                    for _, row in daily_sales.iterrows():
                        d_str = row['날짜_str']
                        events.append({"title": f"매출: {row['매출액']:,.0f}", "start": d_str, "color": "#555555"})
                        events.append({"title": f"이익: {row['순이익']:,.0f}", "start": d_str, "color": "#800020"})
                
                if not events:
                    st.info("💡 캘린더에 표시할 유효한 판매 데이터가 없습니다.")
                else:
                    dynamic_key = f"sales_cal_{len(events)}_{daily_sales['매출액'].sum()}"
                    calendar(events=events, options=cal_options, key=dynamic_key)

        with tab_detail:
                st.markdown("### 주문건별 상세 내역")
                display_cols = ['날짜_str', '구매자명', '상품명', '수량', '예상결제금액', '마켓수수료', '매입단가(1개)', '총매입원가', '적용택배비', '예상순이익', '마진율(%)']
                styled_df = df_calc[display_cols].style.format({
                    '예상결제금액': '{:,.0f}', '마켓수수료': '{:,.0f}', '매입단가(1개)': '{:,.0f}', '총매입원가': '{:,.0f}', '적용택배비': '{:,.0f}',
                    '예상순이익': '{:,.0f}', '마진율(%)': '{:.1f}%'
                })
                try: styled_df = styled_df.background_gradient(subset=['마진율(%)'], cmap='RdYlGn')
                except: pass
                st.dataframe(styled_df, use_container_width=True, hide_index=True)

with tab_cal:
                st.markdown("### 캘린더 뷰 (일별 매출 & 순이익)")
                cal_options = {
                    "headerToolbar": {
                        "left": "prev,next today",
                        "center": "title",
                        "right": "dayGridMonth"
                    },
                    "initialView": "dayGridMonth", 
                    "height": 650, 
                }
                
                # 🔴 1. 쓸모없는 날짜 데이터('NaT', 'nan', 빈칸) 완벽하게 걸러내기
                valid_dates = df_calc.dropna(subset=['날짜_str']).copy()
                valid_dates['날짜_str'] = valid_dates['날짜_str'].astype(str).str.strip()
                
                valid_dates = valid_dates[
                    (valid_dates['날짜_str'] != '') & 
                    (valid_dates['날짜_str'].str.lower() != 'nan') & 
                    (valid_dates['날짜_str'].str.lower() != 'nat')
                ]
                
                events = []
                if not valid_dates.empty:
                    daily_sales = valid_dates.groupby('날짜_str').agg(매출액=('예상결제금액', 'sum'), 순이익=('예상순이익', 'sum')).reset_index()
                    for _, row in daily_sales.iterrows():
                        d_str = row['날짜_str']
                        events.append({"title": f"매출: {row['매출액']:,.0f}", "start": d_str, "color": "#555555"})
                        events.append({"title": f"이익: {row['순이익']:,.0f}", "start": d_str, "color": "#800020"})
                
                if not events:
                    st.info("💡 캘린더에 표시할 유효한 판매 데이터가 없습니다.")
                else:
                    # 🔴 2. 데이터(매출 합계)가 바뀔 때마다 달력이 강제로 다시 그려지도록 스마트 키(key) 생성!
                    dynamic_key = f"sales_cal_{len(events)}_{daily_sales['매출액'].sum()}"
                    calendar(events=events, options=cal_options, key=dynamic_key)

            with tab_detail:
                st.markdown("### 주문건별 상세 내역")
                display_cols = ['날짜_str', '구매자명', '상품명', '수량', '예상결제금액', '마켓수수료', '매입단가(1개)', '총매입원가', '적용택배비', '예상순이익', '마진율(%)']
                styled_df = df_calc[display_cols].style.format({
                    '예상결제금액': '{:,.0f}', '마켓수수료': '{:,.0f}', '매입단가(1개)': '{:,.0f}', '총매입원가': '{:,.0f}', '적용택배비': '{:,.0f}',
                    '예상순이익': '{:,.0f}', '마진율(%)': '{:.1f}%'
                })
                try: styled_df = styled_df.background_gradient(subset=['마진율(%)'], cmap='RdYlGn')
                except: pass
                st.dataframe(styled_df, use_container_width=True, hide_index=True)

# === 옵션 관리 ===
elif menu == "옵션 관리":
    render_page_header("옵션 관리", "제품 등록/수정")
    df_opt, sheet_opt = load_data("옵션관리")
    if not df_opt.empty:
        edited_df = st.data_editor(df_opt, num_rows="dynamic", use_container_width=True)
        if st.button("저장"):
            sheet_opt.clear()
            new_data = [edited_df.columns.values.tolist()] + edited_df.values.tolist()
            sheet_opt.update(values=new_data, range_name="A1")
            st.success("저장됨"); st.rerun()

# === 재고 관리 ===
elif menu == "재고 관리":
    render_page_header("재고 관리", "통합 재고 관리 시스템")
    df_stock, sheet_stock = load_data("재고관리")
    df_opt, _ = load_data("옵션관리")
    tab1, tab2, tab3 = st.tabs(["재고 현황 (자동집계)", "대량 입출고 등록 (엑셀)", "개별 조정 (수동)"])
    
    with tab1:
        if not df_stock.empty and not df_all.empty:
            st.dataframe(df_stock, use_container_width=True)
            
    with tab2:
        st.markdown("### 엑셀로 재고 일괄 등록")
        uploaded_file = st.file_uploader("작성한 엑셀 파일 업로드", type=['xlsx', 'xls', 'csv'], key="stock_up")
        if uploaded_file and st.button("재고 일괄 반영하기", type="primary"):
            st.info("재고 일괄 반영 로직 수행")
            
    with tab3:
        st.markdown("### 개별 상품 입/출고 (블랙박스 작동중)")
        if not df_stock.empty:
            with st.form("manual_stock"):
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1: target_prod = st.selectbox("상품 선택", df_stock['상품명'].unique())
                with col2: action = st.radio("구분", ["입고 (+)", "출고/손실 (-)"], horizontal=True)
                with col3: qty = st.number_input("수량", min_value=1, value=1)
                
                if st.form_submit_button("반영"):
                    client = get_client()
                    fresh_stock_sheet = client.open_by_key(SHEET_ID).worksheet("재고관리")
                    cell = fresh_stock_sheet.find(target_prod)
                    if cell:
                        headers = fresh_stock_sheet.row_values(1)
                        col_idx = next((i + 1 for i, h in enumerate(headers) if '재고' in h or '수량' in h), -1)
                        if col_idx != -1:
                            curr_val = int(pd.to_numeric(fresh_stock_sheet.cell(cell.row, col_idx).value, errors='coerce') or 0)
                            final_qty = curr_val + qty if "입고" in action else curr_val - qty
                            fresh_stock_sheet.update_cell(cell.row, col_idx, final_qty)
                            log_msg = f"{target_prod} {qty}개 {action.split()[0]} 처리 (변경 전: {curr_val} -> 변경 후: {final_qty})"
                            add_log("재고수동조정", log_msg)
                            st.success(f"✅ {target_prod}: {final_qty}개로 변경 및 로그 기록 완료!"); time.sleep(1); st.rerun()

# === 일정 관리 ===
elif menu == "일정 관리":
    render_page_header("일정 관리", "사내 스케줄 및 주요 일정을 등록하고 관리하세요.")
    df_sch, sheet_sch = load_data("일정관리")
    tab_cal, tab_edit = st.tabs(["캘린더 뷰", "일정 등록 및 수정 (수동 편집)"])
    
    with tab_cal:
        if not df_sch.empty:
            events = []
            for _, r in df_sch.iterrows():
                title = str(r.get('일정명', ''))
                time_str = str(r.get('시간', ''))
                if time_str and time_str != 'nan':
                    title = f"[{time_str}] {title}"
                events.append({"title": title, "start": str(r.get('시작일', '')), "color": "#2B3A55"})
            calendar(events=events, options={"height": 650})
        else:
            st.info("💡 아직 등록된 일정이 없습니다.")

    with tab_edit:
        col1, col2 = st.columns([1, 2.5])
        with col1:
            st.markdown("##### ➕ 새 일정 추가")
            with st.form("add_schedule"):
                d_date = st.date_input("날짜", datetime.now())
                d_time = st.time_input("시간")
                d_title = st.text_input("일정명")
                d_desc = st.text_area("상세내용")
                if st.form_submit_button("일정 저장", type="primary"):
                    if sheet_sch: 
                        sheet_sch.append_row([str(d_date), str(d_date), str(d_time), d_title, d_desc])
                        st.success("✅ 새 일정이 저장되었습니다!"); time.sleep(1); st.rerun()
        with col2:
            st.markdown("##### 기존 일정 수정 및 삭제")
            if not df_sch.empty:
                edited_df = st.data_editor(df_sch, num_rows="dynamic", use_container_width=True, key="schedule_editor")
                if st.button("변경된 일정 내용 시트에 반영하기", type="secondary"):
                    with st.spinner("구글 시트에 업데이트 중입니다..."):
                        if sheet_sch:
                            sheet_sch.clear()
                            new_data = [edited_df.columns.values.tolist()] + edited_df.values.tolist()
                            sheet_sch.update(values=new_data, range_name="A1")
                            st.success(" 일정이 성공적으로 수정/삭제되었습니다!"); time.sleep(1); st.rerun()
            else:
                st.write("등록된 일정이 없습니다.")

# === 마케팅 & CRM ===
elif menu == "마케팅 & CRM":
    render_page_header("마케팅 & CRM 통합 센터", "고객 관리부터 광고 성과 측정, AI 카피라이팅까지 한곳에서 관리하세요.")
    
    m_tab1, m_tab2, m_tab3, m_tab4, m_tab5 = st.tabs([
        "1. 고객 CRM 프로필", "2. 광고 효율(ROAS)", "3. AI 카피/네이밍", "4. 리뷰/CS 응대", "5. SNS 콘텐츠 생성"
    ])

    with m_tab1:
        st.markdown("#### 고객 통합 프로필 및 상담")
        if not df_all.empty:
            df_crm = df_all.copy()
            if '결제금액' in df_crm.columns: df_crm['amt'] = pd.to_numeric(df_crm['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
            else: df_crm['amt'] = 0
            
            cust_profile = df_crm.groupby('구매자명').agg({'날짜': ['max', 'count'], 'amt': 'sum'}).reset_index()
            cust_profile.columns = ['고객명', '최근구매일', '구매횟수', '누적금액']
            
            def analyze_cx(row):
                grade = "VIP" if row['누적금액'] >= 500000 else "일반"
                days = (datetime.now() - row['최근구매일']).days if pd.notnull(row['최근구매일']) else 0
                status = "교체주기" if 150 <= days <= 210 else "정상"
                return pd.Series([grade, status, days])
            cust_profile[['등급', '상태', '경과일']] = cust_profile.apply(analyze_cx, axis=1)
            
            c_list, c_detail = st.columns([1, 1.5])
            with c_list:
                st.dataframe(cust_profile[['고객명', '등급', '누적금액', '상태']], use_container_width=True, hide_index=True)
            with c_detail:
                search_nm = st.selectbox("상세 조회할 고객 선택", cust_profile['고객명'].unique())
                sel_data = cust_profile[cust_profile['고객명'] == search_nm].iloc[0]
                st.write(f"**등급:** {sel_data['등급']} | **누적금액:** {sel_data['누적금액']:,.0f}원")
                try:
                    client = get_client(); target_sh = client.open("주문데이터").worksheet("시트1")
                    h = target_sh.row_values(1)
                    if '비고' in h:
                        cell = target_sh.find(search_nm)
                        if cell:
                            current_history = target_sh.cell(cell.row, h.index('비고')+1).value
                            st.text_area("상담 히스토리", value=current_history or "내용 없음", height=100, disabled=True)
                            memo_in = st.text_input("신규 상담 내용 입력")
                            if st.button("저장"):
                                now = datetime.now().strftime('%Y-%m-%d %H:%M')
                                final = f"{current_history}\n[{now}] {memo_in}" if current_history else f"[{now}] {memo_in}"
                                target_sh.update_cell(cell.row, h.index('비고')+1, final)
                                st.success("저장됨")
                except: st.info("히스토리 연동 대기중")

    with m_tab2:
        st.markdown("#### 일일 광고비 입력 및 ROAS 측정")
        today_str = datetime.now().strftime("%Y-%m-%d")
        today_sales = 0
        if '결제금액' in df_all.columns:
            sales_today = df_all[df_all['날짜_str'] == today_str]['결제금액'].astype(str)
            sales_today = sales_today.str.replace(r'[^\d]', '', regex=True)
            today_sales = pd.to_numeric(sales_today, errors='coerce').fillna(0).sum()
        
        with st.form("roas_form"):
            r_col1, r_col2 = st.columns(2)
            with r_col1: naver_spend = st.number_input("네이버 광고비 (원)", min_value=0, step=10000)
            with r_col2: meta_spend = st.number_input("인스타/페북 광고비 (원)", min_value=0, step=10000)
            if st.form_submit_button("ROAS 계산하기", type="primary"):
                total_spend = naver_spend + meta_spend
                c_roas1, c_roas2, c_roas3 = st.columns(3)
                c_roas1.metric("총 광고비 지출", f"{total_spend:,.0f}원")
                c_roas2.metric("오늘 발생한 총 매출(추정)", f"{today_sales:,.0f}원")
                if total_spend > 0:
                    roas = (today_sales / total_spend) * 100
                    roas_icon = "대박!" if roas >= 400 else ("양호" if roas >= 250 else "점검 필요")
                    c_roas3.metric("오늘의 ROAS", f"{roas:,.1f}% {roas_icon}")

    with m_tab3:
        st.markdown("#### 매체 최적화 광고 문구 / 네이밍 생성")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            product_name = st.text_input("상품 특징", value="시그니처 와인 컬러 프리미엄 와플 타월")
            channel = st.selectbox("광고 매체", ["네이버 파워링크", "인스타그램", "카카오 알림톡"])
        with col_c2:
            target_audience = st.selectbox("타겟 고객", ["3040 리빙", "2030 신혼부부/집들이", "전체"])
            tone = st.selectbox("톤앤매너", ["세련된", "재치있는", "전문적인", "긴급/한정수량"])
        
        if st.button("광고 소재 및 네이밍 생성", type="primary"):
            with st.spinner("AI가 소재를 작성 중입니다..."):
                prompt = f"상품: {product_name}, 타겟: {target_audience}, 채널: {channel}, 톤: {tone}. 이 매체에 완벽하게 맞는 광고 카피 3가지와 매력적인 캠페인 네이밍 2가지를 제안해줘."
                try:
                    st.markdown(f"<div style='background-color:#fff; padding:20px; border-radius:12px; border:1px solid #E9ECEF;'>{ask_ai(prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
                except: st.error("AI 연결이 필요합니다.")

    with m_tab4:
        st.markdown("#### 스마트 고객 응대 (리뷰 & CS)")
        sub_tab1, sub_tab2 = st.tabs(["엑셀 리뷰 일괄 답글", "CS 답변 스크립트"])
        with sub_tab1:
            uploaded_review = st.file_uploader("리뷰 엑셀 파일 (.xlsx)", type=['xlsx'])
            if uploaded_review:
                df_rev = pd.read_excel(uploaded_review)
                review_col = st.selectbox("고객 리뷰 내용 열", df_rev.columns.tolist())
                if st.button("AI 답글 생성"):
                    with st.spinner("답글 작성 중..."):
                        try:
                            replies = [ask_ai(f"고객 리뷰: '{row[review_col]}'. 감사와 공감이 담긴 정중한 답글 2문장 작성.") for _, row in df_rev.iterrows()]
                            df_rev['AI_자동답글'] = replies
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df_rev.to_excel(writer, index=False)
                            st.download_button("결과 다운로드", output.getvalue(), "리뷰답글완료.xlsx")
                        except: st.error("AI 연결이 필요합니다.")
        with sub_tab2:
            cs_type = st.radio("문의 유형", ["배송 지연", "제품 교환", "불만", "기타"], horizontal=True)
            cs_detail = st.text_area("고객 문의 내용")
            if st.button("방어 답변 생성") and cs_detail:
                try: st.info(ask_ai(f"유형: {cs_type}, 내용: {cs_detail}\n프리미엄 브랜드에 맞는 정중한 사과와 해결책이 담긴 답변 작성."))
                except: st.error("AI 연결이 필요합니다.")

    with m_tab5:
        st.markdown("#### 📸 AI 블로그/인스타 콘텐츠 및 이미지 생성")
        st.info("주제, 컨셉, 또는 참고 URL을 입력하면 AI가 블로그/인스타그램용 글과 추천 이미지를 생성해 줍니다.")
        sns_type = st.radio("어떤 플랫폼에 올리실 건가요?", ["📸 인스타그램", "📝 네이버 블로그", "💬 스레드 / X"], horizontal=True)
        with st.form("ai_sns_form"):
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                mkt_topic = st.text_input("콘텐츠 주제 또는 제품명", placeholder="예: 봄맞이 프리미엄 호텔 수건 세트")
                mkt_url = st.text_input("참고 URL (선택)", placeholder="예: 상세페이지 또는 참고할 타사 링크")
            with col_m2:
                mkt_tone = st.selectbox("글의 톤앤매너", ["친근하고 발랄하게", "전문적이고 신뢰감 있게", "감성적이고 따뜻하게", "유머러스하고 재치있게"])
                mkt_concept = st.text_input("이미지 컨셉 (선택)", placeholder="예: 따뜻한 햇살이 비치는 화이트 욕실에 놓인 수건")
            mkt_detail = st.text_area("핵심 내용 및 강조할 포인트 (할인 행사, 특장점 등)")
            submit_btn = st.form_submit_button("✨ AI 콘텐츠 및 사진 생성하기", type="primary")

        if submit_btn:
            if not mkt_topic and not mkt_detail: st.warning("주제나 핵심 내용을 최소한 하나는 입력해 주세요!")
            else:
                with st.spinner("AI가 열심히 글을 작성하고 사진을 기획하고 있습니다..."):
                    prompt = f"플랫폼: {sns_type}\n주제: {mkt_topic}\n참고URL: {mkt_url}\n톤앤매너: {mkt_tone}\n핵심내용: {mkt_detail}\n\n위 내용을 바탕으로 {sns_type}에 바로 올릴 수 있는 매력적인 홍보 글을 작성해줘. 해시태그도 5개 이상 넉넉히 포함해줘."
                    try: generated_text = ask_ai(prompt)
                    except: generated_text = "AI 서버와 연결할 수 없습니다. ask_ai 설정을 확인해주세요."
                    st.success("🎉 마케팅 콘텐츠가 성공적으로 생성되었습니다!")
                    tab_text, tab_img = st.tabs(["📝 생성된 텍스트", "🎨 생성된 이미지"])
                    with tab_text:
                        st.markdown(f"**[{sns_type} 맞춤형 텍스트]**")
                        st.text_area("복사해서 바로 SNS에 올려보세요!", generated_text.strip(), height=300)
                    with tab_img:
                        st.markdown(f"**[요청하신 '{mkt_concept}' 컨셉의 추천 이미지]**")
                        st.image("https://images.unsplash.com/photo-1584947937402-28e4e9fbdba8?q=80&w=800&auto=format&fit=crop", caption="AI 생성 이미지 미리보기 (샘플)")

# === AI 비즈니스 센터 ===
elif menu == "AI 비즈니스 센터":
    render_page_header("DUWELL AI 비즈니스 센터", "10년 노하우를 학습한 AI 에이전트 팀이 대표님의 업무를 지원합니다.")
    ai_tab1, ai_tab2, ai_tab3, ai_tab4, ai_tab5, ai_tab6, ai_tab7 = st.tabs([
        "1. MD 신제품", "2. B2B 영업", "3. 리뷰 분석", "4. 경영 브리핑", "5. 마진 시뮬", "6. 이미지 프롬프트", "7. 마켓별 SEO 등록"
    ])

    with ai_tab1:
        st.markdown("#### 신제품 런칭 브리프 생성기")
        with st.form("md_agent_form"):
            new_product_desc = st.text_area("기획 중인 상품 특징 입력", placeholder="예: 프리미엄 와플 직조 수건. 일반 수건보다 건조가 빠르고 먼지가 안 나. 고급 에스테틱 느낌.", height=100)
            if st.form_submit_button("✨ 런칭 기획안 생성", type="primary"):
                if new_product_desc:
                    with st.spinner("MD 에이전트가 기획안을 작성 중입니다..."):
                        agent_prompt = f"""You are an elite Towel Merchandiser for 'DUWELL'. Backed by 10 years of towel industry know-how, write a product launch brief in Korean.
[신제품 특징]: {new_product_desc}
출력 형식: [상품명 아이디어 3개], [핵심 타겟 고객], [강력한 셀링 포인트 3개], [상세페이지 스토리라인]"""
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    with ai_tab2:
        st.markdown("#### B2B 맞춤형 영업 제안서 작성")
        with st.form("b2b_agent_form"):
            col1, col2 = st.columns(2)
            with col1: target_company = st.text_input("제안 대상 (예: A 고급 에스테틱, B 부티크 호텔)")
            with col2: target_product = st.text_input("제안 상품 (예: 프리미엄 와플 수건 세트)")
            sales_points = st.text_area("강조할 소구점 (예: 먼지 없음, 빠른 건조, 고급스러운 디자인)")
            
            if st.form_submit_button("B2B 영업 메일 초안 생성", type="primary"):
                if target_company and target_product:
                    with st.spinner("B2B 영업 에이전트가 제안서를 작성 중입니다..."):
                        agent_prompt = f"""You are an elite B2B Sales Representative for the premium towel brand 'DUWELL'. 
Write a highly professional, polite, and persuasive B2B sales email in Korean.
- 타겟 고객사: {target_company} / 제안 상품: {target_product} / 강조할 포인트: {sales_points}
- 톤앤매너: 10년 경력의 신뢰감, 고급스러움, 상대방 비즈니스에 확실한 도움이 된다는 확신."""
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    with ai_tab3:
        st.markdown("#### 타사 리뷰 기반 마케팅 포인트 추출")
        with st.form("review_agent_form"):
            bad_reviews = st.text_area("경쟁사(일반 수건) 부정적 리뷰 복사/붙여넣기", placeholder="예: 먼지가 너무 날려요. 잘 안 마르고 꿉꿉한 냄새가 나요.", height=100)
            if st.form_submit_button("공격적 마케팅 무기 생성", type="primary"):
                if bad_reviews:
                    with st.spinner("리서치 에이전트가 페인포인트를 분석 중입니다..."):
                        agent_prompt = f"""You are an expert Market Researcher and Copywriter for 'DUWELL'.
Analyze the following negative reviews of competitor's normal towels: "{bad_reviews}"
1. 고객의 핵심 Pain Point 요약 / 2. 이를 완벽히 해결해주는 DUWELL '프리미엄 와플 수건'의 특장점 연결 / 3. 당장 인스타그램 카드뉴스에 쓸 수 있는 후킹 카피 3가지 제안 (한국어)"""
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    with ai_tab4:
        st.markdown("#### 대표님 맞춤형 일일 경영 브리핑 (데이터 연동)")
        if st.button("오늘의 경영 브리핑 생성", type="primary"):
            with st.spinner("비서 에이전트가 데이터를 취합하고 분석 중입니다..."):
                today_str = datetime.now().strftime("%Y-%m-%d")
                today_sales = df_all[df_all['날짜_str'] == today_str]['결제금액'].sum() if not df_all.empty and '결제금액' in df_all.columns else 0
                
                df_stock_temp, _ = load_data("재고관리")
                low_stock_msg = "재고 부족 상품 없음"
                if not df_stock_temp.empty:
                    df_stock_temp['현재재고'] = pd.to_numeric(df_stock_temp['현재재고'], errors='coerce').fillna(0)
                    df_stock_temp['안전재고'] = pd.to_numeric(df_stock_temp['안전재고'], errors='coerce').fillna(0)
                    low_items = df_stock_temp[df_stock_temp['현재재고'] <= df_stock_temp['안전재고']]['상품명'].tolist()
                    if low_items: low_stock_msg = ", ".join(low_items) + " (발주 필요!)"

                df_sch_temp, _ = load_data("일정관리")
                today_schedule = "일정 없음"
                if not df_sch_temp.empty:
                    today_events = df_sch_temp[df_sch_temp['시작일'] == today_str]['일정명'].tolist()
                    if today_events: today_schedule = ", ".join(today_events)

                agent_prompt = f"""You are the Executive Assistant to the CEO of 'DUWELL'. Write a crisp, objective, and encouraging morning briefing in Korean.
[오늘의 데이터] 예상 매출: {today_sales:,.0f}원 / 재고 경고: {low_stock_msg} / 주요 일정: {today_schedule}
위 데이터를 바탕으로: 1. 성과 요약 및 칭찬 한마디 / 2. 오늘 반드시 처리해야 할 액션 아이템 2가지 제안"""
                st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    with ai_tab5:
        st.markdown("#### 기획 상품 마진 시뮬레이터")
        with st.form("margin_agent_form"):
            col1, col2, col3 = st.columns(3)
            with col1: base_price = st.number_input("기존 판매가 (원)", value=15000)
            with col2: discount_rate = st.number_input("할인율 (%)", value=15)
            with col3: product_cost = st.number_input("매입 원가 (원)", value=6000)
            extra_costs = st.text_input("추가 비용 내역 (예: 포장비 1000원, 사은품 500원)")
            market_type = st.selectbox("판매 채널 (수수료율)", ["스마트스토어 (약 5.5%)", "쿠팡 (약 10.8%)", "자사몰 (약 3%)"])

            if st.form_submit_button("적정 판매가 및 마진 계산", type="primary"):
                with st.spinner("재무 에이전트가 마진율을 계산 중입니다..."):
                    agent_prompt = f"""You are a strict and smart Financial Advisor for 'DUWELL'. Calculate the profit margin based on the following data:
- 기존 판매가: {base_price}원 / 기획 할인율: {discount_rate}% / 매입 원가: {product_cost}원 / 추가 비용: {extra_costs} / 판매 채널: {market_type}
1. 최종 예상 판매가, 예상 수수료, 마진 금액, 최종 마진율(%)을 수식과 함께 직관적으로 보여주세요.
2. 마진율이 30% 미만일 경우, 이익을 방어하기 위한 가격 정책이나 세트 구성 아이디어를 제안해주세요. (한국어)"""
                    st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    with ai_tab6:
        st.markdown("#### 📸 고품질 상세페이지 이미지 프롬프트 생성기")
        with st.form("image_prompt_form"):
            col1, col2 = st.columns(2)
            with col1: prod_img_desc = st.text_input("제품의 핵심 특징", placeholder="예: 먼지 없는 프리미엄 와플 직조 타월")
            with col2: img_mood = st.selectbox("원하는 연출 분위기", ["햇살이 들어오는 따뜻하고 아늑한 욕실", "5성급 호텔의 모던하고 어두운 고급 욕실", "깨끗하고 위생적인 에스테틱 샵", "제품의 질감을 극대화한 초근접 마크로 샷"])
            if st.form_submit_button("영문 프롬프트 생성", type="primary"):
                if prod_img_desc:
                    with st.spinner("전문 포토그래퍼 에이전트가 카메라 렌즈와 조명 세팅을 조율 중입니다..."):
                        agent_prompt = f"""You are an elite Commercial Photographer and AI Prompt Engineer specializing in Home & Living products.
I need highly detailed, professional image generation prompts (optimized for Midjourney v6 or DALL-E 3) based on:
- Target Product: {prod_img_desc} / - Mood & Background: {img_mood}
Please create 3 different variations of the shot (e.g., Close-up texture, Lifestyle interior, Wide angle). For each variation, output MUST strictly follow this format:
📌 [Shot Type in Korean]
- 연출 의도: (Korean description of the scene)
- 🇬🇧 Prompt: (English comma-separated prompt containing subject, background, lighting, camera lens like 85mm f/1.8, high-end commercial photography, 8k resolution, photorealistic)"""
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
                else: st.warning("제품 특징을 입력해주세요!")

    with ai_tab7:
        st.markdown("#### 오픈마켓별 SEO 등록 & HTML 상세페이지 생성기")
        with st.form("seo_agent_form"):
            col1, col2 = st.columns(2)
            with col1:
                target_market = st.selectbox("등록할 마켓 선택", ["네이버 스마트스토어", "쿠팡", "자사몰 / 기타 오픈마켓"])
                prod_name = st.text_input("기본 상품명", value="프리미엄 와플 수건")
            with col2:
                target_customer = st.selectbox("메인 타겟", ["3040 리빙/인테리어", "2030 집들이/신혼부부", "답례품/대량구매", "전체 연령대"])
                core_points = st.text_input("핵심 강조 포인트", value="먼지없는, 빠른건조, 고급스러운 질감")
            submit_btn = st.form_submit_button("SEO 데이터 및 HTML 생성", type="primary")

        if submit_btn:
            if prod_name and core_points:
                with st.spinner(f"{target_market} 로직에 맞춰 최적화 데이터와 HTML 코드를 작성 중입니다... (약 10~20초 소요)"):
                    agent_prompt = f"""You are an elite E-commerce Merchandiser, SEO expert, and Web Designer in Korea, working for the premium towel brand 'DUWELL'.
Your task is to generate highly optimized product registration data and a complete HTML detail page for {target_market}.
- Product: {prod_name} / - Target Customer: {target_customer} / - Key Selling Points: {core_points}
Please provide the output strictly in Korean, following this structure:
1. 🏷️ [최적화 상품명 3가지]: Create 3 variations of the product title optimized for {target_market}'s search algorithm.
2. 🔑 [검색 태그/키워드]: Provide exactly 10 highly searched, relevant tags/keywords separated by commas.
3. 📝 [메타 디스크립션/PC·모바일 요약 설명]: A compelling 2-3 sentence description.
4. 💻 [상세페이지 기획 뼈대]: 3-step storyline for the detail page (Hook -> USP explanation -> Closing/Trust).
5. 🌐 [네이버 에디터 완벽 호환 HTML 상세페이지]: Write clean, modern HTML code. MUST use inline CSS on EVERY SINGLE TAG.
- [폰트 크기]: Use `em` or `px`. Do NOT use `vw` or `clamp()`.
- [단어 찢어짐 100% 방지 규칙]: Apply `word-break: keep-all;` to EVERY text tag. For titles or long sentences, insert `<br>` tags at natural semantic pauses.
- Style: #2B3A55 or #800020 for accent colors, centered text, `line-height: 1.6;`.
- Ensure the HTML code is enclosed in a markdown code block."""
                    st.session_state['seo_result_text'] = ask_ai(agent_prompt)
                    st.session_state['seo_prod_name'] = prod_name
            else: st.warning("상품명과 핵심 포인트를 입력해주세요!")

        if 'seo_result_text' in st.session_state:
            st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{st.session_state['seo_result_text'].replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
            st.download_button(
                label="기획안 및 HTML 코드 다운로드 (.txt)",
                data=st.session_state['seo_result_text'],
                file_name=f"DUWELL_상세페이지기획_HTML_{st.session_state.get('seo_prod_name', '기본')}.txt",
                mime="text/plain",
                use_container_width=True
            )
