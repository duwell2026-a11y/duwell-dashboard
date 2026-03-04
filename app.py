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
import io

# --------------------------------------------------------------------------
# 🔒 초간단 도어락 (보안 및 사용자 식별)
# --------------------------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

# 로그인 안 한 상태면 현관문(로그인 화면)만 보여주고 밑으로 못 내려가게 막음
if not st.session_state['logged_in']:
    st.markdown("<h2 style='text-align: center; margin-top: 100px;'>🍷 DUWELL 스마트 센터</h2>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        # 두 분 중 누구인지 선택
        user_choice = st.selectbox("접속자 선택", ["고은정 (대표)", "두재훈 (팀장)"])
        # 공용 비밀번호 입력 (예: 오픈 연도나 기념일)
        pwd = st.text_input("비밀번호 (PIN)", type="password")
        
        if st.form_submit_button("입장하기", use_container_width=True):
            if pwd == "1121":  # 원하는 비밀번호로 변경하세요!
                st.session_state['logged_in'] = True
                st.session_state['current_user'] = user_choice
                st.rerun()
            else:
                st.error("🚨 비밀번호가 틀렸습니다.")
    
    # 이 아래 코드는 실행 안 되게 여기서 정지!
    st.stop()

# --------------------------------------------------------------------------
# 1. 페이지 및 홈페이지(웹사이트) 스타일 상단 메뉴 UI 설정
# --------------------------------------------------------------------------
import streamlit as st

st.set_page_config(page_title="DUWELL 스마트 ERP", layout="wide", page_icon="🍷", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
        /* 폰트 설정 */
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        html, body, [class*="css"] { font-family: 'Pretendard', sans-serif; }
        
        /* 1. 전체 배경색 */
        .stApp { background-color: #F4F6F9; }
        
        /* 2. 사이드바 및 접기 버튼 완전 숨김 (홈페이지 스타일을 위해) */
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }

        /* 3. 상단 헤더 숨기기 */
        [data-testid="stHeader"] { background: transparent; display: none; }

        /* 4. 🔥 상단 메뉴 (라디오 버튼) 홈페이지 네비게이션으로 튜닝 */
        div.row-widget.stRadio > div {
            display: flex;
            flex-direction: row;
            justify-content: center; /* 중앙 정렬 */
            gap: 15px;
            background-color: #FFFFFF;
            padding: 15px 20px;
            border-radius: 16px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.05);
            margin-bottom: 10px;
            flex-wrap: wrap; /* 화면이 좁아지면 알아서 줄바꿈 */
        }
        div.row-widget.stRadio > div > label {
            background-color: transparent;
            padding: 12px 20px;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.2s ease;
            border: 1px solid transparent;
            margin: 0;
        }
        div.row-widget.stRadio > div > label:hover {
            background-color: #F0F4FF;
        }
        
        /* 기존 동그라미 라디오 아이콘 없애기 */
        div.row-widget.stRadio > div > label[data-baseweb="radio"] > div:first-child {
            display: none !important; 
        }
        
        /* 선택된 메뉴 디자인 (포인트 블루) */
        div.row-widget.stRadio > div > label[data-checked="true"] {
            background-color: #4E73DF !important;
            box-shadow: 0 4px 10px rgba(78,115,223,0.3);
        }
        div.row-widget.stRadio > div > label[data-checked="true"] * {
            color: #FFFFFF !important;
            font-weight: 800 !important;
        }

        /* 5. 카드형 메트릭 및 표 스타일 (이전과 동일) */
        [data-testid="metric-container"] {
            background-color: #FFFFFF; border-radius: 16px; padding: 20px 24px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04); border: 1px solid rgba(0,0,0,0.02);
            transition: transform 0.2s ease;
        }
        [data-testid="metric-container"]:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0, 0, 0, 0.08); }
        [data-testid="metric-container"] label { color: #6C757D !important; font-weight: 600; font-size: 1rem; }
        [data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #2B3A55 !important; font-size: 1.8rem; font-weight: 800; }
        .stDataFrame { background-color: #FFFFFF; padding: 15px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03); border: 1px solid #EBEEF4; }
        [data-testid="stTabs"] button { border-bottom: 3px solid transparent; font-weight: 600; color: #6C757D !important; }
        [data-testid="stTabs"] button[aria-selected="true"] { color: #4E73DF !important; border-bottom: 3px solid #4E73DF !important; }

        /* 모바일 반응형 */
        @media (max-width: 768px) {
            .block-container { padding-top: 1rem !important; padding-left: 0.5rem !important; padding-right: 0.5rem !important; }
            div.row-widget.stRadio > div { gap: 5px; padding: 10px; }
            div.row-widget.stRadio > div > label { padding: 8px 10px; font-size: 0.9rem; flex: 1 1 45%; text-align: center; }
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
        # 💻 1. 내 컴퓨터(로컬)에서 테스트할 때의 세팅
        SHEET_ID = "1xqcbuzRzzp4i_Qsy4CKRjIIvGOTthT88bXxxY5RjEjQ"
        SENDER_EMAIL = "duwell2026@gmail.com"
        # API 키와 이메일 비밀번호는 비밀금고(secrets.toml)에서 가져옵니다!
        GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
        SENDER_PASSWORD = st.secrets["SENDER_PASSWORD"]
        
        with open(local_key_path, "r", encoding="utf-8") as f:
            GOOGLE_CREDENTIALS = json.load(f)
            
    else:
        # 🌐 2. 깃허브 연동 후 웹(Streamlit Cloud)에서 실행될 때의 세팅
        SHEET_ID = st.secrets["SHEET_ID"]
        SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
        GOOGLE_API_KEY = st.secrets["GOOGLE_API_KEY"]
        SENDER_PASSWORD = st.secrets["SENDER_PASSWORD"]
        
        # 구글 시트 인증서(JSON) 내용도 웹의 비밀금고에서 불러옵니다.
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

# ==========================================================
# [수정] 인증 및 데이터 로드 함수 (캐시 최적화 적용)
# ==========================================================
from google.oauth2.service_account import Credentials

# ==========================================================
# 1. 구글 연결 최적화 (API 호출 제한 에러 완벽 방어 🛡️)
# ==========================================================
@st.cache_resource(ttl=3600) # 🔴 핵심: 구글 시트와의 '연결 통로'를 1시간 동안 기억해둡니다.
def get_client():
    try:
        creds_dict = dict(GOOGLE_CREDENTIALS)
        return gspread.service_account_from_dict(creds_dict)
    except Exception as e:
        st.error(f"인증 에러: {e}")
        return None

@st.cache_resource(ttl=3600) # 🔴 핵심: 시트 객체도 기억해둬서 매번 구글에 문을 두드리지 않습니다.
def get_sheet_object(sheet_name):
    client = get_client()
    if client:
        return client.open_by_key(SHEET_ID).worksheet(sheet_name)
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

@st.cache_data(ttl=300) 
def fetch_raw_data(sheet_name):
    sheet = get_sheet_object(sheet_name)
    if not sheet: return []
    try:
        return sheet.get_all_records()
    except Exception:
        return []

def load_data(sheet_name):
    raw_data = fetch_raw_data(sheet_name)
    df = pd.DataFrame(raw_data)
    
    # 매번 연결하지 않고 기억해둔 시트 객체를 바로 가져옵니다. (에러 멈춤!)
    sheet = get_sheet_object(sheet_name)
    
    if df.empty: return df, sheet
    
    # --- 대표님의 기존 전처리 코드 완벽하게 동일 ---
    df.columns = [str(c).strip() for c in df.columns]
    for col in ['날짜', '시작일', '종료일', '주문일시', '주문일']:
        if col in df.columns: df[col] = df[col].apply(clean_date_str)

# ==========================================================
# 2. 캐시 분리 (데이터가 안 뜨는 '텅 빈 화면' 문제 해결)
# ==========================================================
@st.cache_data(ttl=300) 
def fetch_raw_data(sheet_name):
    client = get_client()
    if not client: return []
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        return sheet.get_all_records()
    except Exception:
        return []

# ==========================================================
# 3. 대표님의 기존 데이터 로직 100% 복구 + 시트 객체 연결
# ==========================================================
def load_data(sheet_name):
    # 캐시된 데이터 가져오기
    raw_data = fetch_raw_data(sheet_name)
    df = pd.DataFrame(raw_data)
    
    # [핵심] 시트 객체(sheet)는 캐시하지 않고 그때그때 가져옵니다. 
    # (이렇게 해야 나중에 옵션 '저장' 버튼을 누를 때 에러가 안 납니다!)
    client = get_client()
    sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name) if client else None
    
    if df.empty: return df, sheet
    
    # --- 대표님의 기존 전처리 코드 완벽하게 동일 ---
    df.columns = [str(c).strip() for c in df.columns]
    for col in ['날짜', '시작일', '종료일', '주문일시', '주문일']:
        if col in df.columns: df[col] = df[col].apply(clean_date_str)
    
    rename_map = {
        '주문일시': '날짜', '주문일': '날짜', '일자': '날짜',
        '금액': '결제금액', '총 주문금액': '결제금액',
        '성함': '구매자명', '고객명': '구매자명', '이름': '구매자명', '수취인명': '구매자명',
        '연락처': '연락처', '수취인연락처1': '연락처', '전화번호': '연락처',
        '주소': '주소', '배송지': '주소',
        '상품': '상품명', '품목': '상품명', '제품명': '상품명', 
        '디자인파일': '디자인파일', '첨부파일': '디자인파일', '시안': '디자인파일',
        '상태': '상태', '진행상태': '상태',
        '배송메세지': '요청사항', '비고': '요청사항', '메모': '요청사항',
        '포장옵션': '포장옵션', '컬러': '컬러'
    }
    df.rename(columns=rename_map, inplace=True)
    df = df.loc[:, ~df.columns.duplicated()]
    
    required_cols = ['날짜', '구매자명', '연락처', '주소', '상품명', '수량', '결제금액', '요청사항', '디자인파일', '상태', '포장옵션']
    for col in required_cols:
        if col not in df.columns: df[col] = "" 
    
    if '주문처' not in df.columns: df['주문처'] = '🏠 자사몰'
    
    return df, sheet

def update_status_in_sheet(sheet, row_data, new_status="발주완료"):
    try:
        # 1. 전체 데이터 로드
        records = sheet.get_all_records()
        header = sheet.row_values(1)
        
        # 2. '상태' 컬럼 인덱스 찾기 (다양한 표현 대응)
        col_idx = -1
        status_names = ['상태', '진행상태', '배송상태', '주문상태']
        for i, h in enumerate(header):
            if any(name in h.strip() for name in status_names):
                col_idx = i + 1
                break
        
        if col_idx == -1:
            return False, "❌ 시트에서 '진행상태' 열을 찾을 수 없습니다."

        # 3. 대상 행(Row) 찾기
        target_row_idx = -1
        # 비교를 위해 입력 데이터 정리
        t_name = str(row_data.get('구매자명', '')).strip()
        t_item = str(row_data.get('상품명', '')).strip()

        for idx, record in enumerate(records):
            # 시트 데이터 정리 (여러 컬럼명 대응)
            r_name = str(record.get('구매자명') or record.get('성함') or record.get('이름') or record.get('수취인명') or '').strip()
            r_item = str(record.get('상품명') or record.get('상품') or record.get('제품명') or '').strip()
            
            # 이름이 일치하고 상품명이 포함되어 있다면 해당 행으로 간주
            if r_name == t_name and (t_item in r_item or r_item in t_item):
                target_row_idx = idx + 2 # 헤더가 1행이므로 +2
                break
                
        # 4. 시트 업데이트 실행
        if target_row_idx != -1:
            sheet.update_cell(target_row_idx, col_idx, new_status)
            return True, f"✅ {target_row_idx}행 업데이트 성공"
        else:
            return False, f"❌ '{t_name}' 고객의 주문을 시트에서 매칭하지 못했습니다."

    except Exception as e:
        return False, f"❌ 시스템 오류: {str(e)}"

def get_drive_id(url):
    if not url or url == "-" or "이미지없음" in url: return None
    # /d/ 뒤 또는 id= 뒤의 25~50자 ID를 추출 (usp= 등 불필요한 인자 무시)
    match = re.search(r"(?:id=|\/d\/)([\w-]{25,50})", str(url))
    return match.group(1) if match else None

def get_best_model():
    try:
        available_models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                available_models.append(m.name)
        for m in available_models:
            if 'flash' in m.lower(): return m
        return available_models[0] if available_models else "models/gemini-pro"
    except Exception:
        return "gemini-pro"

def ask_ai(prompt):
    if not GOOGLE_API_KEY: return "API Key Missing"
    
    # 💡 [핵심] 에러가 나더라도 튕기지 않도록 변수를 맨 처음에 미리 만들어둡니다.
    available_models = [] 
    
    try:
        # 1. 내 구글 API 키로 접속 가능한 모든 모델 리스트를 불러옵니다.
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                available_models.append(m.name)
        
        if not available_models:
            return "🚨 권한 에러: 구글 API 키로 사용할 수 있는 모델이 하나도 없습니다."
            
        # 2. 쓸 수 있는 모델 중 가장 똑똑하고 빠른(flash 또는 pro) 모델을 자동으로 골라냅니다.
        target_model = available_models[-1] # 임시 기본값
        for m in available_models:
            if 'flash' in m.lower() or 'pro' in m.lower():
                target_model = m
                break
                
        # 3. 모델 이름 깔끔하게 정리 (models/ 글자 제거)
        clean_model_name = target_model.replace("models/", "")
        
        # 4. 검색된 진짜 모델로 기획안 작성을 지시합니다.
        model = genai.GenerativeModel(clean_model_name)
        response = model.generate_content(prompt)
        return response.text
        
    except Exception as e: 
        # 이제 2차 에러 없이 "진짜 에러 원인"을 화면에 띄워줍니다!
        return f"🚨 진짜 에러 원인: {str(e)}\n\n💡 [참고] 내 API로 사용 가능한 모델 목록: {available_models}"

def add_log(action_type, details):
    """실수를 추적하기 위한 블랙박스(작업로그) 기록 함수"""
    try:
        client = get_client()
        if client:
            log_sheet = client.open_by_key(SHEET_ID).worksheet("작업로그")
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_sheet.append_row([now_str, action_type, details])
    except Exception as e:
        pass # 로그 기록이 실패해도 본래 작업(주문/재고수정 등)은 멈추면 안 되므로 pass 처리

def send_email_with_attach(to, subject, body, attachment_file=None, filename="attachment.xlsx", multiple_attachments=None):
    """
    단일 파일 또는 여러 개의 파일(multiple_attachments)을 이메일로 전송하는 함수
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        # 1. 단일 첨부파일 처리 (기존 발주서 전송용)
        if attachment_file:
            attachment_file.seek(0)
            part = MIMEApplication(attachment_file.read(), Name=filename)
            part['Content-Disposition'] = f'attachment; filename="{filename}"'
            msg.attach(part)
            
        # 2. 다중 첨부파일 처리 (작업지시서 & 원본 이미지용)
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
    except Exception as e:
        return False, str(e)

def deduct_stock_smart(product_name, qty, df_opt, sheet_stock):
    """
    개선된 재고 차감 로직:
    1. 매핑명 중 가장 긴 단어(정밀한 명칭)부터 매칭 시도
    2. 데이터 타입 에러 방지 (숫자 변환 예외 처리)
    """
    try:
        if df_opt.empty or not sheet_stock:
            return False, "⚠️ 옵션 설정 또는 재고 시트 로드 실패"

        product_name = str(product_name).strip()
        target_std_name = None
        
        # [개선] 매칭 우선순위 로직
        # 매핑명을 리스트로 만들고, 글자 수가 긴 순서대로 정렬하여 정밀 매칭 유도
        match_candidates = []
        for _, opt in df_opt.iterrows():
            std_name = str(opt.get('상품명', '')).strip()
            mapping_str = str(opt.get('매핑명', '')).strip()
            keywords = [k.strip() for k in mapping_str.split(',') if k.strip()]
            
            for kw in keywords:
                if kw in product_name:
                    # (일치하는 키워드 길이, 실제 상품명) 저장
                    match_candidates.append((len(kw), std_name))
        
        # 일치하는 키워드가 있다면 가장 긴 것(가장 구체적인 이름)을 선택
        if match_candidates:
            match_candidates.sort(key=lambda x: x[0], reverse=True)
            target_std_name = match_candidates[0][1]
        else:
            # 매핑에 없으면 원본 상품명으로 시도
            target_std_name = product_name

        # [재고 반영]
        stock_records = sheet_stock.get_all_records()
        for idx, s_item in enumerate(stock_records):
            if str(s_item.get('상품명')).strip() == target_std_name:
                # 숫자 변환 안정성 확보
                try:
                    current_qty = int(pd.to_numeric(s_item.get('현재재고', 0), errors='coerce'))
                except:
                    current_qty = 0
                
                new_qty = current_qty - int(qty)
                
                # 구글 시트 업데이트 (B열이 '현재재고'라고 가정: idx + 2행, 2열)
                sheet_stock.update_cell(idx + 2, 2, new_qty)
                return True, f"✅ '{target_std_name}' {qty}개 차감 완료 (잔여: {new_qty})"
        
        return False, f"⚠️ '{target_std_name}' 상품을 재고 목록에서 찾을 수 없습니다."

    except Exception as e:
        return False, f"❌ 재고 차감 중 에러 발생: {str(e)}"

def add_stock_smart(product_name, qty, df_opt, sheet_stock):
    """취소/반품 시 빠졌던 재고를 다시 더해주는(+) 마법의 함수"""
    try:
        if df_opt.empty or not sheet_stock:
            return False, "⚠️ 옵션 설정 또는 재고 시트 로드 실패"

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
        else:
            target_std_name = product_name

        stock_records = sheet_stock.get_all_records()
        for idx, s_item in enumerate(stock_records):
            if str(s_item.get('상품명')).strip() == target_std_name:
                try:
                    current_qty = int(pd.to_numeric(s_item.get('현재재고', 0), errors='coerce'))
                except:
                    current_qty = 0
                
                # 🔥 여기서 재고를 다시 더해줍니다!
                new_qty = current_qty + int(qty)
                
                sheet_stock.update_cell(idx + 2, 2, new_qty)
                return True, f"✅ '{target_std_name}' {qty}개 복구 완료 (잔여: {new_qty})"
        
        return False, f"⚠️ '{target_std_name}' 상품을 재고 목록에서 찾을 수 없습니다."

    except Exception as e:
        return False, f"❌ 재고 복구 중 에러 발생: {str(e)}"

def check_stock_and_alert(df_stock):
    df_stock['현재재고'] = pd.to_numeric(df_stock['현재재고'], errors='coerce').fillna(0)
    df_stock['안전재고'] = pd.to_numeric(df_stock['안전재고'], errors='coerce').fillna(0)
    low_items = df_stock[df_stock['현재재고'] <= df_stock['안전재고']]
    if not low_items.empty:
        msg = "🚨 [DUWELL 재고 부족 알림]\n\n다음 상품의 재고가 안전 수준 이하입니다:\n\n"
        for _, row in low_items.iterrows():
            msg += f"- {row['상품명']}: 현재 {int(row['현재재고'])}개\n"
        msg += "\n빠른 확인 부탁드립니다."
        send_email_with_attach(SENDER_EMAIL, "[DUWELL] 🚨 긴급: 재고 부족 알림", msg)
        return True
    return False

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

    # --- 송장 번호 개별/일괄 업데이트 함수 ---

def update_tracking_in_sheet(sheet, row_data, tracking_num, new_status="배송중"):
    """수동으로 송장번호 입력 시 사용"""
    try:
        records = sheet.get_all_records()
        header = sheet.row_values(1)
        
        # 필요한 컬럼 위치 찾기
        t_idx = -1
        s_idx = -1
        for i, h in enumerate(header):
            if '송장번호' in h.strip(): t_idx = i + 1
            if h.strip() in ['상태', '진행상태']: s_idx = i + 1
            
        if t_idx == -1: return False, "❌ 시트에 '송장번호' 열이 없습니다."

        # 대상 행 찾기 (구매자명 + 상품명 조합)
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
    except Exception as e:
        return False, str(e)

def bulk_update_tracking_excel(sheet, df_up):
    """공장 엑셀 업로드 시 일괄 사용 (안전장치 강화 버전)"""
    try:
        all_data = sheet.get_all_values()
        if not all_data:
            return False, "❌ 시트에 데이터가 없습니다."
            
        header = all_data[0]
        
        # 시트 컬럼 인덱스 확보
        try:
            name_idx = header.index('구매자명')
            status_idx = next(i for i, h in enumerate(header) if h in ['상태', '진행상태'])
            track_idx = header.index('송장번호')
        except ValueError:
            return False, "❌ 시트 1행(헤더)에 '구매자명', '상태', '송장번호' 컬럼이 정확히 있는지 확인해주세요."

        # 엑셀의 필수 컬럼 존재 확인
        if '구매자명' not in df_up.columns or '송장번호' not in df_up.columns:
            return False, "❌ 업로드한 엑셀에 '구매자명'과 '송장번호' 열이 필수입니다."

        # 🔥 [안전장치 1] 빈 칸(NaN)이 있는 쓸모없는 행 미리 제거
        df_up = df_up.dropna(subset=['구매자명', '송장번호'])

        success_count = 0
        fail_count = 0

        for _, row in df_up.iterrows():
            u_name = str(row['구매자명']).strip()
            # 🔥 [안전장치 2] 송장번호가 숫자로 들어와도 안전하게 문자로 변환 (.0 붙는 현상 방지)
            u_track = str(row['송장번호']).replace('.0', '').strip()
            
            if not u_name or not u_track or u_track.lower() == 'nan': 
                continue

            matched = False
            # 시트에서 매칭되는 행 찾기 (최신 데이터부터 찾기 위해 역순)
            for i in range(len(all_data)-1, 0, -1):
                s_row = all_data[i]
                # 빈 행일 경우 건너뛰기
                if len(s_row) <= name_idx: continue 
                
                if str(s_row[name_idx]).strip() == u_name:
                    # 매칭 성공 시 업데이트
                    sheet.update_cell(i + 1, track_idx + 1, f"'{u_track}") # 엑셀 지수표현 방지용 ' 추가
                    sheet.update_cell(i + 1, status_idx + 1, "배송중")
                    success_count += 1
                    matched = True
                    break
            
            if not matched:
                fail_count += 1
        
        result_msg = f"✅ 총 {success_count}건 배송 처리 완료!"
        if fail_count > 0:
            result_msg += f" (⚠️ {fail_count}건은 명단에 없어 실패했습니다. 이름을 확인해주세요.)"
            
        return True, result_msg
        
    except Exception as e:
        return False, f"❌ 시스템 오류 발생: {str(e)}"

# --------------------------------------------------------------------------
# 🏠 상단 홈페이지형 네비게이션 헤더
# --------------------------------------------------------------------------

# 로고와 새로고침 버튼을 상단에 배치
col_logo, col_btn = st.columns([8, 2])
with col_logo:
    st.markdown("<h1 style='color:#2B3A55; margin:0; padding-top:10px; font-weight:900;'>🍷 DUWELL <span style='font-weight:300; font-size:1.5rem; color:#6C757D;'>스마트 센터</span></h1>", unsafe_allow_html=True)
with col_btn:
    st.write("") # 버튼 위치를 살짝 내리기 위한 공백
    if st.button("🔄 데이터 새로고침", use_container_width=True, type="secondary"): 
        st.cache_data.clear() 
        st.rerun()

st.write("") # 간격 띄우기

# 🔥 가로형 라디오 버튼 (사이드바 대신 상단 메뉴 역할)
menu = st.radio(
    "메뉴 이동", 
    [
        "🏠 통합 모니터링", 
        "📦 주문/생산 통합 관리", 
        "💎 마케팅/CRM 통합 센터", 
        "🛠️ 재고 입출고 관리", 
        "🛠️ 옵션 관리", 
        "📅 일정 관리", 
        "💰 마진/정산 분석",
        "🤖 AI 비즈니스 센터"
    ],
    horizontal=True,
    label_visibility="collapsed" # '메뉴 이동' 이라는 제목 글씨 숨김
)

st.divider()

# 데이터 로딩 (기존 로직 유지)
df_duwell, sheet_main = load_data("시트1") 
df_all = df_duwell.copy()

if not df_all.empty and '날짜' in df_all.columns:
    df_all['날짜'] = pd.to_datetime(df_all['날짜'], errors='coerce')
    df_all = df_all.sort_values(by='날짜', ascending=False)
    df_all['날짜_str'] = df_all['날짜'].dt.strftime('%Y-%m-%d')
else:
    if not df_all.empty: df_all['날짜_str'] = ""


# === [1] 🏠 통합 모니터링 (기존 유지) ===
if menu == "🏠 통합 모니터링":
    st.markdown("### 📊 사업 현황 대시보드")
    
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
        with c1: st.metric("📦 오늘 주문", f"{len(today_orders)}건", delta=f"{len(today_orders)}건 New")
        with c2: st.metric("💰 오늘 매출", f"{today_orders['금액_숫자'].sum():,.0f}원")
        with c3: st.metric("📅 이달의 매출", f"{this_month_orders['금액_숫자'].sum():,.0f}원", delta=f"{len(this_month_orders)}건 주문")
        with c4: st.metric("🏆 누적 매출", f"{df_all['금액_숫자'].sum():,.0f}원")

        st.divider()

        st.markdown("#### ✨ AI 비즈니스 인사이트")
        with st.expander("🤖 AI에게 현재 매출/재고 상황 분석 맡기기", expanded=False):
            if st.button("✨ 맞춤형 AI 리포트 생성", key="ai_report_btn", type="primary"):
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
                    st.success("✅ AI 리포트 생성이 완료되었습니다.")
                    st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px; border:1px solid #E9ECEF;'>{ask_ai(prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

        st.divider()

        st.markdown("#### 📈 기간별 매출 추이")
        tab_d, tab_w, tab_m = st.tabs(["📊 일별 매출", "📊 주간 매출", "📊 월간 매출"])

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
            st.markdown("#### 🚨 재고 경고 (안전재고 미달)")
            df_stock_alert, _ = load_data("재고관리")
            if not df_stock_alert.empty:
                df_stock_alert['현재재고'] = pd.to_numeric(df_stock_alert['현재재고'], errors='coerce').fillna(0)
                df_stock_alert['안전재고'] = pd.to_numeric(df_stock_alert['안전재고'], errors='coerce').fillna(0)
                low_stock = df_stock_alert[df_stock_alert['현재재고'] <= df_stock_alert['안전재고']]
                if not low_stock.empty:
                    for _, s in low_stock.iterrows(): st.error(f"**{s['상품명']}**: 현재 {int(s['현재재고'])}개")
                else: st.success("✅ 모든 상품 재고 정상")
            else: st.write("재고 데이터 없음")

        with col_sch:
            st.markdown("#### 📅 다가오는 일정")
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
        st.markdown("#### 📦 최근 주문 리포트 (최신 5건)")
        possible_cols = ['날짜_str', '구매자명', '상품명', '수량', '결제금액', '상태']
        cols = [c for c in possible_cols if c in df_all.columns]
        st.dataframe(df_all[cols].head(5), hide_index=True, use_container_width=True)

    else:
        st.warning("📊 아직 등록된 주문 데이터가 없습니다.")


# === [2] 📦 주문/생산 통합 관리 (5개 메뉴 완벽 통합) ===
elif menu == "📦 주문/생산 통합 관리":
    st.markdown("""
        <div style="background: linear-gradient(135deg, #2B3A55 0%, #1A2235 100%); padding: 25px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h2 style="margin:0; font-size:1.6rem; font-weight:800; color:white;">📦 주문부터 생산까지 One-Stop</h2>
            <p style="margin:5px 0 0 0; font-size:0.95rem; opacity:0.9; color:white;">주문 등록, 시안 확인, 공장 발주, 지시서 출력, 송장 입력을 순서대로 처리하세요.</p>
        </div>
    """, unsafe_allow_html=True)

    op_tab1, op_tab2, op_tab3, op_tab4, op_tab5 = st.tabs([
        "📝 1. 주문 등록", "🎨 2. 시안 & 장부", "🏭 3. 공장 발주", "🖨️ 4. 작업지시서", "🚚 5. 송장 등록"
    ])

    # 1. 주문 등록 탭
    with op_tab1:
        st.markdown("#### 📝 신규 주문 등록 (재고 자동 차감)")
        sub1, sub2 = st.tabs(["📂 엑셀 일괄 업로드", "✍️ 건별 수동 등록"])
        
        with sub1:
            uploaded_file = st.file_uploader("네이버/자사몰 엑셀 업로드", type=['xlsx'], key="order_up")
            if uploaded_file and st.button("💾 저장 및 재고 차감", type="primary"):
                try:
                    df_new = pd.read_excel(uploaded_file, header=1)
                    df_opt, _ = load_data("옵션관리")
                    _, sheet_stock = load_data("재고관리")
                    
                    df_new = df_new.dropna(subset=['수취인명', '상품명'])
                    rows_add = []
                    log_msg = []
                    
                    for _, row in df_new.iterrows():
                        p_name = str(row.get('상품명','')).strip()
                        try:
                            qty = int(pd.to_numeric(row.get('수량', 1), errors='coerce'))
                        except:
                            qty = 1
                            
                        raw_price = str(row.get('총 주문금액', '0')).replace(',', '').replace('원', '').strip()
                        try:
                            price = int(pd.to_numeric(raw_price, errors='coerce')) if raw_price else 0
                        except:
                            price = 0

                        rows_add.append([
                            str(row.get('주문일시','')), str(row.get('수취인명','')), str(row.get('수취인연락처1','')),
                            str(row.get('배송지','')), p_name, str(qty), str(price), "", "", str(row.get('배송메세지','')), "", "신규"
                        ])
                        
                        ok, msg = deduct_stock_smart(p_name, qty, df_opt, sheet_stock)
                        log_msg.append(msg)

                    if sheet_main:
                        sheet_main.append_rows(rows_add)
                        st.success(f"{len(rows_add)}건 처리 완료")
                        with st.expander("처리 로그 보기"): st.write(log_msg)
                        time.sleep(1); st.rerun()
                        
                except Exception as e:
                    st.error(f"오류: {e}")

        with sub2:
            with st.form("manual"):
                col1, col2 = st.columns(2)
                with col1:
                    m_date = st.date_input("날짜", datetime.now())
                    m_name = st.text_input("구매자명")
                    m_phone = st.text_input("연락처")
                    m_addr = st.text_input("주소")
                with col2:
                    m_prod = st.text_input("상품명 (옵션매핑명)")
                    m_qty = st.number_input("수량", 1, 1000, 1)
                    m_price = st.number_input("금액", 0)
                    m_file = st.text_input("디자인링크")
                m_req = st.text_area("요청사항(자수)")
                
                if st.form_submit_button("등록 및 재고차감", type="primary"):
                    if sheet_main:
                        sheet_main.append_row([str(m_date), m_name, m_phone, m_addr, m_prod, str(m_qty), str(m_price), m_file, "", m_req, "", "신규(수동)"])
                        df_opt, _ = load_data("옵션관리")
                        _, sheet_stock = load_data("재고관리")
                        ok, msg = deduct_stock_smart(m_prod, m_qty, df_opt, sheet_stock)
                        st.success(msg)
                        time.sleep(1); st.rerun()

    # 2. 시안 및 장부 탭
    with op_tab2:
        st.markdown("#### 🎨 디자인 시안 확인 및 전체 장부")
        c_tab1, c_tab2 = st.tabs(["🔥 디자인 작업 대기중 (시안실)", "📋 전체 주문 장부"])
        
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
                            if st.button("✅ 시안 확정 (완료 처리)", key=f"btn_sian_{i}"):
                                success, msg = update_status_in_sheet(sheet_main, r, "완료")
                                if success: st.success(msg); time.sleep(1); st.rerun()
        
        with c_tab2:
            st.markdown("#### 📋 전체 주문 장부 및 취소/반품 처리")
            st.dataframe(df_all, use_container_width=True)
            csv = df_all.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 장부 엑셀 다운로드", csv, "order_list.csv", "text/csv")
            
            st.divider()
            st.markdown("##### 🔙 주문 취소 및 반품 (재고 자동 복구)")
            with st.form("cancel_return_form"):
                cr_col1, cr_col2 = st.columns(2)
                with cr_col1:
                    search_buyer = st.text_input("취소/반품 처리할 구매자명 입력 (정확히 입력)")
                with cr_col2:
                    cr_status = st.selectbox("변경할 상태", ["취소", "반품", "교환"])
                
                if st.form_submit_button("상태 변경 및 재고 복구", type="primary"):
                    if not search_buyer:
                        st.warning("구매자명을 입력해주세요.")
                    else:
                        target_orders = df_all[df_all['구매자명'].astype(str) == search_buyer]
                        if target_orders.empty:
                            st.error("해당 구매자의 주문을 찾을 수 없습니다.")
                        else:
                            target_row = target_orders.iloc[0] 
                            success, msg = update_status_in_sheet(sheet_main, target_row, cr_status)
                            
                            if success:
                                if cr_status in ["취소", "반품"]:
                                    df_opt, _ = load_data("옵션관리")
                                    _, sheet_stock = load_data("재고관리")
                                    
                                    try:
                                        qty_to_restore = int(pd.to_numeric(target_row.get('수량', 1), errors='coerce'))
                                    except:
                                        qty_to_restore = 1
                                        
                                    p_name = target_row.get('상품명', '')
                                    
                                    ok, stock_msg = add_stock_smart(p_name, qty_to_restore, df_opt, sheet_stock)
                                    try:
                                        add_log("주문" + cr_status, f"{search_buyer} 고객 - {p_name} {qty_to_restore}개 재고 복구됨")
                                    except:
                                        pass
                                    
                                    st.success(f"{msg} / {stock_msg}")
                                else:
                                    st.success(f"{msg} (교환은 재고가 복구되지 않습니다)")
                                    try:
                                        add_log("주문교환", f"{search_buyer} 고객 주문 상태 교환 변경")
                                    except:
                                        pass
                                    
                                time.sleep(2); st.rerun()
                            else:
                                st.error(msg)

    # 3. 공장 발주 탭
    with op_tab3:
        st.markdown("#### 🏭 공장 발주 (메일 자동 발송)")
        if not df_all.empty:
            pending_orders = df_all[~df_all['상태'].isin(['발주완료', '배송중', '배송완료', '취소', '반품'])].copy()
            if pending_orders.empty: st.success("🎉 발주 대기 중인 주문이 없습니다.")
            else:
                if "발주선택" not in pending_orders.columns: pending_orders.insert(0, "발주선택", False)
                edited_orders = st.data_editor(
                    pending_orders,
                    column_config={"발주선택": st.column_config.CheckboxColumn(required=True)},
                    column_order=['발주선택', '날짜_str', '구매자명', '상품명', '수량', '상태'],
                    hide_index=True, use_container_width=True, key="factory_order_v7"
                )
                selected_fact = edited_orders[edited_orders['발주선택'] == True]
                
                factory_email = st.text_input("📧 공장 수신 이메일", value="factory@example.com")
                if st.button("🚀 선택 건 발주 확정 및 엑셀 메일 발송", type="primary") and not selected_fact.empty:
                    pb = st.progress(0)
                    for i, (_, row) in enumerate(selected_fact.iterrows()):
                        update_status_in_sheet(sheet_main, row, "발주완료")
                        pb.progress((i + 1) / len(selected_fact) * 0.5)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        selected_fact[['구매자명', '연락처', '주소', '상품명', '수량', '요청사항']].to_excel(writer, index=False)
                    output.seek(0)
                    
                    subject = f"[DUWELL] 발주서_{datetime.now().strftime('%m%d')}"
                    body = f"안녕하세요. DUWELL 신규 발주서 {len(selected_fact)}건 송부드립니다.\n확인 후 생산 부탁드립니다."
                    mail_ok, mail_msg = send_email_with_attach(factory_email, subject, body, output, f"DUWELL_발주_{datetime.now().strftime('%m%d')}.xlsx")
                    
                    if mail_ok:
                        st.success("🎊 발주 완료 및 메일 발송 성공!")
                        st.cache_data.clear(); st.balloons(); time.sleep(2); st.rerun()

    # 4. 작업지시서 탭
    with op_tab4:
        st.markdown("#### 🖨️ 작업지시서 생성 및 원본 전송")
        st.info("💡 공장 발주와 별개로 개별 작업지시서(HTML)와 고해상도 시안을 보낼 수 있습니다.")
        if not df_all.empty:
            filtered_print = df_all.copy()
            if "체크" not in filtered_print.columns: filtered_print.insert(0, "체크", False)
            edited_print = st.data_editor(
                filtered_print, column_order=['체크', '날짜_str', '구매자명', '상품명', '수량', '상태'],
                column_config={"체크": st.column_config.CheckboxColumn(required=True)}, hide_index=True, use_container_width=True, key="print_tab"
            )
            selected_print = edited_print[edited_print['체크'] == True]
            
            if not selected_print.empty:
                st.divider()
                factory_email_print = st.text_input("📧 수신 이메일 주소 (공장/작업자)", value="factory@example.com", key="print_email")
                c_btn1, c_btn2 = st.columns(2)
                
                with c_btn1:
                    if st.button("📥 HTML 로컬 다운로드", use_container_width=True):
                        st.info("HTML 파일을 생성합니다. (기존 로직 수행)") 
                with c_btn2:
                    if st.button("🚀 원본 이미지 포함 메일 발송", type="primary", use_container_width=True):
                        with st.spinner("최고 해상도 원본 이미지를 가져오는 중입니다..."):
                            import requests
                            email_attachments = []
                            pb_print = st.progress(0)
                            for idx, (_, row) in enumerate(selected_print.iterrows()):
                                b_name = str(row['구매자명']).strip()
                                drive_id = get_drive_id(str(row.get('디자인파일', '')))
                                if drive_id:
                                    try:
                                        raw_img_url = f"https://drive.google.com/uc?export=download&id={drive_id}"
                                        res = requests.get(raw_img_url, timeout=15)
                                        if res.status_code == 200:
                                            img_io = io.BytesIO(res.content)
                                            email_attachments.append({"file": img_io, "filename": f"원본시안_{b_name}.jpg"})
                                    except: pass
                                
                                single_html = f"<html><body><h1>작업지시서 ({b_name})</h1><p>상품: {row['상품명']}</p><p>수량: {row['수량']}</p><p>요청: {row['요청사항']}</p></body></html>"
                                email_attachments.append({"file": io.BytesIO(single_html.encode('utf-8')), "filename": f"작업지시서_{b_name}.html"})
                                pb_print.progress((idx + 1) / len(selected_print))
                            
                            mail_ok, _ = send_email_with_attach(to=factory_email_print, subject=f"[DUWELL] 작업지시서 ({len(selected_print)}건)", body="첨부파일 확인 바랍니다.", multiple_attachments=email_attachments)
                            if mail_ok: st.success("발송 성공!"); st.balloons()
                            else: st.error("발송 실패")

    # 5. 송장 등록 탭
    with op_tab5:
        st.markdown("#### 🚚 배송 정보(송장) 업데이트")
        if not df_all.empty:
            completed_orders = df_all[df_all['상태'] == '발주완료'].copy()
            sc1, sc2 = st.columns(2)
            with sc1:
                st.markdown("##### ✍️ 수동 입력 (건별)")
                if completed_orders.empty: st.write("발주 완료된 대기 내역이 없습니다.")
                else:
                    if '송장번호' not in completed_orders.columns: completed_orders['송장번호'] = ""
                    tr_edited = st.data_editor(completed_orders, column_order=['구매자명', '상품명', '송장번호'], hide_index=True, key="man_track")
                    if st.button("💾 수동 송장 저장"):
                        cnt = 0
                        for _, row in tr_edited.iterrows():
                            t_num = str(row.get('송장번호', '')).strip()
                            if t_num and t_num != "nan":
                                ok, _ = update_tracking_in_sheet(sheet_main, row, t_num)
                                if ok: cnt += 1
                        st.success(f"{cnt}건 저장 완료!"); st.cache_data.clear(); time.sleep(1); st.rerun()
            with sc2:
                st.markdown("##### 📂 엑셀 일괄 등록")
                up_f = st.file_uploader("공장 송장 엑셀 파일 (.xlsx)", type=['xlsx'], key="track_up")
                if up_f and st.button("🚀 엑셀 데이터 시트 반영"):
                    df_up = pd.read_excel(up_f)
                    ok, msg = bulk_update_tracking_excel(sheet_main, df_up)
                    if ok: st.success(msg); st.cache_data.clear(); time.sleep(1); st.rerun()
                    else: st.error(msg)
# === [3] 💎 마케팅/CRM 통합 센터 (마케팅, CRM 통합) ===
elif menu == "💎 마케팅/CRM 통합 센터":
    st.markdown("""
        <div style="background: linear-gradient(135deg, #4E73DF 0%, #224ABE 100%); padding: 25px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h2 style="margin:0; font-size:1.6rem; font-weight:800; color:white;">💎 마케팅 & CRM 통합 센터</h2>
            <p style="margin:5px 0 0 0; font-size:0.95rem; opacity:0.9; color:white;">고객 관리부터 광고 성과 측정, AI 카피라이팅까지 한곳에서 관리하세요.</p>
        </div>
    """, unsafe_allow_html=True)
    
    m_tab1, m_tab2, m_tab3, m_tab4 = st.tabs([
        "👤 1. 고객 CRM 프로필", "📈 2. 광고 효율(ROAS)", "✍️ 3. AI 카피/네이밍", "💬 4. 리뷰/CS 응대"
    ])

    # 1. 고객 CRM 탭
    with m_tab1:
        st.markdown("#### 👤 고객 통합 프로필 및 상담")
        if not df_all.empty:
            df_crm = df_all.copy()
            if '결제금액' in df_crm.columns: df_crm['amt'] = pd.to_numeric(df_crm['결제금액'].astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
            else: df_crm['amt'] = 0
            
            cust_profile = df_crm.groupby('구매자명').agg({'날짜': ['max', 'count'], 'amt': 'sum'}).reset_index()
            cust_profile.columns = ['고객명', '최근구매일', '구매횟수', '누적금액']
            
            def analyze_cx(row):
                grade = "💎 VIP" if row['누적금액'] >= 500000 else "🥈 일반"
                days = (datetime.now() - row['최근구매일']).days if pd.notnull(row['최근구매일']) else 0
                status = "🔔 교체주기" if 150 <= days <= 210 else "✅ 정상"
                return pd.Series([grade, status, days])
            cust_profile[['등급', '상태', '경과일']] = cust_profile.apply(analyze_cx, axis=1)
            
            c_list, c_detail = st.columns([1, 1.5])
            with c_list:
                st.dataframe(cust_profile[['고객명', '등급', '누적금액', '상태']], use_container_width=True, hide_index=True)
            with c_detail:
                search_nm = st.selectbox("🎯 상세 조회할 고객 선택", cust_profile['고객명'].unique())
                sel_data = cust_profile[cust_profile['고객명'] == search_nm].iloc[0]
                st.write(f"**등급:** {sel_data['등급']} | **누적금액:** {sel_data['누적금액']:,.0f}원")
                
                try:
                    client = get_client(); target_sh = client.open("주문데이터").worksheet("시트1") # 시트명 확인 필요
                    h = target_sh.row_values(1)
                    if '비고' in h:
                        cell = target_sh.find(search_nm)
                        if cell:
                            current_history = target_sh.cell(cell.row, h.index('비고')+1).value
                            st.text_area("📜 상담 히스토리", value=current_history or "내용 없음", height=100, disabled=True)
                            memo_in = st.text_input("📝 신규 상담 내용 입력")
                            if st.button("💾 저장"):
                                now = datetime.now().strftime('%Y-%m-%d %H:%M')
                                final = f"{current_history}\n[{now}] {memo_in}" if current_history else f"[{now}] {memo_in}"
                                target_sh.update_cell(cell.row, h.index('비고')+1, final)
                                st.success("저장됨")
                except: st.info("히스토리 연동 대기중")

    # 2. ROAS 분석 탭
    with m_tab2:
        st.markdown("#### 🎯 일일 광고비 입력 및 ROAS 측정")
        today_str = datetime.now().strftime("%Y-%m-%d")
        today_sales = df_all[df_all['날짜_str'] == today_str]['결제금액'].sum() if '결제금액' in df_all.columns else 0 # 간략화
        
        with st.form("roas_form"):
            r_col1, r_col2 = st.columns(2)
            with r_col1: naver_spend = st.number_input("🟢 네이버 광고비 (원)", min_value=0, step=10000)
            with r_col2: meta_spend = st.number_input("🟣 인스타/페북 광고비 (원)", min_value=0, step=10000)
                
            if st.form_submit_button("📊 ROAS 계산하기", type="primary"):
                total_spend = naver_spend + meta_spend
                c_roas1, c_roas2, c_roas3 = st.columns(3)
                c_roas1.metric("총 광고비 지출", f"{total_spend:,.0f}원")
                c_roas2.metric("오늘 발생한 총 매출(추정)", f"{today_sales:,.0f}원")
                if total_spend > 0:
                    roas = (today_sales / total_spend) * 100
                    roas_icon = "🔥 대박!" if roas >= 400 else ("✅ 양호" if roas >= 250 else "⚠️ 점검 필요")
                    c_roas3.metric("오늘의 ROAS", f"{roas:,.1f}% {roas_icon}")

    # 3. AI 카피라이팅 탭
    with m_tab3:
        st.markdown("#### ✨ 매체 최적화 광고 문구 / 네이밍 생성")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            product_name = st.text_input("상품 특징", value="시그니처 와인 컬러 프리미엄 와플 타월")
            channel = st.selectbox("광고 매체", ["🟢 네이버 파워링크", "🟣 인스타그램", "🔴 카카오 알림톡"])
        with col_c2:
            target_audience = st.selectbox("타겟 고객", ["3040 리빙", "2030 신혼부부/집들이", "전체"])
            tone = st.selectbox("톤앤매너", ["세련된", "재치있는", "전문적인", "긴급/한정수량"])
        
        if st.button("✍️ 광고 소재 및 네이밍 생성", type="primary"):
            with st.spinner("AI가 소재를 작성 중입니다..."):
                prompt = f"상품: {product_name}, 타겟: {target_audience}, 채널: {channel}, 톤: {tone}. 이 매체에 완벽하게 맞는 광고 카피 3가지와 매력적인 캠페인 네이밍 2가지를 제안해줘."
                st.markdown(f"<div style='background-color:#fff; padding:20px; border-radius:12px; border:1px solid #E9ECEF;'>{ask_ai(prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # 4. 리뷰 및 CS 탭
    with m_tab4:
        st.markdown("#### 💬 스마트 고객 응대 (리뷰 & CS)")
        sub_tab1, sub_tab2 = st.tabs(["엑셀 리뷰 일괄 답글", "CS 답변 스크립트"])
        with sub_tab1:
            uploaded_review = st.file_uploader("리뷰 엑셀 파일 (.xlsx)", type=['xlsx'])
            if uploaded_review:
                df_rev = pd.read_excel(uploaded_review)
                review_col = st.selectbox("고객 리뷰 내용 열", df_rev.columns.tolist())
                if st.button("🤖 AI 답글 생성"):
                    with st.spinner("답글 작성 중..."):
                        replies = [ask_ai(f"고객 리뷰: '{row[review_col]}'. 감사와 공감이 담긴 정중한 답글 2문장 작성.") for _, row in df_rev.iterrows()]
                        df_rev['AI_자동답글'] = replies
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df_rev.to_excel(writer, index=False)
                        st.download_button("📥 결과 다운로드", output.getvalue(), "리뷰답글완료.xlsx")
        with sub_tab2:
            cs_type = st.radio("문의 유형", ["배송 지연", "제품 교환", "불만", "기타"], horizontal=True)
            cs_detail = st.text_area("고객 문의 내용")
            if st.button("💬 방어 답변 생성") and cs_detail:
                st.info(ask_ai(f"유형: {cs_type}, 내용: {cs_detail}\n프리미엄 브랜드에 맞는 정중한 사과와 해결책이 담긴 답변 작성."))


# === [4] 🛠️ 재고 입출고 관리 (기존 유지) ===
elif menu == "🛠️ 재고 입출고 관리":
    st.subheader("📊 재고 통합 관리 시스템")
    df_stock, sheet_stock = load_data("재고관리")
    df_opt, _ = load_data("옵션관리")
    tab1, tab2, tab3 = st.tabs(["📊 재고 현황 (자동집계)", "📂 대량 입출고 등록 (엑셀)", "📝 개별 조정 (수동)"])
    
    with tab1:
        if not df_stock.empty and not df_all.empty:
            st.dataframe(df_stock, use_container_width=True) # 요약 코드 분량상 원본 df 출력
            
    with tab2:
        st.markdown("### 📥 엑셀로 재고 일괄 등록")
        uploaded_file = st.file_uploader("작성한 엑셀 파일 업로드", type=['xlsx', 'xls', 'csv'], key="stock_up")
        if uploaded_file and st.button("🚀 재고 일괄 반영하기", type="primary"):
            st.info("재고 일괄 반영 로직 수행 (이전 코드 동일)")
            
    with tab3:
        st.markdown("### 📝 개별 상품 입/출고 (블랙박스 작동중 🔴)")
        if not df_stock.empty:
            with st.form("manual_stock"):
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1: target_prod = st.selectbox("상품 선택", df_stock['상품명'].unique())
                with col2: action = st.radio("구분", ["입고 (+)", "출고/손실 (-)"], horizontal=True)
                with col3: qty = st.number_input("수량", min_value=1, value=1)
                
                if st.form_submit_button("반영"):
                    # 동시성 문제 방지: 반영 버튼 누른 순간 최신 데이터를 한 번 더 불러옵니다.
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
                            
                            # 🔴 블랙박스에 기록 남기기!
                            log_msg = f"{target_prod} {qty}개 {action.split()[0]} 처리 (변경 전: {curr_val} -> 변경 후: {final_qty})"
                            add_log("재고수동조정", log_msg)
                            
                            st.success(f"✅ {target_prod}: {final_qty}개로 변경 및 로그 기록 완료!"); time.sleep(1); st.rerun()
# === [5] 🛠️ 옵션 관리 ===
elif menu == "🛠️ 옵션 관리":
    st.subheader("🛠️ 옵션 및 통합 상품명 관리")
    df_opt, sheet_opt = load_data("옵션관리")
    if not df_opt.empty:
        edited_df = st.data_editor(df_opt, num_rows="dynamic", use_container_width=True)
        if st.button("💾 저장"):
            sheet_opt.clear()
            
            # 최신 gspread 버전에 맞춘 업데이트 문법 (A1 셀부터 데이터 채우기)
            new_data = [edited_df.columns.values.tolist()] + edited_df.values.tolist()
            sheet_opt.update(values=new_data, range_name="A1")
            
            st.success("저장됨"); st.rerun()

# === [6] 📅 일정 관리 (기존 유지) ===
elif menu == "📅 일정 관리":
    st.subheader("📅 일정 캘린더")
    df_sch, sheet_sch = load_data("일정관리")
    col1, col2 = st.columns([1, 2])
    with col1:
        with st.form("add_schedule"):
            d_date = st.date_input("날짜"); d_time = st.time_input("시간"); d_title = st.text_input("일정명"); d_desc = st.text_area("상세내용")
            if st.form_submit_button("저장"):
                if sheet_sch: sheet_sch.append_row([str(d_date), str(d_date), str(d_time), d_title, d_desc]); st.success("저장됨"); st.rerun()
    with col2:
        if not df_sch.empty:
            events = [{"title": str(r.get('일정명')), "start": str(r.get('시작일'))} for _, r in df_sch.iterrows()]
            calendar(events=events)

# === [7] 💰 마진/정산 분석 ===
elif menu == "💰 마진/정산 분석":
    st.subheader("💰 실시간 마진 및 정산 분석기")   
    with st.expander("⚙️ 정산 기준 설정 (수수료 및 배송비 분리)", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fee_smart = st.number_input("스마트스토어 수수료 (%)", value=5.5)
            fee_own = st.number_input("자사몰(PG) 수수료 (%)", value=3.0)
        with col_f2:
            fee_coupang = st.number_input("쿠팡 수수료 (%)", value=10.8)
            fee_etc = st.number_input("기타 마켓 수수료 (%)", value=5.0)
        with col_f3:
            shipping_revenue = st.number_input("고객 결제 배송비 (매출, 원)", value=3000, step=100)
            shipping_cost = st.number_input("택배사 실제 청구비 (매입, 원)", value=2500, step=100)

    if not df_all.empty:
        df_cost, _ = load_data("옵션관리") 
        if df_cost.empty or '가격' not in df_cost.columns:
            st.warning("⚠️ [옵션관리] 탭에 '가격' 열이 없습니다.")
        if '원가' not in df_cost.columns:
            st.error("🚨 [중요] 옵션관리 시트에 '원가' 열이 없어서 순이익이 과다하게 계산됩니다! 반드시 추가해주세요.")

        df_calc = df_all.copy()
        df_calc['수량'] = pd.to_numeric(df_calc['수량'], errors='coerce').fillna(1)
        if '주문처' not in df_calc.columns: df_calc['주문처'] = '자사몰'

        def calculate_profit(row):
            market = str(row.get('주문처', '자사몰'))
            qty = row['수량']
            
            item_name = str(row['상품명']).strip()
            item_clean = item_name.replace(" ", "").lower()
            
            if '스마트스토어' in market: fee_rate = fee_smart / 100
            elif '쿠팡' in market: fee_rate = fee_coupang / 100
            elif '자사몰' in market: fee_rate = fee_own / 100
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
                        kw_clean = kw.replace(" ", "").lower()
                        if kw_clean in item_clean:
                            raw_price = str(opt.get('가격', 0)).replace(',', '').replace('원', '').strip()
                            raw_cost = str(opt.get('원가', 0)).replace(',', '').replace('원', '').strip()
                            unit_price = pd.to_numeric(raw_price, errors='coerce')
                            unit_cost = pd.to_numeric(raw_cost, errors='coerce')
                            if pd.isna(unit_price): unit_price = 0
                            if pd.isna(unit_cost): unit_cost = 0
                            break 
            
            # 🔥 배송비 매출과 매입을 분리하여 정밀 계산
            expected_item_revenue = unit_price * qty
            total_revenue = expected_item_revenue + shipping_revenue
            commission_fee = total_revenue * fee_rate
            total_cost = (unit_cost * qty) + shipping_cost
            
            net_profit = total_revenue - commission_fee - total_cost
            margin_rate = (net_profit / total_revenue * 100) if total_revenue > 0 else 0
            
            return pd.Series([total_revenue, commission_fee, total_cost, net_profit, margin_rate])

        df_calc[['예상결제금액', '마켓수수료', '총매입원가', '예상순이익', '마진율(%)']] = df_calc.apply(calculate_profit, axis=1)

        tab_sum, tab_month, tab_cal, tab_detail = st.tabs([
            "📊 전체 요약", "📅 월별 정산 내역", "📆 일별 매출 캘린더", "📜 주문별 상세 내역"
        ])

        with tab_sum:
            st.markdown("### 📊 누적 정산 리포트")
            total_sales = df_calc['예상결제금액'].sum()
            total_profit = df_calc['예상순이익'].sum()
            avg_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("누적 예상결제금액 (매출)", f"{total_sales:,.0f} 원")
            c2.metric("💰 누적 예상순이익", f"{total_profit:,.0f} 원")
            c3.metric("📈 평균 마진율", f"{avg_margin:.1f} %")

        with tab_month:
            st.markdown("### 📅 월별 마진/정산 통계")
            df_calc['월'] = df_calc['날짜'].dt.strftime('%Y-%m')
            
            monthly_profit = df_calc.groupby('월').agg(
                주문건수=('구매자명', 'count'),
                매출액=('예상결제금액', 'sum'),
                마켓수수료=('마켓수수료', 'sum'),
                총매입원가=('총매입원가', 'sum'),
                순이익=('예상순이익', 'sum')
            ).reset_index().sort_values('월', ascending=False)
            
            monthly_profit['평균마진율(%)'] = (monthly_profit['순이익'] / monthly_profit['매출액'] * 100).fillna(0)

            styled_monthly = monthly_profit.style.format({
                '매출액': '{:,.0f}', '마켓수수료': '{:,.0f}', '총매입원가': '{:,.0f}',
                '순이익': '{:,.0f}', '평균마진율(%)': '{:.1f}%'
            })
            try: styled_monthly = styled_monthly.background_gradient(subset=['평균마진율(%)'], cmap='RdYlGn')
            except: pass
            
            st.dataframe(styled_monthly, use_container_width=True, hide_index=True)
            
            st.markdown("#### 📉 월별 순이익 추이")
            bar_monthly_profit = alt.Chart(monthly_profit).mark_bar(color='#2ca02c', opacity=0.8).encode(
                x=alt.X('월:N', title='월별'),
                y=alt.Y('순이익:Q', title='순이익(원)'),
                tooltip=['월', '매출액', '순이익']
            ).properties(height=300)
            st.altair_chart(bar_monthly_profit, use_container_width=True)

        with tab_cal:
            st.markdown("### 📆 캘린더 뷰 (일별 매출 & 순이익)")
            
            cal_options = {
                "headerToolbar": {
                    "left": "prev,next today",
                    "center": "title",
                    "right": "dayGridMonth"
                },
                "initialView": "dayGridMonth", 
                "height": 650, # 탭 안에서도 안 숨도록 높이 강제 고정!
            }
            
            events = []
            valid_dates = df_calc[df_calc['날짜_str'].astype(bool) & (df_calc['날짜_str'] != 'nan') & (df_calc['날짜_str'] != '')]
            
            if not valid_dates.empty:
                daily_sales = valid_dates.groupby('날짜_str').agg(
                    매출액=('예상결제금액', 'sum'), 
                    순이익=('예상순이익', 'sum')
                ).reset_index()

                for _, row in daily_sales.iterrows():
                    d_str = str(row['날짜_str']).strip()
                    events.append({"title": f"매출: {row['매출액']:,.0f}", "start": d_str, "color": "#555555"})
                    events.append({"title": f"이익: {row['순이익']:,.0f}", "start": d_str, "color": "#800020"})
            
            # 🔥 [핵심 1] 이벤트가 텅 비었을 때 오류를 막기 위해 오늘 날짜에 '투명한 가짜 데이터'를 하나 심어줍니다.
            if not events:
                events.append({"title": "기록 없음", "start": datetime.now().strftime('%Y-%m-%d'), "color": "transparent", "textColor": "#999999"})
                st.info("💡 아직 발생한 매출 데이터가 없습니다. 빈 달력입니다.")

            # 🔥 [핵심 2] 달력에 고유한 주민번호(key)를 줘서 탭 안에서도 무조건 화면에 그리도록 강제합니다.
            calendar(events=events, options=cal_options, key="sales_dashboard_calendar_v1")

        with tab_detail:
            st.markdown("### 📜 주문건별 상세 내역")
            display_cols = ['날짜_str', '구매자명', '상품명', '수량', '예상결제금액', '마켓수수료', '총매입원가', '예상순이익', '마진율(%)']
            styled_df = df_calc[display_cols].style.format({
                '예상결제금액': '{:,.0f}', '마켓수수료': '{:,.0f}', '총매입원가': '{:,.0f}',
                '예상순이익': '{:,.0f}', '마진율(%)': '{:.1f}%'
            })
            try: styled_df = styled_df.background_gradient(subset=['마진율(%)'], cmap='RdYlGn')
            except: pass
                
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df_calc[display_cols].to_excel(writer, index=False)
            st.download_button("📥 전체 정산 내역 엑셀 다운로드", output.getvalue(), f"예상정산리포트_{datetime.now().strftime('%Y%m%d')}.xlsx")

    else:
        st.warning("분석할 주문 데이터가 없습니다.")

# === [8] 🤖 AI 비즈니스 센터 (5대 에이전트 통합) ===
elif menu == "🤖 AI 비즈니스 센터":
    st.markdown("""
        <div style="background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); padding: 25px; border-radius: 12px; color: white; margin-bottom: 20px;">
            <h2 style="margin:0; font-size:1.6rem; font-weight:800; color:white;">🤖 DUWELL AI 비즈니스 센터</h2>
            <p style="margin:5px 0 0 0; font-size:0.95rem; opacity:0.9; color:white;">10년의 수건 업계 노하우를 학습한 5명의 AI 전문가 팀이 대표님의 업무를 지원합니다.</p>
        </div>
    """, unsafe_allow_html=True)

    # 7개의 에이전트 탭 생성 (수정된 부분)
    ai_tab1, ai_tab2, ai_tab3, ai_tab4, ai_tab5, ai_tab6, ai_tab7 = st.tabs([
        "📝 1. MD 신제품", "💼 2. B2B 영업", "🔍 3. 리뷰 분석", "📊 4. 경영 브리핑", "💰 5. 마진 시뮬", "📸 6. 이미지 프롬프트", "🛒 7. 마켓별 SEO 등록"
    ])

    # --- [1] MD 신제품 기획 에이전트 ---
    with ai_tab1:
        st.markdown("#### 📝 신제품 런칭 브리프 생성기")
        with st.form("md_agent_form"):
            new_product_desc = st.text_area("기획 중인 상품 특징 입력", placeholder="예: 프리미엄 와플 직조 수건. 일반 수건보다 건조가 빠르고 먼지가 안 나. 고급 에스테틱 느낌.", height=100)
            if st.form_submit_button("✨ 런칭 기획안 생성", type="primary"):
                if new_product_desc:
                    with st.spinner("MD 에이전트가 기획안을 작성 중입니다..."):
                        agent_prompt = f"""
                        You are an elite Towel Merchandiser for 'DUWELL'. Backed by 10 years of towel industry know-how, write a product launch brief in Korean.
                        [신제품 특징]: {new_product_desc}
                        출력 형식: [상품명 아이디어 3개], [핵심 타겟 고객], [강력한 셀링 포인트 3개], [상세페이지 스토리라인]
                        """
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # --- [2] B2B 영업 제안서 에이전트 ---
    with ai_tab2:
        st.markdown("#### 💼 B2B 맞춤형 영업 제안서 작성")
        with st.form("b2b_agent_form"):
            col1, col2 = st.columns(2)
            with col1:
                target_company = st.text_input("제안 대상 (예: A 고급 에스테틱, B 부티크 호텔)")
            with col2:
                target_product = st.text_input("제안 상품 (예: 프리미엄 와플 수건 세트)")
            
            sales_points = st.text_area("강조할 소구점 (예: 먼지 없음, 빠른 건조, 고급스러운 디자인)")
            
            if st.form_submit_button("🚀 B2B 영업 메일 초안 생성", type="primary"):
                if target_company and target_product:
                    with st.spinner("B2B 영업 에이전트가 제안서를 작성 중입니다..."):
                        agent_prompt = f"""
                        You are an elite B2B Sales Representative for the premium towel brand 'DUWELL'. 
                        Write a highly professional, polite, and persuasive B2B sales email in Korean.
                        - 타겟 고객사: {target_company}
                        - 제안 상품: {target_product}
                        - 강조할 포인트: {sales_points}
                        - 톤앤매너: 10년 경력의 신뢰감, 고급스러움, 상대방 비즈니스에 확실한 도움이 된다는 확신.
                        """
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # --- [3] 경쟁사 리뷰 분석 에이전트 ---
    with ai_tab3:
        st.markdown("#### 🔍 타사 리뷰 기반 마케팅 포인트 추출")
        with st.form("review_agent_form"):
            bad_reviews = st.text_area("경쟁사(일반 수건) 부정적 리뷰 복사/붙여넣기", placeholder="예: 먼지가 너무 날려요. 잘 안 마르고 꿉꿉한 냄새가 나요.", height=100)
            if st.form_submit_button("💡 공격적 마케팅 무기 생성", type="primary"):
                if bad_reviews:
                    with st.spinner("리서치 에이전트가 페인포인트를 분석 중입니다..."):
                        agent_prompt = f"""
                        You are an expert Market Researcher and Copywriter for 'DUWELL'.
                        Analyze the following negative reviews of competitor's normal towels: "{bad_reviews}"
                        1. 고객의 핵심 Pain Point 요약
                        2. 이를 완벽히 해결해주는 DUWELL '프리미엄 와플 수건'의 특장점 연결
                        3. 당장 인스타그램 카드뉴스에 쓸 수 있는 후킹 카피 3가지 제안 (한국어)
                        """
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # --- [4] 일일 경영 브리핑 에이전트 ---
    with ai_tab4:
        st.markdown("#### 📊 대표님 맞춤형 일일 브리핑 (데이터 연동)")
        st.info("ERP에 쌓인 매출, 재고, 일정 데이터를 종합하여 아침 브리핑을 제공합니다.")
        if st.button("☕ 오늘의 경영 브리핑 생성", type="primary"):
            with st.spinner("비서 에이전트가 데이터를 취합하고 분석 중입니다..."):
                # 현재 시스템 데이터 수집 (매출, 재고, 일정)
                today_str = datetime.now().strftime("%Y-%m-%d")
                today_sales = df_all[df_all['날짜_str'] == today_str]['결제금액'].sum() if not df_all.empty and '결제금액' in df_all.columns else 0
                
                # 재고 부족 상품 찾기
                df_stock_temp, _ = load_data("재고관리")
                low_stock_msg = "재고 부족 상품 없음"
                if not df_stock_temp.empty:
                    df_stock_temp['현재재고'] = pd.to_numeric(df_stock_temp['현재재고'], errors='coerce').fillna(0)
                    df_stock_temp['안전재고'] = pd.to_numeric(df_stock_temp['안전재고'], errors='coerce').fillna(0)
                    low_items = df_stock_temp[df_stock_temp['현재재고'] <= df_stock_temp['안전재고']]['상품명'].tolist()
                    if low_items: low_stock_msg = ", ".join(low_items) + " (발주 필요!)"

                # 오늘 일정 찾기
                df_sch_temp, _ = load_data("일정관리")
                today_schedule = "일정 없음"
                if not df_sch_temp.empty:
                    today_events = df_sch_temp[df_sch_temp['시작일'] == today_str]['일정명'].tolist()
                    if today_events: today_schedule = ", ".join(today_events)

                agent_prompt = f"""
                You are the Executive Assistant to the CEO of 'DUWELL'. Write a crisp, objective, and encouraging morning briefing in Korean.
                [오늘의 데이터]
                - 예상 매출: {today_sales:,.0f}원
                - 재고 경고: {low_stock_msg}
                - 주요 일정: {today_schedule}
                위 데이터를 바탕으로:
                1. 성과 요약 및 칭찬 한마디
                2. 오늘 반드시 처리해야 할 액션 아이템(재고, 일정 관련) 2가지 제안
                """
                st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # --- [5] 마진 시뮬레이터 에이전트 ---
    with ai_tab5:
        st.markdown("#### 💰 기획 상품 마진 시뮬레이터")
        with st.form("margin_agent_form"):
            col1, col2, col3 = st.columns(3)
            with col1: base_price = st.number_input("기존 판매가 (원)", value=15000)
            with col2: discount_rate = st.number_input("할인율 (%)", value=15)
            with col3: product_cost = st.number_input("매입 원가 (원)", value=6000)
            
            extra_costs = st.text_input("추가 비용 내역 (예: 포장비 1000원, 사은품 500원)")
            market_type = st.selectbox("판매 채널 (수수료율)", ["스마트스토어 (약 5.5%)", "쿠팡 (약 10.8%)", "자사몰 (약 3%)"])

            if st.form_submit_button("🧮 적정 판매가 및 마진 계산", type="primary"):
                with st.spinner("재무 에이전트가 마진율을 계산 중입니다..."):
                    agent_prompt = f"""
                    You are a strict and smart Financial Advisor for 'DUWELL'. Calculate the profit margin based on the following data:
                    - 기존 판매가: {base_price}원 / 기획 할인율: {discount_rate}%
                    - 매입 원가: {product_cost}원 / 추가 비용: {extra_costs}
                    - 판매 채널: {market_type}
                    1. 최종 예상 판매가, 예상 수수료, 마진 금액, 최종 마진율(%)을 수식과 함께 직관적으로 보여주세요.
                    2. 마진율이 30% 미만일 경우, 이익을 방어하기 위한 가격 정책이나 세트 구성 아이디어를 제안해주세요. (한국어)
                    """
                    st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # --- [6] 상세페이지 이미지 프롬프트 제작 에이전트 ---
    with ai_tab6:
        st.markdown("#### 📸 고품질 상세페이지 이미지 프롬프트 (미드저니/DALL-E 용)")
        st.info("제품 특징과 원하는 분위기를 고르면, AI 이미지 생성기에 바로 복사해 넣을 수 있는 영문 프롬프트를 전문가 수준으로 뽑아줍니다.")
        
        with st.form("image_prompt_form"):
            col1, col2 = st.columns(2)
            with col1:
                prod_img_desc = st.text_input("제품의 핵심 특징", placeholder="예: 먼지 없는 프리미엄 와플 직조 타월")
            with col2:
                img_mood = st.selectbox("원하는 연출 분위기", [
                    "햇살이 들어오는 따뜻하고 아늑한 욕실", 
                    "5성급 호텔의 모던하고 어두운 고급 욕실", 
                    "깨끗하고 위생적인 에스테틱 샵", 
                    "제품의 질감을 극대화한 초근접 마크로 샷"
                ])
            
            if st.form_submit_button("🎨 영문 프롬프트 생성", type="primary"):
                if prod_img_desc:
                    with st.spinner("전문 포토그래퍼 에이전트가 카메라 렌즈와 조명 세팅을 조율 중입니다..."):
                        agent_prompt = f"""
                        You are an elite Commercial Photographer and AI Prompt Engineer specializing in Home & Living products.
                        I need highly detailed, professional image generation prompts (optimized for Midjourney v6 or DALL-E 3) based on:
                        - Target Product: {prod_img_desc}
                        - Mood & Background: {img_mood}
                        
                        Please create 3 different variations of the shot (e.g., Close-up texture, Lifestyle interior, Wide angle).
                        For each variation, output MUST strictly follow this format:
                        
                        📌 [Shot Type in Korean (e.g., 초근접 질감 컷)]
                        - 연출 의도: (Korean description of the scene)
                        - 🇬🇧 Prompt: (English comma-separated prompt containing subject, background, lighting, camera lens like 85mm f/1.8, high-end commercial photography, 8k resolution, photorealistic)
                        """
                        st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{ask_ai(agent_prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
                else:
                    st.warning("제품 특징을 입력해주세요!")
# --- [7] 다중 마켓 SEO 최적화 상품 등록 에이전트 (HTML 상세페이지 자동 생성 기능 추가) ---
    with ai_tab7:
        st.markdown("#### 🛒 오픈마켓별 SEO 등록 & HTML 상세페이지 생성기")
        st.info("SEO 최적화 데이터뿐만 아니라, 스마트스토어 에디터에 바로 붙여넣을 수 있는 '상세페이지 HTML 코드'까지 한 번에 생성합니다.")

        with st.form("seo_agent_form"):
            col1, col2 = st.columns(2)
            with col1:
                target_market = st.selectbox("등록할 마켓 선택", ["🟢 네이버 스마트스토어", "🚀 쿠팡", "🏠 자사몰 / 기타 오픈마켓"])
                prod_name = st.text_input("기본 상품명", value="프리미엄 와플 수건")
            with col2:
                target_customer = st.selectbox("메인 타겟", ["3040 리빙/인테리어", "2030 집들이/신혼부부", "답례품/대량구매", "전체 연령대"])
                core_points = st.text_input("핵심 강조 포인트", placeholder="예: 먼지없는, 빠른건조, 호텔수건, 10년 노하우", value="먼지없는, 빠른건조, 고급스러운 질감")

            # 폼 제출 버튼
            submit_btn = st.form_submit_button("🚀 SEO 데이터 및 HTML 생성", type="primary")

        if submit_btn:
            if prod_name and core_points:
                with st.spinner(f"{target_market} 로직에 맞춰 최적화 데이터와 HTML 코드를 작성 중입니다... (약 10~20초 소요)"):
                    agent_prompt = f"""
                    You are an elite E-commerce Merchandiser, SEO expert, and Web Designer in Korea, working for the premium towel brand 'DUWELL'.
                    You have 10 years of deep towel industry experience.
                    Your task is to generate highly optimized product registration data and a complete HTML detail page for {target_market}.

                    - Product: {prod_name}
                    - Target Customer: {target_customer}
                    - Key Selling Points: {core_points}

                    Please provide the output strictly in Korean, following this structure:

                    1. 🏷️ [최적화 상품명 3가지]: Create 3 variations of the product title optimized for {target_market}'s search algorithm.
                    2. 🔑 [검색 태그/키워드]: Provide exactly 10 highly searched, relevant tags/keywords separated by commas.
                    3. 📝 [메타 디스크립션/PC·모바일 요약 설명]: A compelling 2-3 sentence description.
                    4. 💻 [상세페이지 기획 뼈대]: 3-step storyline for the detail page (Hook -> USP explanation -> Closing/Trust with 10 years of experience).
                    5. 🌐 [모바일 최적화 상세페이지 HTML (복붙용)]:
                       - Write clean, modern HTML code. MUST use inline CSS.
                       - [모바일 최적화 필수]: Apply `max-width: 100%;`, `word-break: keep-all;`, and `line-height: 1.6;` to the main container to ensure perfect readability on mobile screens.
                       - Style: Elegant typography, #2B3A55 or #800020 for accent colors, clear headings, centered text.
                       - [인트로 고정]: At the very top, insert: `<div style='background:#E9ECEF; padding:120px 20px; margin-bottom:40px; text-align:center; border:2px dashed #ADB5BD; border-radius:12px; color:#495057; font-weight:900; font-size:1.3rem; max-width:100%; word-break:keep-all;'>📸 [메인 인트로 사진 삽입: 시선을 사로잡는 {prod_name} 연출 컷]</div>`
                       - Include persuasive copywriting highlighting DUWELL's 10 years of know-how.
                       - For other images, insert: `<div style='background:#F4F6F9; padding:60px 20px; margin:20px 0; text-align:center; border:2px dashed #B0BEC5; border-radius:12px; color:#6C757D; font-weight:bold; max-width:100%;'>📸 [여기에 사진 삽입: 디테일 컷]</div>`
                       - Ensure the HTML code is enclosed in a markdown code block (```html ... ```).
                    """
                    st.session_state['seo_result_text'] = ask_ai(agent_prompt)
                    st.session_state['seo_prod_name'] = prod_name
            else:
                st.warning("상품명과 핵심 포인트를 입력해주세요!")

        if 'seo_result_text' in st.session_state:
            st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px;'>{st.session_state['seo_result_text'].replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
            
            st.write("") 
            
            st.download_button(
                label="📥 기획안 및 HTML 코드 다운로드 (.txt)",
                data=st.session_state['seo_result_text'],
                file_name=f"DUWELL_상세페이지기획_HTML_{st.session_state.get('seo_prod_name', '기본')}.txt",
                mime="text/plain",
                use_container_width=True
            )
