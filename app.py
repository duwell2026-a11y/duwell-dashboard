import textwrap
import streamlit as st
import altair as alt
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
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
# 1. 페이지 및 WEHAGO 스타일 모던 UI 설정
# --------------------------------------------------------------------------
import streamlit as st

st.set_page_config(page_title="DUWELL 스마트 ERP", layout="wide", page_icon="🍷")

st.markdown("""
    <style>
        /* 폰트 설정 (프리텐다드) */
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        html, body, [class*="css"] { font-family: 'Pretendard', sans-serif; }
        
        /* 1. 전체 배경색 (모던한 라이트 그레이) */
        .stApp { 
            background-color: #F4F6F9; 
        }
        
        /* 2. 사이드바 스타일 (화이트 배경 + 은은한 그림자) */
        [data-testid="stSidebar"] { 
            background-color: #FFFFFF; 
            border-right: none !important;
            box-shadow: 2px 0 12px rgba(0,0,0,0.05);
        }
        /* 사이드바 글자색을 어둡게 변경 */
        [data-testid="stSidebar"] * { 
            color: #333333 !important; 
            font-weight: 500;
        }
        /* 사이드바 라디오 버튼(메뉴) 호버 효과 */
        div.row-widget.stRadio > div > label {
            padding: 10px 15px;
            border-radius: 10px;
            transition: all 0.2s ease-in-out;
        }
        div.row-widget.stRadio > div > label:hover {
            background-color: #F0F4FF; /* 연한 블루 배경 */
        }

        /* 3. 카드형 메트릭 (상단 요약 지표 박스) */
        [data-testid="metric-container"] {
            background-color: #FFFFFF;
            border-radius: 16px;
            padding: 20px 24px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);
            border: 1px solid rgba(0,0,0,0.02);
            transition: transform 0.2s ease;
        }
        [data-testid="metric-container"]:hover {
            transform: translateY(-2px); /* 마우스 올리면 살짝 위로 뜸 */
            box-shadow: 0 6px 16px rgba(0, 0, 0, 0.08);
        }
        /* 메트릭 폰트 색상 및 크기 조정 */
        [data-testid="metric-container"] label { color: #6C757D !important; font-weight: 600; font-size: 1rem; }
        [data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #2B3A55 !important; font-size: 1.8rem; font-weight: 800; }

        /* 4. 탭(Tab) 메뉴 모던화 (블루 포인트) */
        [data-testid="stTabs"] button {
            background-color: transparent;
            border: none;
            border-bottom: 3px solid transparent;
            padding-bottom: 10px;
            font-weight: 600;
            color: #6C757D !important;
        }
        [data-testid="stTabs"] button[aria-selected="true"] {
            color: #4E73DF !important; /* 메인 포인트 컬러: 모던 블루 */
            border-bottom: 3px solid #4E73DF !important;
        }

        /* 5. 데이터프레임(표) 스타일 (화이트 박스 안에 가두기) */
        .stDataFrame {
            background-color: #FFFFFF;
            padding: 15px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
            border: 1px solid #EBEEF4;
        }

        /* 6. Expander (아코디언 메뉴) 스타일 */
        [data-testid="stExpander"] {
            background-color: #FFFFFF;
            border-radius: 12px;
            border: 1px solid #EBEEF4;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02);
        }

        /* 7. 상단 헤더 숨기기 (더 깔끔하게) */
        [data-testid="stHeader"] { background: transparent; }

        /* 모바일 반응형 유지 */
        @media (max-width: 768px) {
            .block-container { padding-top: 2rem !important; padding-left: 1rem !important; padding-right: 1rem !important; }
            h1 { font-size: 1.6rem !important; }
            [data-testid="metric-container"] { padding: 15px; }
            div.row-widget.stRadio > div { flex-direction: row; flex-wrap: wrap; }
            div.row-widget.stRadio > div > label { width: 48%; margin-bottom: 5px; text-align: center; background-color: #F8F9FA;}
        }

        /* 인쇄 설정 */
        @media print {
            @page { size: A4; margin: 10mm; }
            body, .stApp, .block-container { background: white !important; margin: 0; padding: 0; }
            aside, header, button { display: none !important; }
            [data-testid="metric-container"], .stDataFrame { box-shadow: none !important; border: 1px solid #000 !important; }
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
        SHEET_ID = "1xqcbuzRzzp4i_Qsy4CKRjIIvGOTthT88bXxxY5RjEjQ"
        GOOGLE_API_KEY = "AIzaSyBBReb6mUNBeIGa2n-GJEt-lUphanHq3jg"
        SENDER_EMAIL = "duwell2026@gmail.com"
        SENDER_PASSWORD = "mvxo jzki djzg iwor"
        with open(local_key_path, "r", encoding="utf-8") as f:
            GOOGLE_CREDENTIALS = json.load(f)
    else:
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

def get_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        if GOOGLE_CREDENTIALS:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(GOOGLE_CREDENTIALS, scope)
            return gspread.authorize(creds)
        return None
    except Exception as e: return None

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
def load_data(sheet_name):
    client = get_client()
    if not client: return pd.DataFrame(), None
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        if df.empty: return df, sheet
        
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
        
        # 컬러가 있으면 상품명에 병합
        # if '컬러' in df.columns and '상품명' in df.columns:
          #  df['상품명'] = df.apply(lambda x: f"{x['상품명']} ({x['컬러']})" if pd.notnull(x.get('컬러')) and str(x.get('컬러')).strip() != '' else x['상품명'], axis=1)

        df = df.loc[:, ~df.columns.duplicated()]
        
        # [중요] 필수 컬럼 보장
        required_cols = ['날짜', '구매자명', '연락처', '주소', '상품명', '수량', '결제금액', '요청사항', '디자인파일', '상태', '포장옵션']
        for col in required_cols:
            if col not in df.columns: df[col] = "" 
        
        if '주문처' not in df.columns: df['주문처'] = '🏠 자사몰'
        return df, sheet
    except Exception as e: return pd.DataFrame(), None

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
    try:
        model_name = get_best_model()
        model = genai.GenerativeModel(model_name)
        response = model.generate_content(prompt)
        return response.text
    except: return "AI Error"

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
    """공장 엑셀 업로드 시 일괄 사용"""
    try:
        all_data = sheet.get_all_values()
        header = all_data[0]
        
        # 시트 컬럼 인덱스 확보
        try:
            name_idx = header.index('구매자명')
            item_idx = header.index('상품명')
            status_idx = next(i for i, h in enumerate(header) if h in ['상태', '진행상태'])
            track_idx = header.index('송장번호')
        except:
            return False, "❌ 시트 컬럼명(구매자명, 상품명, 상태, 송장번호)을 확인해주세요."

        success_count = 0
        # 엑셀의 '구매자명'과 '송장번호' 컬럼 존재 확인
        if '구매자명' not in df_up.columns or '송장번호' not in df_up.columns:
            return False, "❌ 엑셀에 '구매자명'과 '송장번호' 열이 있어야 합니다."

        for _, row in df_up.iterrows():
            u_name = str(row['구매자명']).strip()
            u_track = str(row['송장번호']).strip()
            if not u_name or not u_track or u_track == 'nan': continue

            # 시트에서 매칭되는 행 찾기 (최신 데이터부터 찾기 위해 역순)
            for i in range(len(all_data)-1, 0, -1):
                s_row = all_data[i]
                if s_row[name_idx].strip() == u_name:
                    # 매칭 성공 시 업데이트
                    sheet.update_cell(i + 1, track_idx + 1, u_track)
                    sheet.update_cell(i + 1, status_idx + 1, "배송중")
                    success_count += 1
                    break
        
        return True, f"✅ 총 {success_count}건의 배송 상태가 업데이트되었습니다."
    except Exception as e:
        return False, f"오류: {str(e)}"

# --------------------------------------------------------------------------
# 🏠 메인 UI 로직
# --------------------------------------------------------------------------

with st.sidebar:
    st.markdown("<h1 style='color:#800020;'>🍷 DUWELL</h1>", unsafe_allow_html=True)
    if st.button("🔄 데이터 새로고침", type="primary"): 
        st.cache_data.clear()  # 저장된 캐시를 싹 비우고 새로 가져옴
        st.rerun()
    menu = st.radio("메뉴 이동", [
        "🏠 통합 모니터링", "📦 주문 등록/관리", "🖨️ 작업지시서", 
        "🛠️ 재고 입출고 관리", "🏭 공장 발주", "📢 마케팅 센터", 
        "🎨 디자인 시안실", "📅 일정 관리", "📋 주문 장부", "🛠️ 옵션 관리", "💎 고객 CRM 센터", "💰 마진/정산 분석"
    ])

st.markdown(f"<h2 style='color:#333;'>{menu}</h2>", unsafe_allow_html=True)
st.divider()

df_duwell, sheet_main = load_data("시트1") 
df_all = df_duwell.copy()
if not df_all.empty and '날짜' in df_all.columns:
    df_all['날짜'] = pd.to_datetime(df_all['날짜'], errors='coerce')
    df_all = df_all.sort_values(by='날짜', ascending=False)
    df_all['날짜_str'] = df_all['날짜'].dt.strftime('%Y-%m-%d')
else:
    if not df_all.empty: df_all['날짜_str'] = ""

# === [1] 🏠 통합 모니터링 ===
if menu == "🏠 통합 모니터링":
    # 🔥 블루 배너 추가
    st.markdown("""
        <div style="background: linear-gradient(135deg, #4E73DF 0%, #224ABE 100%); padding: 30px; border-radius: 16px; color: white; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(78, 115, 223, 0.4);">
            <h2 style="margin:0; font-size:1.8rem; font-weight:800; color:white;">DUWELL 통합 비즈니스 대시보드</h2>
            <p style="margin:10px 0 0 0; font-size:1rem; opacity:0.9; color:white;">오늘의 매출 현황과 다가오는 일정을 한눈에 확인하세요.</p>
        </div>
    """, unsafe_allow_html=True)
    
    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")
    
    if not df_all.empty:
        # ==========================================
        # 금액 계산 및 파생 변수(주/월) 생성
        # ==========================================
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

        # 📅 주차 및 월별 데이터 생성
        df_all['주'] = df_all['날짜'].dt.strftime('%Y년 %U주차')
        df_all['월'] = df_all['날짜'].dt.strftime('%Y-%m')

        today_orders = df_all[df_all['날짜_str'] == today_str]
        this_month_orders = df_all[df_all['월'] == today.strftime('%Y-%m')]
        
        # 1. 상단 핵심 지표
# (기존 코드) 상단 4개 메트릭 카드
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.metric("📦 오늘 주문", f"{len(today_orders)}건", delta=f"{len(today_orders)}건 New")
        with c2: st.metric("💰 오늘 매출", f"{today_orders['금액_숫자'].sum():,.0f}원")
        with c3: st.metric("📅 이달의 매출", f"{this_month_orders['금액_숫자'].sum():,.0f}원", delta=f"{len(this_month_orders)}건 주문")
        with c4: st.metric("🏆 누적 매출", f"{df_all['금액_숫자'].sum():,.0f}원")

        st.divider()

        # ==========================================
        # 🔥 [부활 및 업그레이드!] 실시간 AI 비즈니스 분석 리포트
        # ==========================================
        st.markdown("#### ✨ AI 비즈니스 인사이트")
        with st.expander("🤖 AI에게 현재 매출/재고 상황 분석 맡기기", expanded=False):
            st.write("현재까지의 주문, 매출, 베스트셀러 및 재고 데이터를 바탕으로 AI가 비즈니스 인사이트를 도출합니다.")
            
            if st.button("✨ 맞춤형 AI 리포트 생성", key="ai_report_btn", type="primary"):
                with st.spinner("데이터를 분석하고 인사이트를 작성 중입니다..."):
                    # 1. AI에게 먹일 데이터 요약하기
                    top_items_dict = df_all['상품명'].value_counts().head(3).to_dict()
                    top_items_str = ", ".join([f"{k}({v}건)" for k, v in top_items_dict.items()])
                    
                    # 2. 실시간 재고 부족 상황 파악
                    df_stock_temp, _ = load_data("재고관리")
                    low_stock_str = "없음 (모두 정상)"
                    if not df_stock_temp.empty:
                        df_stock_temp['현재재고'] = pd.to_numeric(df_stock_temp['현재재고'], errors='coerce').fillna(0)
                        df_stock_temp['안전재고'] = pd.to_numeric(df_stock_temp['안전재고'], errors='coerce').fillna(0)
                        low_stock_items = df_stock_temp[df_stock_temp['현재재고'] <= df_stock_temp['안전재고']]['상품명'].tolist()
                        if low_stock_items:
                            low_stock_str = ", ".join(low_stock_items)

                    # 3. AI 프롬프트 (명령어)
                    prompt = f"""
                    당신은 'DUWELL' (프리미엄 타월 브랜드)의 수석 비즈니스 분석가입니다.
                    아래 제공된 실시간 비즈니스 현황 데이터를 분석하여, 대표님께 보고할 짧고 명확한 인사이트 리포트를 작성해주세요.

                    [현재 실시간 데이터]
                    - 오늘 매출: {today_orders['금액_숫자'].sum():,.0f}원 (주문 {len(today_orders)}건)
                    - 이달 매출: {this_month_orders['금액_숫자'].sum():,.0f}원 (주문 {len(this_month_orders)}건)
                    - 누적 매출: {df_all['금액_숫자'].sum():,.0f}원
                    - 베스트셀러 Top 3: {top_items_str}
                    - 재고 부족 경고 상품: {low_stock_str}

                    [요청 사항]
                    1. 현 상황에 대한 긍정적인 요약 (1문장)
                    2. 재고 부족이나 판매 추이를 기반으로 한 당면 과제 및 조언 (1~2문장)
                    3. 이달 매출을 더 끌어올릴 수 있는 즉시 실행 가능한 마케팅/세일즈 액션 플랜 2가지 (글머리 기호 사용)
                    
                    어조는 전문적이고 희망차며, 모바일에서도 읽기 좋게 이모지를 적절히 섞어 작성해주세요.
                    """
                    
                    # 4. AI 답변 출력
                    ai_result = ask_ai(prompt)
                    st.success("✅ AI 리포트 생성이 완료되었습니다.")
                    st.markdown(f"<div style='background-color:#F8F9FA; padding:20px; border-radius:12px; border:1px solid #E9ECEF;'>{ai_result.replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

        st.divider()
        
        # (기존 코드 계속...)
        # st.markdown("#### 📈 기간별 매출 추이")
        # tab_d, tab_w, tab_m = st.tabs(["📊 일별 매출", "📊 주간 매출", "📊 월간 매출"])

        # 2. 🔥 [업그레이드 1] 기간별 매출 및 주문 추이
        st.markdown("#### 📈 기간별 매출 추이")
        tab_d, tab_w, tab_m = st.tabs(["📊 일별 매출", "📊 주간 매출", "📊 월간 매출"])

        with tab_d:
            df_trend = df_all.groupby('날짜_str')['금액_숫자'].sum().reset_index().sort_values('날짜_str').tail(15)
            line_chart = alt.Chart(df_trend).mark_line(point=True, color='#800020').encode(
                x=alt.X('날짜_str:N', title='날짜', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('금액_숫자:Q', title='매출액'),
                tooltip=['날짜_str', '금액_숫자']
            ).properties(height=300)
            st.altair_chart(line_chart, use_container_width=True)

        with tab_w:
            df_week = df_all.groupby('주').agg(매출액=('금액_숫자', 'sum'), 주문건수=('구매자명', 'count')).reset_index().sort_values('주').tail(10)
            bar_week = alt.Chart(df_week).mark_bar(color='#30343B', opacity=0.9).encode(
                x=alt.X('주:N', title='주차', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('매출액:Q', title='매출액'),
                tooltip=['주', '매출액', '주문건수']
            ).properties(height=300)
            st.altair_chart(bar_week, use_container_width=True)

        with tab_m:
            df_month = df_all.groupby('월').agg(매출액=('금액_숫자', 'sum'), 주문건수=('구매자명', 'count')).reset_index().sort_values('월').tail(12)
            bar_month = alt.Chart(df_month).mark_bar(color='#2ca02c', opacity=0.8).encode(
                x=alt.X('월:N', title='월별', axis=alt.Axis(labelAngle=0)),
                y=alt.Y('매출액:Q', title='매출액'),
                tooltip=['월', '매출액', '주문건수']
            ).properties(height=300)
            st.altair_chart(bar_month, use_container_width=True)

        st.divider()

        # 3. 중간 영역: 재고 및 일정
        col_stock, col_sch = st.columns([1, 1])
        
        with col_stock:
            st.markdown("#### 🚨 재고 경고 (안전재고 미달)")
            df_stock_alert, _ = load_data("재고관리")
            if not df_stock_alert.empty:
                df_stock_alert['현재재고'] = pd.to_numeric(df_stock_alert['현재재고'], errors='coerce').fillna(0)
                df_stock_alert['안전재고'] = pd.to_numeric(df_stock_alert['안전재고'], errors='coerce').fillna(0)
                low_stock = df_stock_alert[df_stock_alert['현재재고'] <= df_stock_alert['안전재고']]
                
                if not low_stock.empty:
                    for _, s in low_stock.iterrows():
                        st.error(f"**{s['상품명']}**: 현재 {int(s['현재재고'])}개")
                else:
                    st.success("✅ 모든 상품 재고 정상")
            else:
                st.write("재고 데이터 없음")

        with col_sch:
            # 🔥 [업그레이드 2] 기간별 일정 관리 탭
            st.markdown("#### 📅 다가오는 일정")
            df_sch, _ = load_data("일정관리")
            
            t_today, t_week, t_month = st.tabs(["오늘", "이번 주", "이번 달"])
            
            if not df_sch.empty:
                # 안전한 날짜 계산을 위해 datetime 타입으로 변환
                df_sch['시작일_dt'] = pd.to_datetime(df_sch['시작일'], errors='coerce')
                
                with t_today:
                    today_sch = df_sch[df_sch['시작일'] == today_str]
                    if not today_sch.empty:
                        for _, r in today_sch.iterrows(): st.info(f"**{r.get('시간','')}** | {r.get('일정명','')}")
                    else: st.write("오늘 예정된 일정이 없습니다.")
                
                with t_week:
                    # 오늘부터 7일 이내 일정
                    week_end = today + timedelta(days=7)
                    week_sch = df_sch[(df_sch['시작일_dt'] >= today) & (df_sch['시작일_dt'] <= week_end)].sort_values('시작일_dt')
                    if not week_sch.empty:
                        for _, r in week_sch.iterrows(): st.success(f"[{r['시작일']}] {r.get('시간','')} | {r.get('일정명','')}")
                    else: st.write("이번 주 예정된 일정이 없습니다.")
                    
                with t_month:
                    # 이번 달 1일부터 말일까지의 일정
                    this_month_str = today.strftime('%Y-%m')
                    month_sch = df_sch[df_sch['시작일'].astype(str).str.startswith(this_month_str, na=False)].sort_values('시작일_dt')
                    if not month_sch.empty:
                        for _, r in month_sch.iterrows(): st.warning(f"[{r['시작일']}] {r.get('일정명','')}")
                    else: st.write("이번 달 일정이 없습니다.")
            else:
                st.write("등록된 일정이 없습니다.")

        st.divider()

        st.markdown("#### 📦 최근 주문 리포트 (최신 5건)")
        possible_cols = ['날짜_str', '구매자명', '상품명', '수량', '결제금액', '상태']
        cols = [c for c in possible_cols if c in df_all.columns]
        st.dataframe(df_all[cols].head(5), hide_index=True, use_container_width=True)

    else:
        st.warning("📊 아직 등록된 주문 데이터가 없습니다.")

# === [2] 📦 주문 등록/관리 ===
elif menu == "📦 주문 등록/관리":
    tab1, tab2 = st.tabs(["📂 엑셀 일괄 등록", "📝 수동 주문 등록"])
    
    with tab1:
        uploaded_file = st.file_uploader("네이버 엑셀 업로드", type=['xlsx'])
        if uploaded_file and st.button("💾 저장 및 재고 차감"):
            try:
                df_new = pd.read_excel(uploaded_file, header=1)
                df_opt, _ = load_data("옵션관리")
                _, sheet_stock = load_data("재고관리")
                
                rows_add = []
                log_msg = []
                
                for _, row in df_new.iterrows():
                    p_name = str(row.get('상품명',''))
                    qty = int(row.get('수량', 1))
                    
                    rows_add.append([
                        str(row.get('주문일시','')), str(row.get('수취인명','')), str(row.get('수취인연락처1','')),
                        str(row.get('배송지','')), p_name, str(qty),
                        str(row.get('총 주문금액','0')), "", "", str(row.get('배송메세지','')), "", "신규"
                    ])
                    ok, msg = deduct_stock_smart(p_name, qty, df_opt, sheet_stock)
                    log_msg.append(msg)

                if sheet_main:
                    sheet_main.append_rows(rows_add)
                    updated_stock, _ = load_data("재고관리")
                    check_stock_and_alert(updated_stock)
                    st.success(f"{len(rows_add)}건 처리 완료")
                    with st.expander("처리 로그"): st.write(log_msg)
                    time.sleep(1); st.rerun()
            except Exception as e: st.error(f"오류: {e}")

    with tab2:
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
                m_req = st.text_area("요청사항(자수)")
                m_file = st.text_input("디자인링크")
            
            if st.form_submit_button("등록 및 재고차감"):
                if sheet_main:
                    sheet_main.append_row([str(m_date), m_name, m_phone, m_addr, m_prod, str(m_qty), str(m_price), m_file, "", m_req, "", "신규(수동)"])
                    df_opt, _ = load_data("옵션관리")
                    _, sheet_stock = load_data("재고관리")
                    ok, msg = deduct_stock_smart(m_prod, m_qty, df_opt, sheet_stock)
                    
                    updated_stock, _ = load_data("재고관리")
                    check_stock_and_alert(updated_stock)
                    
                    st.success(msg)
                    time.sleep(1); st.rerun()

# === [3] 🖨️ 작업지시서 (Base64 이미지 임베딩 & 원본 메일 발송) ===
elif menu == "🖨️ 작업지시서":
    import requests
    import base64
    
    st.subheader("🖨️ 작업지시서 발급 및 메일 전송")
    st.info("💡 지시서를 다운로드하거나, **원본 고해상도 이미지와 함께 공장으로 바로 메일을 발송**할 수 있습니다.")

    if df_all.empty:
        st.warning("주문 데이터가 없습니다.")
    else:
        filtered = df_all.copy()
        for c in ['구매자명', '연락처', '주소', '상품명', '수량', '요청사항', '디자인파일', '희망수령일']:
            if c not in filtered.columns: filtered[c] = "-"

        if "체크" not in filtered.columns:
            filtered.insert(0, "체크", False)
            
        edited = st.data_editor(
            filtered,
            column_order=['체크', '날짜_str', '구매자명', '상품명', '수량', '상태'],
            column_config={"체크": st.column_config.CheckboxColumn(required=True)},
            hide_index=True, use_container_width=True, key="print_base64_final"
        )
        
        selected = edited[edited['체크'] == True]
        
        if not selected.empty:
            st.divider()
            st.markdown("#### 🚀 선택한 작업지시서 처리")
            
            # 메일 발송 입력란
            factory_email = st.text_input("📧 수신 이메일 주소 (공장/작업자)", value="factory@example.com")
            
            col_btn1, col_btn2 = st.columns(2)
            
            # [버튼 1] 기존 다운로드 버튼
            with col_btn1:
                if st.button("📥 작업지시서 HTML 다운로드 (Click)", type="secondary", use_container_width=True):
                    with st.spinner("이미지를 변환하여 지시서를 만드는 중입니다..."):
                        html_content = """
                        <!DOCTYPE html>
                        <html lang="ko">
                        <head>
                            <meta charset="UTF-8">
                            <title>DUWELL 작업지시서</title>
                            <style>
                                @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
                                @page { size: A4; margin: 0; }
                                body { font-family: 'Pretendard', sans-serif; margin: 0; padding: 0; background-color: #eee; -webkit-print-color-adjust: exact; }
                                .page { width: 210mm; height: 297mm; padding: 15mm; background: white; margin: 0 auto 20px auto; box-sizing: border-box; position: relative; page-break-after: always; overflow: hidden; }
                                @media print { body { background: white; } .page { margin: 0; width: 210mm; height: 297mm; page-break-after: always; border: none; } .no-print { display: none; } }
                                h1 { text-align: center; font-size: 28pt; margin: 0 0 20px 0; letter-spacing: 8px; border-bottom: 3px double black; padding-bottom: 10px; }
                                table { width: 100%; border-collapse: collapse; margin-bottom: 15px; border: 2px solid black; }
                                th, td { border: 1px solid black; padding: 6px 10px; font-size: 12pt; vertical-align: middle; }
                                th { background-color: #f2f2f2 !important; width: 18%; text-align: center; font-weight: bold; }
                                .highlight { color: #d63384; font-weight: bold; }
                                .date-highlight { background-color: #ffebeb !important; color: red !important; font-weight: bold; }
                                .img-box { text-align: center; margin-top: 10px; border: 1px solid #ddd; padding: 5px; height: 500px; display: flex; flex-direction: column; justify-content: center; align-items: center; }
                                .img-box img { max-width: 100%; max-height: 480px; object-fit: contain; }
                                .footer { margin-top: 15px; border: 2px dashed red; padding: 10px; background-color: #fff5f5; }
                                .footer h3 { margin: 0 0 5px 0; color: red; font-size: 14pt; }
                                .footer li { font-size: 11pt; line-height: 1.4; }
                            </style>
                        </head>
                        <body>
                        """
                        for _, row in selected.iterrows():
                            raw_url = str(row.get('디자인파일', ''))
                            drive_id = None
                            match = re.search(r"(?:id=|\/d\/)([\w-]{20,})", raw_url)
                            if match: drive_id = match.group(1)

                            img_tag = ""
                            if drive_id:
                                try:
                                    img_url = f"https://drive.google.com/uc?export=view&id={drive_id}"
                                    response = requests.get(img_url, timeout=10)
                                    if response.status_code == 200:
                                        img_b64 = base64.b64encode(response.content).decode('utf-8')
                                        img_tag = f'<div class="img-box"><p style="font-weight:bold; margin:0 0 5px 0; color:#555;">[ 디자인 시안 ]</p><img src="data:image/jpeg;base64,{img_b64}" alt="디자인 시안"></div>'
                                    else:
                                        img_tag = '<div class="img-box" style="color:red;">이미지 로딩 실패</div>'
                                except Exception as e:
                                    img_tag = f'<div class="img-box" style="color:red;">이미지 변환 오류<br><small>{str(e)}</small></div>'
                            else:
                                img_tag = '<div class="img-box" style="color:#aaa;">이미지 파일 없음</div>'

                            req_val = str(row['요청사항']).replace('nan', '없음')
                            target_date = str(row.get('희망수령일', '-')).replace('nan', '-')
                            html_content += f"""
                            <div class="page">
                                <h1>작 업 지 시 서</h1>
                                <table>
                                    <tr><th>구매자</th><td>{row['구매자명']} ({row['연락처']})</td><th>접수일</th><td>{row['날짜_str']}</td></tr>
                                    <tr><th>배송지</th><td colspan="3" style="font-weight:bold; font-size:14pt;">{row['주소']}</td></tr>
                                    <tr><th>상품명</th><td>{row['상품명']}</td><th>수량</th><td style="font-size:20pt; color:red; font-weight:bold;">{row['수량']} 개</td></tr>
                                    <tr><th>요청사항</th><td colspan="3" class="highlight" style="font-size:14pt;">{req_val}</td></tr>
                                    <tr><th class="date-highlight">희망수령일</th><td colspan="3" class="date-highlight" style="font-size:18pt;">{target_date}</td></tr>
                                </table>
                                {img_tag}
                                <div class="footer"><h3>⚠️ 공장 필독 주의사항</h3><ul><li>자수 퀄리티 및 실밥 마감 상태를 철저히 확인해주세요.</li><li>시안에 명시된 자수 위치와 크기를 준수해주세요.</li><li><b>{target_date}</b> 납기 엄수!</li></ul></div>
                            </div>"""
                        html_content += "</body></html>"
                        file_name = f"DUWELL_작업지시서_{datetime.now().strftime('%Y%m%d_%H%M')}.html"
                        st.download_button("📥 생성 완료! 다운로드", data=html_content, file_name=file_name, mime="text/html", type="primary", use_container_width=True)

            # [버튼 2] 원본 이미지와 함께 메일 전송
            with col_btn2:
                if st.button("🚀 원본 이미지 포함 메일 발송", type="primary", use_container_width=True):
                    with st.spinner("최고 해상도 원본 이미지를 드라이브에서 가져오는 중입니다..."):
                        email_attachments = []
                        progress_bar = st.progress(0)
                        
                        for idx, (_, row) in enumerate(selected.iterrows()):
                            buyer_name = str(row['구매자명']).strip()
                            raw_url = str(row.get('디자인파일', ''))
                            
                            drive_id = None
                            match = re.search(r"(?:id=|\/d\/)([\w-]{20,})", raw_url)
                            if match: drive_id = match.group(1)

                            img_tag = '<div class="img-box" style="color:#aaa;">이미지 파일 없음</div>'
                            
                            if drive_id:
                                try:
                                    # 🔥 [핵심] export=download 를 사용해 압축없는 최고 해상도 원본 파일 가져오기
                                    raw_img_url = f"https://drive.google.com/uc?export=download&id={drive_id}"
                                    response = requests.get(raw_img_url, timeout=15)
                                    
                                    if response.status_code == 200:
                                        # 1. 이메일 첨부용으로 원본 이미지 파일 추가 (.jpg로 통일)
                                        img_io = io.BytesIO(response.content)
                                        email_attachments.append({
                                            "file": img_io, 
                                            "filename": f"원본시안_{buyer_name}.jpg"
                                        })
                                        
                                        # 2. 지시서(HTML) 내부 렌더링용 base64
                                        img_b64 = base64.b64encode(response.content).decode('utf-8')
                                        img_tag = f'<div class="img-box"><p style="font-weight:bold; margin:0 0 5px 0; color:#555;">[ 디자인 시안 ]</p><img src="data:image/jpeg;base64,{img_b64}" alt="디자인 시안"></div>'
                                except Exception as e:
                                    pass
                            
                            # 개별 HTML 지시서 생성 (메일 첨부용)
                            req_val = str(row['요청사항']).replace('nan', '없음')
                            target_date = str(row.get('희망수령일', '-')).replace('nan', '-')
                            
                            single_html = f"""
                            <!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>작업지시서_{buyer_name}</title>
                            <style>
                                body {{ font-family: sans-serif; background: white; margin: 0; padding: 20px; }}
                                h1 {{ text-align: center; border-bottom: 2px solid black; padding-bottom: 10px; }}
                                table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
                                th, td {{ border: 1px solid black; padding: 10px; }}
                                th {{ background-color: #f2f2f2; width: 20%; }}
                                .img-box {{ text-align: center; margin-top: 20px; }}
                                .img-box img {{ max-width: 100%; }}
                            </style></head><body>
                                <h1>작 업 지 시 서 ({buyer_name})</h1>
                                <table>
                                    <tr><th>상품명</th><td>{row['상품명']}</td><th>수량</th><td style="color:red; font-weight:bold;">{row['수량']} 개</td></tr>
                                    <tr><th>요청사항</th><td colspan="3" style="color:#d63384; font-weight:bold;">{req_val}</td></tr>
                                    <tr><th>배송지</th><td colspan="3">{row['주소']} (연락처: {row['연락처']})</td></tr>
                                </table>
                                {img_tag}
                            </body></html>
                            """
                            html_io = io.BytesIO(single_html.encode('utf-8'))
                            email_attachments.append({
                                "file": html_io, 
                                "filename": f"작업지시서_{buyer_name}.html"
                            })
                            
                            progress_bar.progress((idx + 1) / len(selected))

                        # 최종 이메일 발송
                        subject = f"[DUWELL] 신규 작업지시서 및 원본시안 송부 ({len(selected)}건)"
                        body = "안녕하세요,\n\n첨부된 작업지시서(HTML)와 최고 해상도의 원본 시안 이미지(JPG)를 확인 후 작업 부탁드립니다.\n\n감사합니다."
                        
                        mail_ok, mail_msg = send_email_with_attach(
                            to=factory_email, 
                            subject=subject, 
                            body=body, 
                            multiple_attachments=email_attachments
                        )
                        
                        if mail_ok:
                            st.success("🎉 성공적으로 공장에 지시서와 원본 이미지가 발송되었습니다!")
                            st.balloons()
                        else:
                            st.error(f"메일 발송 실패: {mail_msg}")

# === [4] 🛠️ 재고 입출고 관리 (대량 등록 기능 추가) ===
elif menu == "🛠️ 재고 입출고 관리":
    st.subheader("📊 재고 통합 관리 시스템")
    
    # 데이터 로드
    df_stock, sheet_stock = load_data("재고관리")
    df_opt, _ = load_data("옵션관리")
    
    # 탭 구성
    tab1, tab2, tab3 = st.tabs(["📊 재고 현황 (자동집계)", "📂 대량 입출고 등록 (엑셀)", "📝 개별 조정 (수동)"])
    
    # [탭 1] 재고 현황 (기존 로직 유지 + 시각화)
    with tab1:
        st.info("💡 **자동 집계**: 현재 재고는 '주문 내역'을 기반으로 실시간 차감된 수치입니다.")
        
        if not df_stock.empty and not df_all.empty:
            stock_summary = []
            
            # 재고 현황 계산
            for _, stock_item in df_stock.iterrows():
                s_name = str(stock_item['상품명']).strip()
                
                # 옵션 매핑 키워드 찾기
                keywords = [s_name]
                if not df_opt.empty and '매핑명' in df_opt.columns:
                    match = df_opt[df_opt['상품명'] == s_name]
                    if not match.empty:
                        keywords += [k.strip() for k in str(match.iloc[0]['매핑명']).split(',') if k.strip()]
                
                # 총 판매량 집계
                total_out = 0
                for _, order in df_all.iterrows():
                    o_name = str(order.get('상품명', ''))
                    o_qty = pd.to_numeric(order.get('수량', 0), errors='coerce')
                    if any(k in o_name for k in keywords):
                        total_out += o_qty
                
                current_stock = int(pd.to_numeric(stock_item.get('현재재고', 0), errors='coerce'))
                safety_stock = int(pd.to_numeric(stock_item.get('안전재고', 0), errors='coerce'))
                
                # 상태 판별
                status = "✅ 정상"
                if current_stock <= 0: status = "❌ 품절"
                elif current_stock <= safety_stock: status = "🚨 부족"
                
                stock_summary.append({
                    "상품명": s_name,
                    "현재 재고": current_stock,
                    "안전 재고": safety_stock,
                    "상태": status,
                    "누적 판매(추정)": int(total_out)
                })
                
            df_summary = pd.DataFrame(stock_summary)
            
            # 스타일링된 데이터프레임 출력
            st.dataframe(
                df_summary.style.map(lambda x: 'color: red; font-weight: bold;' if x in ['🚨 부족', '❌ 품절'] else '', subset=['상태']),
                use_container_width=True, 
                hide_index=True
            )
            
            # 재고 차트
            st.divider()
            st.markdown("#### 📉 재고 수량 차트")
            st.bar_chart(df_summary.set_index('상품명')['현재 재고'])
            
        else:
            st.warning("데이터가 부족합니다.")

    # [탭 2] 대량 입출고 등록 (엑셀 업로드) - 신규 기능
    with tab2:
        st.markdown("### 📥 엑셀로 재고 일괄 등록")
        st.write("입고(추가)하거나 출고(차감)할 내역을 엑셀로 업로드하여 한 번에 처리합니다.")
        
        # 1. 양식 다운로드
        with st.expander("ℹ️ 엑셀 양식 다운로드 및 사용법"):
            st.markdown("""
            1. 아래 표를 복사하여 엑셀에 붙여넣으세요. (헤더명 정확히 일치해야 함)
            2. **구분**: '입고' 또는 '출고'로 입력
            3. **수량**: 숫자만 입력
            """)
            sample_data = pd.DataFrame({
                '상품명': ['와플 타월 (와인)', '호텔 수건 170g', '선물세트 A'],
                '수량': [100, 50, 10],
                '구분': ['입고', '출고', '입고']
            })
            st.dataframe(sample_data, hide_index=True)
            
            # CSV 다운로드 버튼
            csv = sample_data.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 양식 파일 다운로드 (CSV)", csv, "재고등록_양식.csv", "text/csv")

        # 2. 파일 업로드 및 처리
        uploaded_file = st.file_uploader("작성한 엑셀 파일 업로드", type=['xlsx', 'xls', 'csv'])
        
        if uploaded_file:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df_upload = pd.read_csv(uploaded_file)
                else:
                    df_upload = pd.read_excel(uploaded_file)
                
                # 필수 컬럼 확인
                required_cols = ['상품명', '수량', '구분']
                if not all(col in df_upload.columns for col in required_cols):
                    st.error(f"❌ 필수 컬럼이 누락되었습니다. (필요 컬럼: {required_cols})")
                else:
                    st.success(f"{len(df_upload)}건의 데이터를 읽어왔습니다. 아래 내용을 확인하고 '반영하기'를 누르세요.")
                    st.dataframe(df_upload.head(), use_container_width=True)
                    
                    if st.button("🚀 재고 일괄 반영하기", type="primary"):
                        if sheet_stock:
                            # 전체 재고 데이터 로드 (속도를 위해 한 번에 읽음)
                            all_values = sheet_stock.get_all_values()
                            headers = all_values[0]
                            
                            # 컬럼 인덱스 찾기
                            try:
                                idx_name = headers.index('상품명')
                                idx_qty = next(i for i, h in enumerate(headers) if '재고' in h or '수량' in h) # '현재재고' 또는 '수량'
                            except:
                                st.error("❌ 시트에서 '상품명' 또는 '현재재고' 열을 찾을 수 없습니다.")
                                st.stop()

                            # 상품명 -> 행 번호 매핑 (속도 최적화)
                            # (row_index는 0부터 시작하므로, 시트 행번호는 +1)
                            prod_map = {row[idx_name].strip(): i for i, row in enumerate(all_values)}
                            
                            success_count = 0
                            fail_list = []
                            
                            # 데이터 업데이트 로직 (메모리 상에서 수정)
                            for _, row in df_upload.iterrows():
                                p_name = str(row['상품명']).strip()
                                qty = int(row['수량'])
                                type_ = str(row['구분']).strip()
                                
                                if p_name in prod_map:
                                    row_idx = prod_map[p_name]
                                    current_val = all_values[row_idx][idx_qty]
                                    current_qty = int(pd.to_numeric(current_val, errors='coerce') or 0)
                                    
                                    if type_ == '입고':
                                        new_qty = current_qty + qty
                                    elif type_ == '출고':
                                        new_qty = current_qty - qty
                                    else:
                                        new_qty = current_qty # 구분 오타 시 변동 없음
                                    
                                    # 메모리 값 업데이트
                                    all_values[row_idx][idx_qty] = new_qty
                                    success_count += 1
                                else:
                                    fail_list.append(p_name)
                            
                            # 시트에 통째로 다시 쓰기 (Batch Update - 가장 빠름)
                            sheet_stock.update(all_values)
                            
                            st.success(f"✅ 총 {success_count}건 재고 반영 완료!")
                            if fail_list:
                                st.warning(f"⚠️ 다음 상품은 찾을 수 없어 제외됨: {', '.join(fail_list)}")
                            
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("구글 시트 연결 실패")
            except Exception as e:
                st.error(f"처리 중 오류 발생: {e}")

    # [탭 3] 개별 조정 (수동)
    with tab3:
        st.markdown("### 📝 개별 상품 입/출고")
        if not df_stock.empty:
            with st.form("manual_stock"):
                col1, col2, col3 = st.columns([2, 1, 1])
                with col1:
                    target_prod = st.selectbox("상품 선택", df_stock['상품명'].unique())
                with col2:
                    action = st.radio("구분", ["입고 (+)", "출고/손실 (-)"], horizontal=True)
                with col3:
                    qty = st.number_input("수량", min_value=1, value=1)
                
                if st.form_submit_button("반영"):
                    # 해당 상품 찾아서 업데이트
                    cell = sheet_stock.find(target_prod)
                    if cell:
                        # 현재 재고값 가져오기 (B열이라고 가정하지 않고 헤더로 찾음)
                        headers = sheet_stock.row_values(1)
                        col_idx = -1
                        for i, h in enumerate(headers):
                            if '재고' in h or '수량' in h:
                                col_idx = i + 1
                                break
                        
                        if col_idx != -1:
                            curr_val = int(pd.to_numeric(sheet_stock.cell(cell.row, col_idx).value, errors='coerce') or 0)
                            if "입고" in action:
                                final_qty = curr_val + qty
                            else:
                                final_qty = curr_val - qty
                                
                            sheet_stock.update_cell(cell.row, col_idx, final_qty)
                            st.success(f"✅ {target_prod}: {curr_val} -> {final_qty}개로 변경됨")
                            time.sleep(1); st.rerun()
                        else:
                            st.error("재고 컬럼을 찾을 수 없습니다.")
                    else:
                        st.error("상품을 찾을 수 없습니다.")
# === [5] 🏭 공장 발주 (발주/메일/송장 관리 통합) ===
elif menu == "🏭 공장 발주":
    st.subheader("🏭 공장 스마트 발주 및 배송 관리")
    
    # 1. 주문 데이터 필터링
    if not df_all.empty:
        # 미발주: 상태가 신규이거나 발주 완료가 아닌 것들
        pending_orders = df_all[~df_all['상태'].isin(['발주완료', '배송중', '배송완료', '취소'])].copy()
        # 발주완료: 이미 발주가 들어간 것들 (송장 입력 대상)
        completed_orders = df_all[df_all['상태'] == '발주완료'].copy()
    else:
        pending_orders = pd.DataFrame()
        completed_orders = pd.DataFrame()

    tab1, tab2, tab3 = st.tabs(["📦 1. 발주 대상 선택", "📧 2. 메일 전송", "🚚 3. 발주 완료 내역 & 송장 등록"])

    # [탭 1] 발주 대상 선택
    with tab1:
        if pending_orders.empty:
            st.success("🎉 발주 대기 중인 주문이 없습니다.")
        else:
            if "발주선택" not in pending_orders.columns:
                pending_orders.insert(0, "발주선택", False)
            
            edited_orders = st.data_editor(
                pending_orders,
                column_config={"발주선택": st.column_config.CheckboxColumn(required=True)},
                column_order=['발주선택', '날짜_str', '구매자명', '상품명', '수량', '상태'],
                hide_index=True, use_container_width=True, key="factory_order_v5"
            )
            st.session_state['selected_orders'] = edited_orders[edited_orders['발주선택'] == True]

    # [탭 2] 메일 전송
    with tab2:
        selected = st.session_state.get('selected_orders', pd.DataFrame())
        if selected.empty:
            st.warning("⚠️ 발주할 주문을 먼저 선택해 주세요.")
        else:
            st.write(f"✅ 선택된 주문: {len(selected)}건")
            factory_email = st.text_input("공장 이메일 주소", value="factory@example.com")

            if st.button("🚀 발주 확정 및 메일 전송"):
                progress_bar = st.progress(0)
                # 시트 업데이트
                for i, (_, row) in enumerate(selected.iterrows()):
                    update_status_in_sheet(sheet_main, row, "발주완료")
                    progress_bar.progress((i + 1) / len(selected) * 0.5)
                
                # 엑셀 파일 생성 및 전송
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    selected[['구매자명', '연락처', '주소', '상품명', '수량', '요청사항']].to_excel(writer, index=False)
                output.seek(0)
                
                subject = f"[DUWELL] 발주서_{datetime.now().strftime('%m%d')}"
                body = f"신규 발주서 {len(selected)}건 송부드립니다."
                mail_ok, mail_msg = send_email_with_attach(factory_email, subject, body, output, f"DUWELL_발주_{datetime.now().strftime('%m%d')}.xlsx")
                
                if mail_ok:
                    st.success("🎊 발주 완료 및 메일 발송 성공!")
                    st.cache_data.clear()
                    time.sleep(1); st.rerun()

    # [탭 3] 발주 완료 내역 및 송장 등록 (신규 추가)
    with tab3:
        st.markdown("#### 📜 발주 완료 목록")
        if completed_orders.empty:
            st.info("발주 완료된 내역이 없습니다.")
        else:
            # 송장 번호를 입력받기 위한 에디터
            if '송장번호' not in completed_orders.columns:
                completed_orders['송장번호'] = ""
            
            st.write("💡 송장번호 열에 번호를 입력한 후 아래 **'송장번호 일괄 저장'** 버튼을 누르세요.")
            
            # 특정 열만 수정 가능하게 설정
            tracking_edited = st.data_editor(
                completed_orders,
                column_order=['날짜_str', '구매자명', '상품명', '수량', '송장번호'],
                column_config={
                    "송장번호": st.column_config.TextColumn("송장번호 입력", help="숫자만 입력하세요", width="medium")
                },
                disabled=['날짜_str', '구매자명', '상품명', '수량'],
                hide_index=True, use_container_width=True, key="tracking_editor"
            )

            if st.button("💾 송장번호 일괄 저장"):
                with st.spinner("시트에 송장 정보를 기록 중..."):
                    save_count = 0
                    # 번호가 입력된 행만 골라서 업데이트
                    for _, row in tracking_edited.iterrows():
                        t_num = str(row.get('송장번호', '')).strip()
                        if t_num and t_num != "" and t_num != "None":
                            ok, msg = update_tracking_in_sheet(sheet_main, row, t_num)
                            if ok: save_count += 1
                    
                    if save_count > 0:
                        st.success(f"✅ {save_count}건의 송장 번호가 저장되었습니다.")
                        st.cache_data.clear()
                        time.sleep(1); st.rerun()
                    else:
                        st.warning("입력된 송장 번호가 없습니다.")

# === [6] 📢 마케팅 센터 (광고 관리 및 ROAS 통합) ===
elif menu == "📢 마케팅 센터":
    st.markdown("""
        <div style="background: linear-gradient(135deg, #4E73DF 0%, #224ABE 100%); padding: 25px; border-radius: 12px; color: white; margin-bottom: 20px; box-shadow: 0 4px 10px rgba(78, 115, 223, 0.3);">
            <h2 style="margin:0; font-size:1.6rem; font-weight:800; color:white;">📢 퍼포먼스 마케팅 & AI 센터</h2>
            <p style="margin:5px 0 0 0; font-size:0.95rem; opacity:0.9; color:white;">광고 효율(ROAS)을 추적하고, 매체별 최적화된 광고 소재를 생성하세요.</p>
        </div>
    """, unsafe_allow_html=True)
    
    # 🔥 탭에 광고 분석 추가
    m_tab0, m_tab1, m_tab2, m_tab3, m_tab4 = st.tabs([
        "📈 광고 효율(ROAS) 분석", "✍️ 매체별 광고 생성(네이버/인스타)", "💡 브랜드 네이밍", "📅 프로모션 기획", "💬 리뷰/CS 관리"
    ])
    
    # --- [신규 탭] 광고 효율(ROAS) 분석 ---
    with m_tab0:
        st.markdown("#### 🎯 일일 광고비 입력 및 ROAS 측정")
        st.info("💡 오늘 집행한 광고비를 입력하면, ERP에 기록된 **오늘의 매출**과 연동하여 즉시 ROAS를 계산합니다.")
        
        # 오늘 매출 가져오기 (대시보드 로직 활용)
        today_str = datetime.now().strftime("%Y-%m-%d")
        today_sales = 0
        
        if not df_all.empty:
            df_roas = df_all.copy()
            df_roas['기존금액'] = pd.to_numeric(df_roas.get('결제금액', 0).astype(str).str.replace(r'[^\d]', '', regex=True), errors='coerce').fillna(0)
            df_roas['수량_숫자'] = pd.to_numeric(df_roas.get('수량', 1), errors='coerce').fillna(1)
            
            df_opt, _ = load_data("옵션관리")
            def get_price_for_roas(item_name):
                if df_opt.empty or '가격' not in df_opt.columns: return 0
                item_clean = str(item_name).replace(" ", "").lower()
                for _, opt in df_opt.iterrows():
                    std = str(opt.get('상품명', ''))
                    map_str = str(opt.get('매핑명', ''))
                    kws = [k.strip() for k in map_str.split(',') if k.strip()] + [std]
                    for kw in kws:
                        if kw and kw.replace(" ", "").lower() in item_clean:
                            return pd.to_numeric(str(opt.get('가격', 0)).replace(',', '').replace('원', ''), errors='coerce')
                return 0
                
            df_roas['계산단가'] = df_roas['상품명'].apply(get_price_for_roas)
            df_roas['최종금액'] = df_roas.apply(lambda x: x['기존금액'] if x['기존금액'] > 0 else x['계산단가'] * x['수량_숫자'], axis=1)
            
            today_sales = df_roas[df_roas['날짜_str'] == today_str]['최종금액'].sum()
        
        with st.form("roas_form"):
            r_col1, r_col2 = st.columns(2)
            with r_col1:
                naver_spend = st.number_input("🟢 네이버 검색/디스플레이 광고비 (원)", min_value=0, step=10000)
            with r_col2:
                meta_spend = st.number_input("🟣 인스타그램/페이스북 광고비 (원)", min_value=0, step=10000)
                
            if st.form_submit_button("📊 ROAS 계산하기", type="primary"):
                total_spend = naver_spend + meta_spend
                
                c_roas1, c_roas2, c_roas3 = st.columns(3)
                c_roas1.metric("총 광고비 지출", f"{total_spend:,.0f}원")
                c_roas2.metric("오늘 발생한 총 매출", f"{today_sales:,.0f}원")
                
                if total_spend > 0:
                    roas = (today_sales / total_spend) * 100
                    # ROAS가 300% 이상이면 초록색, 이하면 빨간색/회색 느낌으로 이모지 부여
                    roas_icon = "🔥 대박!" if roas >= 400 else ("✅ 양호" if roas >= 250 else "⚠️ 점검 필요")
                    c_roas3.metric("오늘의 ROAS", f"{roas:,.1f}% {roas_icon}")
                else:
                    c_roas3.metric("오늘의 ROAS", "광고비 0원")
                    
                st.caption("※ ROAS(Return On Ad Spend) = (매출액 ÷ 광고비) × 100. 보통 300% 이상을 안정권으로 봅니다.")

    # --- 탭 1: 카피라이팅 (네이버/인스타 특화) ---
    with m_tab1:
        st.markdown("#### ✨ 매체 최적화 광고 문구 생성")
        with st.form("copywriting_form"):
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                # 자연스럽게 브랜드의 컨셉을 예시 및 기본값으로 세팅
                product_name = st.text_input("상품/캠페인 특징", value="시그니처 와인 컬러의 프리미엄 와플 타월, '미스터 와플' 3D 캐릭터가 갓 구워낸 베이커리 컨셉")
                target_audience = st.selectbox("타겟 고객", ["3040 리빙/인테리어 고관여자", "2030 신혼부부/집들이 선물", "호캉스 매니아", "전체"])
            with col_c2:
                channel = st.selectbox("광고 매체 (채널)", [
                    "🟢 네이버 파워링크 (제목 15자 / 설명 45자 제한)", 
                    "🟣 인스타그램 스폰서드 (시각적 후킹 + 감성 해시태그)", 
                    "🔴 카카오 알림톡 (짧고 강렬한 전환 유도)"
                ])
                tone = st.selectbox("톤앤매너", ["세련된/고급스러운", "재치있는/유머러스한 (캐릭터 활용)", "신뢰감 주는/전문적인", "긴급한/혜택 강조 (한정수량)"])
            
            keywords = st.text_input("필수 포함 키워드 (쉼표로 구분)", value="집들이선물, 호텔수건, 와플수건, 흡수력")
            
            if st.form_submit_button("✍️ 광고 소재 생성하기", type="primary") and product_name:
                with st.spinner(f"AI가 {channel.split(' ')[1]}에 최적화된 소재를 작성 중입니다..."):
                    prompt = f"""
                    당신은 'DUWELL' 타월 브랜드의 퍼포먼스 마케터입니다.
                    상품 및 캠페인: {product_name}
                    타겟: {target_audience}
                    채널: {channel}
                    톤앤매너: {tone}
                    필수 키워드: {keywords}
                    
                    요청:
                    선택한 채널의 특성에 완벽하게 맞춘 광고 카피를 3가지 버전으로 제안해주세요.
                    - 네이버 파워링크일 경우: [제목]은 15자 이내, [설명]은 45자 이내로 글자수를 철저히 지켜주세요.
                    - 인스타그램일 경우: 이미지/영상에 들어갈 [이미지 텍스트]와 본문에 적을 [캡션], 그리고 유입을 이끌 [해시태그]를 구분해서 써주세요.
                    """
                    st.success("✅ 소재 생성 완료!")
                    st.markdown(f"<div style='background-color:#fff; padding:20px; border-radius:12px; border:1px solid #E9ECEF;'>{ask_ai(prompt).replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)

    # --- 탭 2: 네이밍 ---
    with m_tab2:
        st.markdown("#### 🏷️ 브랜드/이벤트 네이밍")
        with st.form("naming_form"):
            col_n1, col_n2 = st.columns(2)
            with col_n1:
                category = st.selectbox("분야", ["신상품 이름", "브랜드 슬로건", "이벤트/캠페인 명", "컬러/옵션 명칭"])
                lang_style = st.selectbox("언어 스타일", ["영어 (세련된 느낌)", "한글 (순우리말 느낌)", "합성어 (직관적)", "짧고 톡톡 튀는"])
            with col_n2:
                desc = st.text_area("특징 및 컨셉 설명", height=100)
            if st.form_submit_button("💡 아이디어 제안받기") and desc:
                with st.spinner("창의적인 이름을 고민 중입니다..."):
                    prompt = f"분야: {category}, 스타일: {lang_style}, 특징: {desc}\n부르기 좋은 이름 5가지를 추천이유와 함께 제안해줘."
                    st.info(ask_ai(prompt))

    # --- 탭 3: 프로모션 기획 ---
    with m_tab3:
        st.markdown("#### 📅 매출을 올리는 이벤트 기획")
        with st.form("promo_form"):
            season = st.text_input("이벤트 테마 (예: 런칭 기념, 가정의 달)")
            goal = st.selectbox("목표", ["브랜드 인지도 확보", "초기 리뷰/사진 확보", "매출 극대화", "재구매 유도"])
            discount_type = st.selectbox("혜택 유형", ["할인 (Show me the money)", "증정 (선물 박스/사은품)", "무료배송", "체험단 모집"])
            if st.form_submit_button("📅 기획안 작성"):
                with st.spinner("프로모션을 기획 중입니다..."):
                    st.markdown(ask_ai(f"테마: {season}, 목표: {goal}, 혜택: {discount_type}\n이벤트 타이틀 3개, 구체적인 진행방식, 마케팅 소구포인트를 포함해 기획안 작성해줘."))

    # --- 탭 4: 리뷰/CS 관리 통합 ---
    with m_tab4:
        st.markdown("#### 💬 스마트 고객 응대 (리뷰 & CS)")
        sub_tab1, sub_tab2 = st.tabs(["엑셀 리뷰 일괄 답글", "CS 답변 스크립트"])
        
        with sub_tab1:
            uploaded_review = st.file_uploader("리뷰 엑셀 파일 (.xlsx)", type=['xlsx'])
            if uploaded_review:
                df_rev = pd.read_excel(uploaded_review)
                st.write("데이터 미리보기:", df_rev.head(2))
                review_col = st.selectbox("고객 리뷰 내용이 있는 열", df_rev.columns.tolist())
                if st.button("🤖 AI 답글 생성"):
                    with st.spinner("답글을 작성 중입니다..."):
                        replies = []
                        pb = st.progress(0)
                        for i, row in df_rev.iterrows():
                            prompt = f"고객 리뷰: '{row[review_col]}'. 브랜드 관리자로서 감사와 공감이 담긴 정중한 답글을 2문장으로 작성해."
                            replies.append(ask_ai(prompt))
                            pb.progress((i + 1) / len(df_rev))
                        df_rev['AI_자동답글'] = replies
                        st.success("🎉 생성 완료!")
                        import io
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer: df_rev.to_excel(writer, index=False)
                        st.download_button("📥 결과 파일 다운로드", output.getvalue(), "리뷰답글완료.xlsx")
                        
        with sub_tab2:
            cs_type = st.radio("문의 유형", ["배송 지연", "제품 불량/교환", "포장/상태 불만", "오염/세탁 관련 문의"], horizontal=True)
            cs_detail = st.text_area("고객 문의 내용 (복사 붙여넣기)", height=100)
            if st.button("💬 CS 방어 답변 생성") and cs_detail:
                st.info(ask_ai(f"상황: {cs_type}, 고객 문의: {cs_detail}\n프리미엄 브랜드 CS 매니저로서 진정성 있는 사과와 명확한 해결책이 담긴 답변 작성."))

# === [7] 🎨 디자인 시안실 ===
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

# === [8] 📅 일정 관리 ===
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

# === [9] 📋 주문 장부 ===
elif menu == "📋 주문 장부": 
    st.subheader("📋 전체 주문 장부")
    if not df_all.empty:
        st.dataframe(df_all, use_container_width=True)
        csv = df_all.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 다운로드", csv, "order_list.csv", "text/csv")

# === [10] 🛠️ 옵션 관리 ===
elif menu == "🛠️ 옵션 관리":
    st.subheader("🛠️ 옵션 및 통합 상품명 관리")
    df_opt, sheet_opt = load_data("옵션관리")
    if '매핑명' not in df_opt.columns: st.error("⚠️ '매핑명' 컬럼 필요")
    if not df_opt.empty:
        edited_df = st.data_editor(df_opt, num_rows="dynamic", use_container_width=True)
        if st.button("💾 저장"):
            sheet_opt.clear()
            sheet_opt.update([edited_df.columns.values.tolist()] + edited_df.values.tolist())
            st.success("저장됨"); st.rerun()
    with st.expander("도움말"): st.write("매핑명: 상품명, 별칭1, 별칭2 (쉼표로 구분)")

# === [11] 💎 고객 CRM 센터 (완전 복구됨) ===
elif menu == "💎 고객 CRM 센터":
    st.subheader("💎 고객 통합 프로필 및 상담")
    if not df_all.empty:
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
            
            t1, t2 = st.tabs(["👤 통합 리스트", "🎯 상세 상담 관리"])
            with t1:
                st.dataframe(cust_profile, use_container_width=True)
            with t2:
                search_nm = st.selectbox("고객 선택", cust_profile['고객명'].unique())
                sel_data = cust_profile[cust_profile['고객명'] == search_nm].iloc[0]
                
                c_a, c_b = st.columns(2)
                with c_a:
                    st.markdown(f"### 👤 {search_nm} 프로필")
                    st.write(f"- 등급: {sel_data['등급']}")
                    st.write(f"- 누적 금액: {sel_data['누적금액']:,.0f}원")
                    
                    st.markdown("#### 📜 상담 히스토리")
                    current_history = ""
                    try:
                        client = get_client(); target_sh = client.open("주문데이터").worksheet("시트1")
                        h = target_sh.row_values(1)
                        if '비고' in h:
                            cell = target_sh.find(search_nm)
                            if cell:
                                current_history = target_sh.cell(cell.row, h.index('비고')+1).value
                                st.text_area("기록", value=current_history or "내용 없음", height=150, disabled=True)
                    except: pass
                    
                    memo_in = st.text_area("📝 신규 상담 내용 입력")
                    if st.button("💾 상담 내용 저장"):
                        try:
                            now = datetime.now().strftime('%Y-%m-%d %H:%M')
                            final = f"{current_history}\n[{now}] {memo_in}" if current_history else f"[{now}] {memo_in}"
                            target_sh.update_cell(cell.row, h.index('비고')+1, final)
                            st.success("저장됨"); st.rerun()
                        except: st.error("저장 실패")
                
                with c_b:
                    st.markdown("### 🤖 맞춤 마케팅 제안")
                    if st.button("✨ AI 문구 추천"):
                        msg = ask_ai(f"고객명: {search_nm}, 등급: {sel_data['등급']}. 안부 문자 메시지 작성해줘.")
                        st.write(msg)

        except Exception as e: st.error(f"오류: {e}")

# === [12] 🛠️ 재고 관리 (간편 보기) ===
elif menu == "🛠️ 재고 관리":
    st.info("💡 상세 입출고 내역은 **'🛠️ 재고 입출고 관리'** 메뉴를 이용해주세요.")
    df_stock, _ = load_data("재고관리")
    if not df_stock.empty:
        cols = st.columns(4)
        for i, (idx, row) in enumerate(df_stock.head(4).iterrows()):
            is_low = pd.to_numeric(row['현재재고'], errors='coerce') <= pd.to_numeric(row['안전재고'], errors='coerce')
            cols[i].metric(
                label=row['상품명'], 
                value=f"{row['현재재고']}개", 
                delta="-재고부족" if is_low else "정상", 
                delta_color="inverse" if is_low else "normal"
            )
        st.divider()
        st.bar_chart(df_stock.set_index('상품명')[['현재재고']], height=400)
# === [13] 💰 마진/정산 분석 ===
elif menu == "💰 마진/정산 분석":
    st.subheader("💰 실시간 예상 마진 및 정산 분석기")
    
    with st.expander("⚙️ 정산 기준 설정 (수수료 및 배송비)", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            fee_smart = st.number_input("스마트스토어 수수료 (%)", value=5.5)
            fee_own = st.number_input("자사몰(PG) 수수료 (%)", value=3.0)
        with col_f2:
            fee_coupang = st.number_input("쿠팡 수수료 (%)", value=10.8)
            fee_etc = st.number_input("기타 마켓 수수료 (%)", value=5.0)
        with col_f3:
            shipping_cost = st.number_input("건당 평균 택배비 (원)", value=2500, step=100)

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
            
            # 🔥 [강력해진 검색] 띄어쓰기 무시
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
            
            expected_revenue = unit_price * qty
            commission_fee = expected_revenue * fee_rate
            total_cost = unit_cost * qty
            delivery_fee = shipping_cost
            
            net_profit = expected_revenue - commission_fee - total_cost - delivery_fee
            margin_rate = (net_profit / expected_revenue * 100) if expected_revenue > 0 else 0
            
            return pd.Series([expected_revenue, commission_fee, total_cost, net_profit, margin_rate])

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

# [탭 3] 일별 매출 캘린더
        with tab_cal:
            st.markdown("### 📆 캘린더 뷰 (일별 매출 & 순이익)")
            
            # 1. 날짜가 빈칸이거나 'nan'인 찌꺼기 데이터 확실히 제거
            valid_dates = df_calc[df_calc['날짜_str'].astype(bool) & (df_calc['날짜_str'] != 'nan') & (df_calc['날짜_str'] != '')]
            
            if valid_dates.empty:
                st.info("표시할 정상적인 날짜 데이터가 없습니다.")
            else:
                # 일별로 매출과 이익 합산
                daily_sales = valid_dates.groupby('날짜_str').agg(
                    매출액=('예상결제금액', 'sum'), 
                    순이익=('예상순이익', 'sum')
                ).reset_index()

                events = []
                for _, row in daily_sales.iterrows():
                    d_str = str(row['날짜_str']).strip()
                    
                    # 매출액 이벤트 (회색 뱃지)
                    events.append({
                        "title": f"매출: {row['매출액']:,.0f}",
                        "start": d_str,
                        "color": "#555555"
                    })
                    # 순이익 이벤트 (와인색 뱃지)
                    events.append({
                        "title": f"이익: {row['순이익']:,.0f}",
                        "start": d_str,
                        "color": "#800020"
                    })
                
                # 2. [가장 중요] 캘린더가 보이게 하는 필수 뼈대 옵션 추가!
                cal_options = {
                    "headerToolbar": {
                        "left": "prev,next today",
                        "center": "title",
                        "right": "dayGridMonth"
                    },
                    "initialView": "dayGridMonth", # 한 달 단위로 보여주기
                }
                
                if events: 
                    # 달력 그리기 실행
                    calendar(events=events, options=cal_options)
                else: 
                    st.info("달력에 표시할 이벤트가 없습니다.")
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
