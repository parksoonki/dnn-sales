import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import datetime
import re
import requests
from supabase import create_client, Client

# Supabase 연결 (Streamlit Secrets에서 정보를 가져옴)
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(url, key)

# 데이터 불러오기 함수
def fetch_data():
    # 'sales' 테이블의 모든 데이터를 가져옴 (행이 많을 경우를 대비해 넉넉히 설정)
    response = supabase.table("sales").select("*").execute()
    if len(response.data) > 0:
        return pd.DataFrame(response.data)
    return None

# 앱 시작 시 DB에서 데이터 로드
if st.session_state.main_data is None:
    st.session_state.main_data = fetch_data()



# --- 1. 페이지 기본 설정 ---
st.set_page_config(page_title="영업 종합 대시보드", page_icon="💡", layout="wide")

# --- 2. 초기 상태(Session State) 세팅 ---
if 'main_data' not in st.session_state: st.session_state.main_data = None
if 'cust_data' not in st.session_state: st.session_state.cust_data = None
if 'company_goal' not in st.session_state: st.session_state.company_goal = 500000000 # 5억
if 'rep_goals' not in st.session_state: st.session_state.rep_goals = {}
if 'rep_list' not in st.session_state: st.session_state.rep_list = []
if 'deleted_reps' not in st.session_state: st.session_state.deleted_reps = set()

if 'category_map' not in st.session_state:
    st.session_state.category_map = {
        '파30': 'PAR30', '멀티매입': 'PAR30', 'T5': 'T5', '다운라이트': '다운라이트',
        'U-LAMP': '램프', '벌브': '램프', '촛대구': '램프', 'T-LAMP': '램프', '엠알16': '램프', '볼구': '램프', '미니크립톤': '램프', '인지구': '램프',
        '방습등': '방습등', '다목적등': '방습등', '주차장등': '방습등',
        '벽부등': '벽부등', '사각모던': '벽부등', '원형모던': '벽부등', '직사각COB': '벽부등', '모던투': '벽부등',
        '롱넥': '볼구', '실링팬': '실링팬',
        '십자등': '주택등', '일자등': '주택등', '욕실등': '주택등', '리츠': '주택등', '원형방등': '주택등', '사각방등': '주택등', '시스템': '주택등', '리폼': '주택등',
        '홈LED직하': '직하엣지', '직부': '직부/센서', '센서': '직부/센서', '슬림면조명': '직하엣지', '직하엣지': '직하엣지',
        '투광등': '투광등', '이클립스': '펜던트', '우디': '펜던트', '코모': '펜던트', '심플라인': '펜던트', '듀스': '펜던트', '레체': '펜던트', '벨로스': '펜던트', '베로나': '펜던트',
        '크림': '펜던트', '코튼': '펜던트', '벨라': '펜던트', '루나': '펜던트', '데이지': '펜던트', '글레어': '펜던트', '몽블랑': '펜던트', '프리즘': '펜던트',
        '타워': '펜던트', '샤르망': '펜던트', '이모션': '펜던트', '아크': '펜던트', '로렐': '펜던트', '노블': '펜던트', '물망초': '펜던트', '골드라인': '펜던트', '실버라인': '펜던트',
        '컨버터': '부속', '파레트': '배송비', '퀵': '배송비', '대신화물': '배송비'
    }

def format_krw(val):
    """숫자를 한국식 조, 억, 만 단위로 변환 (예: 1억 5,000만 원)"""
    if not val or val == 0: return "0원"
    sign = "-" if val < 0 else ""
    val = abs(val)
    
    units = [(10**12, '조'), (10**8, '억'), (10**4, '만')]
    result = []
    
    for unit_val, unit_name in units:
        if val >= unit_val:
            quotient = val // unit_val
            result.append(f"{int(quotient):,}{unit_name}")
            val %= unit_val
            
    if val > 0 or not result:
        result.append(f"{int(val):,}")
        
    return sign + " ".join(result) + "원"

def format_short_krw(val):
        if pd.isna(val) or val <= 0: return ""
        uk = val // 100000000
        man = (val % 100000000) // 10000
        if uk > 0 and man > 0: return f"{int(uk)}억 {int(man):,}만"
        elif uk > 0: return f"{int(uk)}억"
        elif man > 0: return f"{int(man):,}만"
        return f"{int(val):,}" # 만원 미만일 경우

def assign_category(product_name, cat_map):
    for keyword, category in cat_map.items():
        if str(keyword) in str(product_name):
            return category
    return '미분류'

def normalize_customer_name(name):
    name = str(name)

    # 괄호 제거
    name = re.sub(r'\(.*?\)', '', name)

    # 회사 접두어 제거
    name = name.replace("주식회사", "")
    name = name.replace("(주)", "")

    # 공백 정리
    name = name.strip()
    return name

# 이월건(적요 날짜) 추출 함수
def extract_month(row):
    if pd.notna(row.get('일자')):
        try: 
            m = pd.to_datetime(str(row['일자']).split(' -')[0]).month
            if pd.notna(m) and 1 <= int(m) <= 12:
                return int(m)
        except: pass
    return 1


# 배송 유형 추출용
def get_delivery_type(text):
    text = str(text)
    if "배송" in text: return "배송"
    elif "화물" in text or "택배" in text: return "화물"
    elif "합바" in text: return "합바"
    elif "퀵" in text: return "퀵"
    elif "직배" in text: return "직배"
    elif "방문" in text: return "방문수령"
    elif "직송" in text: return "업체직송"
    else: return "기타"

# 색온도 추출용
def extract_color_temp(text):
    import re
    match = re.search(r'(\d{2,4})\s*[Kk]', str(text))
    if match:
        raw_num = match.group(1)
        if len(raw_num) == 4:
            return raw_num[:2] + "K"
        else:
            return raw_num + "K"
    return "기타"

# --- 3. 엑셀 데이터 전처리 함수 ---
@st.cache_data
def preprocess_data(df, raw_header_str):
    import pandas as pd
    import numpy as np
    import re
    import datetime
    df = df.copy()

    # --- 1. 파일 전체의 '업로드 출처월' 판단 (A1 셀 또는 데이터 빈도) ---
    source_month = None
    header_str = str(raw_header_str)
    match_date = re.search(r'(\d{4})/(\d{2})/\d{2}\s*~\s*\d{4}/(\d{2})', header_str)
    if match_date:
        source_month = int(match_date.group(3))
    
    if source_month is None or not (1 <= source_month <= 12):
        try:
            temp_months = df['일자'].astype(str).str.extract(r'/\s*(0?[1-9]|1[0-2])\s*/')[0].dropna().astype(int)
            source_month = int(temp_months.mode()[0])
        except:
            source_month = int(datetime.datetime.now().month)

    # --- 2. [핵심] 일자 데이터 정제 및 타입 변환 ---
    df['일자'] = df['일자'].astype(str).str.split(' -').str[0].str.strip()
    df['일자'] = pd.to_datetime(df['일자'], errors='coerce')

    # --- 3. [개선] 행별 매출월 추출 (사용자 규칙 적용) ---
    def extract_month_strictly(row):
        dt = row['일자']
        if pd.isnull(dt): return np.nan
        if dt.day == 1:
            val_long = str(row.get('장문형식1', '')).strip()
            match_carry = re.search(r'(?<!\d)(0?[1-9]|1[0-2])\.(\d{1,2})(?!\s*(?:인치|인찌|동|호|W|V|A|K))', val_long)
            if match_carry:
                m = int(match_carry.group(1))
                return m
        
        return int(dt.month)

    df['월'] = df.apply(extract_month_strictly, axis=1)
    df = df.dropna(subset=['월'])
    
    df['월'] = df['월'].astype(int)
    df['업로드_출처월'] = int(source_month)

    # 3. 숫자형 데이터 변환 (콤마 제거)
    for col in ['수량', '합 계', '공급가액', '단가']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    
    # '합 계' 띄어쓰기 통일
    if '합 계' in df.columns and '합계' not in df.columns:
        df.rename(columns={'합 계': '합계'}, inplace=True)

    # 4. 텍스트 데이터 정제 및 결측치(NaN)를 '미분류'로 처리
    for col in ['판매처명', '담당자명', '품명 및 규격', '카테고리', '적요', '장문형식1']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(['nan', 'NaN', 'None', ''], '미분류')
        else:
            df[col] = '미분류'

    # 5. 완벽한 지역 매핑 (주식회사, 공백 등 무시하고 VLOOKUP)
    if 'cust_data' in st.session_state and st.session_state.cust_data is not None:
        clean_cust_data = {str(k).strip(): v for k, v in st.session_state.cust_data.items() if k}
        # 판매처명도 공백 제거 후 매핑, 실패 시 '미분류'
        df['지역'] = df['판매처명'].astype(str).str.strip().map(clean_cust_data).fillna('미분류')
    else:
        df['지역'] = '미분류'

    # 6. 카테고리(품목) 자동 분류
    if 'category_map' in st.session_state:
        def assign_category(item_name):
            for k, v in st.session_state.category_map.items():
                if k in item_name: return v
            return "기타"
        df['카테고리'] = df['품명 및 규격'].apply(assign_category)

    # 7. 배송유형 추출 ('적요'와 '장문형식1'의 텍스트를 보고 택배/화물/배송 등 판별)
    def get_delivery_type(row):
        txt = str(row.get('적요', '')) + " " + str(row.get('장문형식1', ''))
        if any(k in txt for k in ['퀵']): return '퀵'
        if any(k in txt for k in ['직배']): return '직배'
        if any(k in txt for k in ['배송']): return '배송'
        if any(k in txt for k in ['화물']): return '화물'
        if any(k in txt for k in ['합바']): return '합바'
        if any(k in txt for k in ['수령', '방문']): return '방문수령'
        return '기타'
    df['배송유형'] = df.apply(get_delivery_type, axis=1)

    # 8. 색온도 추출
    def extract_color_temp(item_name):
        name = str(item_name).upper()
        
        # 1. 3색 변환 제품 최우선 처리
        if '3색' in name or '변환' in name: 
            return '3색변환'
            
        # 2. 명확한 숫자+K (예: 57K, 5700K)가 있으면 '주광/주백' 글자 무시하고 무조건 이것을 적용!
        import re
        match = re.search(r'(\d{2,4})\s*K', name)
        if match:
            raw_num = match.group(1)
            # 6500 -> 65K, 57 -> 57K 로 2자리 통일
            if len(raw_num) >= 3:
                return raw_num[:2] + 'K'
            else:
                return raw_num + 'K'
                
        # 3. K 숫자가 아예 안 적혀있고, 한글(주광/주백/전구)만 적힌 경우의 예외 처리
        if '주광' in name: return '65K'
        if '주백' in name: return '40K'
        if '전구' in name: return '30K'
        
        return '기타'
        
    df['색온도'] = df['품명 및 규격'].apply(extract_color_temp)

    # 9. 와트(W) 수 추출
    df['와트(W)'] = df['품명 및 규격'].astype(str).str.upper().str.extract(r'(\d{1,3}W)')[0].fillna('미표기')

    # 10. 사이즈/길이 추출 (엣지등 가로세로, T5 길이, 다운라이트 인치 매칭)
    def extract_size(text):
        text = str(text).upper()
        import re
        # 예: 1280*320, 1200MM, 6인치, 1.2M
        match = re.search(r'(\d+[\*X]\d+|\d+MM|\d+인치|\d+(\.\d+)?M)', text)
        if match: return match.group(1)
        return '미표기'
    df['사이즈'] = df['품명 및 규격'].apply(extract_size)

    # 11. 바디색상 추출
    def extract_body_color(text):
        text = str(text)
        colors = ['블랙', '검정', '화이트', '하얀', '백색', '실버', '골드', '로즈골드', '우드', '크롬', '브론즈']
        for c in colors:
            if c in text:
                if c in ['하얀', '백색']: return '화이트'
                if c == '검정': return '블랙'
                return c
        return '미표기'
    df['바디색상'] = df['품명 및 규격'].apply(extract_body_color)

    if '담당자명' in df.columns:
        unique_reps = [r for r in df['담당자명'].unique() if str(r).strip() not in ['미분류', 'nan', '']]
        if 'rep_list' in st.session_state:
            for rep in unique_reps:
                if rep not in st.session_state.rep_list:
                    st.session_state.rep_list.append(rep)

    return df

# 과거 데이터 업데이트 안전장치
df = st.session_state.main_data
if df is not None and ('카테고리' not in df.columns or '와트(W)' not in df.columns):
    st.session_state.main_data = preprocess_data(df)
    df = st.session_state.main_data


# --- 4. 글로벌 CSS ---
st.markdown("""
<style>
/* 전체 폰트 변경 */
html, body, [class*="css"]  {
    font-family: Pretendard, Noto Sans KR, sans-serif;
}
/* 화면 위쪽 불필요한 여백 최소화 */
.block-container { padding-top: 3.5rem !important; padding-bottom: 2rem !important; }

section[data-testid="stSidebar"] { min-width: 220px !important; max-width: 260px !important; background-color: #F4F5F7; }
div.stRadio > div[role="radiogroup"] > label {
    background-color: transparent !important; border: none !important; padding: 12px 15px !important;
    margin-bottom: 4px; cursor: pointer; border-radius: 8px !important; transition: background-color 0.2s ease;
}
div.stRadio > div[role="radiogroup"] > label p { color: #333333 !important; font-size: 15px !important; font-weight: 500 !important; }
div.stRadio > div[role="radiogroup"] > label:hover { background-color: #E2E8F0 !important; }
div.stRadio > div[role="radiogroup"] > label:has(input:checked) { background-color: #3B82F6 !important; }
div.stRadio > div[role="radiogroup"] > label:has(input:checked) p { color: #FFFFFF !important; font-weight: 700 !important; }
div.stRadio > div[role="radiogroup"] > label > div:first-child { display: none !important; }


/* 1. 버튼 본체: 100% 꽉 채우고 중앙 정렬 환경 구성 */
div.stButton > button[kind="secondary"] {
    width: 100% !important;
    min-height: 110px !important;
    background-color: #F8FAFC !important;
    border: 1px solid #E2E8F0 !important;
    border-radius: 12px !important;
    padding: 10px 0 !important;
    display: flex !important;
    flex-direction: column !important;
    justify-content: center !important;
    align-items: center !important;
    overflow: hidden !important;
}

/* 2. 내부 컨테이너 */
div.stButton > button[kind="secondary"] div[data-testid="stMarkdownContainer"] {
    width: 100% !important;
    display: block !important;
    text-align: center !important;
    margin: 0 !important;
    padding: 0 !important;
}

/* 3. 텍스트: 자동 꺾임 방지(pre) + 완벽한 가로 중앙 정렬(center) */
div.stButton > button[kind="secondary"] p {
    width: 100% !important; /* 💡 문제의 max-content를 100%로 변경하여 버튼 전체 너비를 쓰도록 수정 */
    text-align: center !important; /* 💡 넓어진 100% 공간 안에서 완벽하게 가운데 정렬 수행 */
    white-space: pre !important; /* 💡 칸이 좁아도 마음대로 4줄로 꺾어버리는 현상 방지 (\n만 인식) */

    font-size: 17px !important; /* 모니터 크기에 상관없이 가장 안정적인 폰트 사이즈 */
    letter-spacing: -0.5px !important;
    font-weight: 600 !important;
    color: #475569 !important;
    line-height: 1.5 !important;
    margin: 0 auto !important;
    padding: 0 !important;
}

/* 4. 첫 번째 줄(제목) 스타일 */
div.stButton > button p::first-line {
    font-size: 16x !important;
    font-weight: 500 !important;
    color: #64748B !important;
}
div.stButton > button:hover { border-color: #93C5FD !important; }
</style>
""", unsafe_allow_html=True)

# --- 5. 왼쪽 사이드바 메뉴 ---
st.sidebar.markdown("<br>", unsafe_allow_html=True)
menu = st.sidebar.radio("메뉴 이동", ["📊 종합 대시보드", "👨‍💼 담당자별 분석", "🏢 거래처별 분석", "🗺️ 지역별 현황", "📦 품목별 상세 분석", "⚙️ 데이터 설정"], label_visibility="collapsed")

def calc_mom(current_df, prev_df, col_name='합계'):
    curr_val = current_df[col_name].sum()
    prev_val = prev_df[col_name].sum()
    if prev_val == 0: return curr_val, 0
    diff = curr_val - prev_val
    return diff, (diff / prev_val) * 100

# --- 6. 메뉴별 화면 렌더링 ---
if menu != "⚙️ 데이터 설정" and df is None:
    st.warning("📊 대시보드에 표시할 데이터가 없습니다. [⚙️ 데이터 설정] 탭에서 엑셀을 업로드해주세요.")

elif menu == "📊 종합 대시보드" and df is not None:
    st.title("📊 통합 영업 대시보드")

    # 1. 📅 조회 기간 설정
    st.markdown("### 📅 조회 기간 설정")
    valid_months = sorted(df[df['월'] > 0]['월'].unique())
    if not valid_months: valid_months = [1]
    
    # 평소엔 1개 월만 보이고, 체크하면 기간 범위 선택창이 열림
    is_range = st.checkbox("🔄 기간으로 검색 (시작월 ~ 종료월)", value=False)
    
    if not is_range:
        # 단일 월 선택 (기본 화면 - 가장 최근 달이 기본으로 뜨도록 index 설정)
        sel_month = st.selectbox("📅 조회할 월 선택", options=valid_months, index=len(valid_months)-1, format_func=lambda x: f"{int(x)}월")
        sel_start = sel_month
        sel_end = sel_month
    else:
        # 기간 선택창 (체크 시 등장)
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            sel_start = st.selectbox("🟢 시작 월", options=valid_months, index=0, format_func=lambda x: f"{int(x)}월")
        with col_d2:
            valid_ends = [m for m in valid_months if m >= sel_start]
            if not valid_ends: valid_ends = [sel_start]
            sel_end = st.selectbox("🔴 종료 월", options=valid_ends, index=len(valid_ends)-1, format_func=lambda x: f"{int(x)}월")
    
    st.session_state.saved_start_m = sel_start
    st.session_state.saved_end_m = sel_end
    st.markdown("---")

    curr_df = df[(df['월'] >= sel_start) & (df['월'] <= sel_end)]
    prev_df = df[(df['월'] >= sel_start - 1) & (df['월'] <= sel_end - 1)]
    if not curr_df.empty and pd.notnull(curr_df['일자'].min()):
        prev_year_df = df[(df['일자'] >= curr_df['일자'].min() - pd.DateOffset(years=1)) & (df['일자'] <= curr_df['일자'].max() - pd.DateOffset(years=1))]
    else:
        prev_year_df = pd.DataFrame(columns=df.columns)

    # 2. 핵심 KPI 카드
    total_sales = curr_df['합계'].sum()
    goal_rate = (total_sales / st.session_state.company_goal) * 100 if st.session_state.company_goal > 0 else 0
    mom_diff, mom_pct = calc_mom(curr_df, prev_df)
    yoy_diff, yoy_pct = calc_mom(curr_df, prev_year_df)

    # 일자 + 판매처명을 그룹화하여 1건의 고유 발주(전표)로 계산
    # 1. 당월 객단가 계산 (묶음 기준)
    curr_orders = len(curr_df.groupby(['일자', '판매처명']))
    curr_aov = total_sales / curr_orders if curr_orders > 0 else 0

    # 2. 전월 객단가 계산 (동일한 묶음 기준)
    prev_orders = len(prev_df.groupby(['일자', '판매처명']))
    prev_sales = prev_df['합계'].sum()
    prev_aov = prev_sales / prev_orders if prev_orders > 0 else 0

    # 3. 객단가 증감률 계산
    if prev_aov > 0:
        aov_mom = (curr_aov - prev_aov) / prev_aov * 100
        aov_mom_txt = f"▲{aov_mom:.1f}%" if aov_mom > 0 else f"▼{abs(aov_mom):.1f}%"
    else:
        aov_mom_txt = "-"

    mom_txt = f"▲{mom_pct:.1f}%" if mom_diff > 0 else f"▼{abs(mom_pct):.1f}%" if mom_diff < 0 else "-"
    yoy_txt = f"▲{yoy_pct:.1f}%" if yoy_diff > 0 else f"▼{abs(yoy_pct):.1f}%" if yoy_diff < 0 else "-"

    c1, c2, c3, c4 = st.columns(4)
    c1.button(f"💰 당월 매출액\n{format_krw(total_sales)}\n(전월 대비 {mom_txt})", use_container_width=True, key="main_k1")
    c2.button(f"🎯 당월 목표 달성률\n{goal_rate:.1f}%\n(월 목표 할당량 기준)", use_container_width=True, key="main_k2")
    c3.button(f"🤝 당월 활성 거래처\n{curr_df['판매처명'].nunique():,} 개사\n(이번 달 구매가 발생한 곳)", use_container_width=True, key="main_k3")
    c4.button(f"💸 평균 주문액\n{format_krw(curr_aov)}\n(전월 대비 {aov_mom_txt})", use_container_width=True, key="main_k4")
    st.markdown("---")

    # 2-1. 월간 핵심 요약 브리핑 자동 생성
    if not curr_df.empty:
        date_title = f"{sel_start}~{sel_end}월" if sel_start != sel_end else f"{sel_end}월"
        item_group = curr_df.groupby('품명 및 규격')['합계'].sum()
        top_item_name = item_group.idxmax() if not item_group.empty else "데이터 없음"
        top_item_sales = item_group.max() if not item_group.empty else 0
        top_item_ratio = (top_item_sales / total_sales * 100) if total_sales > 0 else 0

        clean_reg_df = curr_df[~curr_df['지역'].isin(['미분류', 'NA', 'NAN', 'None']) & curr_df['지역'].notna()]
        reg_group = clean_reg_df.groupby('지역')['합계'].sum()
        top_reg_name = reg_group.idxmax() if not reg_group.empty else "데이터 없음"
        top_reg_sales = reg_group.max() if not reg_group.empty else 0
        top_reg_ratio = (top_reg_sales / total_sales * 100) if total_sales > 0 else 0
        
        # 2. 리스크 분석: 상위 5개 거래처 데이터 추출
        top5_cust_data = curr_df.groupby('판매처명')['합계'].sum().nlargest(5)
        top5_names = ", ".join(top5_cust_data.index.tolist())
        top5_sales = top5_cust_data.sum()
        top5_ratio = (top5_sales / total_sales * 100) if total_sales > 0 else 0
        top5_sales_txt = format_krw(top5_sales)
        
        trend_color = "#E11D48" if mom_diff > 0 else "#2563EB" if mom_diff < 0 else "#475569"
        trend_word = "상승세" if mom_diff > 0 else "하락세" if mom_diff < 0 else "유지"
        
        # HTML/CSS 브리핑 박스
        briefing_html = f"""
        <div style="background-color: #F8FAFC; border-left: 5px solid #3B82F6; padding: 20px; border-radius: 8px; margin-bottom: 25px;">
            <h4 style="margin-top: 0; color: #1E293B;">📢 {date_title} 영업 실적 핵심 요약</h4>
            <ul style="margin-bottom: 0; color: #334155; font-size: 15px; line-height: 1.8;">
                <li><b>실적 추이:</b> 당월 매출액은 전월 대비 <span style="color:{trend_color}; font-weight:bold;">{mom_txt}</span> 변화하며 전체적으로 <span style="color:{trend_color}; font-weight:bold;">{trend_word}</span>를 보이고 있습니다.</li>
                <li><b>주도 품목:</b> 가장 높은 매출을 기록 중인 품목은 <b>'{top_item_name}'</b> (매출액: {format_krw(top_item_sales)}, 전체 매출의 <b>{top_item_ratio:.1f}%</b>) 이며, 전체 매출 성장을 주도하고 있습니다.</li>
                <li><b>지역 동향:</b> <b>{top_reg_name}</b> 지역에서 가장 많은 매출이 발생하였으며, 전국 매출 점유율의 <b>{top_reg_ratio:.1f}%</b>를 차지하고 있습니다.</li>
                <li><b>리스크(집중도):</b> 상위 5개 거래처 매출액은 <b>{format_krw(top5_sales)}</b>이며, 전체 매출 의존도는 <b>{top5_ratio:.1f}%</b> 입니다.</li>
                <li style="list-style: none; margin-left: 20px; font-size: 13px; color: #64748B;">
                    📌 <b>상위 5개사:</b> {top5_names}
                </li>
            </ul>
        </div>
        """
        st.markdown(briefing_html, unsafe_allow_html=True)

    # 3. 지도 및 지역 매출 랭킹표
    c_map, c_table = st.columns([6, 4])
    clean_df = curr_df[curr_df['지역'] != '미분류']
    reg_sales = clean_df.groupby('지역')['합계'].sum().reset_index()
    
    with c_map:
        st.markdown("**🗺️ 전국 지역별 매출 현황 (지도)**")
        kor_map = {'서울':'서울특별시','부산':'부산광역시','대구':'대구광역시','인천':'인천광역시','광주':'광주광역시','대전':'대전광역시','울산':'울산광역시','세종':'세종특별자치시','경기':'경기도','강원':'강원도','충북':'충청북도','충남':'충청남도','전북':'전라북도','전남':'전라남도','경북':'경상북도','경남':'경상남도','제주':'제주특별자치도'}
        reg_sales['fullname'] = reg_sales['지역'].map(kor_map)
        try:
            geojson_url = "https://raw.githubusercontent.com/southkorea/southkorea-maps/master/kostat/2013/json/skorea_provinces_geo_simple.json"
            geojson = requests.get(geojson_url).json()
            fig_map = px.choropleth_mapbox(reg_sales.dropna(subset=['fullname']), geojson=geojson, locations='fullname', featureidkey='properties.name', color='합계', color_continuous_scale="Blues", mapbox_style="carto-positron", zoom=5.5, center={"lat": 35.9, "lon": 127.7}, opacity=0.8)
            # 지도의 세로 높이를 450으로 고정
            fig_map.update_layout(margin={"r":0,"t":0,"l":0,"b":0}, height=450)
            fig_map.update_traces(hovertemplate="<b>%{location}</b><br>매출액: %{z:,.0f}원<extra></extra>")
            st.plotly_chart(fig_map, width="stretch", key="main_map_unique")
        except:
            st.warning("지도 데이터를 불러올 수 없습니다.")
            
    with c_table:
        st.markdown("**🥇 지역별 매출 순위**")
        if not prev_df.empty and not curr_df.empty:
            # 1. 지역별 매출 집계 (미분류 제외)
            reg_curr = curr_df[~curr_df['지역'].isin(['미분류', 'NA', 'NAN', 'None']) & curr_df['지역'].notna()].groupby('지역')['합계'].sum()
            reg_prev = prev_df[~prev_df['지역'].isin(['미분류', 'NA', 'NAN', 'None']) & prev_df['지역'].notna()].groupby('지역')['합계'].sum()
            
            # 2. 데이터 병합 및 계산
            reg_table = pd.DataFrame({'당월매출': reg_curr, '전월매출': reg_prev}).fillna(0)
            
            # 전월 매출 비중(%) 계산
            total_curr = reg_table['당월매출'].sum()
            reg_table['매출 비중(%)'] = (reg_table['당월매출'] / total_curr * 100) if total_curr > 0 else 0
            
            reg_table['증감률(%)'] = ((reg_table['당월매출'] - reg_table['전월매출']) / reg_table['전월매출'].replace(0, 1)) * 100
            
            reg_table = reg_table.sort_values('당월매출', ascending=False).reset_index()
            reg_table.insert(0, '순위', range(1, len(reg_table) + 1))
            
            # 4. 스타일링 (성장률 색상 적용)
            def color_growth(val):
                color = '#E11D48' if val > 0 else '#2563EB' if val < 0 else '#475569'
                return f'color: {color}; font-weight: bold;'
            
            styled_table = reg_table.style \
                .format({
                    '당월매출': format_krw, 
                    '전월매출': format_krw, 
                    '매출 비중(%)': '{:.1f}%',
                    '증감률(%)': '{:+.1f}%'
                }) \
                .map(color_growth, subset=['증감률(%)'])
            
            st.dataframe(styled_table, hide_index=True, use_container_width=True, height=450)

    st.markdown("---")

    # 4. YTD 및 경보 시스템
    c_ytd, c_alert = st.columns([1, 1])
    with c_ytd:
        st.markdown("**🎯 연간 누적 달성률**")
        ytd_df = df[(df['월'] >= 1) & (df['월'] <= sel_end)]
        ytd_sales = ytd_df['합계'].sum()
        ytd_goal = st.session_state.company_goal
        ytd_pct = (ytd_sales / ytd_goal) * 100 if ytd_goal > 0 else 0
        fig_gauge = go.Figure(go.Indicator(
            mode = "gauge+number+delta", value = ytd_pct,
            number = {'suffix': "%", 'font': {'size': 35, 'color': '#0F172A', 'family': 'Pretendard'}},
            domain = {'x': [0, 1], 'y': [0, 1]},
            gauge = {'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"}, 'bar': {'color': "#2563EB"}, 'bgcolor': "white", 'borderwidth': 2, 'bordercolor': "#E2E8F0", 'steps': [{'range': [0, 50], 'color': '#FEE2E2'}, {'range': [50, 80], 'color': '#FEF3C7'}, {'range': [80, 100], 'color': '#D1FAE5'}], 'threshold': {'line': {'color': "red", 'width': 4}, 'thickness': 0.75, 'value': 100}}
        ))
        fig_gauge.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig_gauge, width="stretch", key="main_gauge_unique")
        st.markdown(f"<div style='text-align:center; color:#64748B; font-weight:600;'>누적 매출: {format_krw(ytd_sales)} / 연간 목표: {format_krw(ytd_goal)}</div>", unsafe_allow_html=True)

    with c_alert:
        st.markdown("**🚨 미거래 경고 거래처**")
        with st.container():
            col_t1, col_t2 = st.columns(2)
            with col_t1: 
                risk_rev_m = st.number_input("기준 월매출 (단위: 만원)", value=st.session_state.get('risk_rev_m', 300), step=50, key="risk_rev_main")
                risk_revenue = risk_rev_m * 10000
                st.caption(f"👉 **적용 대상: 월 {risk_rev_m:,.0f}만 원 이상 거래처**")
            with col_t2: 
                risk_days = st.number_input("기준 미거래일 (일)", value=st.session_state.get('risk_days', 45), step=5, key="risk_day_main")
                st.caption(f"👉 **적용 조건: 최근 {risk_days}일 발주 없음**")
        
        if not df.empty and pd.notnull(df['일자'].max()):
            base_date = df['일자'].max()
            churn_analysis = df.groupby('판매처명').agg(총매출=('합계', 'sum'), 최근거래일=('일자', 'max'), 거래월수=('월', 'nunique')).reset_index()
            churn_analysis['월평균매출'] = churn_analysis['총매출'] / churn_analysis['거래월수'].replace(0, 1)
            churn_analysis['미거래일수'] = (base_date - churn_analysis['최근거래일']).dt.days
            risk_group = churn_analysis[(churn_analysis['월평균매출'] >= risk_revenue) & (churn_analysis['미거래일수'] >= risk_days)].sort_values('월평균매출', ascending=False)
            risk_group = risk_group[~risk_group['판매처명'].str.contains('OEM', na=False)].sort_values('월평균매출', ascending=False)
            
            if not risk_group.empty:
                show_risk = risk_group[['판매처명', '미거래일수', '월평균매출']].copy()
                st.dataframe(show_risk.style.format({'월평균매출': format_krw, '미거래일수': '{:.0f}일 경과'}), hide_index=True, use_container_width=True, height=320)
            else:
                st.success(f"✅ 특이사항 없음 (기준: 월 {risk_revenue//10000:,.0f}만 / {risk_days}일)")

    st.markdown("---")
    
    # (급등/급락) 알림 섹션 - 좌/우 분리형
    st.markdown("**🚨 전월 대비 특이사항 (매출 급등 / 급락 거래처)**")
    # 특이사항 표시를 위한 동적 필터
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        spike_limit = st.number_input("📈 급등 필터 (증감액 만원 이상)", value=1000, step=100, key="spike_main") * 10000
    with f_col2:
        drop_limit = st.number_input("📉 급락 필터 (감소액 만원 이상)", value=500, step=100, key="drop_main") * 10000
        
    c_up, c_down = st.columns(2)
    up_html, down_html = "", ""
    
    if not prev_df.empty and not curr_df.empty:
        cust_curr = curr_df.groupby('판매처명')['합계'].sum()
        cust_prev = prev_df.groupby('판매처명')['합계'].sum()
        cust_merged = pd.DataFrame({'당월': cust_curr, '전월': cust_prev}).fillna(0)
        
        # OEM 제외 및 전월 100만 원 이상 필터 고정 적용
        cust_merged = cust_merged[(~cust_merged.index.str.contains('OEM', na=False)) & (cust_merged['전월'] >= 1000000)]
        cust_merged['증감액'] = cust_merged['당월'] - cust_merged['전월']
        
        cust_merged['증감률'] = (cust_merged['증감액'] / cust_merged['전월']) * 100
        
        # 독립된 필터 적용
        spikes = cust_merged[cust_merged['증감액'] >= spike_limit].sort_values('증감액', ascending=False)
        drops = cust_merged[cust_merged['증감액'] <= -drop_limit].sort_values('증감액')
        
        for c, r in spikes.iterrows():
            up_html += f"<div style='color:#B91C1C; background-color:#FEF2F2; padding:10px; margin-bottom:5px; border-left:4px solid #DC2626;'>📈 <b>{c}</b><br>당월: {r['당월']/10000:,.0f}만원 (전월: {r['전월']/10000:,.0f}만원 / <b>▲{r['증감액']/10000:,.0f}만원</b> / {r['증감률']:,.0f}%)</div>"
        for c, r in drops.iterrows():
            down_html += f"<div style='color:#1D4ED8; background-color:#EFF6FF; padding:10px; margin-bottom:5px; border-left:4px solid #2563EB;'>📉 <b>{c}</b><br>당월: {r['당월']/10000:,.0f}만원 (전월: {r['전월']/10000:,.0f}만원 / <b>▼{abs(r['증감액'])/10000:,.0f}만원</b> / {r['증감률']:,.0f}%)</div>"

    with c_up: st.markdown(up_html if up_html else "✅ 급등 업체가 없습니다.", unsafe_allow_html=True)
    with c_down: st.markdown(down_html if down_html else "✅ 급락 업체가 없습니다.", unsafe_allow_html=True)

    st.markdown("---")

    # 5. 월별 매출 추이
    st.markdown("**📈 월별 매출 추이**")
    main_trend = df[df['월'] > 0].groupby('월')['합계'].sum().reset_index()
    main_trend = main_trend[main_trend['합계'] > 0].sort_values('월')
    main_trend['월'] = main_trend['월'].astype(str) + '월'
    main_trend['합계_억'] = main_trend['합계'] / 100000000

    main_trend['합계_표기'] = main_trend['합계'].apply(format_krw)
    main_trend['합계_축약표기'] = main_trend['합계'].apply(format_short_krw)

    fig_main_trend = px.line(main_trend, x='월', y='합계_억', custom_data=['합계_표기'], markers=True, color_discrete_sequence=['#2563EB'])
    fig_main_trend.update_traces(
        mode='lines+markers+text', 
        text=main_trend['합계_축약표기'],
        line=dict(width=5), 
        textposition='top center', 
        textfont=dict(color='black', size=14, weight='bold'), 
        hovertemplate="<b>%{x}</b><br>매출액: %{customdata[0]}<extra></extra>"
    )
    fig_main_trend.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)],range=[-0.5, 11.5])
    fig_main_trend.update_yaxes(title="매출액 (단위: 억원)")
    fig_main_trend.update_layout(template="plotly_white")
    st.plotly_chart(fig_main_trend, width="stretch", key="main_trend")

    c_m1, c_m2 = st.columns(2)
    with c_m1:
        st.markdown("**🏆 매출 Top 10 거래처**")
        top_cust = curr_df.groupby('판매처명')['합계'].sum().nlargest(10).reset_index()
        top_cust['합계_천만'] = top_cust['합계'] / 10000000
        top_cust['합계_풀표기'] = top_cust['합계'].apply(format_krw)
        top_cust['합계_축약표기'] = top_cust['합계'].apply(format_short_krw)
        
        fig_t1 = px.bar(top_cust, x='합계_천만', y='판매처명', orientation='h', text='합계_축약표기', custom_data=['합계_풀표기'], color_discrete_sequence=['#3B82F6'])

        fig_t1.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>총 매출액: %{customdata[0]}<extra></extra>" # 툴팁에는 합계_표기 노출
        )
        fig_t1.update_yaxes(categoryorder='total ascending')
        fig_t1.update_xaxes(title="매출액 (단위: 천만원)")
        fig_t1.update_layout(template="plotly_white")
        st.plotly_chart(fig_t1, width="stretch", key="m_chart_1")
        
    with c_m2:
        st.markdown("**📦 품목 Top 10 (금액 기준)**")
        top_item = curr_df.groupby('품명 및 규격')['합계'].sum().nlargest(10).reset_index()
        top_item['합계_천만'] = top_item['합계'] / 10000000
        top_item['합계_풀표기'] = top_item['합계'].apply(format_krw)
        top_item['합계_축약표기'] = top_item['합계'].apply(format_short_krw)
        
        fig_t2 = px.bar(top_item, x='합계_천만', y='품명 및 규격', orientation='h', text='합계_축약표기', custom_data=['합계_풀표기'], color_discrete_sequence=['#3B82F6'])
        fig_t2.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        fig_t2.update_yaxes(categoryorder='total ascending')
        fig_t2.update_xaxes(title="매출액 (단위: 천만원)")
        fig_t2.update_layout(template="plotly_white")
        st.plotly_chart(fig_t2, width="stretch", key="m_chart_2")

    c_m3, c_m4 = st.columns(2)
    with c_m3:
        st.markdown("**📊 품목별 매출 비중**")
        pie_df = curr_df[~curr_df['카테고리'].isin(['기타', '배송비'])].groupby('카테고리')['합계'].sum().reset_index()
        pie_df['합계_표기'] = pie_df['합계'].apply(format_krw)
        
        fig_t3 = px.pie(pie_df, names='카테고리', values='합계', hole=0.4, custom_data=['합계_표기'])
        fig_t3.update_traces(
            texttemplate='<b>%{label}</b><br><b>%{percent}</b>', 
            textfont=dict(size=14, color='white'), 
            hovertemplate="<b>%{label}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig_t3, width="stretch", key="m_chart_3")
    with c_m4:
        st.markdown("**🚚 배송 유형 비중 (수량)**")
        del_df = curr_df[curr_df['배송유형'] != '기타']
        fig_t4 = px.pie(del_df.groupby('배송유형')['수량'].sum().reset_index(), names='배송유형', values='수량', hole=0.4)
        fig_t4.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(size=14, color='white'), hovertemplate="<b>%{label}</b><br>수량: %{value:,.0f}개<extra></extra>")
        st.plotly_chart(fig_t4, width="stretch", key="m_chart_4")

    # 색온도 매핑 제거 (추출된 65K, 3색변환을 그대로 사용)
    light_colors = {'65K': '#00BCD4', '57K': '#0288D1', '40K': '#9E9E9E', '30K': '#F57C00', '27K': '#E65100', '3색변환': '#9C27B0', '기타': '#EAEAEA'}

    c_m5, c_m6 = st.columns(2)
    with c_m5:
        st.markdown("**💡 색온도별 출고 비중 (수량)**")
        fig_t5 = px.pie(curr_df, names='색온도', values='수량', color='색온도', color_discrete_map=light_colors)
        fig_t5.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(size=14, color='white'), hovertemplate="<b>%{label}</b><br>수량: %{value:,.0f}개<extra></extra>")
        st.plotly_chart(fig_t5, width="stretch", key="m_chart_5")
    with c_m6:
        st.markdown("**📈 색온도 출고 추이 (수량 기준)**")
        color_trend = df[df['월'] > 0].copy()
        color_trend = color_trend.groupby(['월', '색온도'])['수량'].sum().reset_index()
        color_trend['월'] = color_trend['월'].astype(str) + '월'
        
        fig_t6 = px.line(color_trend, x='월', y='수량', color='색온도', color_discrete_map=light_colors, markers=True)
        fig_t6.update_traces(
            mode='lines+markers', 
            line=dict(width=4), 
            hovertemplate="<b>%{fullData.name}</b><br>수량: %{y:,.0f}개<extra></extra>"
        )
        fig_t6.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)],range=[-0.5, 11.5])
        fig_t6.update_layout(template="plotly_white")
        st.plotly_chart(fig_t6, width="stretch", key="m_chart_6")

    st.markdown("**📈 품목별 출고 추이 (수량 기준 Top 5 핵심 품목)**")
    top_5_cats = curr_df.groupby('카테고리')['수량'].sum().nlargest(5).index.tolist()
    df_cat = df[(df['월'] > 0) & (df['카테고리'].isin(top_5_cats))].copy()
    cat_qty_trend = df_cat.groupby(['월', '카테고리'])['수량'].sum().reset_index()
    cat_qty_trend = cat_qty_trend[cat_qty_trend['수량'] > 0].sort_values('월')
    cat_qty_trend['월'] = cat_qty_trend['월'].astype(str) + '월'
    
    fig_t7 = px.line(cat_qty_trend, x='월', y='수량', color='카테고리', markers=True)
    fig_t7.update_traces(mode='lines+markers', line=dict(width=4), hovertemplate="<b>%{fullData.name}</b><br>수량: %{y:,.0f}개<extra></extra>")
    fig_t7.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)], range=[-0.5, 11.5])
    fig_t7.update_layout(template="plotly_white")
    st.plotly_chart(fig_t7, width="stretch", key="m_chart_7")

    st.markdown("---")
    st.markdown("**👥 담당자별 매출 현황 (금액 기준)**")
    if not curr_df.empty:
        clean_rep_df = curr_df[~curr_df['담당자명'].isin(st.session_state.deleted_reps)]
        rep_sales_df = clean_rep_df.groupby('담당자명')['합계'].sum().reset_index().sort_values('합계', ascending=False)
        rep_sales_df['합계_천만'] = rep_sales_df['합계'] / 10000000
        rep_sales_df['합계_풀표기'] = rep_sales_df['합계'].apply(format_krw)
        rep_sales_df['합계_축약표기'] = rep_sales_df['합계'].apply(format_short_krw)
        
        fig_rep_main = px.bar(
            rep_sales_df, 
            x='합계_천만', 
            y='담당자명', 
            orientation='h', 
            text='합계_축약표기', 
            custom_data=['합계_풀표기'], 
            color_discrete_sequence=['#6366F1']
        )
        fig_rep_main.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(size=13),
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig_rep_main.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="매출액 (단위: 천만원)"), width="stretch")

elif menu == "👨‍💼 담당자별 분석" and df is not None:
    valid_months = sorted(df[df['월'] > 0]['월'].unique())
    rep_list = [r for r in df['담당자명'].unique() if str(r).strip() not in ['미분류', 'nan', 'NAN', 'NA', 'None', ''] and r not in st.session_state.deleted_reps]

    is_range_rep = st.checkbox("🔄 기간으로 검색 (시작월 ~ 종료월)", value=False, key="range_rep")
    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 3, 3])
    
    if not is_range_rep:
        with col_f1:
            sel_month = st.selectbox("📅 월 선택", valid_months, index=len(valid_months)-1, format_func=lambda x: f"{int(x)}월", key="rep_m")
            sel_start, sel_end = sel_month, sel_month
    else:
        with col_f1: sel_start = st.selectbox("🟢 시작 월", valid_months, key="rep_s", format_func=lambda x: f"{int(x)}월")
        with col_f2: sel_end = st.selectbox("🔴 종료 월", [m for m in valid_months if m >= sel_start], key="rep_e", format_func=lambda x: f"{int(x)}월")

    with col_f3: t2_rep = st.selectbox("👨‍💼 담당자 선택", rep_list)
    # NA, 미분류 제거된 지역 리스트
    clean_regs = ['전체'] + sorted([
        str(r) for r in df[df['담당자명']==t2_rep]['지역'].unique() 
        if pd.notnull(r) and str(r).strip().upper() not in ['미분류', 'NAN', 'NA', 'NONE', '']
    ])
    with col_f4: sel_reg = st.selectbox("🗺️ 지역 필터", clean_regs)

    curr_df_all = df[(df['월'] >= sel_start) & (df['월'] <= sel_end)]
    curr_df = curr_df_all[curr_df_all['담당자명'] == t2_rep]
    if sel_reg != '전체': curr_df = curr_df[curr_df['지역'] == sel_reg]
    prev_df = df[(df['월'] >= sel_start - 1) & (df['월'] <= sel_end - 1) & (df['담당자명'] == t2_rep)]
    if sel_reg != '전체': prev_df = prev_df[prev_df['지역'] == sel_reg]

    rep_sales = curr_df['합계'].sum()
    mom_diff, mom_pct = calc_mom(curr_df, prev_df)
    total_sales_all = curr_df_all['합계'].sum() if sel_reg == '전체' else curr_df_all[curr_df_all['지역'] == sel_reg]['합계'].sum()
    share_rate = (rep_sales / total_sales_all * 100) if total_sales_all > 0 else 0
    
    mom_sign = "+" if mom_pct > 0 else ""
    mom_display = f"{mom_sign}{mom_pct:.1f}% \n ({format_krw(mom_diff)})" if mom_diff != 0 else "-"

    if mom_diff > 0:
        mom_display = f"▲ {format_krw(abs(mom_diff))} (+{mom_pct:.1f}%)"
    elif mom_diff < 0:
        mom_display = f"▼ {format_krw(abs(mom_diff))} ({mom_pct:.1f}%)"
    else:
        mom_display = "-"

    # 💰 KPI 카드
    c1, c2, c3, c4 = st.columns(4)
    c1.button(f"💰 담당자 매출액\n{format_krw(rep_sales)}", use_container_width=True, key="r_k1")
    c2.button(f"📈 전월 대비\n{mom_display}", use_container_width=True, key="r_k2")
    c3.button(f"📊 전체 매출 내 비중\n{share_rate:.1f}%", use_container_width=True, key="r_k3")
    c4.button(f"🤝 활성 거래처 수\n{curr_df['판매처명'].nunique():,}개", use_container_width=True, key="r_k4")

    # 💡 3번 사항 반영: 상승/하락에 따른 동적 색상 지정
    trend_color_rep = "#E11D48" if mom_diff > 0 else "#2563EB" if mom_diff < 0 else "#475569"
    trend_word_rep = "상승세" if mom_diff > 0 else "하락세" if mom_diff < 0 else "유지"

    st.markdown("---")
    
    if not curr_df.empty:
        # 거래처별 매출 집계 및 요약 지표
        cust_group = curr_df.groupby('판매처명')['합계'].sum()
        rep_top_cust = cust_group.idxmax()
        rep_top_cust_sales = cust_group.max() 
        rep_top_cust_ratio = (rep_top_cust_sales / rep_sales * 100) if rep_sales > 0 else 0
        
        active_cust_count = curr_df['판매처명'].nunique()
        avg_cust_sales = rep_sales / active_cust_count if active_cust_count > 0 else 0
        
        ytd_rep_df = df[(df['월'] >= 1) & (df['월'] <= sel_end) & (df['담당자명'] == t2_rep)]
        ytd_rep_sales = ytd_rep_df['합계'].sum()
        rep_goal = st.session_state.rep_goals.get(t2_rep, 100000000)
        
        gap_sales = rep_goal - ytd_rep_sales
        if gap_sales > 0:
            gap_msg = f"연간 목표({format_krw(rep_goal)}) 달성까지 <b>{format_krw(gap_sales)}</b> 남았습니다."
        else:
            gap_msg = f"연간 목표를 <b>{format_krw(abs(gap_sales))} 초과 달성</b>하였습니다."
        
        # 💡 HTML 텍스트 내에 trend_color_rep 변수를 적용하여 빨강/파랑 자동 색상화
        briefing_rep_html = f"""
        <div style="background-color: #F0F9FF; border-left: 5px solid #0EA5E9; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h4 style="margin-top: 0; color: #0C4A6E;">👨‍💼 {t2_rep} 담당자 성과 요약</h4>
            <ul style="margin-bottom: 0; color: #075985; font-size: 14px; line-height: 1.6;">
                <li>당월 총 매출액은 <b>{format_krw(rep_sales)}</b>이며, 전월 대비 <span style="color:{trend_color_rep}; font-weight:bold;">{mom_display}</span>로 <span style="color:{trend_color_rep}; font-weight:bold;">{trend_word_rep}</span>입니다.</li>
                <li>관리 중인 거래처 중 <b>'{rep_top_cust}'</b> (매출액: {format_krw(rep_top_cust_sales)}, 담당자 내 매출 비중: <b>{rep_top_cust_ratio:.1f}%</b>)에서 가장 높은 실적을 보이고 있습니다.</li>
                <li>담당 거래처 1개사당 평균 매출은 <b>{format_krw(avg_cust_sales)}</b>이며, 전체 매출 내 비중은 <b>{share_rate:.1f}%</b> 입니다.</li>
                <li style="color: #0369A1; font-weight: 500;"><b>목표 진척도:</b> {gap_msg}</li>
            </ul>
        </div>
        """
        st.markdown(briefing_rep_html, unsafe_allow_html=True)
    
    # 지도와 담당자 개인 YTD 게이지 배치
    c_map, c_ytd = st.columns([1, 1])
    with c_map:
        st.markdown("**🗺️ 담당자 관할 지역 (지도)**")
        reg_sales = curr_df[curr_df['지역'] != '미분류'].groupby('지역')['합계'].sum().reset_index()
        kor_map = {'서울':'서울특별시','부산':'부산광역시','대구':'대구광역시','인천':'인천광역시','광주':'광주광역시','대전':'대전광역시','울산':'울산광역시','세종':'세종특별자치시','경기':'경기도','강원':'강원도','충북':'충청북도','충남':'충청남도','전북':'전라북도','전남':'전라남도','경북':'경상북도','경남':'경상남도','제주':'제주특별자치도'}
        reg_sales['fullname'] = reg_sales['지역'].map(kor_map)
        try:
            geojson_url = "https://raw.githubusercontent.com/southkorea/southkorea-maps/master/kostat/2013/json/skorea_provinces_geo_simple.json"
            geojson = requests.get(geojson_url).json()
            fig_map = px.choropleth_mapbox(reg_sales.dropna(subset=['fullname']), geojson=geojson, locations='fullname', featureidkey='properties.name', color='합계', color_continuous_scale="Blues", mapbox_style="carto-positron", zoom=5.5, center={"lat": 35.9, "lon": 127.7}, opacity=0.8)
            fig_map.update_layout(margin={"r":0,"t":0,"l":0,"b":0})
            fig_map.update_traces(hovertemplate="<b>%{location}</b><br>매출액: %{z:,.0f}원<extra></extra>")
            st.plotly_chart(fig_map, width="stretch", key="rep_map")
        except:
            st.warning("지도 로드 실패")

    with c_ytd:
        st.markdown(f"**🎯 [{t2_rep}] 개인 연간 목표 달성률 (YTD)**")
        ytd_rep_df = df[(df['월'] >= 1) & (df['월'] <= sel_end) & (df['담당자명'] == t2_rep)]
        ytd_rep_sales = ytd_rep_df['합계'].sum()
        # 아직 개별 목표가 없다면 임시로 1억으로 세팅하여 모양을 보여줌
        rep_goal = st.session_state.rep_goals.get(t2_rep, 100000000) 
        ytd_rep_pct = (ytd_rep_sales / rep_goal) * 100 if rep_goal > 0 else 0

        fig_rep_gauge = go.Figure(go.Indicator(
            mode = "gauge+number+delta", value = ytd_rep_pct,
            number = {'suffix': "%", 'font': {'size': 40, 'color': '#0F172A', 'family': 'Pretendard'}},
            domain = {'x': [0, 1], 'y': [0, 1]},
            gauge = {
                'axis': {'range': [None, 100], 'tickwidth': 1, 'tickcolor': "darkblue"},
                'bar': {'color': "#059669"}, # 영업사원 목표는 산뜻한 초록색 계열로
                'bgcolor': "white", 'borderwidth': 2, 'bordercolor': "#E2E8F0",
                'steps': [{'range': [0, 50], 'color': '#FEE2E2'}, {'range': [50, 80], 'color': '#FEF3C7'}, {'range': [80, 100], 'color': '#D1FAE5'}],
                'threshold': {'line': {'color': "red", 'width': 4}, 'thickness': 0.75, 'value': 100}
            }
        ))
        fig_rep_gauge.update_layout(height=280, margin=dict(l=20, r=20, t=20, b=20))
        st.plotly_chart(fig_rep_gauge, width="stretch", key="rep_gauge")
        st.markdown(f"<div style='text-align:center; color:#64748B; font-weight:600;'>누적 매출: {format_krw(ytd_rep_sales)} / 개인 할당량: {format_krw(rep_goal)}</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(f"**🚨 [{t2_rep}] 담당 거래처 특이사항 (급등/급락)**")
    f_col1, f_col2 = st.columns(2)
    with f_col1: rep_spike_limit = st.number_input("📈 급등 필터 (증감액 만원 이상)", value=1000, step=100, key="rep_spike") * 10000
    with f_col2: rep_drop_limit = st.number_input("📉 급락 필터 (감소액 만원 이상)", value=500, step=100, key="rep_drop") * 10000
        
    c_up, c_down = st.columns(2)
    up_html, down_html = "", ""
    
    if not prev_df.empty and not curr_df.empty:
        c_curr = curr_df.groupby('판매처명')['합계'].sum()
        c_prev = prev_df.groupby('판매처명')['합계'].sum()
        merged = pd.DataFrame({'당월': c_curr, '전월': c_prev}).fillna(0)
        
        merged = merged[(~merged.index.str.contains('OEM', na=False)) & (merged['전월'] >= 1000000)]
        merged['증감액'] = merged['당월'] - merged['전월']
        
        merged['증감률'] = (merged['증감액'] / merged['전월']) * 100
        
        spikes = merged[merged['증감액'] >= rep_spike_limit].sort_values('증감액', ascending=False)
        drops = merged[merged['증감액'] <= -rep_drop_limit].sort_values('증감액')

        for c, r in spikes.iterrows():
            up_html += f"<div style='color:#B91C1C; background-color:#FEF2F2; padding:10px; margin-bottom:5px; border-left:4px solid #DC2626;'>📈 <b>{c}</b><br>당월: {r['당월']/10000:,.0f}만 (전월: {r['전월']/10000:,.0f}만 / <b>▲{r['증감액']/10000:,.0f}만</b> / {r['증감률']:,.0f}%)</div>"
        for c, r in drops.iterrows():
            down_html += f"<div style='color:#1D4ED8; background-color:#EFF6FF; padding:10px; margin-bottom:5px; border-left:4px solid #2563EB;'>📉 <b>{c}</b><br>당월: {r['당월']/10000:,.0f}만 (전월: {r['전월']/10000:,.0f}만 / <b>▼{abs(r['증감액'])/10000:,.0f}만</b> / {r['증감률']:,.0f}%)</div>"

    with c_up: st.markdown(up_html if up_html else "✅ 설정한 기준 이상의 급등 업체가 없습니다.", unsafe_allow_html=True)
    with c_down: st.markdown(down_html if down_html else "✅ 설정한 기준 이상의 급락 업체가 없습니다.", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**📈 담당자 월별 매출 추이**")
    rep_trend = df[(df['담당자명'] == t2_rep) & (df['월'] > 0)]
    if sel_reg != '전체': rep_trend = rep_trend[rep_trend['지역'] == sel_reg]
    rep_trend = rep_trend.groupby('월')['합계'].sum().reset_index()
    rep_trend = rep_trend[rep_trend['합계'] > 0].sort_values('월') # 빈 달 방지
    
    rep_trend['월'] = rep_trend['월'].astype(str) + '월'
    rep_trend['합계_표기'] = rep_trend['합계'].apply(format_krw)
    rep_trend['합계_축약표기'] = rep_trend['합계'].apply(format_short_krw)
    rep_trend['합계_억'] = rep_trend['합계'] / 100000000
    
    fig_rep = px.line(rep_trend, x='월', y='합계_억', custom_data=['합계_표기'], markers=True, color_discrete_sequence=['#2563EB'])
    
    fig_rep.update_traces(
        mode='lines+markers+text', 
        text=rep_trend['합계_축약표기'], 
        line=dict(width=5), 
        textposition='top center', 
        textfont=dict(color='black', size=13, weight='bold'),
        hovertemplate="<b>%{x}월</b><br>매출액: %{customdata[0]}<extra></extra>"
    )
    # y축 제목 추가 (단위: 억원)
    fig_rep.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)], range=[-0.5, 11.5])
    fig_rep.update_layout(yaxis_title="매출액 (단위: 억원)", template="plotly_white", margin=dict(t=20, b=0, l=0, r=0))
    st.plotly_chart(fig_rep, width="stretch")
    
    st.markdown("---")
    c5, c6 = st.columns(2)
    with c5:
        st.markdown("**🏆 매출 Top 10 거래처**")
        
        top_c = curr_df.groupby('판매처명')['합계'].sum().nlargest(10).reset_index()
        top_c['합계_표기'] = top_c['합계'].apply(format_krw)
        top_c['합계_축약표기'] = top_c['합계'].apply(format_short_krw) # 💡 축약표기 추가
        top_c['합계_천만'] = top_c['합계'] / 10000000
        
        fig1 = px.bar(top_c, x='합계_천만', y='판매처명', orientation='h', text='합계_축약표기', custom_data=['합계_표기'], color_discrete_sequence=['#3B82F6'])
        fig1.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        # x축 제목 추가 (단위: 천만원)
        st.plotly_chart(fig1.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="매출액 (단위: 천만원)"), width="stretch")
    with c6:
        st.markdown("**📊 품목별 매출 비중**")
        # 기타, 배송비 제외
        pie_df = curr_df[~curr_df['카테고리'].isin(['기타', '배송비'])]
        fig_pie1 = px.pie(pie_df.groupby('카테고리')['합계'].sum().reset_index(), names='카테고리', values='합계', hole=0.4)
        fig_pie1.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(color='white', size=14), hovertemplate="<b>%{label}</b><br>매출액: %{value:,.0f}원<extra></extra>")
        st.plotly_chart(fig_pie1, width="stretch")

    st.markdown("---")
    c7, c8 = st.columns(2)
    with c7:
        st.markdown("**📦 품목 Top 10 (금액 기준)**")
        top_i = curr_df.groupby('품명 및 규격')['합계'].sum().nlargest(10).reset_index()
        top_i['합계_표기'] = top_i['합계'].apply(format_krw)
        top_i['합계_축약표기'] = top_i['합계'].apply(format_short_krw)
        top_i['합계_천만'] = top_i['합계'] / 10000000
        
        fig2 = px.bar(top_i, x='합계_천만', y='품명 및 규격', orientation='h', text='합계_축약표기', custom_data=['합계_표기'])
        fig2.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig2.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="매출액 (단위: 천만원)"), width="stretch")
    
    with c8:
        st.markdown("**🔢 품목 Top 10 (수량 기준)**")
        top_qty = curr_df.groupby('품명 및 규격')['수량'].sum().nlargest(10).reset_index()
        fig_q = px.bar(top_qty, x='수량', y='품명 및 규격', orientation='h', text='수량', color_discrete_sequence=['#10B981'])
        fig_q.update_traces(
            texttemplate='<b>%{text:,.0f}개</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>수량: %{text:,.0f}개<extra></extra>"
        )
        st.plotly_chart(fig_q.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="수량 (개)"), width="stretch")

    st.markdown("---")

    # 담당자의 품목별 100% 누적 비중 추이 차트
    st.markdown(f"**📈 [{t2_rep}] 품목별 매출 비중 변화 추이**")
    
    # 전체 기간 동안의 Top 5 카테고리를 먼저 추출하여 고정
    rep_df_valid = df[(df['담당자명'] == t2_rep) & (df['월'] > 0) & (~df['카테고리'].isin(['배송비', '기타']))]
    top_5_cats = rep_df_valid.groupby('카테고리')['합계'].sum().nlargest(5).index.tolist()
    
    # Top 5에 속하지 않는 카테고리는 모두 '기타(Top 5 외)'로 묶음
    rep_df_valid = rep_df_valid.copy()
    rep_df_valid['카테고리_표시'] = rep_df_valid['카테고리'].apply(lambda x: x if x in top_5_cats else '기타(Top 5 외)')
    
    cat_rep_trend = rep_df_valid.groupby(['월', '카테고리_표시'])['합계'].sum().reset_index()
    cat_rep_trend['월명'] = cat_rep_trend['월'].astype(str) + '월'
    
    monthly_tot_rep = cat_rep_trend.groupby('월명')['합계'].transform('sum')
    cat_rep_trend['비중'] = (cat_rep_trend['합계'] / monthly_tot_rep * 100).fillna(0)
    cat_rep_trend['텍스트'] = cat_rep_trend.apply(lambda r: f"<b>{r['카테고리_표시']}</b><br>{r['비중']:.0f}%" if r['비중'] >= 10 else "", axis=1)

    category_order = top_5_cats + ['기타(Top 5 외)']

    fig_rep_cat = px.bar(cat_rep_trend.sort_values('월'), x='월명', y='비중', color='카테고리_표시', 
                         custom_data=['카테고리_표시', '합계'], text='텍스트',
                         category_orders={"카테고리_표시": category_order},
                         labels={'월명': '월', '카테고리_표시': '카테고리'})
                         
    fig_rep_cat.update_traces(
        textposition='inside', 
        textfont=dict(color='white', size=13, weight='bold'), 
        hovertemplate="<b>%{customdata[0]}</b><br>매출액: %{customdata[1]:,.0f}원<br>비중: %{y:.1f}%<extra></extra>"
    )
    
    fig_rep_cat.update_layout(
        barmode='stack', template="plotly_white", yaxis_title="매출 비중 (%)", height=450,
        font=dict(family="Pretendard", weight="bold", color="#334155") # 전체 폰트 굵게
    )
    fig_rep_cat.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)], range=[-0.5, 11.5])
    st.plotly_chart(fig_rep_cat, use_container_width=True)

elif menu == "🏢 거래처별 분석" and df is not None:
    valid_months = sorted(df[df['월'] > 0]['월'].unique())
    is_range_cust = st.checkbox("🔄 기간으로 검색 (시작월 ~ 종료월)", value=False, key="range_cust")
    
    col_c1, col_c2, col_c3 = st.columns([2, 2, 6])
    if not is_range_cust:
        with col_c1:
            sel_month = st.selectbox("📅 월 선택", valid_months, index=len(valid_months)-1, format_func=lambda x: f"{int(x)}월", key="cust_m")
            sel_start, sel_end = sel_month, sel_month
    else:
        with col_c1: sel_start = st.selectbox("🟢 시작 월", valid_months, format_func=lambda x: f"{int(x)}월", key="cust_s")
        with col_c2: sel_end = st.selectbox("🔴 종료 월", [m for m in valid_months if m >= sel_start], format_func=lambda x: f"{int(x)}월", key="cust_e")

    with col_c3: 
        t3_cust = st.selectbox("🔍 거래처명 선택", sorted(df['판매처명'].unique()))
    
    curr_df_all = df[(df['월'] >= sel_start) & (df['월'] <= sel_end)]
    curr_df = curr_df_all[curr_df_all['판매처명'] == t3_cust]
    prev_df = df[(df['월'] >= sel_start - 1) & (df['월'] <= sel_end - 1) & (df['판매처명'] == t3_cust)]
    
    cust_sales = curr_df['합계'].sum()
    mom_diff, mom_pct = calc_mom(curr_df, prev_df)
    
    grade = "A (VIP)" if cust_sales >= 30000000 else "B (우수)" if cust_sales >= 10000000 else "C (일반)" if cust_sales >= 3000000 else "D (소액)"
    rep_name = curr_df['담당자명'].iloc[0] if not curr_df.empty else "미지정"
    region_name = curr_df['지역'].iloc[0] if not curr_df.empty else "미지정"
    
    rank_df = curr_df_all.groupby('판매처명')['합계'].sum().rank(ascending=False, method='min')
    total_rank = int(rank_df.get(t3_cust, 0))
    reg_df = curr_df_all[curr_df_all['지역'] == region_name].groupby('판매처명')['합계'].sum().rank(ascending=False, method='min')
    reg_rank = int(reg_df.get(t3_cust, 0))
    share_total = (cust_sales / curr_df_all['합계'].sum() * 100) if curr_df_all['합계'].sum() > 0 else 0

    st.markdown(f"#### 📌 [{t3_cust}] 거래처 핵심 요약")
    
    # 전월 대비 계산
    cust_sales = curr_df['합계'].sum()
    prev_sales = prev_df['합계'].sum()
    mom_diff = cust_sales - prev_sales
    mom_pct = (mom_diff / prev_sales * 100) if prev_sales > 0 else 0
    mom_txt = f"▲{mom_diff:,.0f}원 (+{mom_pct:.1f}%)" if mom_diff > 0 else f"▼{abs(mom_diff):,.0f}원 ({mom_pct:.1f}%)" if mom_diff < 0 else "-"
    
    # 랭킹 및 점유율 계산
    grade = "A (VIP)" if cust_sales >= 30000000 else "B (우수)" if cust_sales >= 10000000 else "C (일반)" if cust_sales >= 3000000 else "D (소액)"
    rep_name = curr_df['담당자명'].iloc[0] if not curr_df.empty else "미지정"
    region_name = curr_df['지역'].iloc[0] if not curr_df.empty else "미지정"
    
    total_rank = int(curr_df_all.groupby('판매처명')['합계'].sum().rank(ascending=False, method='min').get(t3_cust, 0))
    reg_rank = int(curr_df_all[curr_df_all['지역'] == region_name].groupby('판매처명')['합계'].sum().rank(ascending=False, method='min').get(t3_cust, 0))
    
    total_comp_sales = curr_df_all['합계'].sum()
    share_total = (cust_sales / total_comp_sales * 100) if total_comp_sales > 0 else 0
    
    rep_total_sales = curr_df_all[curr_df_all['담당자명'] == rep_name]['합계'].sum()
    share_rep = (cust_sales / rep_total_sales * 100) if rep_total_sales > 0 else 0

    s1, s2, s3 = st.columns(3)
    s1.info(f"**👨‍💼 주 담당자**\n### {rep_name}\n(지역: {region_name})")
    
    mom_diff_txt = format_krw(mom_diff)
    mom_display = f"▲{mom_diff_txt}" if mom_diff > 0 else f"▼{format_krw(abs(mom_diff))}" if mom_diff < 0 else "-"
    s2.success(f"**💰 당월 매출액**\n### {format_krw(cust_sales)}\n(전월 대비: {mom_display} / {mom_pct:.1f}%)")
    
    s3.markdown(f"""
    <div title="※ 기준: VIP(3천만↑), 우수(1천만↑), 일반(3백만↑), 소액(3백만↓)" 
         style="background-color: #FFF3CD; color: #856404; padding: 1.2rem; border-radius: 0.5rem; border: 1px solid #ffeeba;">
        <div style="font-weight: bold; margin-bottom: 3rem;">🏆 매출 등급</div>
        <div style="font-size: 1.8rem; font-weight: bold;">{grade}</div>
    </div>
    """, unsafe_allow_html=True)

    s4, s5, s6 = st.columns(3)
    s4.error(f"**🥇 매출 순위**\n### 전국 {total_rank}위\n({region_name} 내 {reg_rank}위)")
    s5.error(f"**🏢 전체 매출 점유율**\n### {share_total:.2f}%\n(전체 매출 대비)")
    s6.error(f"**🤝 담당자({rep_name}) 내 비중**\n### {share_rep:.2f}%\n(담당자 총매출 대비)")

    st.markdown("---")

    # 📈 연간 매출액 추이 (전체 흐름)
    st.markdown("**📈 연간 매출액 추이 (전체 흐름)**")
    cust_all = df[df['판매처명'] == t3_cust].groupby('월')['합계'].sum().reset_index()
    cust_all['월명'] = cust_all['월'].astype(str) + "월"
    cust_all['합계_억'] = cust_all['합계'] / 100000000
    cust_all['합계_풀표기'] = cust_all['합계'].apply(format_krw) 
    cust_all['합계_축약표기'] = cust_all['합계'].apply(format_short_krw)
    
    fig_c = px.line(cust_all, x='월명', y='합계_억', custom_data=['합계_풀표기'], markers=True)
    fig_c.update_traces(
        mode='lines+markers+text', 
        line=dict(width=5), 
        text=cust_all['합계_축약표기'], 
        textposition='top center', 
        textfont=dict(color='black', size=13, weight='bold'),
        hovertemplate="<b>%{x}</b><br>매출액: %{customdata[0]}<extra></extra>"
    )
    fig_c.update_yaxes(title="매출액 (단위: 억원)")
    fig_c.update_xaxes(title="월", type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)],range=[-0.5, 11.5])
    st.plotly_chart(fig_c.update_layout(template="plotly_white"), width="stretch")
    st.markdown("---")

    # 3. 품목별 비중 변화 (100% 누적 - 툴팁 정리)
    st.markdown("**📊 품목별 비중 변화 (100% 누적 수량 기준)**")
    cat_qty_trend = df[(df['판매처명'] == t3_cust) & (df['월'] > 0) & (~df['카테고리'].isin(['배송비', '기타']))].groupby(['월', '카테고리'])['수량'].sum().reset_index()
    cat_qty_trend['월'] = cat_qty_trend['월'].astype(str) + '월'
    
    monthly_total_cust = cat_qty_trend.groupby('월')['수량'].transform('sum')
    cat_qty_trend['비중'] = (cat_qty_trend['수량'] / monthly_total_cust * 100).fillna(0)
    
    cat_qty_trend['텍스트'] = cat_qty_trend.apply(lambda r: f"<b>{r['카테고리']}</b><br>{r['비중']:.0f}%" if r['비중'] >= 10 else "", axis=1)

    fig_cat_trend = px.bar(cat_qty_trend, x='월', y='비중', color='카테고리', custom_data=['카테고리', '수량'], text='텍스트')
    
    fig_cat_trend.update_traces(
        textposition='inside',
        textangle=0,
        textfont=dict(color='white', size=12),
        hovertemplate="<b>%{customdata[0]}</b><br>수량: %{customdata[1]:,.0f}개<br>비중: %{y:.1f}%<extra></extra>"
    )

    fig_cat_trend.update_layout(barmode='stack', template="plotly_white", yaxis_title="수량 비중 (%)")
    fig_cat_trend.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)],range=[-0.5, 11.5])
    st.plotly_chart(fig_cat_trend, width="stretch")

    st.markdown("---")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**🍩 품목별 매출 비중**")
        # 💡 기타, 배송비 제외
        pie_df = curr_df[~curr_df['카테고리'].isin(['기타', '배송비'])]
        fig_p1 = px.pie(pie_df.groupby('카테고리')['합계'].sum().reset_index(), names='카테고리', values='합계', hole=0.4)
        fig_p1.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(color='white', size=14), hovertemplate="<b>%{label}</b><br>매출액: %{value:,.0f}원<extra></extra>")
        st.plotly_chart(fig_p1, width="stretch")
    with c4:
        st.markdown("**🚚 배송 유형 비중**")
        fig_p2 = px.pie(curr_df[curr_df['배송유형'] != '기타'].groupby('배송유형')['수량'].sum().reset_index(), names='배송유형', values='수량', hole=0.4)
        fig_p2.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(color='white', size=14), hovertemplate="<b>%{label}</b><br>수량: %{value:,.0f}개<extra></extra>")
        st.plotly_chart(fig_p2, width="stretch")

    st.markdown("---")
    c5, c6 = st.columns(2)
    with c5:
        st.markdown("**📦 품목 Top 10 (금액)**")

        top_b1 = curr_df.groupby('품명 및 규격')['합계'].sum().nlargest(10).reset_index()
        top_b1['합계_표기'] = top_b1['합계'].apply(format_krw)
        top_b1['합계_축약표기'] = top_b1['합계'].apply(format_short_krw)
        top_b1['합계_천만'] = top_b1['합계'] / 10000000
        
        fig_b1 = px.bar(top_b1, x='합계_천만', y='품명 및 규격', orientation='h', text='합계_축약표기', custom_data=['합계_표기'])
        fig_b1.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig_b1.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="매출액 (단위: 천만원)"), width="stretch")
    with c6:
        st.markdown("**🔢 품목 Top 10 (수량)**")
        fig_b2 = px.bar(curr_df.groupby('품명 및 규격')['수량'].sum().nlargest(10).reset_index(), x='수량', y='품명 및 규격', orientation='h', text='수량', color_discrete_sequence=['#059669'])
        fig_b2.update_traces(texttemplate='<b>%{text:,.0f}개</b>', textposition='inside', textfont=dict(color='white', size=13), hovertemplate="<b>%{y}</b><br>수량: %{text:,.0f}개<extra></extra>")
        st.plotly_chart(fig_b2.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white"), width="stretch")

elif menu == "🗺️ 지역별 현황" and df is not None:
    st.title("🗺️ 지역별 통합 현황 분석")
    
    valid_months = sorted(df[df['월'] > 0]['월'].unique())
    is_range_reg = st.checkbox("🔄 기간으로 검색 (시작월 ~ 종료월)", value=False, key="range_reg")

    col_r1, col_r2, col_r3 = st.columns([2, 2, 5])

    if not is_range_reg:
        with col_r1:
            sel_month = st.selectbox("📅 월 선택", valid_months, index=len(valid_months)-1, format_func=lambda x: f"{int(x)}월", key="reg_m")
            sel_start, sel_end = sel_month, sel_month
    else:
        with col_r1: sel_start = st.selectbox("🟢 시작 월", valid_months, index=0, format_func=lambda x: f"{int(x)}월", key="reg_st")
        with col_r2: sel_end = st.selectbox("🔴 종료 월", [m for m in valid_months if m >= sel_start], index=len([m for m in valid_months if m >= sel_start])-1, format_func=lambda x: f"{int(x)}월", key="reg_en")
        
    reg_list = ['전국'] + sorted([str(r) for r in df['지역'].unique() if str(r).upper() not in ['미분류', 'NAN', 'NONE', 'NA']])
    with col_r3: t4_reg = st.selectbox("🗺️ 지역 선택", reg_list)
    
    curr_df_all = df[(df['월'] >= sel_start) & (df['월'] <= sel_end)]
    
    if t4_reg == '전국':
        curr_df = curr_df_all.copy()
        prev_df = df[(df['월'] >= sel_start - 1) & (df['월'] <= sel_end - 1)]
    else:
        curr_df = curr_df_all[curr_df_all['지역'] == t4_reg]
        prev_df = df[(df['월'] >= sel_start - 1) & (df['월'] <= sel_end - 1) & (df['지역'] == t4_reg)]
    
    reg_sales = curr_df['합계'].sum()
    mom_diff, mom_pct = calc_mom(curr_df, prev_df)
    reg_share = (reg_sales / curr_df_all['합계'].sum() * 100) if curr_df_all['합계'].sum() > 0 else 0
    reg_cust_cnt = curr_df['판매처명'].nunique()
    
    mom_display_reg = f"{mom_pct:+.1f}% \n ({format_krw(mom_diff)})" if mom_diff != 0 else "-"

    if mom_diff > 0:
        mom_display_reg = f"▲ {format_krw(abs(mom_diff))} (+{mom_pct:.1f}%)"
    elif mom_diff < 0:
        mom_display_reg = f"▼ {format_krw(abs(mom_diff))} ({mom_pct:.1f}%)"
    else:
        mom_display_reg = "-"

    c1, c2, c3, c4 = st.columns(4)
    c1.button(f"💰 지역 매출액\n{format_krw(reg_sales)}", use_container_width=True, key="reg_k1")
    c2.button(f"📈 전월 대비\n{mom_display_reg}", use_container_width=True, key="reg_k2")
    c3.button(f"🤝 활성 거래처 수\n{reg_cust_cnt:,}개", use_container_width=True, key="reg_k3")
    c4.button(f"📊 매출 점유율\n{reg_share:.1f}%", use_container_width=True, key="reg_k4")

    st.markdown("---")

    c_map_area, c_list_area = st.columns([5, 5]) 
    
    with c_map_area:
        st.markdown(f"**🗺️ [{t4_reg}] 매출액 분포 지도**")
        kor_map = {'서울':'서울특별시','부산':'부산광역시','대구':'대구광역시','인천':'인천광역시','광주':'광주광역시','대전':'대전광역시','울산':'울산광역시','세종':'세종특별자치시','경기':'경기도','강원':'강원도','충북':'충청북도','충남':'충청남도','전북':'전라북도','전남':'전라남도','경북':'경상북도','경남':'경상남도','제주':'제주특별자치도'}
        geojson_url = "https://raw.githubusercontent.com/southkorea/southkorea-maps/master/kostat/2013/json/skorea_provinces_geo_simple.json"
        
        try:
            geojson = requests.get(geojson_url).json()
            # 선택된 지역(curr_df)만 필터링하여 지도에 색상 표시
            map_data = curr_df[~curr_df['지역'].isin(['미분류', 'NA', 'NAN'])].groupby('지역')['합계'].sum().reset_index()
            map_data['fullname'] = map_data['지역'].map(kor_map)
            
            fig_map_main = px.choropleth_mapbox(
                map_data.dropna(), geojson=geojson, locations='fullname', featureidkey='properties.name', 
                color='합계', color_continuous_scale="Blues", mapbox_style="carto-positron", 
                zoom=5.5, # 💡 Zoom 고정 (확대 방지)
                center={"lat": 35.9, "lon": 127.7}, opacity=0.8
            )
            fig_map_main.update_traces(hovertemplate="<b>%{location}</b><br>매출액: %{z:,.0f}원<extra></extra>")
            fig_map_main.update_layout(margin={"r":0,"t":0,"l":0,"b":0}, height=550)
            st.plotly_chart(fig_map_main, width="stretch", key="region_main_map_v2")
        except:
            st.warning("지도 데이터를 불러올 수 없습니다.")

    with c_list_area:
        st.markdown(f"**🥇 [{t4_reg}] 지역 거래처 매출 순위**")
        if not curr_df.empty:
            loc_curr = curr_df[~curr_df['판매처명'].isin(['미분류', 'NAN', 'NA'])].groupby('판매처명')['합계'].sum()
            loc_prev = prev_df[~prev_df['판매처명'].isin(['미분류', 'NAN', 'NA'])].groupby('판매처명')['합계'].sum()
            
            loc_merged = pd.DataFrame({'당월매출': loc_curr, '전월매출': loc_prev}).fillna(0)
            
            # 성장률 계산
            loc_merged['성장률(%)'] = ((loc_merged['당월매출'] - loc_merged['전월매출']) / loc_merged['전월매출'].replace(0, 1)) * 100
            # 전월 매출이 없으면 '신규'로 표시하기 위한 플래그
            loc_merged.loc[(loc_merged['당월매출'] > 0) & (loc_merged['전월매출'] == 0), '성장률(%)'] = 999.9 
            
            loc_merged = loc_merged.sort_values('당월매출', ascending=False).reset_index()
            loc_merged.insert(0, '순위', range(1, len(loc_merged) + 1))
            
            def format_growth(val):
                if val == 999.9: return '✨신규'
                return f'{val:+.1f}%'

            def color_growth(val):
                if val == 999.9: return 'color: red; font-weight: bold;'
                # 💡 플러스(+) 빨강, 마이너스(-) 파랑 적용
                color = 'red' if val > 0 else 'blue' if val < 0 else 'black'
                return f'color: {color}; font-weight: bold;'
                
            styled_list = loc_merged.style \
                .format({'당월매출': format_krw, '전월매출': format_krw, '성장률(%)': format_growth}) \
                .map(color_growth, subset=['성장률(%)'])
            
            st.dataframe(styled_list, hide_index=True, use_container_width=True, height=550)

    st.markdown("---")
    
    # 💡 3. 월별 매출액(막대) + 성장률(선) 콤보 차트 
    st.markdown(f"**📊 [{t4_reg}] 월별 매출 및 성장률 추이**")
    
    df_all_reg = df if t4_reg == '전국' else df[df['지역'] == t4_reg]
    
    # 매출 데이터가 있는 마지막 달을 찾아 그 달까지만 X축 범위를 설정 (4월 이후 숨기기)
    last_active_month = df_all_reg[df_all_reg['합계'] > 0]['월'].max() if not df_all_reg.empty else 12
    all_months = pd.DataFrame({'월': list(range(1, 13))})
    
    reg_trend_raw = df_all_reg[df_all_reg['월'] > 0].groupby('월')['합계'].sum().reset_index()
    reg_trend = pd.merge(all_months, reg_trend_raw, on='월', how='left').fillna(0)
    
    reg_trend = reg_trend.sort_values('월')
    # 성장률 계산 (1월 데이터가 있다면 2월에 정상 표시됨)
    reg_trend['전월합계'] = reg_trend['합계'].shift(1)
    reg_trend['성장률'] = ((reg_trend['합계'] - reg_trend['전월합계']) / reg_trend['전월합계'].replace(0, np.nan)) * 100
    reg_trend.loc[reg_trend['월'] == 1, '성장률'] = 0.0
    reg_trend['월명'] = reg_trend['월'].astype(str) + '월'
    
    # 💡 [수정] 매출이 0인 달은 -100%로 보이지 않게 NaN 처리
    reg_trend['합계_억'] = reg_trend['합계'] / 100000000
    reg_trend.loc[reg_trend['합계'] <= 0, '합계_억'] = np.nan
    reg_trend.loc[reg_trend['합계'] <= 0, '성장률'] = np.nan
    
    reg_trend['합계_표기'] = reg_trend['합계'].apply(format_krw)
    
    # 막대그래프 텍스트용 축약 포맷
    def format_short(val):
        if val <= 0 or pd.isna(val): return ""
        uk = val // 100000000
        man = (val % 100000000) // 10000
        if uk > 0 and man > 0: return f"{int(uk)}억 {int(man):,}만"
        elif uk > 0: return f"{int(uk)}억"
        elif man > 0: return f"{int(man):,}만"
        return ""

    reg_trend['합계_축약'] = reg_trend['합계'].apply(format_short)
    
    max_sales = reg_trend['합계_억'].max() if pd.notna(reg_trend['합계_억'].max()) else 1
    abs_max_growth = reg_trend['성장률'].abs().max() if pd.notna(reg_trend['성장률'].abs().max()) else 10

    from plotly.subplots import make_subplots
    fig_combo = make_subplots(specs=[[{"secondary_y": True}]])
    
    # 매출액 막대 그래프
    fig_combo.add_trace(
        go.Bar(
            x=reg_trend['월명'], 
            y=reg_trend['합계_억'], 
            name="매출액 (억원)", 
            marker_color="#A5B4FC", 
            text=reg_trend['합계_축약'],
            textposition='inside',      
            insidetextanchor='middle',  
            textangle=0,                
            textfont=dict(color='#312E81', size=14, family="Pretendard", weight='bold'),
            customdata=reg_trend['합계_표기'],
            hovertemplate="<b>%{x}</b><br>매출액: %{customdata}<extra></extra>"
        ),
        secondary_y=False,
    )
    
    # 성장률 꺾은선 그래프
    growth_colors = ['#E11D48' if pd.notna(v) and v > 0 else '#2563EB' if pd.notna(v) and v < 0 else '#64748B' for v in reg_trend['성장률']]
    
    fig_combo.add_trace(
        go.Scatter(
            x=reg_trend['월명'], 
            y=reg_trend['성장률'], 
            name="성장률 (%)", 
            mode='lines+markers+text',
            line=dict(color="#334155", width=4), 
            connectgaps=False, 
            marker=dict(size=12, color=growth_colors, symbol='circle', line=dict(color='white', width=3)),
            text=[f"{v:+.1f}%" if pd.notna(v) else "" for v in reg_trend['성장률']],
            textposition='top center',
            textfont=dict(color=growth_colors, size=15, weight='bold'), 
            hovertemplate="<b>%{x}</b><br>성장률: %{y:+.1f}%<extra></extra>"
        ),
        secondary_y=True,
    )
    
    fig_combo.update_layout(
        template="plotly_white", 
        margin=dict(t=50, b=20, l=0, r=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
        height=550,
        hovermode="x unified"
    )
    
    fig_combo.update_yaxes(range=[0, max_sales * 1.6], secondary_y=False, gridcolor='#F1F5F9')
    fig_combo.update_yaxes(range=[-abs_max_growth * 2.5, abs_max_growth * 2.5], secondary_y=True, showgrid=False, showticklabels=False)
    fig_combo.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)], range=[-0.5, 11.5])
    
    st.plotly_chart(fig_combo, width="stretch", key="reg_combo_chart_filtered")

    st.markdown("---")

    c5, c6 = st.columns(2)
    with c5:
        st.markdown("**🏆 당월 매출 Top 10 거래처**")

        top_c_reg = curr_df.groupby('판매처명')['합계'].sum().nlargest(10).reset_index()
        top_c_reg['합계_표기'] = top_c_reg['합계'].apply(format_krw)
        top_c_reg['합계_축약표기'] = top_c_reg['합계'].apply(format_short_krw)
        top_c_reg['합계_천만'] = top_c_reg['합계'] / 10000000
        
        fig2 = px.bar(top_c_reg, x='합계_천만', y='판매처명', orientation='h', text='합계_축약표기', custom_data=['합계_표기'])
        fig2.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig2.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="매출액 (단위: 천만원)"), width="stretch")

    st.markdown("---")

    with c6:
        st.markdown("**📊 당월 품목별 매출 비중**")
        pie_df_reg = curr_df[~curr_df['카테고리'].isin(['기타', '배송비', '부속'])]
        pie_df_grouped = pie_df_reg.groupby('카테고리')['합계'].sum().reset_index()
        pie_df_grouped['합계_표기'] = pie_df_grouped['합계'].apply(format_krw)
        
        fig3 = px.pie(pie_df_grouped, names='카테고리', values='합계', hole=0.4, custom_data=['합계_표기'])
        fig3.update_traces(
            texttemplate='<b>%{label}</b><br><b>%{percent}</b>', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{label}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig3, width="stretch")

    c7, c8 = st.columns(2)
    with c7:
        st.markdown("**📦 당월 품목 Top 10 (금액 기준)**")
        top_i_reg = curr_df.groupby('품명 및 규격')['합계'].sum().nlargest(10).reset_index()
        top_i_reg['합계_표기'] = top_i_reg['합계'].apply(format_krw)
        top_i_reg['합계_축약표기'] = top_i_reg['합계'].apply(format_short_krw)
        top_i_reg['합계_천만'] = top_i_reg['합계'] / 10000000
        
        fig4 = px.bar(top_i_reg, x='합계_천만', y='품명 및 규격', orientation='h', text='합계_축약표기', custom_data=['합계_표기'])
        fig4.update_traces(
            texttemplate='<b>%{text}</b>', 
            textposition='inside', 
            textfont=dict(color='white', size=13), 
            hovertemplate="<b>%{y}</b><br>매출액: %{customdata[0]}<extra></extra>"
        )
        st.plotly_chart(fig4.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white", xaxis_title="매출액 (단위: 천만원)"), width="stretch")

    with c8:
        st.markdown("**🔢 당월 품목 Top 10 (수량 기준)**")
        fig5 = px.bar(curr_df.groupby('품명 및 규격')['수량'].sum().nlargest(10).reset_index(), x='수량', y='품명 및 규격', orientation='h', text='수량', color_discrete_sequence=['#059669'])
        fig5.update_traces(texttemplate='<b>%{text:,.0f}개</b>', textposition='inside', textfont=dict(color='white', size=13), hovertemplate="<b>%{y}</b><br>수량: %{text:,.0f}개<extra></extra>")
        st.plotly_chart(fig5.update_yaxes(categoryorder='total ascending').update_layout(template="plotly_white"), width="stretch")

    st.markdown("---")
    st.markdown("**📈 품목별 출고 추이 (수량 비중 Top 5) - 연간 흐름**")
    
    df_filtered = df_all_reg[~df_all_reg['카테고리'].isin(['기타', '배송비'])]
    top_5_cats = df_filtered.groupby('카테고리')['수량'].sum().nlargest(5).index.tolist()
    
    df_cat = df_filtered[(df_filtered['월'] > 0) & (df_filtered['카테고리'].isin(top_5_cats))].copy()
    df_cat['카테고리명'] = df_cat['카테고리']
    
    cat_qty_trend2_raw = df_cat.groupby(['월', '카테고리명'])['수량'].sum().reset_index()
    if not cat_qty_trend2_raw.empty:
        pivot_df = cat_qty_trend2_raw.pivot(index='월', columns='카테고리명', values='수량').reindex(range(1, 13)).fillna(0).reset_index()
        cat_qty_trend2 = pivot_df.melt(id_vars='월', value_name='수량').sort_values('월')
    else:
        cat_qty_trend2 = pd.DataFrame({'월': range(1,13), '카테고리명': '기타', '수량': 0})
        
    cat_qty_trend2['월'] = cat_qty_trend2['월'].astype(str) + '월'

    monthly_total = cat_qty_trend2.groupby('월')['수량'].transform('sum')
    cat_qty_trend2['비중'] = (cat_qty_trend2['수량'] / monthly_total * 100).fillna(0)
    cat_qty_trend2['텍스트'] = cat_qty_trend2.apply(lambda r: f"<b>{r['카테고리명']}</b><br>{r['비중']:.0f}%" if r['비중'] >= 10 else "", axis=1)

    fig6 = px.bar(cat_qty_trend2, x='월', y='비중', color='카테고리명', custom_data=['카테고리명', '수량'], text='텍스트')
    fig6.update_layout(barmode='stack', template="plotly_white", yaxis_title="수량 비중 (%)")
    fig6.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)])
    fig6.update_traces(
        textposition='inside',
        textfont=dict(color='white', size=13), # 통일감을 위해 글씨를 흰색으로 변경
        hovertemplate="<b>%{customdata[0]}</b><br>수량: %{customdata[1]:,.0f}개<extra></extra>"
    )
    
    st.plotly_chart(fig6, width="stretch")

    light_colors = {'65K': '#00BCD4', '57K': '#0288D1', '40K': '#9E9E9E', '30K': '#F57C00', '27K': '#E65100', '3색변환': '#9C27B0', '기타': '#EAEAEA'}

    st.markdown("---")

    c9, c10 = st.columns(2)
    with c9:
        st.markdown("**💡 당월 색온도 출고 현황 (수량)**")
        fig7 = px.pie(curr_df, names='색온도', values='수량', color='색온도', color_discrete_map=light_colors)
        fig7.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(color='white', size=13), hovertemplate="<b>%{label}</b><br>수량: %{value:,.0f}개<extra></extra>")
        st.plotly_chart(fig7, width="stretch")

    st.markdown("---")

    with c10:
        st.markdown("**📈 색온도 출고 추이 (연간 수량 흐름)**")
        color_trend = df_all_reg[df_all_reg['월'] > 0].groupby(['월', '색온도'])['수량'].sum().reset_index()
        
        color_trend['월명'] = color_trend['월'].astype(str) + '월' 

        max_m = color_trend['월'].max() if not color_trend.empty else 1
        color_trend['텍스트'] = color_trend.apply(lambda r: f"{r['수량']:,.0f}개" if r['월'] == max_m else "", axis=1)

        fig8 = px.line(color_trend, x='월명', y='수량', color='색온도', color_discrete_map=light_colors, custom_data=['수량'], markers=True)
        fig8.update_traces(mode='lines+markers', line=dict(width=4), hovertemplate="<b>%{x}</b><br>수량: %{customdata[0]:,.0f}개<extra></extra>")
    
        fig8.update_xaxes(title="월", type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)],range=[-0.5, 11.5])
        st.plotly_chart(fig8.update_layout(template="plotly_white", margin=dict(r=20)), width="stretch")

    c11, c12 = st.columns(2)
    with c11:
        st.markdown("**🚚 당월 배송 현황 (수량)**")
        pie_del_data = curr_df[curr_df['배송유형'] != '기타'].groupby('배송유형')['수량'].sum().reset_index()
        fig9 = px.pie(pie_del_data, names='배송유형', values='수량', hole=0.4)
        fig9.update_traces(texttemplate='<b>%{label}</b><br><b>%{percent}</b>', textfont=dict(color='white', size=13), hovertemplate="<b>%{label}</b><br>수량: %{value:,.0f}개<extra></extra>")
        st.plotly_chart(fig9, width="stretch")

    with c12:
        st.markdown("**📈 배송 추이 (연간 월별 수량)**")
        del_filtered = df_all_reg[(df_all_reg['월'] > 0) & (df_all_reg['배송유형'] != '기타')]
        del_trend = del_filtered.groupby(['월', '배송유형'])['수량'].sum().reset_index()
        
        del_trend['월명'] = del_trend['월'].astype(str) + '월'

        max_m_del = del_trend['월'].max() if not del_trend.empty else 1
        del_trend['텍스트'] = del_trend.apply(lambda r: f"{r['수량']:,.0f}개" if r['월'] == max_m_del else "", axis=1)

        fig10 = px.line(del_trend, x='월명', y='수량', color='배송유형', custom_data=['수량'], markers=True)
        fig10.update_traces(mode='lines+markers', line=dict(width=4), hovertemplate="<b>%{x}</b><br>수량: %{customdata[0]:,.0f}개<extra></extra>")
    
        fig10.update_xaxes(title="월", type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)],range=[-0.5, 11.5])
        st.plotly_chart(fig10.update_layout(template="plotly_white", margin=dict(r=20)), width="stretch")

elif menu == "📦 품목별 상세 분석" and df is not None:
    st.title("📦 품목별 상세 규격 및 수요처 분석")

    if 'item_search_key' not in st.session_state:
        st.session_state['item_search_key'] = '🔍 전체 품목 검색'
        
    def reset_search():
        st.session_state['item_search_key'] = '🔍 전체 품목 검색'

    # --- 1. 상단 컨트롤 바 ---
    is_period = st.checkbox("📅 기간으로 조회 (시작월~종료월)", value=False)
    
    c_ctrl1, c_ctrl2, c_ctrl3, c_ctrl4 = st.columns([2, 2, 4, 2]) 
    
    valid_months = sorted(df[df['월'] > 0]['월'].unique())
    with c_ctrl1:
        if is_period:
            sel_start = st.selectbox("시작월", valid_months, index=0, format_func=lambda x: f"{int(x)}월")
        else:
            sel_month = st.selectbox("조회월 선택", valid_months, index=len(valid_months)-1, format_func=lambda x: f"{int(x)}월")
            sel_start = sel_end = sel_month
            
    with c_ctrl2:
        if is_period:
            sel_end = st.selectbox("종료월", valid_months, index=len(valid_months)-1, format_func=lambda x: f"{int(x)}월")

    with c_ctrl3:
        valid_items = ['🔍 전체 품목 검색'] + sorted([str(x) for x in df['품명 및 규격'].unique() if str(x) not in ['nan', 'None', '']])
        sel_item = st.selectbox("🔍 품목/규격 검색", valid_items, key='item_search_key')

    with c_ctrl4:
        valid_cats = ['전체'] + sorted([str(c) for c in df['카테고리'].unique() if str(c) not in ['미분류', 'nan', '기타', '배송비']])
        sel_cat = st.selectbox("📂 품목군 선택", valid_cats, on_change=reset_search)

    # --- 2. 데이터 필터링 로직 ---
    curr_df_p = df[(df['월'] >= sel_start) & (df['월'] <= sel_end)]
    prev_df_p = df[(df['월'] >= sel_start - 1) & (df['월'] <= sel_end - 1)]
    
    current_search = st.session_state['item_search_key']

    if current_search != '🔍 전체 품목 검색':
        target_curr = curr_df_p[curr_df_p['품명 및 규격'] == current_search]
        target_prev = prev_df_p[prev_df_p['품명 및 규격'] == current_search]
        analysis_title = f"🔍 {current_search}"
    elif sel_cat != '전체':
        target_curr = curr_df_p[curr_df_p['카테고리'] == sel_cat]
        target_prev = prev_df_p[prev_df_p['카테고리'] == sel_cat]
        analysis_title = f"📂 {sel_cat}"
    else:
        target_curr = curr_df_p.copy()
        target_prev = prev_df_p.copy()
        analysis_title = "🏢 전체 품목"

    target_curr = target_curr[target_curr['합계'] > 0]
    target_prev = target_prev[target_prev['합계'] > 0]

    st.markdown("---")

    # --- 3. KPI 카드 ---
    total_q = target_curr['수량'].sum()
    total_a = target_curr['합계'].sum()
    cust_n = target_curr['판매처명'].nunique()
    
    k1, k2, k3 = st.columns(3)
    k1.button(f"📦 총 출고량\n{total_q:,.0f}개",use_container_width=True, key="p_k1_v9")
    k2.button(f"💰 총 매출액\n{format_krw(total_a)}",use_container_width=True, key="p_k2_v9")
    k3.button(f"🏢 구매 거래처\n{cust_n}개사",use_container_width=True, key="p_k3_v9")

    st.markdown("---")

    # --- 4. [좌] 지도 | [우] 지역별 랭킹 ---
    row1_c1, row1_c2 = st.columns([5, 5])
    
    target_curr['지역_str'] = target_curr['지역'].fillna('미분류').astype(str).str.upper()
    target_prev['지역_str'] = target_prev['지역'].fillna('미분류').astype(str).str.upper()
    
    target_curr_reg = target_curr[~target_curr['지역_str'].isin(['미분류', 'NA', 'NAN', 'NONE'])]
    target_prev_reg = target_prev[~target_prev['지역_str'].isin(['미분류', 'NA', 'NAN', 'NONE'])]

    with row1_c1:
        st.markdown(f"**🗺️ {analysis_title} 지역별 수요 분포 지도**")
        if not target_curr_reg.empty:
            kor_map = {'서울':'서울특별시','부산':'부산광역시','대구':'대구광역시','인천':'인천광역시','광주':'광주광역시','대전':'대전광역시','울산':'울산광역시','세종':'세종특별자치시','경기':'경기도','강원':'강원도','충북':'충청북도','충남':'충청남도','전북':'전라북도','전남':'전라남도','경북':'경상북도','경남':'경상남도','제주':'제주특별자치도'}
            geojson_url = "https://raw.githubusercontent.com/southkorea/southkorea-maps/master/kostat/2013/json/skorea_provinces_geo_simple.json"
            try:
                geojson = requests.get(geojson_url).json()
                reg_map_data = target_curr_reg.groupby('지역').agg({'합계':'sum', '수량':'sum'}).reset_index()
                reg_map_data['fullname'] = reg_map_data['지역'].map(kor_map)
                
                fig_p_map = px.choropleth_mapbox(
                    reg_map_data.dropna(), geojson=geojson, locations='fullname', featureidkey='properties.name', 
                    color='합계', color_continuous_scale="Blues", mapbox_style="carto-positron", 
                    zoom=5.0, center={"lat": 35.9, "lon": 127.7}, opacity=0.8
                )
                fig_p_map.update_traces(hovertemplate="<b>%{location}</b><br>출고량: %{customdata:,.0f}개<extra></extra>",
                                        customdata=reg_map_data['수량'])
                fig_p_map.update_layout(margin={"r":0,"t":0,"l":0,"b":0}, height=450, coloraxis_showscale=False)
                st.plotly_chart(fig_p_map, width="stretch")
            except:
                st.info("지도를 불러오는 중입니다...")

    with row1_c2:
        st.markdown(f"**🥇 {analysis_title} 지역별 매출 순위 및 증감**")
        if not target_curr_reg.empty:
            r_curr_val = target_curr_reg.groupby('지역')['합계'].sum() # 매출액
            r_curr_qty = target_curr_reg.groupby('지역')['수량'].sum() # 당월 수량
            r_prev_qty = target_prev_reg.groupby('지역')['수량'].sum() # 전월 수량
            
            r_table = pd.DataFrame({
                '매출': r_curr_val, 
                '당월 수량': r_curr_qty, 
                '전월 수량': r_prev_qty
            }).fillna(0)
            
            r_table['증감(개)'] = r_table['당월 수량'] - r_table['전월 수량']
            r_table['증감(%)'] = ((r_table['당월 수량'] - r_table['전월 수량']) / r_table['전월 수량'].replace(0, 1)) * 100
            
            r_table = r_table.sort_values('매출', ascending=False).reset_index()
            r_table.insert(0, '순위', range(1, len(r_table) + 1))
            
            st.dataframe(
                r_table.style.format({
                    '매출': format_krw, 
                    '당월 수량': '{:,.0f}개', '전월 수량': '{:,.0f}개', 
                    '증감(개)': '{:+,.0f}개', '증감(%)': '{:+.1f}%'
                }).map(lambda x: 'color: #E11D48; font-weight: bold;' if x > 0 else 'color: #2563EB; font-weight: bold;' if x < 0 else '', subset=['증감(개)', '증감(%)']),
                use_container_width=True, hide_index=True, height=450
            )

    st.markdown("---")

    # --- 5. 주요 구매처 TOP 20 ---
    st.markdown(f"**🏢 {analysis_title} 주요 구매처 TOP 20 (수량순)**")
    if not target_curr.empty:
        curr_agg = target_curr.groupby(['판매처명', '지역']).agg({'수량':'sum', '합계':'sum'}).reset_index()
        prev_agg = target_prev.groupby(['판매처명', '지역'])['수량'].sum().reset_index()
        
        top_20 = pd.merge(curr_agg, prev_agg, on=['판매처명', '지역'], how='left', suffixes=('_당월', '_전월')).fillna(0)
        top_20['증감(%)'] = ((top_20['수량_당월'] - top_20['수량_전월']) / top_20['수량_전월'].replace(0, 1)) * 100
        top_20 = top_20.nlargest(20, '수량_당월').reset_index(drop=True)
        top_20.insert(0, '순위', range(1, len(top_20) + 1))
        
        top_20_show = top_20[['순위', '판매처명', '지역', '합계', '수량_당월', '수량_전월', '증감(%)']]
        top_20_show.columns = ['순위', '판매처명', '지역', '매출', '당월 수량', '전월 수량', '증감(%)']
        
        st.dataframe(
            top_20_show.style.format({
                '매출': format_krw, 
                '당월 수량': '{:,.0f}개', 
                '전월 수량': '{:,.0f}개', 
                '증감(%)': '{:+.1f}%'
            }).map(lambda x: 'color: #E11D48; font-weight: bold;' if x > 0 else 'color: #2563EB; font-weight: bold;' if x < 0 else '', subset=['증감(%)']), 
            use_container_width=True, hide_index=True
        )

    st.markdown("---")

    st.markdown(f"**📑 {analysis_title} 세부 품목 성과 및 ABC 분석 (선택 기간)**")
    st.caption("※ A등급(핵심): 누적 매출 상위 80% / B등급(일반): 80~95% / C등급(비주력): 하위 5%")
    
    if not target_curr.empty:
        item_curr = target_curr.groupby('품명 및 규격').agg({'합계': 'sum', '수량': 'sum'}).reset_index()
        item_prev = target_prev.groupby('품명 및 규격').agg({'합계': 'sum', '수량': 'sum'}).reset_index()
        
        # 💡 [핵심] 당월 기준 최근 3개월 데이터 추출
        start_3m = max(1, sel_end - 2)
        target_3m = df[(df['월'] >= start_3m) & (df['월'] <= sel_end)]
        if current_search != '🔍 전체 품목 검색': target_3m = target_3m[target_3m['대표품명'] == current_search]
        elif sel_cat != '전체': target_3m = target_3m[target_3m['카테고리'] == sel_cat]
        
        item_3m = target_3m.groupby('품명 및 규격').agg({'합계': 'sum', '수량': 'sum'}).reset_index()
        item_3m.columns = ['품명 및 규격', '3개월_매출액', '3개월_수량']
        
        # 데이터 병합
        item_merged = pd.merge(item_curr, item_prev, on='품명 및 규격', how='left', suffixes=('_당월', '_전월')).fillna(0)
        item_merged = pd.merge(item_merged, item_3m, on='품명 및 규격', how='left').fillna(0)
        
        item_merged = item_merged.sort_values('합계_당월', ascending=False).reset_index(drop=True)
        item_merged.insert(0, '순위', range(1, len(item_merged) + 1))
        
        item_merged['증감(%)'] = ((item_merged['합계_당월'] - item_merged['합계_전월']) / item_merged['합계_전월'].replace(0, 1)) * 100
        total_cat_sales = item_merged['합계_당월'].sum()
        item_merged['누적매출'] = item_merged['합계_당월'].cumsum()
        item_merged['누적비율(%)'] = (item_merged['누적매출'] / total_cat_sales * 100)
        
        item_merged['ABC 등급'] = pd.cut(item_merged['누적비율(%)'], bins=[0, 80, 95, 101], labels=['A등급 (핵심)', 'B등급 (일반)', 'C등급 (비주력)'])
        
        # 표출 컬럼에 3개월 데이터 포함
        item_show = item_merged[['순위', '품명 및 규격', 'ABC 등급', '3개월_매출액', '3개월_수량', '합계_당월', '합계_전월', '증감(%)']].copy()
        item_show.columns = ['순위', '품목명', 'ABC 등급', '최근 3개월 매출', '최근 3개월 수량', '당월 매출액', '전월 매출액', '증감(%)']
        
        st.dataframe(
            item_show.style.format({
                '최근 3개월 매출': format_krw, '최근 3개월 수량': '{:,.0f}개',
                '당월 매출액': format_krw, '전월 매출액': format_krw,
                '증감(%)': '{:+.1f}%'
            }).map(lambda x: 'color: #E11D48; font-weight: bold;' if x > 0 else 'color: #2563EB; font-weight: bold;' if x < 0 else '', subset=['증감(%)'])
              .map(lambda x: 'background-color: #FEF2F2; font-weight: bold; color: #B91C1C;' if 'A등급' in str(x) else 'background-color: #F8FAFC;' if 'B등급' in str(x) else 'color: #94A3B8;', subset=['ABC 등급']),
            use_container_width=True, hide_index=True, height=450
        )
    else:
        st.info("해당 기간/조건에 데이터가 없습니다.")

    st.markdown("---")

    # 💡 [핵심] 규격별 탭 분리 (비중 + 연간 추이 병렬 배치)
    st.markdown(f"**📊 {analysis_title} 상세 옵션 심층 분석**")
    
    trend_raw = df.copy()
    if current_search != '🔍 전체 품목 검색': trend_raw = trend_raw[trend_raw['대표품명'] == current_search]
    elif sel_cat != '전체': trend_raw = trend_raw[trend_raw['카테고리'] == sel_cat]
    trend_raw['월명'] = trend_raw['월'].astype(str) + '월'

    tab_s, tab_w, tab_c = st.tabs(["📏 사이즈 분석", "⚡ 와트(W) 분석", "💡 색온도 분석"])
    
    def render_spec_tab(col_name, title, color_hex, color_map=None):
        c_left, c_right = st.columns([4, 6])
        with c_left:
            st.markdown(f"**✔️ 당월 {title} 비중 (수량)**")
            dist = target_curr[target_curr[col_name] != '미표기'].groupby(col_name)['수량'].sum().sort_values().reset_index()
            if not dist.empty:
                # 💡 color_map이 있으면 각 항목별 색상을 다르게, 없으면 단일 색상 적용
                if color_map:
                    fig_d = px.bar(dist, x='수량', y=col_name, orientation='h', color=col_name, color_discrete_map=color_map)
                else:
                    fig_d = px.bar(dist, x='수량', y=col_name, orientation='h', color_discrete_sequence=[color_hex])
                    
                fig_d.update_traces(
                    texttemplate='%{x:,.0f}개', textposition='inside', 
                    hovertemplate="<b>%{y}</b>: %{x:,.0f}개<extra></extra>"
                )
                fig_d.update_layout(
                    template="plotly_white", height=350, xaxis_visible=False, yaxis_title=None, margin=dict(t=0, b=0, l=0, r=0),
                    font=dict(family="Pretendard", weight="bold", color="#334155"), showlegend=False
                )
                st.plotly_chart(fig_d, use_container_width=True)
                
        with c_right:
            st.markdown(f"**📈 연간 {title} 출고 추이 (수량)**")
            trend = trend_raw[trend_raw[col_name] != '미표기'].groupby(['월', '월명', col_name])['수량'].sum().reset_index()
            if not trend.empty:

                if color_map:
                    fig_t = px.line(trend.sort_values('월'), x='월명', y='수량', color=col_name, color_discrete_map=color_map, markers=True)
                else:
                    fig_t = px.line(trend.sort_values('월'), x='월명', y='수량', color=col_name, markers=True)
                    
                fig_t.update_traces(
                    mode='lines+markers', line=dict(width=3),
                    hovertemplate="<b>%{x} %{fullData.name}</b>: %{y:,.0f}개<extra></extra>"
                )
                fig_t.update_layout(
                    template="plotly_white", height=350, margin=dict(t=0, b=0, l=0, r=0), yaxis_title="수량(개)", xaxis_title=None,
                    font=dict(family="Pretendard", weight="bold", color="#334155"), showlegend=(color_map is not None)
                )
                fig_t.update_xaxes(type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)], range=[-0.5, 11.5])
                st.plotly_chart(fig_t, use_container_width=True)

    with tab_s: render_spec_tab('사이즈', '사이즈별', '#818CF8')
    with tab_w: render_spec_tab('와트(W)', '와트(W)별', '#FBBF24')
    with tab_c: 
        target_curr['색온도_str'] = target_curr['색온도'].replace('기타', '미표기')
        trend_raw['색온도_str'] = trend_raw['색온도'].replace('기타', '미표기')
        c_colors = {'65K': '#00BCD4', '57K': '#0288D1', '40K': '#9E9E9E', '30K': '#F57C00', '27K': '#E65100', '3색변환': '#9C27B0', '미표기': '#EAEAEA'}
        render_spec_tab('색온도', '색온도별', '#FCA5A5', color_map=c_colors)

    st.markdown("---")

    # --- 7. 월별 매출액 추이 ---
    st.markdown(f"**📈 {analysis_title} 월별 매출액 추이**")
    trend_raw = df.copy()
    
    if current_search != '🔍 전체 품목 검색':
        trend_raw = trend_raw[trend_raw['품명 및 규격'] == current_search]
    elif sel_cat != '전체':
        trend_raw = trend_raw[trend_raw['카테고리'] == sel_cat]
    
    max_m_active = trend_raw[trend_raw['합계'] > 0]['월'].max() if not trend_raw.empty else 1
    trend_data = trend_raw.groupby('월')['합계'].sum().reset_index()
    trend_data['월명'] = trend_data['월'].astype(str) + '월'
    
    def short_fmt_val(val):
        if val == 0: return ""
        if val >= 100000000: return f"{val/100000000:.1f}억"
        if val >= 10000: return f"{val/10000:,.0f}만"
        return f"{val:,.0f}"

    fig_trend = px.bar(trend_data, x='월명', y='합계', text=trend_data['합계'].apply(short_fmt_val), color_discrete_sequence=['#3B82F6'])
    
    fig_trend.update_traces(
        textposition='inside', 
        textfont=dict(color='white', weight='bold', size=15), 
        hovertemplate="매출액: %{y:,.0f}원<extra></extra>"
    )
    fig_trend.update_layout(template="plotly_white", height=400, yaxis_visible=False, xaxis_title=None, margin=dict(t=10))
    
    fig_trend.update_xaxes(title="월", type='category', categoryorder='array', categoryarray=[f"{i}월" for i in range(1, 13)], range=[-0.5, 11.5])
    st.plotly_chart(fig_trend, width="stretch")

elif menu == "⚙️ 데이터 설정":
    st.title("⚙️ 데이터 업로드 및 설정")
    
    # 탭을 5개로 분리 (목표 관리 옆에 거래처 리스트 배치)
    tab_up, tab_edit, tab_cat, tab_goal, tab_cust = st.tabs([
        "📤 엑셀 업로드", "✏️ 수동 데이터 수정", "🏷️ 카테고리 설정", "🎯 목표 및 영업사원 관리", "🏢 거래처 리스트 (지역)"
    ])
    
    with tab_up:
        st.subheader("1. 📊 판매현황 데이터 업로드")
        
        uploaded_file = st.file_uploader("판매현황 엑셀 파일(.xlsx) 드래그 앤 드롭", type=['xlsx'], key="sales_up")
        if uploaded_file is not None:
            if st.button("🔄 저장하기"):
                try:
                    # A1 셀 값 미리 읽기
                    raw_header = pd.read_excel(uploaded_file, header=None, nrows=1).iloc[0, 0]
                    uploaded_file.seek(0)

                    try:
                        new_df = pd.read_excel(uploaded_file, header=1)
                    except Exception:
                        uploaded_file.seek(0)
                        html_dfs = pd.read_html(uploaded_file.getvalue(), encoding='utf-8')
                        new_df = html_dfs[0]
                        new_df.columns = new_df.iloc[0]
                        new_df = new_df[1:]

                    # 전처리 함수 호출 (A1 셀 문자열 전달)
                    new_df = preprocess_data(new_df, raw_header)
                    
                    # 이번 파일의 출처월 확인
                    current_source = int(new_df['업로드_출처월'].iloc[0])
                    supabase.table("sales").delete().eq("업로드_출처월", current_source).execute()
                    
                    # 2. 새 데이터 DB에 삽입 (딕셔너리 리스트 형태로 변환 후 업로드)
                    rows = new_df.to_dict(orient='records')
                    supabase.table("sales").insert(rows).execute()
                    
                    # 3. 로컬 세션 상태 업데이트 및 새로고침
                    st.session_state.main_data = fetch_data()
                    st.success(f"✅ {current_source}월 데이터가 저장되었습니다.")

                except Exception as e:
                    st.error(f"오류: {e}")

        st.markdown("---")
        st.markdown("#### 🗑️ 월 데이터 삭제")
        if st.session_state.main_data is not None and not st.session_state.main_data.empty:
            avail_months = sorted([m for m in st.session_state.main_data['월'].unique() if m > 0])
            
            col_del1, col_del2, col_del3 = st.columns([2, 2, 6])
            with col_del1:
                month_to_delete = st.selectbox("삭제할 월 선택", avail_months, format_func=lambda x: f"{int(x)}월")
            with col_del2:
                st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                
                if st.button(f"{int(month_to_delete)}월 데이터 삭제"):
                    # '월' 컬럼이 아니라 '업로드_출처월' 컬럼을 기준으로 삭제
                    st.session_state.main_data = st.session_state.main_data[
                        st.session_state.main_data['업로드_출처월'] != month_to_delete
                    ]
                    st.success(f"✅ {int(month_to_delete)}월 파일에서 업로드된 모든 데이터가 삭제되었습니다.")

    with tab_edit:
        st.subheader("✏️ 수동 데이터 수정 (월별)")
        
        if st.session_state.main_data is not None:
            st.session_state.main_data = st.session_state.main_data.reset_index(drop=True)
            
            edit_month = st.selectbox("수정할 월 선택", options=[m for m in sorted(st.session_state.main_data['월'].unique()) if m>0], format_func=lambda x: f"{int(x)}월")
            
            # 수정이 필요한 데이터만 필터링하는 토글 스위치
            focus_mode = st.toggle("🚨 수정 요망 데이터 (미분류, 기타, NA)만 모아보기")
            
            mask = st.session_state.main_data['월'] == edit_month
            target_df = st.session_state.main_data[mask].copy()
            
            if focus_mode:
                # 지역이 미분류/NA거나, 카테고리/배송유형이 기타인 것만 쏙 골라냄
                focus_mask = (target_df['지역'].isin(['미분류', 'NA', 'NAN', 'None'])) | (target_df['카테고리'] == '기타') | (target_df['배송유형'] == '기타')
                target_df = target_df[focus_mask]
                if target_df.empty:
                    st.success("🎉 선택하신 월에는 미분류나 기타로 지정된 데이터가 없습니다! 아주 깔끔하네요.")

            if '일자' in target_df.columns:
                target_df['일자'] = pd.to_datetime(target_df['일자'], errors='coerce').dt.strftime('%Y-%m-%d')

            ordered_cols = [
                '일자', '품목코드', '품명 및 규격', '카테고리', '색온도', 
                '수량', '단가', '공급가액', '부가세', '합계', 
                '판매처명', '지역', '담당자명', '적요', '장문형식1', '배송유형'
            ]
            final_cols = [c for c in ordered_cols if c in target_df.columns] + [c for c in target_df.columns if c not in ordered_cols]
            target_df = target_df[final_cols]
            
            edited_df = st.data_editor(
                target_df, 
                num_rows="dynamic", 
                use_container_width=True,
                height=600,
                column_config={
                    "수량": st.column_config.NumberColumn("수량", format="%d"),
                    "단가": st.column_config.NumberColumn("단가", format="%d"),
                    "합계": st.column_config.NumberColumn("합계", format="%d"),
                    "월": None 
                }
            )

            if st.button("💾 수정 내용 저장", type="secondary", use_container_width=True, key="btn_save_edit"):
                if '일자' in edited_df.columns:
                    edited_df['일자'] = pd.to_datetime(edited_df['일자'], errors='coerce')
                
                st.session_state.main_data = st.session_state.main_data.drop(target_df.index)
                st.session_state.main_data = pd.concat([st.session_state.main_data, edited_df]).reset_index(drop=True)
                
                st.success("✅ 데이터가 수정 되었습니다!")
                st.rerun()
        else:
            st.info("먼저 엑셀 데이터를 업로드해주세요.")

    with tab_cat:
        st.subheader("🏷️ 품목 카테고리 키워드 설정")
        cat_df = pd.DataFrame(list(st.session_state.category_map.items()), columns=['키워드', '카테고리명'])
        edited_cat_df = st.data_editor(cat_df, num_rows="dynamic", use_container_width=True)
        if st.button("카테고리 설정 저장 및 데이터 재분류"):
            new_map = dict(zip(edited_cat_df['키워드'], edited_cat_df['카테고리명']))
            st.session_state.category_map = new_map
            if st.session_state.main_data is not None:
                st.session_state.main_data['카테고리'] = st.session_state.main_data['품명 및 규격'].apply(lambda x: assign_category(x, st.session_state.category_map))
            st.success("✅ 카테고리 매핑 완료!")

    with tab_goal:
        st.subheader("🎯 회사 및 담당자별 매출 관리")

        # --- 1. 전사 연간 목표 설정 ---
        st.markdown("#### 🏢 회사 연간 매출 목표")
        
        c_g1, c_g2 = st.columns([4, 6])
        with c_g1:
            company_goal_input = st.number_input(
                "회사 연간 총 목표 금액 (단위: 만 원)", 
                value=int(st.session_state.company_goal // 10000), 
                step=100,
                help="1000을 입력하면 1000만 원으로 반영됩니다."
            )
            st.session_state.company_goal = company_goal_input * 10000
        with c_g2:
            st.info(f"💡 현재 설정된 전사 연간 목표: **{format_krw(st.session_state.company_goal)}**")

        st.markdown("---")

        # --- 2. 담당자별 연간 목표 관리 ---
        st.markdown("#### 👨‍💼 담당자별 연간 목표 설정")
        
        # 데이터가 있으면 자동으로 담당자 명단을 세션(dict)에 추가
        if df is not None and '담당자명' in df.columns:
            current_reps = [r for r in df['담당자명'].dropna().unique() if str(r).strip() not in ['', 'nan', '미분류', 'None']]
            for rep in current_reps:
                if rep not in st.session_state.rep_goals and rep not in st.session_state.deleted_reps:
                    st.session_state.rep_goals[rep] = 100000000

        if st.session_state.rep_goals:
            cols = st.columns(3)
            for idx, rep in enumerate(list(st.session_state.rep_goals.keys())):
                with cols[idx % 3]:
                    with st.container(border=True): # 테두리 박스 적용
                        col_in1, col_in2 = st.columns([8, 2])
                        with col_in1:
                            current_val = int(st.session_state.rep_goals[rep] // 10000)
                            new_val = st.number_input(f"👤 {rep} (만 원)", value=current_val, step=50, key=f"goal_{rep}")
                            st.session_state.rep_goals[rep] = new_val * 10000
                        with col_in2:
                            st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
                            # 🗑️ 휴지통 삭제 버튼
                            if st.button("🗑️", key=f"del_{rep}", help=f"{rep} 데이터 삭제"):
                                del st.session_state.rep_goals[rep]
                                st.session_state.deleted_reps.add(rep)
                                st.rerun() # 삭제 즉시 새로고침
                                
                        st.markdown(f"<div style='text-align:right; color:#2563EB; font-weight:bold; font-size:15px;'>목표: {format_krw(st.session_state.rep_goals.get(rep, 0))}</div>", unsafe_allow_html=True)
        else:
            st.info("판매 데이터를 업로드하면 담당자 목록이 자동으로 생성됩니다.")

    # 5번째 거래처 리스트 탭 추가 및 J열(10번째 열) 인식 로직 적용
    with tab_cust:
        st.subheader("🏢 거래처 리스트 및 주소 관리")
        
        # 주소와 지역을 포함한 마스터 데이터 관리
        if 'cust_master' not in st.session_state:
            st.session_state.cust_master = pd.DataFrame(columns=['거래처명', '주소(J열)', '매핑된 지역'])

        cust_file = st.file_uploader("거래처 리스트 업로드 (.xlsx)", type=['xlsx'], key="cust_up")
        
        if cust_file is not None:
            if st.button("📥 거래처 리스트 불러오기 및 저장"):
                try:
                    c_df = pd.read_excel(cust_file, header=1)
                    cust_col = next((c for c in c_df.columns if c in ['판매처명', '거래처명', '상호', '고객사']), c_df.columns[0])
                    
                    # 주소(J열) 및 지역 추출
                    addresses = c_df.iloc[:, 9].astype(str).str.strip() if len(c_df.columns) >= 10 else c_df.iloc[:, -1].astype(str).str.strip()
                    front_two = addresses.str[:2]
                    special_combo = addresses.str[0] + addresses.str[2:3]
                    regions = np.where(front_two.isin(['전라', '경상', '충청']), special_combo, front_two)
                    
                    # 마스터 데이터 프레임 생성
                    new_master = pd.DataFrame({
                        '거래처명': c_df[cust_col].astype(str).str.strip(),
                        '주소(J열)': addresses,
                        '매핑된 지역': regions
                    }).drop_duplicates('거래처명')
                    
                    st.session_state.cust_master = new_master
                    st.session_state.cust_data = dict(zip(new_master['거래처명'], new_master['매핑된 지역']))
                    
                    if st.session_state.main_data is not None:
                        clean_map = {str(k).strip(): v for k, v in st.session_state.cust_data.items()}
                        st.session_state.main_data['지역'] = st.session_state.main_data['판매처명'].astype(str).str.strip().map(clean_map).fillna('미분류')
                    
                    st.success("✅ 거래처 리스트 저장완료")
                except Exception as e:
                    st.error(f"파일 읽기 오류: {e}")

        # 수동 편집 및 목록 보기 (세로 스크롤 가능하도록 높이 설정)
        edited_master = st.data_editor(
            st.session_state.cust_master, 
            num_rows="dynamic", 
            use_container_width=True,
            height=500,
            key="master_editor"
        )
        
        if st.button("💾 저장"):
            st.session_state.cust_master = edited_master
            st.session_state.cust_data = dict(zip(edited_master['거래처명'], edited_master['매핑된 지역']))
            # 판매 데이터 동기화
            if st.session_state.main_data is not None:
                clean_map = {str(k).strip(): v for k, v in st.session_state.cust_data.items()}
                st.session_state.main_data['지역'] = st.session_state.main_data['판매처명'].astype(str).str.strip().map(clean_map).fillna('미분류')
            st.success("✅ 거래처 리스트가 저장되었습니다.")