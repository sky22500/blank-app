import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import datetime
import calendar
import re
import json
import gspread

# 1. 페이지 설정
st.set_page_config(page_title="마케팅 부문 대시보드", layout="wide")

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
default_target_date = f"{yesterday.month}/{yesterday.day}"

# HTML 안전 렌더링 헬퍼 함수
def render_html(html_str):
    if hasattr(st, "html"):
        st.html(html_str)
    else:
        st.markdown(html_str, unsafe_allow_html=True)

# ----------------- 구글 시트 동기화 엔진 -----------------
@st.cache_resource
def get_gspread_client():
    try:
        if "gcp_service_account" not in st.secrets:
            return None
        creds = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds and "\\n" in creds["private_key"]:
            creds["private_key"] = creds["private_key"].replace("\\n", "\n")
        return gspread.service_account_from_dict(creds)
    except Exception as e:
        return None

def get_db_worksheet():
    gc = get_gspread_client()
    if not gc:
        return None
    try:
        sheet_url = st.secrets["sheet"]["url"]
        sh = gc.open_by_url(sheet_url)
        try:
            ws = sh.worksheet("DASHBOARD_DB")
        except:
            try:
                ws = sh.add_worksheet(title="DASHBOARD_DB", rows=100, cols=5)
            except:
                ws = sh.get_worksheet(0)
        return ws
    except Exception as e:
        return None

def save_data_to_sheet(payload):
    ws = get_db_worksheet()
    if not ws:
        return False, "구글 시트 연동 실패: Secrets 설정을 확인하세요."
    try:
        json_str = json.dumps(payload, ensure_ascii=False)
        CHUNK_SIZE = 30000
        chunks = [[json_str[i:i + CHUNK_SIZE]] for i in range(0, len(json_str), CHUNK_SIZE)]
        ws.clear()
        ws.update(chunks, 'A1')
        load_data_from_sheet.clear()
        return True, "성공"
    except Exception as e:
        return False, str(e)

@st.cache_data(ttl=60)
def load_data_from_sheet():
    ws = get_db_worksheet()
    if not ws:
        return None
    try:
        col_vals = ws.col_values(1)
        if not col_vals:
            return None
        json_str = "".join(col_vals)
        if not json_str.strip():
            return None
        data = json.loads(json_str)
        return data
    except Exception as e:
        return None

def sync_session_from_sheet():
    sheet_data = load_data_from_sheet()
    if sheet_data:
        st.session_state['target_date'] = sheet_data.get('target_date', default_target_date)
        if 'kpi_data' in sheet_data:
            st.session_state['kpi_data'] = sheet_data['kpi_data']
        if 'chart_data' in sheet_data:
            st.session_state['chart_data'] = sheet_data['chart_data']
        if 'inventory_data' in sheet_data:
            st.session_state['inventory_data'] = sheet_data['inventory_data']
        if 'fcst_data' in sheet_data:
            st.session_state['fcst_data'] = sheet_data['fcst_data']
        if 'mapping_data' in sheet_data:
            st.session_state['mapping_data'] = sheet_data['mapping_data']
        if 'actual_by_day' in sheet_data:
            raw_act = sheet_data['actual_by_day']
            st.session_state['actual_by_day'] = {int(k) if str(k).isdigit() else k: v for k, v in raw_act.items()}
        for k in ['_ship_plan_by_model', '_ship_plan_by_site', '_ship_actual_by_model', '_ship_actual_by_site']:
            if k in sheet_data:
                st.session_state[k] = sheet_data[k]
        st.session_state['data_loaded'] = True

# 세션 상태 초기화 및 기본값 설정
if 'target_date' not in st.session_state:
    st.session_state['target_date'] = default_target_date

if 'mapping_data' not in st.session_state: 
    st.session_state['mapping_data'] = {'gh_map': {}, 'suffix_map': {}, 'model_map': {}, 'gh_to_model': {}}
else:
    st.session_state['mapping_data'].setdefault('gh_map', {})
    st.session_state['mapping_data'].setdefault('suffix_map', {})
    st.session_state['mapping_data'].setdefault('model_map', {})
    st.session_state['mapping_data'].setdefault('gh_to_model', {})

if 'fcst_data' not in st.session_state:
    st.session_state['fcst_data'] = {'rows': [], 'date_cols': []}

if 'actual_by_day' not in st.session_state:
    st.session_state['actual_by_day'] = {}

if 'active_cal_popup' not in st.session_state:
    st.session_state['active_cal_popup'] = None

default_kpis = {
    'daily': {'prod_p': 0, 'prod_a': 0, 'ship_p': 0, 'ship_a': 0, 'rev_p': 0.0, 'rev_a': 0.0},
    'weekly': {'prod_p': 0, 'prod_a': 0, 'ship_p': 0, 'ship_a': 0, 'rev_p': 0.0, 'rev_a': 0.0, 'pure_rev_a': 0.0, 'carryover_rev': 0.0},
    'monthly': {'prod_p': 0, 'prod_a': 0, 'ship_p': 0, 'ship_a': 0, 'rev_p': 0.0, 'rev_a': 0.0, 'pure_rev_a': 0.0, 'carryover_rev': 0.0}
}
if 'kpi_data' not in st.session_state:
    st.session_state['kpi_data'] = default_kpis
else:
    for period in ['daily', 'weekly', 'monthly']:
        if period not in st.session_state['kpi_data']:
            st.session_state['kpi_data'][period] = default_kpis[period]

default_charts = {
    'daily_models': [], 'daily_prod_p': [], 'daily_prod_a': [], 
    'weekly_weeks': ['W45', 'W46', 'W47', 'W48'], 
    'weekly_prod_p': [0]*4, 'weekly_prod_a': [0]*4,
    'weekly_ship_p': [0]*4, 'weekly_ship_a': [0]*4,
    'weekly_rev_p': [0.0]*4, 'weekly_rev_a': [0.0]*4,
    'monthly_months': ['11월'], 
    'monthly_prod_p': [0], 'monthly_prod_a': [0],
    'monthly_ship_p': [0], 'monthly_ship_a': [0],
    'monthly_rev_p': [0.0], 'monthly_rev_a': [0.0],
    'ship_by_model': [], 'ship_by_site': []
}
if 'chart_data' not in st.session_state:
    st.session_state['chart_data'] = default_charts
else:
    for k, v in default_charts.items():
        if k not in st.session_state['chart_data']:
            st.session_state['chart_data'][k] = v

if 'inventory_data' not in st.session_state:
    st.session_state['inventory_data'] = {'qty': 0, 'days': 0.0}

# 접속 시 구글 시트에 기존 저장된 데이터가 있다면 자동 동기화
if not st.session_state.get('data_loaded', False):
    sync_session_from_sheet()

# 2. 안전 유틸 및 포맷팅 함수
def safe_cell(sheet, r, c):
    if r < 1 or c < 1 or r > sheet.max_row or c > sheet.max_column:
        return None
    return sheet.cell(r, c).value

def clean_key(val):
    if pd.isna(val) or val is None: return ""
    return re.sub(r'[^A-Z0-9]', '', str(val).upper())

def format_k(val):
    if val is None or pd.isna(val):
        return "0K"
    try:
        v = float(val)
    except:
        return str(val)
    if v == 0:
        return "0K"
    k_val = round(v / 1000.0, 1)
    if k_val.is_integer():
        return f"{int(k_val):,}K"
    else:
        return f"{k_val:,.1f}K"

def get_gh_suffix(code):
    s = clean_key(code)
    m = re.search(r'GH(?:86|96)?(\d+.*)$', s)
    if m: return m.group(1)
    m2 = re.search(r'(\d+[A-Z0-9]*)$', s)
    if m2: return m2.group(1)
    return s

def parse_target_day_info(target_str):
    s = str(target_str).strip()
    m = re.search(r'\b0?(\d{1,2})[-/.]0?(\d{1,2})\b', s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m2 = re.search(r'(\d{1,2})일?', s)
    if m2:
        return 11, int(m2.group(1))
    return 11, 2

def resolve_model_name(gh_code, fallback_model, gh_to_model):
    gh_clean = clean_key(gh_code)
    gh_suf = get_gh_suffix(gh_code)
    if gh_clean in gh_to_model: return gh_to_model[gh_clean]
    if gh_suf in gh_to_model: return gh_to_model[gh_suf]
    m = re.search(r'CAM_([A-Z0-9]+)', str(fallback_model).upper())
    if m:
        extracted = m.group(1)
        return extracted[:-1] if extracted.endswith('E') else extracted
    clean_m = str(fallback_model)[:-1] if str(fallback_model).endswith('E') else str(fallback_model)
    return clean_m.strip()

def is_match_date(cell_val, target_str):
    if not cell_val: return False
    t_m, t_d = parse_target_day_info(target_str)
    if isinstance(cell_val, (datetime.datetime, datetime.date)):
        return cell_val.month == t_m and cell_val.day == t_d
    s = str(cell_val).strip().upper()
    m = re.search(r'\b0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b', s)
    if m: return int(m.group(1)) == t_m and int(m.group(2)) == t_d
    m2 = re.search(r'\b0?(\d{1,2})\s*[-/.]\s*([A-Z]{3})\b', s)
    if m2:
        try:
            d = int(m2.group(1))
            mon = datetime.datetime.strptime(m2.group(2), '%b').month
            return mon == t_m and d == t_d
        except: pass
    m3 = re.search(r'^0?(\d{1,2})일?$', s)
    if m3: return int(m3.group(1)) == t_d
    return False

@st.cache_data
def convert_df(df):
    return df.to_csv(index=False).encode('utf-8-sig')

# 계획 폰트 일치(초록색) 및 계획 대비 실적 차이(+파랑 / -빨강) 통합 KPI 카드
def render_kpi_card(title, icon, actual_val, plan_val, is_rev=False, sub_caption=None):
    if is_rev:
        act_str = f"{actual_val:,.2f}억원"
        plan_str = f"{plan_val:,.2f}억원"
        diff = actual_val - plan_val
        if diff > 0:
            diff_str = f"+{diff:,.2f}억원"
            diff_color = "#0044ff"
        elif diff < 0:
            diff_str = f"-{abs(diff):,.2f}억원"
            diff_color = "#ef4444"
        else:
            diff_str = "0.00억원"
            diff_color = "#6b7280"
    else:
        act_str = format_k(actual_val)
        plan_str = format_k(plan_val)
        diff = actual_val - plan_val
        if diff > 0:
            diff_str = f"+{format_k(diff)}"
            diff_color = "#0044ff"
        elif diff < 0:
            diff_str = f"-{format_k(abs(diff))}"
            diff_color = "#ef4444"
        else:
            diff_str = "0K"
            diff_color = "#6b7280"

    sub_html = f"<div style='font-size: 11px; color: #6b7280; font-weight: 700; margin-top: 6px; border-top: 1px dashed #e5e7eb; padding-top: 4px;'>{sub_caption}</div>" if sub_caption else ""

    html = f"""
    <div style="background-color: #ffffff; border: 3px solid #111111; padding: 15px 18px; box-shadow: 6px 6px 0px #0044ff; margin-bottom: 8px; border-radius: 4px;">
        <div style="font-weight: 800; color: #111111; text-transform: uppercase; font-size: 14px; margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
            <span>{icon}</span> <span>{title}</span>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: flex-end; gap: 10px;">
            <div>
                <div style="font-size: 11px; font-weight: 800; color: #6b7280; margin-bottom: 2px;">실적</div>
                <div style="font-size: 26px; font-weight: 900; color: #111111; line-height: 1.1;">{act_str}</div>
            </div>
            <div style="text-align: center;">
                <div style="font-size: 11px; font-weight: 800; color: #059669; margin-bottom: 2px;">계획</div>
                <div style="font-size: 26px; font-weight: 900; color: #059669; line-height: 1.1;">{plan_str}</div>
            </div>
            <div style="text-align: right;">
                <div style="font-size: 11px; font-weight: 800; color: #6b7280; margin-bottom: 2px;">계획 대비 차이</div>
                <div style="font-size: 20px; font-weight: 900; color: {diff_color}; line-height: 1.1;">{diff_str}</div>
            </div>
        </div>
        {sub_html}
    </div>
    """
    render_html(html)

def render_achievement_gauge(title, actual, plan, unit="K"):
    rate = (actual / plan * 100.0) if plan > 0 else 0.0
    if rate >= 100:
        bar_color = "linear-gradient(90deg, #0044ff, #2563eb)"
        status_color = "#0044ff"
    elif rate >= 80:
        bar_color = "linear-gradient(90deg, #10b981, #059669)"
        status_color = "#059669"
    else:
        bar_color = "linear-gradient(90deg, #f59e0b, #d97706)"
        status_color = "#d97706"

    bar_width = min(max(rate, 0.0), 100.0)

    html = f"""
    <div style="background-color: #ffffff; border: 2.5px solid #111111; padding: 10px 14px; box-shadow: 4px 4px 0px #111111; margin-top: -4px; margin-bottom: 20px; border-radius: 4px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-size: 13px; font-weight: 800; color: #444444;">{title} 달성률</span>
            <span style="font-size: 15px; font-weight: 900; color: {status_color};">{rate:,.1f}%</span>
        </div>
        <div style="background-color: #e5e7eb; border: 1.5px solid #111111; border-radius: 6px; height: 13px; overflow: hidden;">
            <div style="background: {bar_color}; width: {bar_width:,.1f}%; height: 100%; border-radius: 4px;"></div>
        </div>
    </div>
    """
    render_html(html)

# 팝업 대화상자 함수
if hasattr(st, "dialog"):
    @st.dialog("📋 일자별 공급 상세 내역 (전체 모델)", width="large")
    def show_day_detail_dialog(d_num, m_num, y_num, d_info):
        items = d_info['items']
        tot_q = d_info['total']
        tot_rev = d_info['total_rev']
        unique_models = sorted(list(set(it['model'] for it in items)))
        
        st.markdown(f"#### 📅 {y_num}년 {m_num}월 {d_num}일 공급 스케줄")
        
        cards_html = f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px;">
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">📱 전체 모델 수</div>
                <div style="font-size: 20px; font-weight: 900; color: #111111; margin-top: 3px;">{len(unique_models)}종</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">📦 총 공급 품목(GH)</div>
                <div style="font-size: 20px; font-weight: 900; color: #111111; margin-top: 3px;">{len(items)}건</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">🚚 총 공급 수량</div>
                <div style="font-size: 20px; font-weight: 900; color: #0044ff; margin-top: 3px;">{format_k(tot_q)}</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">💰 총 예상 매출액</div>
                <div style="font-size: 20px; font-weight: 900; color: #10b981; margin-top: 3px;">{tot_rev:,.2f}억원</div>
            </div>
        </div>
        """
        render_html(cards_html)
        st.divider()
        
        if items:
            df_pop = pd.DataFrame(items)
            df_pop['공급수량'] = df_pop['qty'].apply(format_k)
            df_pop['단가(원)'] = df_pop['price'].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
            df_pop['예상매출(억원)'] = df_pop['rev_ok'].apply(lambda x: f"{x:,.2f}억원")
            df_pop = df_pop.rename(columns={'model': '모델명', 'gh_code': 'GH-CODE', 'site': '출하 SITE'})
            df_pop = df_pop.sort_values(by='qty', ascending=False)
            
            st.dataframe(
                df_pop[['모델명', 'GH-CODE', '출하 SITE', '공급수량', '단가(원)', '예상매출(억원)']], 
                width="stretch", 
                hide_index=True
            )
            
            st.download_button(
                label=f"📥 {m_num}월 {d_num}일 공급 상세 다운로드 (CSV)",
                data=convert_df(df_pop[['모델명', 'GH-CODE', '출하 SITE', '공급수량', '단가(원)', '예상매출(억원)']]),
                file_name=f"fcst_supply_{m_num}M_{d_num}D_detail.csv",
                mime="text/csv",
                key=f"dl_popup_{m_num}_{d_num}"
            )
        else:
            st.info("해당 일자에 예정된 공급 일정이 없습니다.")

# 3-1. 단가 매핑 파싱 엔진
def parse_mapping_excel(file_bytes):
    try:
        df_temp = pd.read_excel(file_bytes, header=None)
        header_idx = 0
        for i, row in df_temp.head(15).iterrows():
            row_str = "".join([str(x).upper() for x in row.values])
            if 'GH' in row_str or '단가' in row_str or 'PRICE' in row_str or '모델' in row_str:
                header_idx = i; break
                
        df = pd.read_excel(file_bytes, header=header_idx)
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        gh_col = next((c for c in df.columns if 'GH' in c), None)
        mdl_col = next((c for c in df.columns if '모델' in c or 'MODEL' in c), None)
        prc_col = next((c for c in df.columns if '단가' in c or 'PRICE' in c or '금액' in c or 'USD' in c or '$' in c), None)
        
        mapping = {'gh_map': {}, 'suffix_map': {}, 'model_map': {}, 'gh_to_model': {}}
        if prc_col:
            for _, row in df.iterrows():
                p_val = str(row[prc_col]).replace('$', '').replace(',', '').strip()
                try: price = float(p_val)
                except: price = 0.0
                
                mdl_val = str(row[mdl_col]).strip() if mdl_col and pd.notna(row[mdl_col]) else ""
                if gh_col and pd.notna(row[gh_col]):
                    gh_raw = str(row[gh_col])
                    gh_k = clean_key(gh_raw)
                    gh_suf = get_gh_suffix(gh_raw)
                    if gh_k:
                        mapping['gh_map'][gh_k] = price
                        mapping['suffix_map'][gh_suf] = price
                        if mdl_val:
                            mapping['gh_to_model'][gh_k] = mdl_val
                            mapping['gh_to_model'][gh_suf] = mdl_val
                if mdl_val:
                    mdl_k = clean_key(mdl_val)
                    if mdl_k: mapping['model_map'][mdl_k] = price
        return mapping
    except: pass
    return None

# 3-2. 생산 실적 엑셀 파싱 엔진
def parse_production_excel(file_bytes, target_date_str):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    parsed = {
        'daily_prod_plan': 0, 'daily_prod_actual': 0,
        'weekly_labels': [], 'weekly_plan': [], 'weekly_actual': [],
        'monthly_plan': 0, 'monthly_actual': 0, 'monthly_months': ['11월'],
        'inventory_qty': 0, 'inventory_days': 0.0,
        'daily_models': [], 'daily_prod_p': [], 'daily_prod_a': []
    }
    summary_sheet_name = next((s for s in wb.sheetnames if '요약' in s), wb.sheetnames[0])
    sheet = wb[summary_sheet_name]
    
    for r in range(1, 30):
        if '당일완제품재고' in "".join([str(safe_cell(sheet, r, c) or "") for c in range(1, 10)]).replace(" ", ""):
            try:
                parsed['inventory_qty'] = int(float(str(safe_cell(sheet, r+2, 2) or safe_cell(sheet, r+2, 3) or 0).replace(',', '')))
                parsed['inventory_days'] = float(safe_cell(sheet, r+3, 2) or safe_cell(sheet, r+3, 3) or 0.0)
            except: pass
            break

    summary_start = 6 
    for r in range(1, 20):
        if '구분' in "".join([str(safe_cell(sheet, r, c) or "") for c in range(1, 10)]).replace(" ", ""):
            summary_start = r; break

    col_map = {}
    weekly_cols = {}
    w_start_col = None
    w_first_label = 'W45'
    
    for r in range(summary_start, summary_start + 3):
        for c in range(1, sheet.max_column + 1):
            val = safe_cell(sheet, r, c)
            if not val: continue
            if is_match_date(val, target_date_str) and 'target' not in col_map:
                col_map['target'] = c
            vs = str(val).replace(" ", "").upper()
            if re.fullmatch(r'W\d+', vs) and w_start_col is None:
                w_start_col = c; w_first_label = vs
            if 'W' in vs and '합계' in vs:
                m = re.search(r'(W\d+)', vs)
                if m and m.group(1) not in weekly_cols: weekly_cols[m.group(1)] = c
            if '월합계' in vs and '월합계' not in col_map:
                col_map['월합계'] = c
                m = re.search(r'(\d+월)', vs)
                if m: parsed['monthly_months'] = [m.group(1)]

    row_plan, row_act = None, None
    for r in range(summary_start + 1, summary_start + 15):
        val_a = str(safe_cell(sheet, r, 1) or "").replace(" ", "").strip()
        val_b = str(safe_cell(sheet, r, 2) or "").replace(" ", "").strip()
        if (val_a == '계획' or val_b == '계획') and row_plan is None: row_plan = r
        if (val_a == '실적' or val_b == '실적') and row_act is None: row_act = r
            
    def get_val(r, c):
        if not r or not c: return 0
        try:
            v = float(str(safe_cell(sheet, r, c) or 0).replace(",", "").replace("-", "0"))
            if 0 < abs(v) < 10000: v *= 1000
            return int(v)
        except: return 0
        
    parsed['daily_prod_plan'] = get_val(row_plan, col_map.get('target'))
    parsed['daily_prod_actual'] = get_val(row_act, col_map.get('target'))
    
    labels, wp, wa = [], [], []
    seen_labels = set()
    if w_start_col and weekly_cols:
        min_w_col = min(weekly_cols.values())
        w_first_plan, w_first_act = 0, 0
        for c in range(w_start_col, min_w_col):
            w_first_plan += get_val(row_plan, c)
            w_first_act += get_val(row_act, c)
        labels.append(w_first_label); wp.append(w_first_plan); wa.append(w_first_act)
        seen_labels.add(w_first_label)
        
    for lbl, c_idx in sorted(weekly_cols.items(), key=lambda x: x[1]):
        if lbl in seen_labels: continue
        labels.append(lbl); wp.append(get_val(row_plan, c_idx)); wa.append(get_val(row_act, c_idx))
        seen_labels.add(lbl)
        
    parsed['weekly_labels'] = labels if labels else ['W45', 'W46', 'W47', 'W48']
    parsed['weekly_plan'] = wp if wp else [0]*len(parsed['weekly_labels'])
    parsed['weekly_actual'] = wa if wa else [0]*len(parsed['weekly_labels'])
    
    m_plan = get_val(row_plan, col_map.get('월합계'))
    m_act = get_val(row_act, col_map.get('월합계'))
    if m_act == 0 and sum(wa) > 0: m_act = sum(wa)
    parsed['monthly_plan'] = m_plan
    parsed['monthly_actual'] = m_act

    model_sheet = next((s for s in wb.sheetnames if '모델' in s), None)
    if model_sheet:
        ms = wb[model_sheet]
        hr, pr, ar = -1, -1, -1
        for r in range(1, 10):
            if '주요모델별' in "".join([str(safe_cell(ms, r, c) or "") for c in range(1, 6)]).replace(" ", ""): hr = r; break
        if hr != -1:
            models, c_idx = [], []
            for c in range(2, ms.max_column):
                mn = str(safe_cell(ms, hr, c) or "").strip()
                if mn and mn.upper() != 'TOTAL' and 'None' not in mn: models.append(mn); c_idx.append(c)
            for r in range(hr + 1, hr + 15):
                rt = "".join([str(safe_cell(ms, r, c) or "") for c in range(1, 4)]).replace(" ", "")
                if '당일계획' in rt: pr = r
                if '당일실적' in rt: ar = r
            if pr != -1 and ar != -1:
                for c in c_idx:
                    try: parsed['daily_prod_p'].append(int(float(str(safe_cell(ms, pr, c) or 0).replace(",", "").replace("-", "0"))))
                    except: parsed['daily_prod_p'].append(0)
                    try: parsed['daily_prod_a'].append(int(float(str(safe_cell(ms, ar, c) or 0).replace(",", "").replace("-", "0"))))
                    except: parsed['daily_prod_a'].append(0)
                parsed['daily_models'] = models
    return parsed

# 3-3. 출하 실적 파싱 및 일자별 실적 맵(day_actual_totals) 생성 엔진
def parse_new_shipment_actual_excel(file_bytes, target_date_str, mapping_data):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    sheet_name = next((s for s in wb.sheetnames if any(k in s for k in ['출하', '일별', '실적', '11월'])), wb.sheetnames[0])
    sheet = wb[sheet_name]

    mapping_data.setdefault('gh_map', {})
    mapping_data.setdefault('suffix_map', {})
    mapping_data.setdefault('model_map', {})
    mapping_data.setdefault('gh_to_model', {})

    t_month, t_day = parse_target_day_info(target_date_str)

    gh_col, model_col, site_col, gubun_col = None, None, None, None
    header_row = 1
    
    for r in range(1, 5):
        for c in range(1, 15):
            val = str(safe_cell(sheet, r, c) or "").replace(" ", "").upper()
            if any(k in val for k in ['품번', 'GH']) and not gh_col: gh_col = c; header_row = r
            if any(k in val for k in ['모델명', 'MODEL']) and not model_col: model_col = c; header_row = r
            if any(k in val for k in ['업체명', 'SITE', '거래선']) and not site_col: site_col = c; header_row = r
            if any(k in val for k in ['구분', 'TYPE']) and not gubun_col: gubun_col = c; header_row = r

    gh_col, model_col, site_col, gubun_col = gh_col or 1, model_col or 3, site_col or 4, gubun_col or 6

    carryover_col = 6
    for c in range(1, 12):
        v_h1 = str(safe_cell(sheet, 1, c) or "") + str(safe_cell(sheet, 2, c) or "")
        if ('이월' in v_h1 or '전월' in v_h1) and ('KRW' in v_h1 or '원화' in v_h1 or '금액' in v_h1):
            carryover_col = c; break

    day_map = {}
    for c in range(1, sheet.max_column + 1):
        for r in range(1, header_row + 2):
            v = safe_cell(sheet, r, c)
            if not v: continue
            vs = str(v).replace(" ", "")
            m = re.search(r'^(\d{1,2})일?$', vs)
            if m:
                d = int(m.group(1))
                if 1 <= d <= 31 and d not in day_map:
                    qty_c, krw_c = c, c + 1
                    for sub_c in [c, c + 1, c + 2]:
                        sub_v = str(safe_cell(sheet, r + 1, sub_c) or "").replace(" ", "").upper()
                        if '수량' in sub_v: qty_c = sub_c
                        if 'KRW' in sub_v or '원화' in sub_v: krw_c = sub_c
                    day_map[d] = {'qty': qty_c, 'krw': krw_c}
                    break

    daily_ship_a, daily_rev_a = 0, 0.0
    weekly_ship_actual, weekly_rev_actual = {}, {}
    monthly_ship_actual, monthly_rev_actual = 0, 0.0
    model_data, site_data = {}, {}
    curr_gh, curr_model, curr_site = "", "", ""

    learned_prices = {}
    learned_suffix_prices = {}
    learned_model_prices = {}
    total_carryover_krw = 0.0
    
    day_actual_totals = {d: {'qty': 0, 'rev': 0.0, 'items': []} for d in day_map}

    for r in range(header_row + 2, sheet.max_row + 1):
        v_gh = safe_cell(sheet, r, gh_col)
        v_mdl = safe_cell(sheet, r, model_col)
        v_ste = safe_cell(sheet, r, site_col)
        clean_gbn = str(safe_cell(sheet, r, gubun_col) or "").replace(" ", "").strip()

        if v_gh: curr_gh = str(v_gh).strip()
        if v_mdl: curr_model = str(v_mdl).strip()
        if v_ste: curr_site = str(v_ste).strip()

        val_f = str(safe_cell(sheet, r, carryover_col) or 0).replace(',', '').strip()
        try: f_krw = float(val_f) if val_f else 0.0
        except: f_krw = 0.0

        is_summary = any(k in (str(v_gh) + str(v_mdl) + str(v_ste) + clean_gbn) for k in ['합계', '누계', '지지품', '총계', '총합', '차이'])
        if f_krw > 0 and not is_summary:
            total_carryover_krw += f_krw

        if is_summary: continue

        if clean_gbn == '실적':
            base_model = curr_model[:-1] if curr_model.endswith('E') else curr_model
            gh_clean = clean_key(curr_gh)
            gh_suf = get_gh_suffix(curr_gh)
            
            for d, cols in day_map.items():
                q_raw = str(safe_cell(sheet, r, cols['qty']) or 0).replace(',', '').replace('-', '0').strip()
                k_raw = str(safe_cell(sheet, r, cols['krw']) or 0).replace(',', '').replace('-', '0').strip()
                try: q_val = int(float(q_raw)) if q_raw else 0
                except: q_val = 0
                try: k_val = float(k_raw) if k_raw else 0.0
                except: k_val = 0.0

                rev_ok = k_val / 100_000_000.0

                day_actual_totals[d]['qty'] += q_val
                day_actual_totals[d]['rev'] += rev_ok
                if q_val > 0:
                    day_actual_totals[d]['items'].append({
                        'model': base_model, 'gh_code': curr_gh, 'site': curr_site,
                        'qty': q_val, 'price': (k_val / q_val) if q_val > 0 else 0.0,
                        'rev_ok': rev_ok
                    })

                if q_val > 0 and k_val > 0:
                    u_price = k_val / q_val
                    learned_prices[gh_clean] = u_price
                    learned_suffix_prices[gh_suf] = u_price
                    if base_model: learned_model_prices[clean_key(base_model)] = u_price

                try: w_lbl = f"W{datetime.date(2026, t_month, d).isocalendar()[1]}"
                except: w_lbl = 'W45'

                if d <= t_day:
                    weekly_ship_actual[w_lbl] = weekly_ship_actual.get(w_lbl, 0) + q_val
                    weekly_rev_actual[w_lbl] = weekly_rev_actual.get(w_lbl, 0.0) + rev_ok
                    monthly_ship_actual += q_val
                    monthly_rev_actual += rev_ok

                if d == t_day:
                    daily_ship_a += q_val
                    daily_rev_a += rev_ok

                    if base_model not in model_data: model_data[base_model] = {'Actual': 0, 'Rev_Act': 0.0}
                    model_data[base_model]['Actual'] += q_val
                    model_data[base_model]['Rev_Act'] += rev_ok

                    if curr_site not in site_data: site_data[curr_site] = {'Actual': 0, 'Rev_Act': 0.0}
                    site_data[curr_site]['Actual'] += q_val
                    site_data[curr_site]['Rev_Act'] += rev_ok

    if total_carryover_krw == 0:
        for r in range(max(1, sheet.max_row - 15), sheet.max_row + 1):
            val_f = str(safe_cell(sheet, r, carryover_col) or 0).replace(',', '').strip()
            try:
                v_num = float(val_f)
                if v_num > 100000000:
                    total_carryover_krw = max(total_carryover_krw, v_num)
            except: pass

    carryover_rev_ok = round(total_carryover_krw / 100_000_000.0, 2)

    for k, v in learned_prices.items():
        if k not in mapping_data['gh_map'] or mapping_data['gh_map'][k] == 0: mapping_data['gh_map'][k] = v
    for k, v in learned_suffix_prices.items():
        if k not in mapping_data['suffix_map'] or mapping_data['suffix_map'][k] == 0: mapping_data['suffix_map'][k] = v
    for k, v in learned_model_prices.items():
        if k not in mapping_data['model_map'] or mapping_data['model_map'][k] == 0: mapping_data['model_map'][k] = v

    return {
        'daily_ship_a': daily_ship_a,
        'daily_rev_a': round(daily_rev_a, 2),
        'carryover_rev_ok': carryover_rev_ok,
        'day_actual_totals': day_actual_totals,
        'weekly_ship_actual': weekly_ship_actual,
        'weekly_rev_actual': {k: round(v, 2) for k, v in weekly_rev_actual.items()},
        'monthly_ship_actual': monthly_ship_actual,
        'monthly_rev_actual': round(monthly_rev_actual, 2),
        'actual_by_model': model_data,
        'actual_by_site': site_data
    }

# 3-4. RTF Main 출하 계획 및 매출 계획 파싱 엔진
def parse_rtf_main_plan_excel(file_bytes, target_date_str, mapping_data, exchange_rate=1350.0):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    sheet = wb['Main'] if 'Main' in wb.sheetnames else wb.active

    t_m, t_d = parse_target_day_info(target_date_str)
    target_date_obj = datetime.date(2026, t_m, t_d)
    target_iso_week = f"W{target_date_obj.isocalendar()[1]}"
    target_month_lbl = f"{t_m}월"

    date_cols = []
    fcst_date_cols = []
    for c in range(1, sheet.max_column + 1):
        v = safe_cell(sheet, 4, c)
        if isinstance(v, (datetime.datetime, datetime.date)):
            y = v.year if v.year > 2000 else 2026
            d_obj = datetime.date(y, v.month, v.day)
            w_lbl = f"W{d_obj.isocalendar()[1]}"
            m_lbl = f"{d_obj.month}월"
            d_str = d_obj.strftime('%m/%d')
            date_cols.append((c, d_obj, w_lbl, m_lbl, d_str))
            fcst_date_cols.append(d_str)
        elif v:
            s_val = str(v).strip().upper()
            if not re.search(r'\bW\d+\b', s_val):
                m_dt = re.search(r'\b0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b', s_val)
                if m_dt:
                    d_obj = datetime.date(2026, int(m_dt.group(1)), int(m_dt.group(2)))
                    d_str = d_obj.strftime('%m/%d')
                    date_cols.append((c, d_obj, f"W{d_obj.isocalendar()[1]}", f"{d_obj.month}월", d_str))
                    fcst_date_cols.append(d_str)

    site_col = 3
    gh_col = 4
    gubun_col = 9

    gh_map = mapping_data.get('gh_map', {})
    suffix_map = mapping_data.get('suffix_map', {})
    gh_to_model = mapping_data.get('gh_to_model', {})
    model_map = mapping_data.get('model_map', {})

    known_prices = [p for p in list(gh_map.values()) + list(suffix_map.values()) + list(model_map.values()) if p > 0]
    avg_price = (sum(known_prices) / len(known_prices)) if known_prices else 8800.0

    daily_ship_p = 0
    daily_rev_p = 0.0
    weekly_ship_plan, weekly_rev_plan = {}, {}
    monthly_ship_plan, monthly_rev_plan = {}, {}
    plan_by_model, plan_by_site = {}, {}
    fcst_rows = []

    for r in range(5, sheet.max_row + 1):
        v_gubun = str(safe_cell(sheet, r, gubun_col) or "").replace(" ", "").strip()
        site = str(safe_cell(sheet, r, site_col) or "").strip()
        gh_code = str(safe_cell(sheet, r, gh_col) or "").strip()

        if not v_gubun or v_gubun not in ['공급', '공 급']: continue
        if not site and not gh_code: continue

        gh_clean = clean_key(gh_code)
        gh_suf = get_gh_suffix(gh_code)
        proj_val = str(safe_cell(sheet, r, 6) or "").strip()
        model_name = resolve_model_name(gh_code, proj_val, gh_to_model)

        price = gh_map.get(gh_clean, 0.0)
        if price == 0: price = suffix_map.get(gh_suf, 0.0)
        if price == 0 and model_name: price = model_map.get(clean_key(model_name), 0.0)
        if price == 0:
            for k, v in gh_map.items():
                if len(k) >= 5 and (k in gh_clean or gh_clean in k or gh_suf in k):
                    price = v; break
        if price == 0:
            price = avg_price

        price_krw = price if price > 100 else price * exchange_rate

        row_fcst = {
            'SITE': site,
            'GH_CODE': gh_code,
            '모델명': model_name,
            '단가_원': price_krw
        }
        row_fcst_total = 0

        for c_idx, d_obj, w_lbl, m_lbl, d_str in date_cols:
            q_raw = str(safe_cell(sheet, r, c_idx) or 0).replace(',', '').replace('-', '0').strip()
            try: q_val = int(float(q_raw)) if q_raw else 0
            except: q_val = 0

            row_fcst[d_str] = q_val
            row_fcst_total += q_val

            rev_val = (q_val * price_krw) / 100_000_000.0

            weekly_ship_plan[w_lbl] = weekly_ship_plan.get(w_lbl, 0) + q_val
            weekly_rev_plan[w_lbl] = weekly_rev_plan.get(w_lbl, 0.0) + rev_val
            monthly_ship_plan[m_lbl] = monthly_ship_plan.get(m_lbl, 0) + q_val
            monthly_rev_plan[m_lbl] = monthly_rev_plan.get(m_lbl, 0.0) + rev_val

            if d_obj == target_date_obj:
                daily_ship_p += q_val
                daily_rev_p += rev_val
                
                if model_name:
                    if model_name not in plan_by_model: plan_by_model[model_name] = {'Plan_Qty': 0, 'Plan_Rev': 0.0}
                    plan_by_model[model_name]['Plan_Qty'] += q_val
                    plan_by_model[model_name]['Plan_Rev'] += rev_val

                if site:
                    if site not in plan_by_site: plan_by_site[site] = {'Plan_Qty': 0, 'Plan_Rev': 0.0}
                    plan_by_site[site]['Plan_Qty'] += q_val
                    plan_by_site[site]['Plan_Rev'] += rev_val

        row_fcst['합계'] = row_fcst_total
        if row_fcst_total > 0:
            fcst_rows.append(row_fcst)

    return {
        'daily_ship_p': daily_ship_p,
        'daily_rev_p': round(daily_rev_p, 2),
        'cur_week_ship_p': weekly_ship_plan.get(target_iso_week, 0),
        'cur_week_rev_p': round(weekly_rev_plan.get(target_iso_week, 0.0), 2),
        'cur_month_ship_p': monthly_ship_plan.get(target_month_lbl, 0),
        'cur_month_rev_p': round(monthly_rev_plan.get(target_month_lbl, 0.0), 2),
        'weekly_ship_plan': weekly_ship_plan,
        'weekly_rev_plan': {k: round(v, 2) for k, v in weekly_rev_plan.items()},
        'monthly_ship_plan': monthly_ship_plan,
        'monthly_rev_plan': {k: round(v, 2) for k, v in monthly_rev_plan.items()},
        'plan_by_model': plan_by_model,
        'plan_by_site': plan_by_site,
        'fcst_rows': fcst_rows,
        'fcst_date_cols': fcst_date_cols
    }

# 4. CSS 스타일링
render_html("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;700;900&display=swap');
    html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
    .stApp { background-color: #f4f4f0; color: #111111; }
    section[data-testid="stSidebar"] { background-color: #111111; border-right: 4px solid #0044ff; }
    section[data-testid="stSidebar"] * { color: #ffffff !important; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border: 2px solid #111111; padding: 10px 24px; font-weight: 800; color: #111111; box-shadow: 4px -2px 0px #0044ff; }
    .stTabs [aria-selected="true"] { background-color: #0044ff !important; color: #ffffff !important; }
    .date-badge {
        background-color: #ffffff;
        border: 2.5px solid #111111;
        padding: 8px 18px;
        display: inline-block;
        font-weight: 800;
        font-size: 15px;
        margin-bottom: 18px;
        box-shadow: 4px 4px 0px #0044ff;
    }
    div[data-testid="stDialog"] div[role="dialog"] {
        max-width: 900px !important;
        width: 90% !important;
    }
    
    /* 캘린더 일별 카드 및 호버 팝업 스타일 */
    .cal-card-wrapper {
        position: relative;
        margin-bottom: 6px;
    }
    .cal-card {
        min-height: 105px;
        border-radius: 4px;
        padding: 8px 9px;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
        cursor: pointer;
    }
    .cal-card-wrapper:hover .cal-card {
        transform: translateY(-2px);
        box-shadow: 4px 4px 0px #0044ff !important;
    }
    .cal-hover-popup {
        display: none;
        position: absolute;
        top: 100%;
        width: 330px;
        background-color: #ffffff;
        border: 2.5px solid #111111;
        border-radius: 6px;
        padding: 12px;
        box-shadow: 6px 6px 0px #0044ff;
        z-index: 999999;
        pointer-events: none;
    }
    .cal-card-wrapper:hover .cal-hover-popup {
        display: block;
    }
    .pos-left { left: 0; transform: none; }
    .pos-center { left: 50%; transform: translateX(-50%); }
    .pos-right { right: 0; left: auto; transform: none; }
</style>
""")

# 5. 사이드바
with st.sidebar:
    st.markdown("### 🔄 데이터 동기화")
    if st.button("구글 시트 최신 데이터 불러오기", width="stretch"):
        load_data_from_sheet.clear()
        sync_session_from_sheet()
        st.success("동기화 완료!")
        st.rerun()
    st.markdown("---")
    menu = st.radio("📂 Folder", (
        "🏠 현황", "1. 모델별 출하실적", "2. 모델별 생산실적", 
        "3. 출하 SITE별 실적", "4. 고객 FCST 조회", "5. 완제품 재고 조회", 
        "6. 캘린더", "7. 데이터 관리 (Admin)"
    ))

def create_active_chart(title, categories, plan_data, actual_data, name1="계획", name2="실적", unit="개"):
    fig = go.Figure()
    if not categories: categories, plan_data, actual_data = ["-"], [0], [0]
    
    if unit == "억원":
        txt_p = [f"{v:,.2f}" if v > 0 else "" for v in plan_data]
        txt_a = [f"{v:,.2f}" if v > 0 else "" for v in actual_data]
    else:
        txt_p = [format_k(v) if v > 0 else "" for v in plan_data]
        txt_a = [format_k(v) if v > 0 else "" for v in actual_data]
        
    fig.add_trace(go.Bar(x=categories, y=plan_data, name=name1, marker_color='#d1d5db', marker_line=dict(width=1.5, color='#111111'), text=txt_p, textposition='outside'))
    fig.add_trace(go.Bar(x=categories, y=actual_data, name=name2, marker_color='#0044ff', marker_line=dict(width=1.5, color='#111111'), text=txt_a, textposition='outside'))
    fig.update_layout(title=dict(text=f"<b>{title} ({unit})</b>", font=dict(size=14, color='#111111')), barmode='group', plot_bgcolor='rgba(0,0,0,0)', hovermode='x unified', margin=dict(l=20, r=20, t=45, b=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    return fig

# 6. 메인 분기 로직
if menu == "7. 데이터 관리 (Admin)":
    st.title("⚙️ DATA MANAGEMENT")
    if st.text_input("비밀번호", type="password") == "2250":
        c_set1, c_set2 = st.columns(2)
        with c_set1: target_date_input = st.text_input("📅 조회 기준 날짜 지정 (예: 11/2)", default_target_date)
        with c_set2: ex_rate_input = st.number_input("💵 적용 환율 (원/USD)", value=1350.0, step=10.0)
        
        col1, col2 = st.columns(2)
        with col1: prod_file = st.file_uploader("1. 생산 실적 파일", type=['xlsx'], key="admin_prod_file")
        with col2: ship_actual_file = st.file_uploader("2. 출하 실적 파일 (신규 양식)", type=['xlsx'], key="admin_ship_actual_file")

        col3, col4 = st.columns(2)
        with col3: ship_plan_file = st.file_uploader("3. 출하 계획 파일 (RTF Main 시트)", type=['xlsx'], key="admin_ship_plan_file")
        with col4: map_rev_file = st.file_uploader("4. 단가 매핑 파일 (선택)", type=['xlsx'], key="admin_map_rev_file")
            
        if st.button("🚀 데이터 분석 및 반영", type="primary"):
            st.session_state['target_date'] = target_date_input

            # 1. 단가 매핑 반영
            if map_rev_file:
                nmap = parse_mapping_excel(map_rev_file)
                if nmap: 
                    st.session_state['mapping_data'].update(nmap)
                    st.success(f"✅ 단가표 인식 완료! (GH-CODE {len(nmap['gh_map'])}개 등록)")

            # 2. 생산 실적 반영
            if prod_file:
                with st.spinner("생산 데이터를 분석하고 있습니다..."):
                    try:
                        p = parse_production_excel(prod_file, target_date_input)
                        st.session_state['kpi_data']['daily']['prod_p'] = p['daily_prod_plan']
                        st.session_state['kpi_data']['daily']['prod_a'] = p['daily_prod_actual']
                        st.session_state['chart_data']['daily_models'] = p['daily_models']
                        st.session_state['chart_data']['daily_prod_p'] = p['daily_prod_p']
                        st.session_state['chart_data']['daily_prod_a'] = p['daily_prod_a']
                        
                        st.session_state['chart_data']['weekly_weeks'] = list(dict.fromkeys(p['weekly_labels']))
                        st.session_state['chart_data']['weekly_prod_p'] = p['weekly_plan']
                        st.session_state['chart_data']['weekly_prod_a'] = p['weekly_actual']
                        st.session_state['kpi_data']['weekly']['prod_p'] = sum(p['weekly_plan'])
                        st.session_state['kpi_data']['weekly']['prod_a'] = sum(p['weekly_actual'])
                        
                        st.session_state['chart_data']['monthly_months'] = p['monthly_months']
                        st.session_state['chart_data']['monthly_prod_p'] = [p['monthly_plan']]
                        st.session_state['chart_data']['monthly_prod_a'] = [p['monthly_actual']]
                        st.session_state['kpi_data']['monthly']['prod_p'] = p['monthly_plan']
                        st.session_state['kpi_data']['monthly']['prod_a'] = p['monthly_actual']
                        st.session_state['inventory_data']['qty'] = p['inventory_qty']
                        st.session_state['inventory_data']['days'] = p['inventory_days']
                        st.session_state['data_loaded'] = True
                        st.success(f"✅ 생산 데이터 반영 완료! (Daily 계획: {format_k(p['daily_prod_plan'])} / 실적: {format_k(p['daily_prod_actual'])})")
                    except Exception as e: st.error(f"생산 파싱 오류: {e}")
            
            # 3. 출하 실적 및 F열 이월 매출 자동 반영
            if ship_actual_file:
                with st.spinner("출하 실적 및 F열 전월 이월 매출을 분석하고 있습니다..."):
                    try:
                        s_act = parse_new_shipment_actual_excel(ship_actual_file, target_date_input, st.session_state['mapping_data'])
                        if s_act:
                            carryover = s_act['carryover_rev_ok']
                            
                            st.session_state['kpi_data']['daily']['ship_a'] = s_act['daily_ship_a']
                            st.session_state['kpi_data']['daily']['rev_a'] = s_act['daily_rev_a']
                            st.session_state['actual_by_day'] = s_act['day_actual_totals']
                            
                            w_labels = st.session_state['chart_data'].get('weekly_weeks', ['W45', 'W46', 'W47', 'W48'])
                            w_ship_a = [s_act['weekly_ship_actual'].get(w, 0) for w in w_labels]
                            
                            w_rev_a = []
                            for idx, w in enumerate(w_labels):
                                pure_w = s_act['weekly_rev_actual'].get(w, 0.0)
                                if idx == 0:
                                    w_rev_a.append(round(pure_w + carryover, 2))
                                else:
                                    w_rev_a.append(round(pure_w, 2))
                                    
                            pure_weekly_rev = sum(s_act['weekly_rev_actual'].values())
                            st.session_state['chart_data']['weekly_ship_a'] = w_ship_a
                            st.session_state['chart_data']['weekly_rev_a'] = w_rev_a
                            st.session_state['kpi_data']['weekly']['ship_a'] = sum(w_ship_a)
                            st.session_state['kpi_data']['weekly']['pure_rev_a'] = round(pure_weekly_rev, 2)
                            st.session_state['kpi_data']['weekly']['carryover_rev'] = carryover
                            st.session_state['kpi_data']['weekly']['rev_a'] = round(pure_weekly_rev + carryover, 2)
                            
                            pure_monthly_rev = s_act['monthly_rev_actual']
                            st.session_state['kpi_data']['monthly']['ship_a'] = s_act['monthly_ship_actual']
                            st.session_state['kpi_data']['monthly']['pure_rev_a'] = pure_monthly_rev
                            st.session_state['kpi_data']['monthly']['carryover_rev'] = carryover
                            st.session_state['kpi_data']['monthly']['rev_a'] = round(pure_monthly_rev + carryover, 2)
                            st.session_state['chart_data']['monthly_ship_a'] = [s_act['monthly_ship_actual']]
                            st.session_state['chart_data']['monthly_rev_a'] = [round(pure_monthly_rev + carryover, 2)]
                            
                            st.session_state['_ship_actual_by_model'] = s_act['actual_by_model']
                            st.session_state['_ship_actual_by_site'] = s_act['actual_by_site']
                            
                            st.session_state['data_loaded'] = True
                            st.success(f"✅ 출하 실적 반영 완료! (Daily 실적: {format_k(s_act['daily_ship_a'])} / {s_act['daily_rev_a']:,.2f}억원 | F열 이월 매출: {carryover:,.2f}억원)")
                    except Exception as e: st.error(f"출하 실적 파싱 오류: {e}")

            # 4. 출하 계획 및 매출 계획 반영 (RTF Main)
            if ship_plan_file:
                with st.spinner("출하 및 매출 계획(수량 x 단가)을 계산하고 있습니다..."):
                    try:
                        s_plan = parse_rtf_main_plan_excel(ship_plan_file, target_date_input, st.session_state['mapping_data'], ex_rate_input)
                        if s_plan:
                            st.session_state['kpi_data']['daily']['ship_p'] = s_plan['daily_ship_p']
                            st.session_state['kpi_data']['daily']['rev_p'] = s_plan['daily_rev_p']
                            st.session_state['kpi_data']['weekly']['ship_p'] = s_plan['cur_week_ship_p']
                            st.session_state['kpi_data']['weekly']['rev_p'] = s_plan['cur_week_rev_p']
                            st.session_state['kpi_data']['monthly']['ship_p'] = s_plan['cur_month_ship_p']
                            st.session_state['kpi_data']['monthly']['rev_p'] = s_plan['cur_month_rev_p']

                            w_labels = st.session_state['chart_data'].get('weekly_weeks', ['W45', 'W46', 'W47', 'W48'])
                            st.session_state['chart_data']['weekly_ship_p'] = [s_plan['weekly_ship_plan'].get(w, 0) for w in w_labels]
                            st.session_state['chart_data']['weekly_rev_p'] = [s_plan['weekly_rev_plan'].get(w, 0.0) for w in w_labels]

                            m_labels = st.session_state['chart_data'].get('monthly_months', ['11월'])
                            st.session_state['chart_data']['monthly_ship_p'] = [s_plan['monthly_ship_plan'].get(m, 0) for m in m_labels]
                            st.session_state['chart_data']['monthly_rev_p'] = [s_plan['monthly_rev_plan'].get(m, 0.0) for m in m_labels]

                            st.session_state['_ship_plan_by_model'] = s_plan['plan_by_model']
                            st.session_state['_ship_plan_by_site'] = s_plan['plan_by_site']
                            
                            st.session_state['fcst_data'] = {
                                'rows': s_plan.get('fcst_rows', []),
                                'date_cols': s_plan.get('fcst_date_cols', [])
                            }

                            st.session_state['data_loaded'] = True
                            st.success(f"✅ 출하 & 매출 계획 반영 완료! (Daily 계획: {format_k(s_plan['daily_ship_p'])} / {s_plan['daily_rev_p']:,.2f}억원)")
                    except Exception as e: st.error(f"출하 계획 파싱 오류: {e}")

            # 5. 모델별 / SITE별 계획+실적 통합 병합
            p_models = st.session_state.get('_ship_plan_by_model', {})
            a_models = st.session_state.get('_ship_actual_by_model', {})
            all_models = sorted(list(set(list(p_models.keys()) + list(a_models.keys()))))
            merged_models = []
            for m in all_models:
                p_info = p_models.get(m, {'Plan_Qty': 0, 'Plan_Rev': 0.0})
                a_info = a_models.get(m, {'Actual': 0, 'Rev_Act': 0.0})
                p_qty = p_info['Plan_Qty']
                p_rev = round(p_info['Plan_Rev'], 2)
                a_qty = a_info['Actual']
                a_rev = round(a_info['Rev_Act'], 2)
                
                qty_rate = round((a_qty / p_qty * 100), 1) if p_qty > 0 else 0.0
                rev_rate = round((a_rev / p_rev * 100), 1) if p_rev > 0 else 0.0
                
                merged_models.append({
                    '모델명': m,
                    '출하계획(개)': p_qty,
                    '출하계획(K)': format_k(p_qty),
                    '출하실적(개)': a_qty,
                    '출하실적(K)': format_k(a_qty),
                    '수량달성률(%)': qty_rate,
                    '수량차질(K)': format_k(p_qty - a_qty),
                    '매출계획(억원)': p_rev,
                    '매출실적(억원)': a_rev,
                    '매출달성률(%)': rev_rate
                })
            st.session_state['chart_data']['ship_by_model'] = merged_models

            p_sites = st.session_state.get('_ship_plan_by_site', {})
            a_sites = st.session_state.get('_ship_actual_by_site', {})
            all_sites = sorted(list(set(list(p_sites.keys()) + list(a_sites.keys()))))
            merged_sites = []
            for s in all_sites:
                p_info = p_sites.get(s, {'Plan_Qty': 0, 'Plan_Rev': 0.0})
                a_info = a_sites.get(s, {'Actual': 0, 'Rev_Act': 0.0})
                p_qty = p_info['Plan_Qty']
                p_rev = round(p_info['Plan_Rev'], 2)
                a_qty = a_info['Actual']
                a_rev = round(a_info['Rev_Act'], 2)
                
                qty_rate = round((a_qty / p_qty * 100), 1) if p_qty > 0 else 0.0
                rev_rate = round((a_rev / p_rev * 100), 1) if p_rev > 0 else 0.0
                
                merged_sites.append({
                    'SITE': s,
                    '출하계획(개)': p_qty,
                    '출하계획(K)': format_k(p_qty),
                    '출하실적(개)': a_qty,
                    '출하실적(K)': format_k(a_qty),
                    '수량달성률(%)': qty_rate,
                    '수량차질(K)': format_k(p_qty - a_qty),
                    '매출계획(억원)': p_rev,
                    '매출실적(억원)': a_rev,
                    '매출달성률(%)': rev_rate
                })
            st.session_state['chart_data']['ship_by_site'] = merged_sites

            # 6. 구글 스프레드시트에 영구 동기화
            payload = {
                'target_date': st.session_state['target_date'],
                'kpi_data': st.session_state['kpi_data'],
                'chart_data': st.session_state['chart_data'],
                'inventory_data': st.session_state['inventory_data'],
                'fcst_data': st.session_state['fcst_data'],
                'mapping_data': st.session_state['mapping_data'],
                'actual_by_day': st.session_state['actual_by_day'],
                '_ship_plan_by_model': st.session_state.get('_ship_plan_by_model', {}),
                '_ship_plan_by_site': st.session_state.get('_ship_plan_by_site', {}),
                '_ship_actual_by_model': st.session_state.get('_ship_actual_by_model', {}),
                '_ship_actual_by_site': st.session_state.get('_ship_actual_by_site', {})
            }
            with st.spinner("구글 스프레드시트에 데이터를 동기화하고 있습니다..."):
                ok, msg = save_data_to_sheet(payload)
                if ok:
                    st.success("☁️ 구글 스프레드시트에 영구 저장 완료! 이제 핸드폰과 모든 팀원의 화면에도 동일하게 반영됩니다.")
                else:
                    st.warning(f"⚠️ 구글 시트 저장 안내: {msg} (현재 세션에만 반영되었습니다)")

elif menu == "🏠 현황":
    st.title("📊 실시간 현황")
    
    cur_target_date = st.session_state.get('target_date', default_target_date)
    render_html(f"""
    <div class="date-badge">
        📅 조회 기준일자 : <span style="color: #0044ff;">{cur_target_date}</span>
    </div>
    """)

    kpi = st.session_state.get('kpi_data', default_kpis)
    cdata = st.session_state.get('chart_data', default_charts)

    tab_monthly, tab_weekly, tab_daily = st.tabs(["🗓️ Monthly", "📊 Weekly", "📅 Daily"])

    # 1. Monthly 탭
    with tab_monthly:
        m1, m2, m3 = st.columns(3)
        with m1: 
            render_kpi_card("월간 생산 누계", "🏭", kpi['monthly']['prod_a'], kpi['monthly']['prod_p'])
            render_achievement_gauge("생산", kpi['monthly']['prod_a'], kpi['monthly']['prod_p'], unit="K")
        with m2: 
            render_kpi_card("월간 출하 누계", "🚚", kpi['monthly']['ship_a'], kpi['monthly']['ship_p'])
            render_achievement_gauge("출하", kpi['monthly']['ship_a'], kpi['monthly']['ship_p'], unit="K")
        with m3: 
            rev_sub = f"당월 실적 {kpi['monthly']['pure_rev_a']:,.2f}억 + 이월 매출 {kpi['monthly']['carryover_rev']:,.2f}억"
            render_kpi_card("월간 매출 누계", "💰", kpi['monthly']['rev_a'], kpi['monthly']['rev_p'], is_rev=True, sub_caption=rev_sub)
            render_achievement_gauge("매출", kpi['monthly']['rev_a'], kpi['monthly']['rev_p'], unit="억원")

        mch1, mch2, mch3 = st.columns(3)
        m_months = cdata.get('monthly_months', ['11월'])
        with mch1:
            st.plotly_chart(create_active_chart("월별 생산 실적", m_months, cdata.get('monthly_prod_p', [0]), cdata.get('monthly_prod_a', [0]), unit="K"), width="stretch")
        with mch2:
            st.plotly_chart(create_active_chart("월별 출하 실적", m_months, cdata.get('monthly_ship_p', [0]), cdata.get('monthly_ship_a', [0]), unit="K"), width="stretch")
        with mch3:
            st.plotly_chart(create_active_chart("월별 매출 실적", m_months, cdata.get('monthly_rev_p', [0.0]), cdata.get('monthly_rev_a', [0.0]), unit="억원"), width="stretch")

    # 2. Weekly 탭
    with tab_weekly:
        wk1, wk2, wk3 = st.columns(3)
        with wk1: 
            render_kpi_card("주간 생산 누계", "🏭", kpi['weekly']['prod_a'], kpi['weekly']['prod_p'])
            render_achievement_gauge("생산", kpi['weekly']['prod_a'], kpi['weekly']['prod_p'], unit="K")
        with wk2: 
            render_kpi_card("주간 출하 누계", "🚚", kpi['weekly']['ship_a'], kpi['weekly']['ship_p'])
            render_achievement_gauge("출하", kpi['weekly']['ship_a'], kpi['weekly']['ship_p'], unit="K")
        with wk3: 
            rev_sub = f"당월 실적 {kpi['weekly']['pure_rev_a']:,.2f}억 + 이월 매출 {kpi['weekly']['carryover_rev']:,.2f}억"
            render_kpi_card("주간 매출 누계", "💰", kpi['weekly']['rev_a'], kpi['weekly']['rev_p'], is_rev=True, sub_caption=rev_sub)
            render_achievement_gauge("매출", kpi['weekly']['rev_a'], kpi['weekly']['rev_p'], unit="억원")

        wch1, wch2, wch3 = st.columns(3)
        w_weeks = cdata.get('weekly_weeks', ['W45', 'W46', 'W47', 'W48'])
        with wch1:
            st.plotly_chart(create_active_chart("주차별 생산 실적", w_weeks, cdata.get('weekly_prod_p', [0]*len(w_weeks)), cdata.get('weekly_prod_a', [0]*len(w_weeks)), unit="K"), width="stretch")
        with wch2:
            st.plotly_chart(create_active_chart("주차별 출하 실적", w_weeks, cdata.get('weekly_ship_p', [0]*len(w_weeks)), cdata.get('weekly_ship_a', [0]*len(w_weeks)), unit="K"), width="stretch")
        with wch3:
            st.plotly_chart(create_active_chart("주차별 매출 실적", w_weeks, cdata.get('weekly_rev_p', [0.0]*len(w_weeks)), cdata.get('weekly_rev_a', [0.0]*len(w_weeks)), unit="억원"), width="stretch")

    # 3. Daily 탭
    with tab_daily:
        k1, k2, k3 = st.columns(3)
        with k1: 
            render_kpi_card("생산 (TOTAL)", "🏭", kpi['daily']['prod_a'], kpi['daily']['prod_p'])
            render_achievement_gauge("생산", kpi['daily']['prod_a'], kpi['daily']['prod_p'], unit="K")
        with k2: 
            render_kpi_card("출하 (TOTAL)", "🚚", kpi['daily']['ship_a'], kpi['daily']['ship_p'])
            render_achievement_gauge("출하", kpi['daily']['ship_a'], kpi['daily']['ship_p'], unit="K")
        with k3: 
            render_kpi_card("매출 (TOTAL)", "💰", kpi['daily']['rev_a'], kpi['daily']['rev_p'], is_rev=True)
            render_achievement_gauge("매출", kpi['daily']['rev_a'], kpi['daily']['rev_p'], unit="억원")

        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            st.plotly_chart(create_active_chart("생산 실적 (TOTAL)", ["TOTAL"], [kpi['daily']['prod_p']], [kpi['daily']['prod_a']], unit="K"), width="stretch")
        with ch2:
            st.plotly_chart(create_active_chart("출하 실적 (TOTAL)", ["TOTAL"], [kpi['daily']['ship_p']], [kpi['daily']['ship_a']], unit="K"), width="stretch")
        with ch3:
            st.plotly_chart(create_active_chart("매출 실적 (TOTAL)", ["TOTAL"], [kpi['daily']['rev_p']], [kpi['daily']['rev_a']], unit="억원"), width="stretch")

elif menu == "1. 모델별 출하실적":
    st.subheader("📦 1. 모델별 출하 및 매출 현황 (계획 vs 실적)")
    df_ship = pd.DataFrame(st.session_state['chart_data'].get('ship_by_model', []))
    if not df_ship.empty:
        df_ship = df_ship.sort_values('출하실적(개)', ascending=False)
        
        tab_qty, tab_rev = st.tabs(["🚚 출하 수량 현황 (K)", "💰 매출액 현황 (억원)"])
        with tab_qty:
            st.plotly_chart(create_active_chart("모델별 출하 수량 (계획 vs 실적)", df_ship['모델명'].tolist(), df_ship['출하계획(개)'].tolist(), df_ship['출하실적(개)'].tolist(), unit="K"), width="stretch")
        with tab_rev:
            st.plotly_chart(create_active_chart("모델별 매출액 (계획 vs 실적)", df_ship['모델명'].tolist(), df_ship['매출계획(억원)'].tolist(), df_ship['매출실적(억원)'].tolist(), unit="억원"), width="stretch")
            
        col1, col2 = st.columns([0.8, 0.2])
        with col2: st.download_button(label="📥 엑셀(CSV) 다운로드", data=convert_df(df_ship), file_name='model_shipment_performance.csv', mime='text/csv')
        st.dataframe(df_ship[['모델명', '출하계획(K)', '출하실적(K)', '수량달성률(%)', '수량차질(K)', '매출계획(억원)', '매출실적(억원)', '매출달성률(%)']], width="stretch", hide_index=True)
    else: st.info("출하 실적 및 계획 데이터를 업로드해 주세요.")
        
elif menu == "2. 모델별 생산실적":
    st.subheader("🛠️ 2. 모델별 생산실적")
    cdata = st.session_state['chart_data']
    st.plotly_chart(create_active_chart("모델별 생산 현황 (계획 vs 실적)", cdata.get('daily_models', []), cdata.get('daily_prod_p', []), cdata.get('daily_prod_a', []), unit="K"), width="stretch")
    
    df_prod = pd.DataFrame({"모델명": cdata.get('daily_models', []), "생산계획": cdata.get('daily_prod_p', []), "생산실적": cdata.get('daily_prod_a', [])})
    if not df_prod.empty:
        df_prod['생산계획(K)'] = df_prod['생산계획'].apply(format_k)
        df_prod['생산실적(K)'] = df_prod['생산실적'].apply(format_k)
        df_prod['달성률(%)'] = (df_prod['생산실적'] / df_prod['생산계획'].replace(0, 1) * 100).round(1)
        df_prod['차질수량(K)'] = (df_prod['생산계획'] - df_prod['생산실적']).apply(format_k)
        col1, col2 = st.columns([0.8, 0.2])
        with col2: st.download_button(label="📥 엑셀(CSV) 다운로드", data=convert_df(df_prod), file_name='model_production.csv', mime='text/csv')
        st.dataframe(df_prod[['모델명', '생산계획(K)', '생산실적(K)', '달성률(%)', '차질수량(K)']], width="stretch", hide_index=True)

elif menu == "3. 출하 SITE별 실적":
    st.subheader("📍 3. 출하 SITE별 출하 및 매출 현황 (계획 vs 실적)")
    df_site = pd.DataFrame(st.session_state['chart_data'].get('ship_by_site', []))
    if not df_site.empty:
        df_site = df_site.sort_values('출하실적(개)', ascending=False)
        
        tab_s_qty, tab_s_rev = st.tabs(["🚚 SITE별 출하 수량 (K)", "💰 SITE별 매출액 (억원)"])
        with tab_s_qty:
            st.plotly_chart(create_active_chart("SITE별 출하 수량 (계획 vs 실적)", df_site['SITE'].tolist(), df_site['출하계획(개)'].tolist(), df_site['출하실적(개)'].tolist(), unit="K"), width="stretch")
        with tab_s_rev:
            st.plotly_chart(create_active_chart("SITE별 매출액 (계획 vs 실적)", df_site['SITE'].tolist(), df_site['매출계획(억원)'].tolist(), df_site['매출실적(억원)'].tolist(), unit="억원"), width="stretch")
            
        col1, col2 = st.columns([0.8, 0.2])
        with col2: st.download_button(label="📥 엑셀(CSV) 다운로드", data=convert_df(df_site), file_name='site_shipment_performance.csv', mime='text/csv')
        st.dataframe(df_site[['SITE', '출하계획(K)', '출하실적(K)', '수량달성률(%)', '수량차질(K)', '매출계획(억원)', '매출실적(억원)', '매출달성률(%)']], width="stretch", hide_index=True)
    else: st.info("출하 실적 및 계획 데이터를 업로드해 주세요.")

elif menu == "4. 고객 FCST 조회":
    st.title("📈 4. 고객 FCST 조회 (모델별 공급 일정)")
    fcst_data = st.session_state.get('fcst_data', {'rows': [], 'date_cols': []})
    fcst_rows = fcst_data.get('rows', [])
    date_cols = fcst_data.get('date_cols', [])

    if not fcst_rows:
        st.info("💡 출하 계획 파일(RTF Main 시트)을 **[7. 데이터 관리 (Admin)]** 메뉴에서 업로드 및 분석해 주세요.")
    else:
        df_fcst = pd.DataFrame(fcst_rows)
        
        total_fcst_qty = int(df_fcst['합계'].sum())
        model_count = int(df_fcst['모델명'].nunique())
        site_count = int(df_fcst['SITE'].nunique())
        top_model = df_fcst.groupby('모델명')['합계'].sum().idxmax()
        top_model_qty = int(df_fcst.groupby('모델명')['합계'].sum().max())
        
        m1, m2, m3, m4 = st.columns(4)
        with m1: st.metric("📦 총 공급 계획 물량", f"{format_k(total_fcst_qty)}")
        with m2: st.metric("📱 계획 대상 모델 수", f"{model_count} 개 모델")
        with m3: st.metric("🏭 공급 대상 SITE 수", f"{site_count} 개 거점")
        with m4: st.metric("👑 최대 공급 모델", f"{top_model}", f"{format_k(top_model_qty)}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        with f1:
            all_models = sorted(df_fcst['모델명'].unique().tolist())
            selected_models = st.multiselect("🔍 모델명 필터 (다중 선택 가능)", options=all_models, default=[])
        with f2:
            all_sites = sorted(df_fcst['SITE'].unique().tolist())
            selected_sites = st.multiselect("🏢 출하 SITE 필터", options=all_sites, default=[])
        with f3:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            hide_zero_cols = st.checkbox("공급 물량이 있는 날짜만 표시", value=True)

        df_filtered = df_fcst.copy()
        if selected_models:
            df_filtered = df_filtered[df_filtered['모델명'].isin(selected_models)]
        if selected_sites:
            df_filtered = df_filtered[df_filtered['SITE'].isin(selected_sites)]

        if hide_zero_cols:
            valid_dates = [c for c in date_cols if c in df_filtered.columns and df_filtered[c].sum() > 0]
        else:
            valid_dates = [c for c in date_cols if c in df_filtered.columns]

        tab_by_model, tab_by_detail, tab_charts = st.tabs([
            "📊 모델별 공급 일정 (통합)", 
            "📋 품목 / SITE별 상세 공급 일정", 
            "📈 공급 계획 추이 차트"
        ])

        with tab_by_model:
            st.markdown("##### 📌 모델별 공급 계획 일정 (일자별 합계, 단위: K)")
            if not df_filtered.empty and valid_dates:
                df_model_pivot = df_filtered.groupby('모델명')[valid_dates].sum()
                df_model_pivot['총 공급계획(개)'] = df_model_pivot.sum(axis=1)
                df_model_pivot = df_model_pivot.sort_values(by='총 공급계획(개)', ascending=False).reset_index()
                
                df_model_disp = df_model_pivot.copy()
                df_model_disp['총 공급계획'] = df_model_disp['총 공급계획(개)'].apply(format_k)
                for col_d in valid_dates:
                    df_model_disp[col_d] = df_model_disp[col_d].apply(format_k)
                
                cols_order = ['모델명', '총 공급계획'] + valid_dates
                df_model_disp = df_model_disp[cols_order]
                
                c_btn1, c_btn2 = st.columns([0.8, 0.2])
                with c_btn2:
                    st.download_button(
                        label="📥 모델별 일정 CSV 다운로드",
                        data=convert_df(df_model_pivot),
                        file_name='fcst_supply_by_model.csv',
                        mime='text/csv'
                    )
                st.dataframe(df_model_disp, width="stretch", hide_index=True)
            else:
                st.warning("선택한 조건에 해당하는 공급 일정이 없습니다.")

        with tab_by_detail:
            st.markdown("##### 📌 품목코드(GH-CODE) 및 출하 SITE별 상세 공급 일정 (단위: K)")
            if not df_filtered.empty and valid_dates:
                detail_cols = ['모델명', 'GH_CODE', 'SITE', '합계'] + valid_dates
                df_detail = df_filtered[detail_cols].sort_values(by='합계', ascending=False)
                
                df_detail_disp = df_detail.copy()
                df_detail_disp['합계'] = df_detail_disp['합계'].apply(format_k)
                for col_d in valid_dates:
                    df_detail_disp[col_d] = df_detail_disp[col_d].apply(format_k)
                
                c_btn1, c_btn2 = st.columns([0.8, 0.2])
                with c_btn2:
                    st.download_button(
                        label="📥 상세 일정 CSV 다운로드",
                        data=convert_df(df_detail),
                        file_name='fcst_supply_detailed.csv',
                        mime='text/csv'
                    )
                st.dataframe(df_detail_disp, width="stretch", hide_index=True)
            else:
                st.warning("선택한 조건에 해당하는 공급 일정이 없습니다.")

        with tab_charts:
            st.markdown("##### 📌 모델별 공급 물량 및 일자별 공급 추이")
            ch_col1, ch_col2 = st.columns(2)
            
            with ch_col1:
                top10_models = df_filtered.groupby('모델명')['합계'].sum().sort_values(ascending=False).head(10)
                fig_top10 = go.Figure()
                fig_top10.add_trace(go.Bar(
                    x=top10_models.index.tolist(),
                    y=top10_models.values.tolist(),
                    marker_color='#0044ff',
                    marker_line=dict(width=1.5, color='#111111'),
                    text=[format_k(v) for v in top10_models.values],
                    textposition='outside'
                ))
                fig_top10.update_layout(
                    title="<b>TOP 10 모델별 공급 계획 물량 (단위: K)</b>",
                    plot_bgcolor='rgba(0,0,0,0)',
                    hovermode='x unified',
                    margin=dict(l=20, r=20, t=45, b=20)
                )
                st.plotly_chart(fig_top10, width="stretch")
                
            with ch_col2:
                daily_trend = df_filtered[valid_dates].sum()
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Bar(
                    x=daily_trend.index.tolist(),
                    y=daily_trend.values.tolist(),
                    marker_color='#10b981',
                    marker_line=dict(width=1.5, color='#111111'),
                    text=[format_k(v) if v > 0 else "" for v in daily_trend.values],
                    textposition='outside'
                ))
                fig_trend.update_layout(
                    title="<b>일자별 전체 공급 계획 추이 (단위: K)</b>",
                    plot_bgcolor='rgba(0,0,0,0)',
                    hovermode='x unified',
                    margin=dict(l=20, r=20, t=45, b=20)
                )
                st.plotly_chart(fig_trend, width="stretch")

elif menu == "5. 완제품 재고 조회":
    st.subheader("🏢 5. 완제품 재고 조회")
    st.metric("📦 당일 완제품 총 재고", f"{format_k(st.session_state['inventory_data']['qty'])}", f"재고일수: {st.session_state['inventory_data']['days']} 일", delta_color="off")

elif menu == "6. 캘린더":
    st.title("🗓️ 6. 월간 FCST 공급 캘린더")
    fcst_data = st.session_state.get('fcst_data', {'rows': [], 'date_cols': []})
    fcst_rows = fcst_data.get('rows', [])
    actual_by_day = st.session_state.get('actual_by_day', {})
    
    if not fcst_rows:
        st.info("💡 출하 계획 파일(RTF Main 시트)을 **[7. 데이터 관리 (Admin)]** 메뉴에서 업로드 및 분석해 주세요.")
    else:
        t_m, t_d = parse_target_day_info(st.session_state.get('target_date', default_target_date))

        # 1. 상단 월 선택 바
        c_m1, c_m2 = st.columns([1.2, 2.8])
        with c_m1:
            month_options = [10, 11, 12, 1]
            month_labels = {10: "2026년 10월", 11: "2026년 11월 (당월)", 12: "2026년 12월", 1: "2027년 1월"}
            default_idx = month_options.index(t_m) if t_m in month_options else 1
            selected_m = st.selectbox("📅 조회 대상 월 선택", options=month_options, format_func=lambda x: month_labels.get(x, f"{x}월"), index=default_idx)
            
        cur_year = 2027 if selected_m == 1 else 2026

        cal_instance = calendar.Calendar(firstweekday=6) # 일요일 시작 (SUN)
        month_matrix = cal_instance.monthdayscalendar(cur_year, selected_m)
        num_days = calendar.monthrange(cur_year, selected_m)[1]

        day_supply_map = {d: {'total': 0, 'total_rev': 0.0, 'items': []} for d in range(1, num_days + 1)}

        for r_item in fcst_rows:
            m_name = r_item['모델명']
            gh = r_item['GH_CODE']
            site = r_item['SITE']
            price = r_item.get('단가_원', 0.0)

            for d in range(1, num_days + 1):
                d_key = f"{selected_m:02d}/{d:02d}"
                d_key_alt = f"{selected_m}/{d}"
                q = r_item.get(d_key, r_item.get(d_key_alt, 0))
                if q > 0:
                    rev_val = (q * price) / 100_000_000.0 if price > 0 else 0.0
                    day_supply_map[d]['total'] += q
                    day_supply_map[d]['total_rev'] += rev_val
                    day_supply_map[d]['items'].append({
                        'model': m_name,
                        'gh_code': gh,
                        'site': site,
                        'qty': q,
                        'price': price,
                        'rev_ok': rev_val
                    })

        # 2. 당월 요약 지표 카드
        m_total_qty = sum(day_supply_map[d]['total'] for d in day_supply_map)
        m_total_rev = sum(day_supply_map[d]['total_rev'] for d in day_supply_map)
        supply_days_cnt = sum(1 for d in day_supply_map if day_supply_map[d]['total'] > 0)
        
        m_model_totals = {}
        for d in day_supply_map:
            for it in day_supply_map[d]['items']:
                m_model_totals[it['model']] = m_model_totals.get(it['model'], 0) + it['qty']
                
        top_m_name = max(m_model_totals, key=m_model_totals.get) if m_model_totals else "-"
        top_m_qty = m_model_totals.get(top_m_name, 0)

        with c_m2:
            render_html(f"""
            <div style="background-color: #ffffff; border: 2.5px solid #111111; padding: 10px 16px; box-shadow: 4px 4px 0px #0044ff; display: flex; justify-content: space-around; align-items: center; border-radius: 4px;">
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 800; color: #666;">당월 총 공급계획</div>
                    <div style="font-size: 18px; font-weight: 900; color: #111111;">{format_k(m_total_qty)}</div>
                </div>
                <div style="border-left: 1.5px solid #e5e7eb; height: 32px;"></div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 800; color: #0044ff;">당월 총 예상 매출</div>
                    <div style="font-size: 18px; font-weight: 900; color: #0044ff;">{m_total_rev:,.2f} 억원</div>
                </div>
                <div style="border-left: 1.5px solid #e5e7eb; height: 32px;"></div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 800; color: #666;">공급 예정 일수</div>
                    <div style="font-size: 18px; font-weight: 900; color: #111111;">{supply_days_cnt} 일 / {num_days}일</div>
                </div>
                <div style="border-left: 1.5px solid #e5e7eb; height: 32px;"></div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 800; color: #666;">최대 공급 모델</div>
                    <div style="font-size: 15px; font-weight: 900; color: #10b981;">{top_m_name} ({format_k(top_m_qty)})</div>
                </div>
            </div>
            """)

        st.markdown("<br>", unsafe_allow_html=True)

        # 팝업 대화상자 미지원 환경 대비 Fallback 렌더링
        if st.session_state.get('active_cal_popup'):
            p_day, p_month, p_year = st.session_state['active_cal_popup']
            p_info = day_supply_map.get(p_day, {'total': 0, 'total_rev': 0.0, 'items': []})
            p_items = p_info['items']
            p_unique_models = sorted(list(set(it['model'] for it in p_items)))
            
            with st.container():
                render_html(f"""
                <div style="background-color: #ffffff; border: 3px solid #0044ff; padding: 18px; box-shadow: 6px 6px 0px #111111; margin-bottom: 25px; border-radius: 6px;">
                    <h3 style="margin: 0 0 14px 0; color: #111111; border-bottom: 2px solid #111111; padding-bottom: 8px;">
                        📅 {p_year}년 {p_month}월 {p_day}일 공급 상세 내역 (전체 모델)
                    </h3>
                    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 14px;">
                        <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 8px; border-radius: 4px; text-align: center;">
                            <div style="font-size: 11px; font-weight: 800; color: #6b7280;">📱 전체 모델 수</div>
                            <div style="font-size: 18px; font-weight: 900; color: #111111;">{len(p_unique_models)}종</div>
                        </div>
                        <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 8px; border-radius: 4px; text-align: center;">
                            <div style="font-size: 11px; font-weight: 800; color: #6b7280;">📦 공급 품목(GH)</div>
                            <div style="font-size: 18px; font-weight: 900; color: #111111;">{len(p_items)}건</div>
                        </div>
                        <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 8px; border-radius: 4px; text-align: center;">
                            <div style="font-size: 11px; font-weight: 800; color: #6b7280;">🚚 총 공급 수량</div>
                            <div style="font-size: 18px; font-weight: 900; color: #0044ff;">{format_k(p_info['total'])}</div>
                        </div>
                        <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 8px; border-radius: 4px; text-align: center;">
                            <div style="font-size: 11px; font-weight: 800; color: #6b7280;">💰 총 예상 매출액</div>
                            <div style="font-size: 18px; font-weight: 900; color: #10b981;">{p_info['total_rev']:,.2f}억원</div>
                        </div>
                    </div>
                </div>
                """)
                
                if p_items:
                    df_popup = pd.DataFrame(p_items)
                    df_popup['공급수량'] = df_popup['qty'].apply(format_k)
                    df_popup['단가(원)'] = df_popup['price'].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
                    df_popup['예상매출(억원)'] = df_popup['rev_ok'].apply(lambda x: f"{x:,.2f}억원")
                    df_popup = df_popup.rename(columns={'model': '모델명', 'gh_code': 'GH-CODE', 'site': '출하 SITE'})
                    df_popup = df_popup.sort_values(by='qty', ascending=False)
                    st.dataframe(df_popup[['모델명', 'GH-CODE', '출하 SITE', '공급수량', '단가(원)', '예상매출(억원)']], width="stretch", hide_index=True)
                
                btn_close, _ = st.columns([0.2, 0.8])
                with btn_close:
                    if st.button("✖️ 팝업 닫기", key="btn_close_fallback_popup", use_container_width=True):
                        st.session_state['active_cal_popup'] = None
                        st.rerun()

        # 3. 달력 요일 헤더 렌더링
        days_headers = [
            ("<span style='color: #ef4444;'>일 (SUN)</span>", "#ef4444"),
            ("월 (MON)", "#111111"),
            ("화 (TUE)", "#111111"),
            ("수 (WED)", "#111111"),
            ("목 (THU)", "#111111"),
            ("금 (FRI)", "#111111"),
            ("<span style='color: #0044ff;'>토 (SAT)</span>", "#0044ff")
        ]
        
        h_cols = st.columns(7)
        for i, (h_title, text_color) in enumerate(days_headers):
            with h_cols[i]:
                st.markdown(clean_html(f"""
                <div style="background-color: #ffffff; border: 2px solid #111111; padding: 6px; text-align: center; font-weight: 900; font-size: 13px; box-shadow: 2px 2px 0px #111111; margin-bottom: 6px; border-radius: 4px;">
                    {h_title}
                </div>
                """), unsafe_allow_html=True)

        # 4. 달력 날짜 그리드 셀 렌더링 (호버 시 확장 팝업 + 모델명 칩 숨김 + 실적/계획 차이 표시)
        for week in month_matrix:
            w_cols = st.columns(7)
            for day_idx, d_num in enumerate(week):
                with w_cols[day_idx]:
                    if d_num == 0:
                        st.markdown(clean_html("<div style='height: 110px; background: rgba(0,0,0,0.02); border: 1px dashed #d1d5db; border-radius: 4px; margin-bottom: 6px;'></div>"), unsafe_allow_html=True)
                    else:
                        d_info = day_supply_map[d_num]
                        plan_q = d_info['total']
                        plan_rev = d_info['total_rev']
                        
                        # 기준일자 이전 여부 판별
                        is_target = (selected_m == t_m and d_num == t_d)
                        is_past = (selected_m < t_m) or (selected_m == t_m and d_num <= t_d)
                        
                        # 일자별 실적 데이터 조회
                        act_day_data = actual_by_day.get(d_num, {'qty': 0, 'rev': 0.0, 'items': []})
                        act_q = act_day_data['qty']
                        act_rev = act_day_data['rev']
                        
                        # 카드 스타일 지정
                        if is_target:
                            border_css = "3px solid #0044ff; box-shadow: 4px 4px 0px #0044ff;"
                            bg_css = "#f0f5ff"
                            badge_target = "<span style='background: #0044ff; color: #ffffff; font-size: 10px; font-weight: 800; padding: 1px 4px; border-radius: 3px; margin-left: 4px;'>기준일</span>"
                        else:
                            border_css = "2px solid #111111; box-shadow: 3px 3px 0px #111111;"
                            bg_css = "#ffffff" if (plan_q > 0 or act_q > 0) else "#fafafa"
                            badge_target = ""

                        if day_idx == 0: day_color = "#ef4444"
                        elif day_idx == 6: day_color = "#0044ff"
                        else: day_color = "#111111"

                        # [핵심] 일별 보드 표시: 기준일자 이전 vs 미래 일자 분기 (수량/매출액만 깔끔하게 노출)
                        if is_past:
                            diff_q = act_q - plan_q
                            diff_rev = act_rev - plan_rev
                            
                            if diff_q > 0:
                                diff_q_tag = f"<span style='color: #0044ff; font-weight: 900;'>+{format_k(diff_q)}</span>"
                            elif diff_q < 0:
                                diff_q_tag = f"<span style='color: #ef4444; font-weight: 900;'>-{format_k(abs(diff_q))}</span>"
                            else:
                                diff_q_tag = "<span style='color: #6b7280; font-weight: 800;'>0K</span>"

                            if diff_rev > 0:
                                diff_rev_tag = f"<span style='color: #0044ff; font-weight: 900;'>+{diff_rev:,.2f}억</span>"
                            elif diff_rev < 0:
                                diff_rev_tag = f"<span style='color: #ef4444; font-weight: 900;'>-{abs(diff_rev):,.2f}억</span>"
                            else:
                                diff_rev_tag = "<span style='color: #6b7280; font-weight: 800;'>0.00억</span>"

                            board_content_html = f"""
                            <div style="display: flex; flex-direction: column; gap: 4px; margin-top: 4px;">
                                <div style="font-size: 11px; font-weight: 800; color: #111111; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #4b5563;">실적 수량:</span>
                                    <span><b>{format_k(act_q)}</b> ({diff_q_tag})</span>
                                </div>
                                <div style="font-size: 11px; font-weight: 800; color: #111111; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #4b5563;">실적 매출:</span>
                                    <span><b>{act_rev:,.2f}억</b> ({diff_rev_tag})</span>
                                </div>
                            </div>
                            """
                        else:
                            board_content_html = f"""
                            <div style="display: flex; flex-direction: column; gap: 4px; margin-top: 4px;">
                                <div style="font-size: 11px; font-weight: 800; color: #111111; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #4b5563;">계획 수량:</span>
                                    <span style="color: #111111;"><b>{format_k(plan_q)}</b></span>
                                </div>
                                <div style="font-size: 11px; font-weight: 800; color: #111111; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #4b5563;">계획 매출:</span>
                                    <span style="color: #0044ff;"><b>{plan_rev:,.2f}억원</b></span>
                                </div>
                            </div>
                            """

                        # [핵심] 호버(Hover) 시 확장되는 팝업창 HTML 구축
                        unique_models_cnt = len(set(it['model'] for it in d_info['items']))
                        items_sorted = sorted(d_info['items'], key=lambda x: x['qty'], reverse=True)
                        
                        hover_list_html = ""
                        for it in items_sorted[:6]:
                            q_txt = format_k(it['qty'])
                            hover_list_html += f"<div style='font-size: 11px; color: #111827; padding: 2px 0; border-bottom: 1px dotted #e5e7eb; display: flex; justify-content: space-between;'><span>• {it['model']} ({it['site']})</span><span><b>{q_txt}</b> / {it['rev_ok']:,.2f}억</span></div>"
                        if len(items_sorted) > 6:
                            hover_list_html += f"<div style='font-size: 10px; color: #6b7280; font-weight: 800; text-align: center; margin-top: 4px;'>+ 외 {len(items_sorted)-6}개 품목 더 있음</div>"

                        if day_idx <= 1: pos_class = "pos-left"
                        elif day_idx >= 5: pos_class = "pos-right"
                        else: pos_class = "pos-center"

                        hover_popup_html = f"""
                        <div class="cal-hover-popup {pos_class}">
                            <div style="font-size: 13px; font-weight: 900; color: #111111; border-bottom: 2px solid #111111; padding-bottom: 4px; margin-bottom: 8px;">
                                📅 {selected_m}월 {d_num}일 공급 스케줄 (전체 {unique_models_cnt}종)
                            </div>
                            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px; margin-bottom: 8px;">
                                <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 5px; border-radius: 3px; text-align: center;">
                                    <div style="font-size: 10px; font-weight: 800; color: #6b7280;">총 계획수량</div>
                                    <div style="font-size: 14px; font-weight: 900; color: #0044ff;">{format_k(plan_q)}</div>
                                </div>
                                <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 5px; border-radius: 3px; text-align: center;">
                                    <div style="font-size: 10px; font-weight: 800; color: #6b7280;">총 예상매출</div>
                                    <div style="font-size: 14px; font-weight: 900; color: #10b981;">{plan_rev:,.2f}억원</div>
                                </div>
                            </div>
                            <div style="max-height: 140px; overflow-y: auto;">
                                {hover_list_html if hover_list_html else "<div style='font-size: 11px; color: #9ca3af; text-align: center;'>공급 일정 없음</div>"}
                            </div>
                        </div>
                        """

                        full_cell_html = f"""
                        <div class="cal-card-wrapper">
                            <div class="cal-card" style="background: {bg_css}; border: {border_css};">
                                <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1.5px solid #e5e7eb; padding-bottom: 3px; margin-bottom: 3px;">
                                    <span style="font-weight: 900; font-size: 15px; color: {day_color};">{d_num}{badge_target}</span>
                                    <span style="font-size: 10px; font-weight: 800; color: #6b7280;">계획 {format_k(plan_q)}</span>
                                </div>
                                {board_content_html}
                            </div>
                            {hover_popup_html}
                        </div>
                        """
                        st.markdown(clean_html(full_cell_html), unsafe_allow_html=True)

                        btn_label = f"🔍 상세 ({unique_models_cnt}종)" if (plan_q > 0 or act_q > 0) else f"🔍 {d_num}일"
                        if st.button(btn_label, key=f"btn_pop_{selected_m}_{d_num}", use_container_width=True):
                            if hasattr(st, "dialog"):
                                show_day_detail_dialog(d_num, selected_m, cur_year, d_info)
                            else:
                                st.session_state['active_cal_popup'] = (d_num, selected_m, cur_year)
                                st.rerun()

        st.markdown("<br><hr style='border: 1px solid #111111;'>", unsafe_allow_html=True)

        # 5. 하단 일자별 공급 추이 차트
        st.markdown(f"### 📈 {selected_m}월 일자별 공급 수량 및 예상 매출 추이")
        trend_days = [f"{selected_m}/{d:02d}" for d in range(1, num_days + 1)]
        trend_qtys = [day_supply_map[d]['total'] for d in range(1, num_days + 1)]
        trend_revs = [round(day_supply_map[d]['total_rev'], 2) for d in range(1, num_days + 1)]
        
        fig_cal_trend = go.Figure()
        fig_cal_trend.add_trace(go.Bar(
            x=trend_days,
            y=trend_qtys,
            name="공급수량(K)",
            marker_color=['#0044ff' if (d == t_d and selected_m == t_m) else '#d1d5db' for d in range(1, num_days + 1)],
            marker_line=dict(width=1.5, color='#111111'),
            text=[format_k(q) if q > 0 else "" for q in trend_qtys],
            textposition='outside'
        ))
        fig_cal_trend.add_trace(go.Scatter(
            x=trend_days,
            y=trend_revs,
            name="예상매출(억원)",
            yaxis="y2",
            mode="lines+markers",
            line=dict(color="#ef4444", width=2.5),
            marker=dict(size=6, color="#ef4444")
        ))
        fig_cal_trend.update_layout(
            title=f"<b>{selected_m}월 일자별 공급 계획 수량 (막대) 및 예상 매출액 (꺾은선)</b>",
            plot_bgcolor='rgba(0,0,0,0)',
            hovermode='x unified',
            margin=dict(l=20, r=20, t=45, b=20),
            xaxis=dict(tickangle=-45),
            yaxis=dict(title="공급수량 (개)"),
            yaxis2=dict(title="예상매출 (억원)", overlaying="y", side="right"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_cal_trend, use_container_width=True)
