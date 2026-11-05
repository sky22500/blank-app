import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl
import datetime
import calendar
import re
import json
import gspread
from google.oauth2.service_account import Credentials

# 1. 페이지 설정
st.set_page_config(page_title="마케팅 부문 실적 대시보드", layout="wide", initial_sidebar_state="expanded")

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
default_target_date = f"{yesterday.month}/{yesterday.day}"

def render_html(html_str):
    if hasattr(st, "html"):
        st.html(html_str)
    else:
        st.markdown(html_str, unsafe_allow_html=True)

# 구글 스프레드시트 연동 클라이언트 및 저장소 설정
@st.cache_resource
def get_gspread_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    except Exception:
        return None

def get_db_worksheet():
    gc = get_gspread_client()
    if not gc:
        return None
    sheet_url = st.secrets.get("sheet", {}).get("url")
    if not sheet_url:
        return None
    try:
        sh = gc.open_by_url(sheet_url)
        try:
            ws = sh.worksheet("DB_STATE")
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title="DB_STATE", rows=100, cols=20)
        return ws
    except Exception:
        return None

@st.cache_data(ttl=60)
def fetch_db_state_cached():
    ws = get_db_worksheet()
    if not ws:
        return None
    try:
        data = ws.get_all_values()
        if not data:
            return None
        state_dict = {}
        for row in data:
            if not row or not row[0]:
                continue
            k = row[0]
            full_json = "".join(row[1:])
            if full_json:
                try:
                    state_dict[k] = json.loads(full_json)
                except Exception:
                    pass
        return state_dict
    except Exception:
        return None

def save_state_to_gsheet(state_dict):
    ws = get_db_worksheet()
    if not ws:
        return False
    try:
        rows = []
        for k, v in state_dict.items():
            s = json.dumps(v, ensure_ascii=False)
            chunks = [s[i:i+40000] for i in range(0, len(s), 40000)]
            rows.append([k] + chunks)
        ws.clear()
        ws.update(rows)
        fetch_db_state_cached.clear()
        return True
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")
        return False

# 세션 상태 기본값 초기화
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

if 'data_loaded' not in st.session_state: 
    st.session_state['data_loaded'] = False

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

if not st.session_state.get('data_loaded'):
    saved_state = fetch_db_state_cached()
    if saved_state:
        for k, v in saved_state.items():
            st.session_state[k] = v

# 유틸 함수
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

# [최상위 하이엔드] 핀테크 / 프리미엄 SaaS 스타일 KPI 카드
def render_kpi_card(title, icon, actual_val, plan_val, is_rev=False, sub_caption=None):
    if is_rev:
        act_str = f"{actual_val:,.2f}억"
        plan_str = f"{plan_val:,.2f}억"
        diff = actual_val - plan_val
        if diff > 0:
            diff_str = f"+{diff:,.2f}억"
            diff_color = "#10b981"
            bg_accent = "linear-gradient(90deg, #10b981, #059669)"
        elif diff < 0:
            diff_str = f"-{abs(diff):,.2f}억"
            diff_color = "#f43f5e"
            bg_accent = "linear-gradient(90deg, #f43f5e, #e11d48)"
        else:
            diff_str = "0.00억"
            diff_color = "#94a3b8"
            bg_accent = "linear-gradient(90deg, #cbd5e1, #94a3b8)"
    else:
        act_str = format_k(actual_val)
        plan_str = format_k(plan_val)
        diff = actual_val - plan_val
        if diff > 0:
            diff_str = f"+{format_k(diff)}"
            diff_color = "#4f46e5"
            bg_accent = "linear-gradient(90deg, #6366f1, #4f46e5)"
        elif diff < 0:
            diff_str = f"-{format_k(abs(diff))}"
            diff_color = "#f43f5e"
            bg_accent = "linear-gradient(90deg, #f43f5e, #e11d48)"
        else:
            diff_str = "0K"
            diff_color = "#94a3b8"
            bg_accent = "linear-gradient(90deg, #cbd5e1, #94a3b8)"

    sub_html = f"""
    <div style="font-size: 11.5px; color: #64748b; font-weight: 500; margin-top: 14px; padding-top: 10px; border-top: 1px dashed #e2e8f0; display: flex; align-items: center; gap: 6px;">
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color: #94a3b8;"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
        {sub_caption}
    </div>
    """ if sub_caption else ""

    html = f"""
    <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; padding: 22px 24px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03); margin-bottom: 12px; position: relative; overflow: hidden; transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);">
        <div style="position: absolute; top: 0; left: 0; width: 100%; height: 4px; background: {bg_accent};"></div>
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px;">
            <div style="display: flex; align-items: center; gap: 10px;">
                <div style="width: 32px; height: 32px; border-radius: 10px; background: #f8fafc; border: 1px solid #e2e8f0; display: flex; align-items: center; justify-content: center; font-size: 16px; box-shadow: 0 1px 2px rgba(0,0,0,0.02);">
                    {icon}
                </div>
                <span style="font-size: 14px; font-weight: 800; color: #0f172a; letter-spacing: -0.5px;">{title}</span>
            </div>
            <div style="background: rgba(255,255,255,0); color: {diff_color}; font-size: 13px; font-weight: 800; letter-spacing: -0.3px; display: flex; align-items: center; gap: 4px;">
                {diff_str}
            </div>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: baseline;">
            <div>
                <div style="font-size: 11px; font-weight: 700; color: #94a3b8; letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 4px;">Actual</div>
                <div style="font-size: 32px; font-weight: 900; color: #0f172a; line-height: 1; letter-spacing: -1px;">{act_str}</div>
            </div>
            <div style="text-align: right;">
                <div style="font-size: 11px; font-weight: 700; color: #94a3b8; letter-spacing: 0.5px; text-transform: uppercase; margin-bottom: 4px;">Target</div>
                <div style="font-size: 20px; font-weight: 700; color: #64748b; line-height: 1; letter-spacing: -0.5px;">{plan_str}</div>
            </div>
        </div>
        {sub_html}
    </div>
    """
    render_html(html)

def render_achievement_gauge(title, actual, plan, unit="K"):
    rate = (actual / plan * 100.0) if plan > 0 else 0.0
    if rate >= 100:
        bar_gradient = "linear-gradient(90deg, #4f46e5, #6366f1)"
        badge_color = "#4f46e5"
        shadow_color = "rgba(79, 70, 229, 0.3)"
    elif rate >= 80:
        bar_gradient = "linear-gradient(90deg, #059669, #10b981)"
        badge_color = "#059669"
        shadow_color = "rgba(5, 150, 105, 0.3)"
    else:
        bar_gradient = "linear-gradient(90deg, #d97706, #f59e0b)"
        badge_color = "#d97706"
        shadow_color = "rgba(217, 119, 6, 0.3)"

    bar_width = min(max(rate, 0.0), 100.0)

    html = f"""
    <div style="background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 14px 18px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02); margin-top: -6px; margin-bottom: 24px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <span style="font-size: 12px; font-weight: 700; color: #64748b; letter-spacing: -0.3px;">{title} 목표 달성률</span>
            <span style="font-size: 14px; font-weight: 900; color: {badge_color}; letter-spacing: -0.5px;">{rate:,.1f}%</span>
        </div>
        <div style="background-color: #f1f5f9; border-radius: 999px; height: 6px; width: 100%; box-shadow: inset 0 1px 2px rgba(0,0,0,0.05);">
            <div style="background: {bar_gradient}; width: {bar_width:,.1f}%; height: 100%; border-radius: 999px; box-shadow: 0 0 8px {shadow_color}; transition: width 1s ease-in-out;"></div>
        </div>
    </div>
    """
    render_html(html)

if hasattr(st, "dialog"):
    @st.dialog("📊 일자별 공급 상세 스펙", width="large")
    def show_day_detail_dialog(d_num, m_num, y_num, d_info):
        items = d_info['items']
        tot_q = d_info['total']
        tot_rev = d_info['total_rev']
        unique_models = sorted(list(set(it['model'] for it in items)))
        
        st.markdown(f"### 🗓️ {y_num}. {m_num:02d}. {d_num:02d} 공급 데이터")
        
        cards_html = f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px;">
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 16px; border-radius: 12px; text-align: center;">
                <div style="font-size: 11px; font-weight: 700; color: #64748b; letter-spacing: -0.3px;">대상 모델</div>
                <div style="font-size: 22px; font-weight: 900; color: #0f172a; margin-top: 4px;">{len(unique_models)}종</div>
            </div>
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 16px; border-radius: 12px; text-align: center;">
                <div style="font-size: 11px; font-weight: 700; color: #64748b; letter-spacing: -0.3px;">공급 품목</div>
                <div style="font-size: 22px; font-weight: 900; color: #0f172a; margin-top: 4px;">{len(items)}건</div>
            </div>
            <div style="background: #eff6ff; border: 1px solid #bfdbfe; padding: 16px; border-radius: 12px; text-align: center;">
                <div style="font-size: 11px; font-weight: 700; color: #3b82f6; letter-spacing: -0.3px;">총 수량</div>
                <div style="font-size: 22px; font-weight: 900; color: #1d4ed8; margin-top: 4px;">{format_k(tot_q)}</div>
            </div>
            <div style="background: #ecfdf5; border: 1px solid #a7f3d0; padding: 16px; border-radius: 12px; text-align: center;">
                <div style="font-size: 11px; font-weight: 700; color: #10b981; letter-spacing: -0.3px;">예상 매출</div>
                <div style="font-size: 22px; font-weight: 900; color: #047857; margin-top: 4px;">{tot_rev:,.2f}억</div>
            </div>
        </div>
        """
        render_html(cards_html)
        
        if items:
            df_pop = pd.DataFrame(items)
            df_pop['공급수량'] = df_pop['qty'].apply(format_k)
            df_pop['단가(원)'] = df_pop['price'].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
            df_pop['예상매출(억원)'] = df_pop['rev_ok'].apply(lambda x: f"{x:,.2f}억원")
            df_pop = df_pop.rename(columns={'model': '모델명', 'gh_code': 'GH-CODE', 'site': '출하 SITE'})
            df_pop = df_pop.sort_values(by='qty', ascending=False)
            
            st.dataframe(
                df_pop[['모델명', 'GH-CODE', '출하 SITE', '공급수량', '단가(원)', '예상매출(억원)']], 
                use_container_width=True, 
                hide_index=True
            )
            
            st.download_button(
                label=f"📥 {m_num}/{d_num} 데이터 다운로드 (CSV)",
                data=convert_df(df_pop[['모델명', 'GH-CODE', '출하 SITE', '공급수량', '단가(원)', '예상매출(억원)']]),
                file_name=f"fcst_supply_{m_num}M_{d_num}D_detail.csv",
                mime="text/csv",
                key=f"dl_popup_{m_num}_{d_num}",
                use_container_width=True
            )
        else:
            st.info("해당 일자에 예정된 공급 일정이 없습니다.")

# (파싱 엔진 3-1 ~ 3-4 기존 동일 유지)
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
    col_map, weekly_cols = {}, {}
    w_start_col, w_first_label = None, 'W45'
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
    labels, wp, wa, seen_labels = [], [], [], set()
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
    learned_prices, learned_suffix_prices, learned_model_prices = {}, {}, {}
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

def parse_rtf_main_plan_excel(file_bytes, target_date_str, mapping_data, exchange_rate=1350.0):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    sheet = wb['Main'] if 'Main' in wb.sheetnames else wb.active
    t_m, t_d = parse_target_day_info(target_date_str)
    target_date_obj = datetime.date(2026, t_m, t_d)
    target_iso_week = f"W{target_date_obj.isocalendar()[1]}"
    target_month_lbl = f"{t_m}월"
    date_cols, fcst_date_cols = [], []
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
    site_col, gh_col, gubun_col = 3, 4, 9
    gh_map = mapping_data.get('gh_map', {})
    suffix_map = mapping_data.get('suffix_map', {})
    gh_to_model = mapping_data.get('gh_to_model', {})
    model_map = mapping_data.get('model_map', {})
    known_prices = [p for p in list(gh_map.values()) + list(suffix_map.values()) + list(model_map.values()) if p > 0]
    avg_price = (sum(known_prices) / len(known_prices)) if known_prices else 8800.0
    daily_ship_p, daily_rev_p = 0, 0.0
    weekly_ship_plan, weekly_rev_plan, monthly_ship_plan, monthly_rev_plan = {}, {}, {}, {}
    plan_by_model, plan_by_site, fcst_rows = {}, {}, []
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
        if price == 0: price = avg_price
        price_krw = price if price > 100 else price * exchange_rate
        row_fcst = {'SITE': site, 'GH_CODE': gh_code, '모델명': model_name, '단가_원': price_krw}
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
        if row_fcst_total > 0: fcst_rows.append(row_fcst)
    return {
        'daily_ship_p': daily_ship_p, 'daily_rev_p': round(daily_rev_p, 2),
        'cur_week_ship_p': weekly_ship_plan.get(target_iso_week, 0),
        'cur_week_rev_p': round(weekly_rev_plan.get(target_iso_week, 0.0), 2),
        'cur_month_ship_p': monthly_ship_plan.get(target_month_lbl, 0),
        'cur_month_rev_p': round(monthly_rev_plan.get(target_month_lbl, 0.0), 2),
        'weekly_ship_plan': weekly_ship_plan, 'weekly_rev_plan': {k: round(v, 2) for k, v in weekly_rev_plan.items()},
        'monthly_ship_plan': monthly_ship_plan, 'monthly_rev_plan': {k: round(v, 2) for k, v in monthly_rev_plan.items()},
        'plan_by_model': plan_by_model, 'plan_by_site': plan_by_site,
        'fcst_rows': fcst_rows, 'fcst_date_cols': fcst_date_cols
    }

# 4. [프리미엄 SaaS] 초고도화 CSS
render_html("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    
    html, body, [class*="css"] {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif !important;
        letter-spacing: -0.2px;
    }
    
    /* 앱 배경: 매우 은은한 오프화이트/회백색 */
    .stApp { 
        background-color: #f8fafc; 
        color: #0f172a; 
    }
    
    /* 사이드바는 기존 코드 그대로 유지 (터치 X) */
    section[data-testid="stSidebar"] { 
        background-color: #0f172a; 
        border-right: 1px solid #1e293b; 
    }
    section[data-testid="stSidebar"] * { 
        color: #cbd5e1 !important; 
    }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {
        color: #ffffff !important;
    }
    
    /* 탭 디자인: iOS 세그먼트 컨트롤 스타일 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #e2e8f0;
        padding: 4px;
        border-radius: 12px;
        border: none;
        margin-bottom: 24px;
    }
    .stTabs [data-baseweb="tab"] { 
        background-color: transparent; 
        border: none; 
        border-radius: 8px; 
        padding: 10px 24px; 
        font-weight: 700; 
        font-size: 14.5px;
        color: #64748b; 
        box-shadow: none;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .stTabs [aria-selected="true"] { 
        background-color: #ffffff !important; 
        color: #0f172a !important; 
        box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
    }
    
    /* 날짜 배지: 투명도 글래스 뱃지 */
    .date-badge {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.5);
        padding: 8px 18px;
        display: inline-flex;
        align-items: center;
        gap: 8px;
        font-weight: 700;
        font-size: 14px;
        color: #475569;
        margin-bottom: 24px;
        border-radius: 30px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    /* 캘린더 인터랙티브 효과 */
    .cal-card-wrapper {
        position: relative;
        margin-bottom: 10px;
    }
    .cal-card {
        min-height: 110px;
        border-radius: 12px;
        padding: 12px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        cursor: pointer;
    }
    .cal-card-wrapper:hover .cal-card {
        transform: translateY(-4px) scale(1.01);
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.08), 0 10px 10px -5px rgba(0, 0, 0, 0.04) !important;
        border-color: #cbd5e1 !important;
        z-index: 10;
    }
    
    /* 애플 스타일 호버 팝업 (블러 & 섀도우) */
    .cal-hover-popup {
        display: none;
        position: absolute;
        top: calc(100% + 5px);
        width: 340px;
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(0,0,0,0.08);
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 20px 40px -8px rgba(0, 0, 0, 0.15);
        z-index: 999999;
        pointer-events: none;
    }
    .cal-card-wrapper:hover .cal-hover-popup {
        display: block;
        animation: fadeIn 0.2s ease-out forwards;
    }
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .pos-left { left: 0; transform: none; }
    .pos-center { left: 50%; transform: translateX(-50%); }
    .pos-right { right: 0; left: auto; transform: none; }
</style>
""")

# 5. 사이드바 (기존 코드 100% 유지)
with st.sidebar:
    st.markdown("<div style='font-size: 18px; font-weight: 800; color: #ffffff; margin-bottom: 16px;'>📊 CAMMSYS BI</div>", unsafe_allow_html=True)
    menu = st.radio("FOLDER", (
        "🏠 현황", "1. 모델별 출하실적", "2. 모델별 생산실적", 
        "3. 출하 SITE별 실적", "4. 고객 FCST 조회", "5. 완제품 재고 조회", 
        "6. 캘린더", "7. 데이터 관리 (Admin)"
    ))
    
    st.markdown("<hr style='border-color: #334155; margin: 24px 0 16px 0;'>", unsafe_allow_html=True)
    if st.button("🔄 최신 실적 동기화", use_container_width=True):
        fetch_db_state_cached.clear()
        latest = fetch_db_state_cached()
        if latest:
            for k, v in latest.items():
                st.session_state[k] = v
            st.success("클라우드 최신 실적이 반영되었습니다.")
            st.rerun()
        else:
            st.info("동기화할 구글 시트 저장 데이터가 없습니다.")

# [프리미엄 SaaS] 차트 렌더링 (Ocean Indigo & Crisp Teal 컬러)
def create_active_chart(title, categories, plan_data, actual_data, name1="계획", name2="실적", unit="개"):
    fig = go.Figure()
    if not categories: categories, plan_data, actual_data = ["-"], [0], [0]
    
    if unit == "억원":
        txt_p = [f"{v:,.2f}" if v > 0 else "" for v in plan_data]
        txt_a = [f"{v:,.2f}" if v > 0 else "" for v in actual_data]
    else:
        txt_p = [format_k(v) if v > 0 else "" for v in plan_data]
        txt_a = [format_k(v) if v > 0 else "" for v in actual_data]
        
    # 계획 바 (세련된 스틸 톤)
    fig.add_trace(go.Bar(
        x=categories, y=plan_data, name=name1, 
        marker_color='#cbd5e1', 
        marker_line=dict(width=0), 
        text=txt_p, textposition='outside', textfont=dict(size=11, color='#94a3b8', family='Pretendard')
    ))
    # 실적 바 (고급스러운 인디고 톤)
    fig.add_trace(go.Bar(
        x=categories, y=actual_data, name=name2, 
        marker_color='#4f46e5', 
        marker_line=dict(width=0), 
        text=txt_a, textposition='outside', textfont=dict(size=12, color='#0f172a', family='Pretendard', weight='bold')
    ))
    fig.update_layout(
        title=dict(text=f"<b>{title}</b> <span style='font-size:12px; font-weight:500; color:#64748b;'>({unit})</span>", font=dict(size=16, color='#0f172a', family='Pretendard')), 
        barmode='group', bargap=0.3, bargroupgap=0.1,
        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
        hovermode='x unified', margin=dict(l=10, r=10, t=50, b=20), 
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1, font=dict(size=12, color='#475569')),
        yaxis=dict(showgrid=True, gridcolor='#f1f5f9', zeroline=False, tickfont=dict(color='#64748b')),
        xaxis=dict(showgrid=False, tickfont=dict(color='#475569', weight='bold'))
    )
    return fig

# 6. 메인 로직
if menu == "7. 데이터 관리 (Admin)":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px;'>⚙️ Data Management</h1>", unsafe_allow_html=True)
    if st.text_input("관리자 비밀번호", type="password") == "2250":
        c_set1, c_set2 = st.columns(2)
        with c_set1: target_date_input = st.text_input("📅 조회 기준 날짜 지정 (예: 11/2)", default_target_date)
        with c_set2: ex_rate_input = st.number_input("💵 적용 환율 (원/USD)", value=1350.0, step=10.0)
        
        col1, col2 = st.columns(2)
        with col1: prod_file = st.file_uploader("1. 생산 실적 파일", type=['xlsx'], key="admin_prod_file")
        with col2: ship_actual_file = st.file_uploader("2. 출하 실적 파일 (신규 양식)", type=['xlsx'], key="admin_ship_actual_file")

        col3, col4 = st.columns(2)
        with col3: ship_plan_file = st.file_uploader("3. 출하 계획 파일 (RTF Main 시트)", type=['xlsx'], key="admin_ship_plan_file")
        with col4: map_rev_file = st.file_uploader("4. 단가 매핑 파일 (선택)", type=['xlsx'], key="admin_map_rev_file")
            
        if st.button("🚀 데이터 분석 및 클라우드 반영", type="primary"):
            st.session_state['target_date'] = target_date_input

            if map_rev_file:
                nmap = parse_mapping_excel(map_rev_file)
                if nmap: 
                    st.session_state['mapping_data'].update(nmap)
                    st.success(f"✅ 단가표 인식 완료 (GH-CODE {len(nmap['gh_map'])}개 등록)")

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
                        st.success(f"✅ 생산 데이터 분석 완료 (Daily 계획: {format_k(p['daily_prod_plan'])} / 실적: {format_k(p['daily_prod_actual'])})")
                    except Exception as e: st.error(f"생산 파싱 오류: {e}")
            
            if ship_actual_file:
                with st.spinner("출하 실적 및 이월 매출을 분석하고 있습니다..."):
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
                            st.success(f"✅ 출하 실적 분석 완료 (Daily 실적: {format_k(s_act['daily_ship_a'])} / {s_act['daily_rev_a']:,.2f}억원)")
                    except Exception as e: st.error(f"출하 실적 파싱 오류: {e}")

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
                            st.success(f"✅ 계획 데이터 분석 완료 (Daily 계획: {format_k(s_plan['daily_ship_p'])} / {s_plan['daily_rev_p']:,.2f}억원)")
                    except Exception as e: st.error(f"출하 계획 파싱 오류: {e}")

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
                    '모델명': m, '출하계획(개)': p_qty, '출하계획(K)': format_k(p_qty),
                    '출하실적(개)': a_qty, '출하실적(K)': format_k(a_qty), '수량달성률(%)': qty_rate,
                    '수량차질(K)': format_k(p_qty - a_qty), '매출계획(억원)': p_rev,
                    '매출실적(억원)': a_rev, '매출달성률(%)': rev_rate
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
                    'SITE': s, '출하계획(개)': p_qty, '출하계획(K)': format_k(p_qty),
                    '출하실적(개)': a_qty, '출하실적(K)': format_k(a_qty), '수량달성률(%)': qty_rate,
                    '수량차질(K)': format_k(p_qty - a_qty), '매출계획(억원)': p_rev,
                    '매출실적(억원)': a_rev, '매출달성률(%)': rev_rate
                })
            st.session_state['chart_data']['ship_by_site'] = merged_sites

            with st.spinner("구글 스프레드시트에 동기화 중..."):
                sync_keys = [
                    'target_date', 'kpi_data', 'chart_data', 'inventory_data',
                    'fcst_data', 'actual_by_day', 'mapping_data', 'data_loaded',
                    '_ship_plan_by_model', '_ship_actual_by_model',
                    '_ship_plan_by_site', '_ship_actual_by_site'
                ]
                save_payload = {k: st.session_state[k] for k in sync_keys if k in st.session_state}
                is_saved = save_state_to_gsheet(save_payload)
                if is_saved:
                    st.success("🎉 클라우드 동기화 완료! 이제 모든 사용자의 화면에 최신 실적이 표시됩니다.")
                else:
                    st.warning("⚠️ 분석은 완료되었으나 클라우드 저장에 실패했습니다.")

elif menu == "🏠 현황":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px; margin-bottom: 8px;'>Overview</h1>", unsafe_allow_html=True)
    
    cur_target_date = st.session_state.get('target_date', default_target_date)
    render_html(f"""
    <div class="date-badge">
        <span style="font-size: 16px;">📅</span> 
        <span>Reporting Date</span>
        <span style="color: #4f46e5; font-weight: 900;">{cur_target_date}</span>
    </div>
    """)

    kpi = st.session_state.get('kpi_data', default_kpis)
    cdata = st.session_state.get('chart_data', default_charts)

    tab_monthly, tab_weekly, tab_daily = st.tabs(["🗓️ Monthly 누계", "📊 Weekly 주차별", "📅 Daily 당일"])

    with tab_monthly:
        m1, m2, m3 = st.columns(3)
        with m1: 
            render_kpi_card("월간 생산 누계", "🏭", kpi['monthly']['prod_a'], kpi['monthly']['prod_p'])
            render_achievement_gauge("생산 목표", kpi['monthly']['prod_a'], kpi['monthly']['prod_p'], unit="K")
        with m2: 
            render_kpi_card("월간 출하 누계", "🚚", kpi['monthly']['ship_a'], kpi['monthly']['ship_p'])
            render_achievement_gauge("출하 목표", kpi['monthly']['ship_a'], kpi['monthly']['ship_p'], unit="K")
        with m3: 
            rev_sub = f"당월 실적 {kpi['monthly']['pure_rev_a']:,.2f}억 / 전월 이월 {kpi['monthly']['carryover_rev']:,.2f}억"
            render_kpi_card("월간 매출 누계", "💳", kpi['monthly']['rev_a'], kpi['monthly']['rev_p'], is_rev=True, sub_caption=rev_sub)
            render_achievement_gauge("매출 목표", kpi['monthly']['rev_a'], kpi['monthly']['rev_p'], unit="억원")

        mch1, mch2, mch3 = st.columns(3)
        m_months = cdata.get('monthly_months', ['11월'])
        with mch1:
            st.plotly_chart(create_active_chart("월별 생산 추이", m_months, cdata.get('monthly_prod_p', [0]), cdata.get('monthly_prod_a', [0]), unit="K"), use_container_width=True)
        with mch2:
            st.plotly_chart(create_active_chart("월별 출하 추이", m_months, cdata.get('monthly_ship_p', [0]), cdata.get('monthly_ship_a', [0]), unit="K"), use_container_width=True)
        with mch3:
            st.plotly_chart(create_active_chart("월별 매출 추이", m_months, cdata.get('monthly_rev_p', [0.0]), cdata.get('monthly_rev_a', [0.0]), unit="억원"), use_container_width=True)

    with tab_weekly:
        wk1, wk2, wk3 = st.columns(3)
        with wk1: 
            render_kpi_card("주간 생산 누계", "🏭", kpi['weekly']['prod_a'], kpi['weekly']['prod_p'])
            render_achievement_gauge("생산 목표", kpi['weekly']['prod_a'], kpi['weekly']['prod_p'], unit="K")
        with wk2: 
            render_kpi_card("주간 출하 누계", "🚚", kpi['weekly']['ship_a'], kpi['weekly']['ship_p'])
            render_achievement_gauge("출하 목표", kpi['weekly']['ship_a'], kpi['weekly']['ship_p'], unit="K")
        with wk3: 
            rev_sub = f"당주 실적 {kpi['weekly']['pure_rev_a']:,.2f}억 / 이월 {kpi['weekly']['carryover_rev']:,.2f}억"
            render_kpi_card("주간 매출 누계", "💳", kpi['weekly']['rev_a'], kpi['weekly']['rev_p'], is_rev=True, sub_caption=rev_sub)
            render_achievement_gauge("매출 목표", kpi['weekly']['rev_a'], kpi['weekly']['rev_p'], unit="억원")

        wch1, wch2, wch3 = st.columns(3)
        w_weeks = cdata.get('weekly_weeks', ['W45', 'W46', 'W47', 'W48'])
        with wch1:
            st.plotly_chart(create_active_chart("주차별 생산 추이", w_weeks, cdata.get('weekly_prod_p', [0]*len(w_weeks)), cdata.get('weekly_prod_a', [0]*len(w_weeks)), unit="K"), use_container_width=True)
        with wch2:
            st.plotly_chart(create_active_chart("주차별 출하 추이", w_weeks, cdata.get('weekly_ship_p', [0]*len(w_weeks)), cdata.get('weekly_ship_a', [0]*len(w_weeks)), unit="K"), use_container_width=True)
        with wch3:
            st.plotly_chart(create_active_chart("주차별 매출 추이", w_weeks, cdata.get('weekly_rev_p', [0.0]*len(w_weeks)), cdata.get('weekly_rev_a', [0.0]*len(w_weeks)), unit="억원"), use_container_width=True)

    with tab_daily:
        k1, k2, k3 = st.columns(3)
        with k1: 
            render_kpi_card("당일 생산 실적", "🏭", kpi['daily']['prod_a'], kpi['daily']['prod_p'])
            render_achievement_gauge("생산 목표", kpi['daily']['prod_a'], kpi['daily']['prod_p'], unit="K")
        with k2: 
            render_kpi_card("당일 출하 실적", "🚚", kpi['daily']['ship_a'], kpi['daily']['ship_p'])
            render_achievement_gauge("출하 목표", kpi['daily']['ship_a'], kpi['daily']['ship_p'], unit="K")
        with k3: 
            render_kpi_card("당일 매출 실적", "💳", kpi['daily']['rev_a'], kpi['daily']['rev_p'], is_rev=True)
            render_achievement_gauge("매출 목표", kpi['daily']['rev_a'], kpi['daily']['rev_p'], unit="억원")

        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            st.plotly_chart(create_active_chart("생산 실적 (당일)", ["Today"], [kpi['daily']['prod_p']], [kpi['daily']['prod_a']], unit="K"), use_container_width=True)
        with ch2:
            st.plotly_chart(create_active_chart("출하 실적 (당일)", ["Today"], [kpi['daily']['ship_p']], [kpi['daily']['ship_a']], unit="K"), use_container_width=True)
        with ch3:
            st.plotly_chart(create_active_chart("매출 실적 (당일)", ["Today"], [kpi['daily']['rev_p']], [kpi['daily']['rev_a']], unit="억원"), use_container_width=True)

elif menu == "1. 모델별 출하실적":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px;'>📦 모델별 실적</h1>", unsafe_allow_html=True)
    df_ship = pd.DataFrame(st.session_state['chart_data'].get('ship_by_model', []))
    if not df_ship.empty:
        df_ship = df_ship.sort_values('출하실적(개)', ascending=False)
        
        tab_qty, tab_rev = st.tabs(["🚚 출하 수량 현황 (K)", "💰 매출액 현황 (억원)"])
        with tab_qty:
            st.plotly_chart(create_active_chart("주요 모델별 출하 수량", df_ship['모델명'].tolist(), df_ship['출하계획(개)'].tolist(), df_ship['출하실적(개)'].tolist(), unit="K"), use_container_width=True)
        with tab_rev:
            st.plotly_chart(create_active_chart("주요 모델별 매출액", df_ship['모델명'].tolist(), df_ship['매출계획(억원)'].tolist(), df_ship['매출실적(억원)'].tolist(), unit="억원"), use_container_width=True)
            
        col1, col2 = st.columns([0.8, 0.2])
        with col2: st.download_button(label="📥 엑셀(CSV) 다운로드", data=convert_df(df_ship), file_name='model_shipment_performance.csv', mime='text/csv', use_container_width=True)
        st.dataframe(df_ship[['모델명', '출하계획(K)', '출하실적(K)', '수량달성률(%)', '수량차질(K)', '매출계획(억원)', '매출실적(억원)', '매출달성률(%)']], use_container_width=True, hide_index=True)
    else: st.info("데이터가 없습니다.")
        
elif menu == "2. 모델별 생산실적":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px;'>🛠️ 생산 실적</h1>", unsafe_allow_html=True)
    cdata = st.session_state['chart_data']
    st.plotly_chart(create_active_chart("주요 모델별 생산 실적 현황", cdata.get('daily_models', []), cdata.get('daily_prod_p', []), cdata.get('daily_prod_a', []), unit="K"), use_container_width=True)
    
    df_prod = pd.DataFrame({"모델명": cdata.get('daily_models', []), "생산계획": cdata.get('daily_prod_p', []), "생산실적": cdata.get('daily_prod_a', [])})
    if not df_prod.empty:
        df_prod['생산계획(K)'] = df_prod['생산계획'].apply(format_k)
        df_prod['생산실적(K)'] = df_prod['생산실적'].apply(format_k)
        df_prod['달성률(%)'] = (df_prod['생산실적'] / df_prod['생산계획'].replace(0, 1) * 100).round(1)
        df_prod['차질수량(K)'] = (df_prod['생산계획'] - df_prod['생산실적']).apply(format_k)
        col1, col2 = st.columns([0.8, 0.2])
        with col2: st.download_button(label="📥 엑셀 다운로드", data=convert_df(df_prod), file_name='model_production.csv', mime='text/csv', use_container_width=True)
        st.dataframe(df_prod[['모델명', '생산계획(K)', '생산실적(K)', '달성률(%)', '차질수량(K)']], use_container_width=True, hide_index=True)

elif menu == "3. 출하 SITE별 실적":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px;'>📍 SITE 실적</h1>", unsafe_allow_html=True)
    df_site = pd.DataFrame(st.session_state['chart_data'].get('ship_by_site', []))
    if not df_site.empty:
        df_site = df_site.sort_values('출하실적(개)', ascending=False)
        
        tab_s_qty, tab_s_rev = st.tabs(["🚚 SITE별 출하 수량 (K)", "💰 SITE별 매출액 (억원)"])
        with tab_s_qty:
            st.plotly_chart(create_active_chart("거점별 출하 수량 현황", df_site['SITE'].tolist(), df_site['출하계획(개)'].tolist(), df_site['출하실적(개)'].tolist(), unit="K"), use_container_width=True)
        with tab_s_rev:
            st.plotly_chart(create_active_chart("거점별 매출액 현황", df_site['SITE'].tolist(), df_site['매출계획(억원)'].tolist(), df_site['매출실적(억원)'].tolist(), unit="억원"), use_container_width=True)
            
        col1, col2 = st.columns([0.8, 0.2])
        with col2: st.download_button(label="📥 엑셀 다운로드", data=convert_df(df_site), file_name='site_shipment_performance.csv', mime='text/csv', use_container_width=True)
        st.dataframe(df_site[['SITE', '출하계획(K)', '출하실적(K)', '수량달성률(%)', '수량차질(K)', '매출계획(억원)', '매출실적(억원)', '매출달성률(%)']], use_container_width=True, hide_index=True)
    else: st.info("데이터가 없습니다.")

elif menu == "4. 고객 FCST 조회":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px;'>📈 FCST 현황</h1>", unsafe_allow_html=True)
    fcst_data = st.session_state.get('fcst_data', {'rows': [], 'date_cols': []})
    fcst_rows = fcst_data.get('rows', [])
    date_cols = fcst_data.get('date_cols', [])

    if not fcst_rows:
        st.info("💡 출하 계획 파일(RTF Main 시트)을 Admin 메뉴에서 업로드 및 분석해 주세요.")
    else:
        df_fcst = pd.DataFrame(fcst_rows)
        
        total_fcst_qty = int(df_fcst['합계'].sum())
        model_count = int(df_fcst['모델명'].nunique())
        site_count = int(df_fcst['SITE'].nunique())
        top_model = df_fcst.groupby('모델명')['합계'].sum().idxmax()
        top_model_qty = int(df_fcst.groupby('모델명')['합계'].sum().max())
        
        m1, m2, m3, m4 = st.columns(4)
        with m1: st.metric("📦 총 공급 계획 물량", f"{format_k(total_fcst_qty)}")
        with m2: st.metric("📱 대상 모델 수", f"{model_count} 개 모델")
        with m3: st.metric("🏭 공급 거점 수", f"{site_count} 개 SITE")
        with m4: st.metric("👑 최대 공급 모델", f"{top_model}", f"{format_k(top_model_qty)}")
        
        st.markdown("<br>", unsafe_allow_html=True)
        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        with f1:
            all_models = sorted(df_fcst['모델명'].unique().tolist())
            selected_models = st.multiselect("🔍 모델명 필터", options=all_models, default=[])
        with f2:
            all_sites = sorted(df_fcst['SITE'].unique().tolist())
            selected_sites = st.multiselect("🏢 출하 SITE 필터", options=all_sites, default=[])
        with f3:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            hide_zero_cols = st.checkbox("물량이 있는 일자만 표시", value=True)

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
            "📊 모델별 통합 일정", 
            "📋 SITE별 상세 공급 내역", 
            "📈 공급 추이 차트"
        ])

        with tab_by_model:
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
                        label="📥 CSV 다운로드",
                        data=convert_df(df_model_pivot),
                        file_name='fcst_supply_by_model.csv',
                        mime='text/csv',
                        use_container_width=True
                    )
                st.dataframe(df_model_disp, use_container_width=True, hide_index=True)
            else:
                st.warning("선택한 조건에 해당하는 일정이 없습니다.")

        with tab_by_detail:
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
                        label="📥 CSV 다운로드",
                        data=convert_df(df_detail),
                        file_name='fcst_supply_detailed.csv',
                        mime='text/csv',
                        use_container_width=True
                    )
                st.dataframe(df_detail_disp, use_container_width=True, hide_index=True)
            else:
                st.warning("선택한 조건에 해당하는 일정이 없습니다.")

        with tab_charts:
            ch_col1, ch_col2 = st.columns(2)
            with ch_col1:
                top10_models = df_filtered.groupby('모델명')['합계'].sum().sort_values(ascending=False).head(10)
                fig_top10 = go.Figure()
                fig_top10.add_trace(go.Bar(
                    x=top10_models.index.tolist(),
                    y=top10_models.values.tolist(),
                    marker_color='#4f46e5',
                    text=[format_k(v) for v in top10_models.values],
                    textposition='outside',
                    textfont=dict(size=11, color='#0f172a', family='Pretendard', weight='bold')
                ))
                fig_top10.update_layout(
                    title="<b>TOP 10 모델별 공급 물량</b> <span style='font-size:12px; font-weight:500; color:#64748b;'>(단위: K)</span>",
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    hovermode='x unified', margin=dict(l=10, r=10, t=40, b=10),
                    yaxis=dict(showgrid=True, gridcolor='#f1f5f9')
                )
                st.plotly_chart(fig_top10, use_container_width=True)
                
            with ch_col2:
                daily_trend = df_filtered[valid_dates].sum()
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Bar(
                    x=daily_trend.index.tolist(),
                    y=daily_trend.values.tolist(),
                    marker_color='#0d9488',
                    text=[format_k(v) if v > 0 else "" for v in daily_trend.values],
                    textposition='outside',
                    textfont=dict(size=11, color='#0f172a', family='Pretendard', weight='bold')
                ))
                fig_trend.update_layout(
                    title="<b>일자별 공급 추이</b> <span style='font-size:12px; font-weight:500; color:#64748b;'>(단위: K)</span>",
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    hovermode='x unified', margin=dict(l=10, r=10, t=40, b=10),
                    yaxis=dict(showgrid=True, gridcolor='#f1f5f9')
                )
                st.plotly_chart(fig_trend, use_container_width=True)

elif menu == "5. 완제품 재고 조회":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px;'>🏢 재고 현황</h1>", unsafe_allow_html=True)
    st.metric("📦 당일 완제품 총 재고", f"{format_k(st.session_state['inventory_data']['qty'])}", f"재고일수: {st.session_state['inventory_data']['days']} 일", delta_color="off")

elif menu == "6. 캘린더":
    st.markdown("<h1 style='font-weight: 900; color: #0f172a; letter-spacing: -1px; margin-bottom: 8px;'>🗓️ 공급 캘린더</h1>", unsafe_allow_html=True)
    fcst_data = st.session_state.get('fcst_data', {'rows': [], 'date_cols': []})
    fcst_rows = fcst_data.get('rows', [])
    actual_by_day = st.session_state.get('actual_by_day', {})
    
    if not fcst_rows:
        st.info("💡 출하 계획을 Admin 메뉴에서 업로드 및 분석해 주세요.")
    else:
        t_m, t_d = parse_target_day_info(st.session_state.get('target_date', default_target_date))

        c_m1, c_m2 = st.columns([1.2, 2.8])
        with c_m1:
            month_options = [10, 11, 12, 1]
            month_labels = {10: "2026년 10월", 11: "2026년 11월 (당월)", 12: "2026년 12월", 1: "2027년 1월"}
            default_idx = month_options.index(t_m) if t_m in month_options else 1
            selected_m = st.selectbox("📅 조회 대상 월", options=month_options, format_func=lambda x: month_labels.get(x, f"{x}월"), index=default_idx)
            
        cur_year = 2027 if selected_m == 1 else 2026
        cal_instance = calendar.Calendar(firstweekday=6)
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
                        'model': m_name, 'gh_code': gh, 'site': site, 'qty': q, 'price': price, 'rev_ok': rev_val
                    })

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
            <div style="background: #ffffff; border: 1px solid #e2e8f0; padding: 14px 20px; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); display: flex; justify-content: space-around; align-items: center;">
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">당월 총 공급계획</div>
                    <div style="font-size: 20px; font-weight: 900; color: #0f172a; margin-top: 4px; letter-spacing: -0.5px;">{format_k(m_total_qty)}</div>
                </div>
                <div style="border-left: 1px solid #f1f5f9; height: 32px;"></div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 700; color: #4f46e5; text-transform: uppercase; letter-spacing: 0.5px;">당월 예상 매출</div>
                    <div style="font-size: 20px; font-weight: 900; color: #4f46e5; margin-top: 4px; letter-spacing: -0.5px;">{m_total_rev:,.2f} 억원</div>
                </div>
                <div style="border-left: 1px solid #f1f5f9; height: 32px;"></div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">공급 예정 일수</div>
                    <div style="font-size: 20px; font-weight: 900; color: #0f172a; margin-top: 4px; letter-spacing: -0.5px;">{supply_days_cnt}일 <span style="font-size: 14px; font-weight: 600; color: #cbd5e1;">/ {num_days}일</span></div>
                </div>
                <div style="border-left: 1px solid #f1f5f9; height: 32px;"></div>
                <div style="text-align: center;">
                    <div style="font-size: 11px; font-weight: 700; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.5px;">최대 공급 모델</div>
                    <div style="font-size: 16px; font-weight: 900; color: #0d9488; margin-top: 4px; letter-spacing: -0.3px;">{top_m_name} ({format_k(top_m_qty)})</div>
                </div>
            </div>
            """)

        st.markdown("<br>", unsafe_allow_html=True)

        days_headers = [
            ("<span style='color: #f43f5e;'>SUN</span>", "#f43f5e"),
            ("MON", "#64748b"),
            ("TUE", "#64748b"),
            ("WED", "#64748b"),
            ("THU", "#64748b"),
            ("FRI", "#64748b"),
            ("<span style='color: #4f46e5;'>SAT</span>", "#4f46e5")
        ]
        
        h_cols = st.columns(7)
        for i, (h_title, text_color) in enumerate(days_headers):
            with h_cols[i]:
                render_html(f"""
                <div style="background: transparent; padding: 4px; text-align: center; font-weight: 800; font-size: 12px; color: #94a3b8; letter-spacing: 1px; margin-bottom: 4px;">
                    {h_title}
                </div>
                """)

        for week in month_matrix:
            w_cols = st.columns(7)
            for day_idx, d_num in enumerate(week):
                with w_cols[day_idx]:
                    if d_num == 0:
                        render_html("<div style='height: 110px; background: transparent; border-radius: 12px; margin-bottom: 10px;'></div>")
                    else:
                        d_info = day_supply_map[d_num]
                        plan_q = d_info['total']
                        plan_rev = d_info['total_rev']
                        
                        is_target = (selected_m == t_m and d_num == t_d)
                        is_past = (selected_m < t_m) or (selected_m == t_m and d_num <= t_d)
                        
                        act_day_data = actual_by_day.get(d_num, {'qty': 0, 'rev': 0.0, 'items': []})
                        act_q = act_day_data['qty']
                        act_rev = act_day_data['rev']
                        
                        if is_target:
                            border_css = "1.5px solid #4f46e5; box-shadow: 0 4px 12px rgba(79, 70, 229, 0.15);"
                            bg_css = "#eef2ff"
                            badge_target = "<span style='background: #4f46e5; color: #ffffff; font-size: 10px; font-weight: 800; padding: 2px 6px; border-radius: 20px; margin-left: 6px; letter-spacing: -0.2px;'>TODAY</span>"
                        else:
                            border_css = "1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.02);"
                            bg_css = "#ffffff" if (plan_q > 0 or act_q > 0) else "#f8fafc"
                            badge_target = ""

                        if day_idx == 0: day_color = "#f43f5e"
                        elif day_idx == 6: day_color = "#4f46e5"
                        else: day_color = "#0f172a"

                        if is_past:
                            diff_q = act_q - plan_q
                            diff_rev = act_rev - plan_rev
                            
                            if diff_q > 0: diff_q_tag = f"<span style='color: #4f46e5; font-weight: 800;'>+{format_k(diff_q)}</span>"
                            elif diff_q < 0: diff_q_tag = f"<span style='color: #f43f5e; font-weight: 800;'>-{format_k(abs(diff_q))}</span>"
                            else: diff_q_tag = "<span style='color: #94a3b8; font-weight: 700;'>0K</span>"

                            if diff_rev > 0: diff_rev_tag = f"<span style='color: #4f46e5; font-weight: 800;'>+{diff_rev:,.2f}억</span>"
                            elif diff_rev < 0: diff_rev_tag = f"<span style='color: #f43f5e; font-weight: 800;'>-{abs(diff_rev):,.2f}억</span>"
                            else: diff_rev_tag = "<span style='color: #94a3b8; font-weight: 700;'>0.00억</span>"

                            board_content_html = f"""
                            <div style="display: flex; flex-direction: column; gap: 4px; margin-top: 8px;">
                                <div style="font-size: 11.5px; font-weight: 700; color: #0f172a; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #64748b;">실적</span>
                                    <span><span style="font-size: 13px; font-weight: 900;">{format_k(act_q)}</span> <span style="font-size: 10px;">({diff_q_tag})</span></span>
                                </div>
                                <div style="font-size: 11.5px; font-weight: 700; color: #0f172a; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #64748b;">매출</span>
                                    <span><span style="font-size: 13px; font-weight: 900;">{act_rev:,.2f}</span> <span style="font-size: 10px;">({diff_rev_tag})</span></span>
                                </div>
                            </div>
                            """
                        else:
                            board_content_html = f"""
                            <div style="display: flex; flex-direction: column; gap: 4px; margin-top: 8px;">
                                <div style="font-size: 11.5px; font-weight: 700; color: #0f172a; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #64748b;">계획</span>
                                    <span style="font-size: 13px; font-weight: 900; color: #0f172a;">{format_k(plan_q)}</span>
                                </div>
                                <div style="font-size: 11.5px; font-weight: 700; color: #0f172a; display: flex; justify-content: space-between; align-items: baseline;">
                                    <span style="color: #64748b;">매출</span>
                                    <span style="font-size: 13px; font-weight: 900; color: #4f46e5;">{plan_rev:,.2f}억</span>
                                </div>
                            </div>
                            """

                        unique_models_cnt = len(set(it['model'] for it in d_info['items']))
                        items_sorted = sorted(d_info['items'], key=lambda x: x['qty'], reverse=True)
                        
                        hover_list_html = ""
                        for it in items_sorted[:6]:
                            q_txt = format_k(it['qty'])
                            hover_list_html += f"<div style='font-size: 11px; color: #0f172a; padding: 4px 0; border-bottom: 1px solid rgba(0,0,0,0.04); display: flex; justify-content: space-between;'><span style='font-weight: 600;'>{it['model']} <span style='color: #64748b;'>({it['site']})</span></span><span><b style='color: #4f46e5;'>{q_txt}</b> / {it['rev_ok']:,.2f}억</span></div>"
                        if len(items_sorted) > 6:
                            hover_list_html += f"<div style='font-size: 10px; color: #94a3b8; font-weight: 800; text-align: center; margin-top: 6px;'>+ {len(items_sorted)-6} ITEMS MORE</div>"

                        if day_idx <= 1: pos_class = "pos-left"
                        elif day_idx >= 5: pos_class = "pos-right"
                        else: pos_class = "pos-center"

                        hover_popup_html = f"""
                        <div class="cal-hover-popup {pos_class}">
                            <div style="font-size: 14px; font-weight: 900; color: #0f172a; border-bottom: 1px solid rgba(0,0,0,0.06); padding-bottom: 8px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center;">
                                <span>{selected_m}. {d_num} Schedule</span>
                                <span style="font-size: 11px; background: #e2e8f0; color: #475569; padding: 2px 8px; border-radius: 20px;">{unique_models_cnt} Models</span>
                            </div>
                            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin-bottom: 12px;">
                                <div style="background: rgba(255,255,255,0.5); border: 1px solid rgba(0,0,0,0.04); padding: 8px; border-radius: 8px; text-align: center;">
                                    <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Qty</div>
                                    <div style="font-size: 15px; font-weight: 900; color: #4f46e5; margin-top: 2px;">{format_k(plan_q)}</div>
                                </div>
                                <div style="background: rgba(255,255,255,0.5); border: 1px solid rgba(0,0,0,0.04); padding: 8px; border-radius: 8px; text-align: center;">
                                    <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Rev</div>
                                    <div style="font-size: 15px; font-weight: 900; color: #0d9488; margin-top: 2px;">{plan_rev:,.2f}억</div>
                                </div>
                            </div>
                            <div style="max-height: 150px; overflow-y: auto;">
                                {hover_list_html if hover_list_html else "<div style='font-size: 11px; color: #94a3b8; text-align: center; padding: 10px;'>일정 없음</div>"}
                            </div>
                        </div>
                        """

                        full_cell_html = f"""
                        <div class="cal-card-wrapper">
                            <div class="cal-card" style="background: {bg_css}; border: {border_css};">
                                <div style="display: flex; justify-content: space-between; align-items: baseline; border-bottom: 1px solid #f1f5f9; padding-bottom: 4px;">
                                    <span style="font-weight: 900; font-size: 16px; color: {day_color}; letter-spacing: -0.5px;">{d_num}{badge_target}</span>
                                    <span style="font-size: 10.5px; font-weight: 700; color: #94a3b8;">{format_k(plan_q)}</span>
                                </div>
                                {board_content_html}
                            </div>
                            {hover_popup_html}
                        </div>
                        """
                        render_html(full_cell_html)

                        btn_label = f"🔍 {unique_models_cnt}종 상세" if (plan_q > 0 or act_q > 0) else f"🔍 {d_num}일"
                        if st.button(btn_label, key=f"btn_pop_{selected_m}_{d_num}", use_container_width=True):
                            if hasattr(st, "dialog"):
                                show_day_detail_dialog(d_num, selected_m, cur_year, d_info)
                            else:
                                st.session_state['active_cal_popup'] = (d_num, selected_m, cur_year)
                                st.rerun()

        st.divider()

        st.markdown(f"### 📈 {selected_m}월 일자별 공급 추이 및 예상 매출")
        trend_days = [f"{selected_m}/{d:02d}" for d in range(1, num_days + 1)]
        trend_qtys = [day_supply_map[d]['total'] for d in range(1, num_days + 1)]
        trend_revs = [round(day_supply_map[d]['total_rev'], 2) for d in range(1, num_days + 1)]
        
        fig_cal_trend = go.Figure()
        fig_cal_trend.add_trace(go.Bar(
            x=trend_days, y=trend_qtys, name="공급수량(K)",
            marker_color=['#4f46e5' if (d == t_d and selected_m == t_m) else '#cbd5e1' for d in range(1, num_days + 1)],
            text=[format_k(q) if q > 0 else "" for q in trend_qtys],
            textposition='outside', textfont=dict(size=10, color='#64748b', family='Pretendard', weight='bold')
        ))
        fig_cal_trend.add_trace(go.Scatter(
            x=trend_days, y=trend_revs, name="예상매출(억원)", yaxis="y2",
            mode="lines+markers", line=dict(color="#f43f5e", width=2.5), marker=dict(size=6, color="#f43f5e")
        ))
        fig_cal_trend.update_layout(
            title=dict(text=f"<b>{selected_m}월 일자별 공급 계획 수량 및 예상 매출액</b>", font=dict(size=16, color='#0f172a', family='Pretendard')),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            hovermode='x unified', margin=dict(l=10, r=10, t=50, b=10),
            xaxis=dict(tickangle=-45, showgrid=False, tickfont=dict(color='#64748b')),
            yaxis=dict(title="공급수량 (개)", showgrid=True, gridcolor='#f1f5f9', tickfont=dict(color='#64748b')),
            yaxis2=dict(title="예상매출 (억원)", overlaying="y", side="right", showgrid=False, tickfont=dict(color='#f43f5e')),
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1)
        )
        st.plotly_chart(fig_cal_trend, use_container_width=True)
