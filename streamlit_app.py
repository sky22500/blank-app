import os
import re
import json
import shutil
import statistics
import datetime
import calendar
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# 1. 페이지 설정 및 영구 저장소 경로 설정
st.set_page_config(page_title="마케팅 부문 대시보드", layout="wide")

STATE_VERSION = 5
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
STORE_DIR = os.path.join(BASE_DIR, ".dashboard_store")
STATE_JSON_PATH = os.path.join(STORE_DIR, "dashboard_state.json")
CACHED_EXCEL_PATHS = {
    "prod": os.path.join(STORE_DIR, "latest_prod.xlsx"),
    "integrated": os.path.join(STORE_DIR, "latest_integrated_sales.xlsx"),
    "map_rev": os.path.join(STORE_DIR, "latest_map_rev.xlsx"),
}

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
default_target_date = f"{yesterday.month}/{yesterday.day}"


def render_html(html_str):
    if hasattr(st, "html"):
        st.html(html_str)
    else:
        st.markdown(html_str, unsafe_allow_html=True)


# 2. 날짜 · 포맷 · 병합셀 유틸리티 함수
def safe_cell(sheet, r, c):
    if r < 1 or c < 1 or r > sheet.max_row or c > sheet.max_column:
        return None
    return sheet.cell(r, c).value


def build_merged_lookup(sheet):
    lookup = {}
    merged_ranges = getattr(sheet, "merged_cells", None)
    if merged_ranges:
        for rng in merged_ranges.ranges:
            top_val = sheet.cell(rng.min_row, rng.min_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    lookup[(r, c)] = top_val
    return lookup


def merged_cell(sheet, merged_map, r, c):
    if r < 1 or c < 1 or r > sheet.max_row or c > sheet.max_column:
        return None
    if (r, c) in merged_map:
        return merged_map[(r, c)]
    return sheet.cell(r, c).value


def norm_txt(val):
    if val is None:
        return ""
    return re.sub(r"\s+", "", str(val)).upper()


def clean_key(val):
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).upper())


def extract_num(val):
    if val is None:
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s in ("", "-", "—", "–") or s.startswith("#"):
        return 0.0
    is_neg = False
    if any(sym in s for sym in ("△", "▲", "▼")) or (s.startswith("(") and s.endswith(")")):
        is_neg = True
    s_clean = s.replace(",", "")
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s_clean)
    if not m:
        return None
    num = float(m.group(0))
    if is_neg and num > 0:
        num = -num
    return num


def format_k(val):
    if val is None or pd.isna(val):
        return "0K"
    try:
        v = float(val)
    except Exception:
        return str(val)
    if v == 0:
        return "0K"
    k_val = round(v / 1000.0, 1)
    if k_val.is_integer():
        return f"{int(k_val):,}K"
    return f"{k_val:,.1f}K"


def fmt_signed_k(val):
    if val is None or pd.isna(val) or val == 0:
        return "0K"
    sign = "-" if val < 0 else ""
    return f"{sign}{format_k(abs(val))}"


def fmt_amt_ok(val, always_decimal=True):
    if val is None or pd.isna(val):
        return "0.0억" if always_decimal else "0억"
    v = float(val)
    if not always_decimal and round(v, 1).is_integer():
        return f"{int(round(v)):,}억"
    return f"{v:,.1f}억"


def get_gh_suffix(code):
    s = clean_key(code)
    m = re.search(r"GH(?:86|96)?(\d+.*)$", s)
    if m:
        return m.group(1)
    m2 = re.search(r"(\d+[A-Z0-9]*)$", s)
    if m2:
        return m2.group(1)
    return s


def clamp_date(y, m, d):
    m = min(max(int(m), 1), 12)
    max_d = calendar.monthrange(int(y), m)[1]
    d = min(max(int(d), 1), max_d)
    return datetime.date(int(y), m, d)


def infer_year_for_month(m, ref_date=None):
    ref = ref_date or yesterday
    if ref.month in (11, 12) and m in (1, 2):
        return ref.year + 1
    if ref.month in (1, 2) and m in (11, 12):
        return ref.year - 1
    return ref.year


def parse_target_ymd(target_str, ref_date=None):
    ref = ref_date or yesterday
    s = str(target_str or "").strip()

    m_full = re.search(r"\b(\d{2,4})\s*[-/.]\s*0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s)
    if m_full:
        y_raw = int(m_full.group(1))
        y = (2000 + y_raw) if y_raw < 100 else y_raw
        dt = clamp_date(y, int(m_full.group(2)), int(m_full.group(3)))
        return dt.year, dt.month, dt.day

    m_md = re.search(r"\b0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s)
    if m_md:
        mon = int(m_md.group(1))
        day = int(m_md.group(2))
        y = infer_year_for_month(mon, ref)
        dt = clamp_date(y, mon, day)
        return dt.year, dt.month, dt.day

    m_d = re.search(r"(\d{1,2})일?", s)
    if m_d:
        dt = clamp_date(ref.year, ref.month, int(m_d.group(1)))
        return dt.year, dt.month, int(m_d.group(1))

    return ref.year, ref.month, ref.day


def parse_excel_date_header(val, default_year, prev_month=None):
    """엑셀 날짜 헤더(datetime, '11/25', '11/4 (We)', '26-11-25' 등)를 (date, next_year, month)로 변환"""
    if val is None:
        return None, default_year, prev_month
    if isinstance(val, (datetime.datetime, datetime.date)):
        y = val.year if val.year >= 2000 else default_year
        dt = clamp_date(y, val.month, val.day)
        return dt, dt.year, dt.month

    s = str(val).strip()
    if not s or "~" in s or re.search(r"\bW\d+\b", s.upper()):
        return None, default_year, prev_month

    m_ymd = re.search(r"^(\d{2,4})\s*[-/.]\s*0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})", s)
    if m_ymd:
        yr_raw = int(m_ymd.group(1))
        yr = (2000 + yr_raw) if yr_raw < 100 else yr_raw
        dt = clamp_date(yr, int(m_ymd.group(2)), int(m_ymd.group(3)))
        return dt, dt.year, dt.month

    m_md = re.search(r"^0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})(?:\s*\([A-Za-z가-힣]+\))?$", s)
    if m_md:
        mon, day = int(m_md.group(1)), int(m_md.group(2))
        if 1 <= mon <= 12 and 1 <= day <= 31:
            yr = default_year
            if prev_month is not None and mon < prev_month and (prev_month - mon) >= 6:
                yr += 1
            dt = clamp_date(yr, mon, day)
            return dt, yr, mon

    return None, default_year, prev_month


def get_iso_week_key(d_obj):
    iso_y, iso_w, _ = d_obj.isocalendar()
    return f"{iso_y}-W{iso_w:02d}"


def get_iso_week_label(d_obj):
    return f"W{d_obj.isocalendar()[1]}"


def get_month_iso_weeks(year, month):
    num_days = calendar.monthrange(year, month)[1]
    weeks = []
    for d in range(1, num_days + 1):
        w_lbl = get_iso_week_label(datetime.date(year, month, d))
        if w_lbl not in weeks:
            weeks.append(w_lbl)
    return weeks


def resolve_model_name(gh_code, fallback_model, gh_to_model):
    gh_clean = clean_key(gh_code)
    gh_suf = get_gh_suffix(gh_code)
    if gh_clean and gh_clean in gh_to_model:
        return gh_to_model[gh_clean]
    if gh_suf and gh_suf in gh_to_model:
        return gh_to_model[gh_suf]
    fb_str = str(fallback_model or "").strip()
    if not fb_str or fb_str.upper() in ("NONE", "#N/A", "N/A"):
        return gh_code or "기타"
    m = re.search(r"CAM_([A-Z0-9]+)", fb_str.upper())
    if m:
        extracted = m.group(1)
        return extracted[:-1] if extracted.endswith("E") else extracted
    return fb_str


def is_match_date(cell_val, target_str):
    if not cell_val:
        return False
    _, t_m, t_d = parse_target_ymd(target_str)
    if isinstance(cell_val, (datetime.datetime, datetime.date)):
        return cell_val.month == t_m and cell_val.day == t_d
    s = str(cell_val).strip().upper()
    m_ymd = re.search(r"\b\d{2,4}\s*[-/.]\s*0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s)
    if m_ymd:
        return int(m_ymd.group(1)) == t_m and int(m_ymd.group(2)) == t_d
    m = re.search(r"\b0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s)
    if m:
        return int(m.group(1)) == t_m and int(m.group(2)) == t_d
    m2 = re.search(r"\b0?(\d{1,2})\s*[-/.]\s*([A-Z]{3})\b", s)
    if m2:
        try:
            d = int(m2.group(1))
            mon = datetime.datetime.strptime(m2.group(2), "%b").month
            return mon == t_m and d == t_d
        except Exception:
            pass
    m3 = re.search(r"^0?(\d{1,2})일?$", s)
    if m3:
        return int(m3.group(1)) == t_d
    return False


@st.cache_data
def convert_df(df):
    return df.to_csv(index=False).encode("utf-8-sig")


# 3. 기본 상태 정의 및 서버 영구 저장/복구 로직
def get_default_state():
    t_y, t_m, t_d = parse_target_ymd(default_target_date)
    init_weeks = get_month_iso_weeks(t_y, t_m)
    cur_w = get_iso_week_label(datetime.date(t_y, t_m, t_d))
    return {
        "state_version": STATE_VERSION,
        "target_date": default_target_date,
        "ex_rate": 1330.0,
        "detected_ex_rate": None,
        "cur_week_label": cur_w,
        "last_updated_at": None,
        "data_loaded": False,
        "loaded_sheet_info": {},
        "mapping_data": {"gh_map": {}, "suffix_map": {}, "model_map": {}, "gh_to_model": {}},
        "fcst_data": {"rows": [], "date_cols": []},
        "shipment_report": {
            "sheet_name": "",
            "month_label": "",
            "ref_date_label": "",
            "summary_rows": [],
            "detail_rows": [],
        },
        "plan_by_date": {},
        "actual_by_date": {},
        "actual_months_loaded": [],
        "inventory_data": {"qty": 0, "days": 0.0},
        "kpi_data": {
            "daily": {"prod_p": 0, "prod_a": 0, "ship_p": 0, "ship_a": 0, "rev_p": 0.0, "rev_a": 0.0},
            "weekly": {
                "prod_p": 0,
                "prod_a": 0,
                "ship_p": 0,
                "ship_a": 0,
                "rev_p": 0.0,
                "rev_a": 0.0,
                "pure_rev_a": 0.0,
                "carryover_rev": 0.0,
            },
            "monthly": {
                "prod_p": 0,
                "prod_a": 0,
                "prod_sheet_total": 0,
                "prod_cur_w_add": 0,
                "ship_p": 0,
                "ship_a": 0,
                "rev_p": 0.0,
                "rev_a": 0.0,
                "target_ship": 0,
                "target_rev": 0.0,
                "pure_rev_a": 0.0,
                "carryover_rev": 0.0,
            },
        },
        "chart_data": {
            "daily_models": [],
            "daily_prod_p": [],
            "daily_prod_a": [],
            "weekly_weeks": init_weeks,
            "weekly_prod_p": [0] * len(init_weeks),
            "weekly_prod_a": [0] * len(init_weeks),
            "weekly_ship_p": [0] * len(init_weeks),
            "weekly_ship_a": [0] * len(init_weeks),
            "weekly_rev_p": [0.0] * len(init_weeks),
            "weekly_rev_a": [0.0] * len(init_weeks),
            "monthly_months": [f"{t_m}월"],
            "monthly_prod_p": [0],
            "monthly_prod_a": [0],
            "monthly_ship_p": [0],
            "monthly_ship_a": [0],
            "monthly_rev_p": [0.0],
            "monthly_rev_a": [0.0],
            "ship_by_model": [],
            "ship_by_model_period": {"daily": [], "weekly": [], "monthly": []},
        },
    }


PERSIST_KEYS = list(get_default_state().keys())


def _json_default(obj):
    if isinstance(obj, (datetime.date, datetime.datetime)):
        return obj.isoformat()
    if hasattr(obj, "item"):
        return obj.item()
    return str(obj)


def save_state_to_disk():
    os.makedirs(STORE_DIR, exist_ok=True)
    st.session_state["state_version"] = STATE_VERSION
    payload = {k: st.session_state.get(k) for k in PERSIST_KEYS}
    tmp_path = STATE_JSON_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=_json_default)
    os.replace(tmp_path, STATE_JSON_PATH)
    st.session_state["_loaded_mtime"] = os.path.getmtime(STATE_JSON_PATH)


def load_state_from_disk():
    if not os.path.exists(STATE_JSON_PATH):
        return False, None
    try:
        with open(STATE_JSON_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
        saved_ver = saved.get("state_version")
        defaults = get_default_state()
        for k, def_val in defaults.items():
            if k in saved and saved[k] is not None:
                if isinstance(def_val, dict) and isinstance(saved[k], dict):
                    merged = def_val.copy()
                    merged.update(saved[k])
                    st.session_state[k] = merged
                else:
                    st.session_state[k] = saved[k]
            else:
                st.session_state.setdefault(k, def_val)
        st.session_state["_loaded_mtime"] = os.path.getmtime(STATE_JSON_PATH)
        return True, saved_ver
    except Exception:
        return False, None


def reset_persisted_store():
    if os.path.exists(STORE_DIR):
        shutil.rmtree(STORE_DIR, ignore_errors=True)
    defaults = get_default_state()
    for k, v in defaults.items():
        st.session_state[k] = v
    st.session_state["_loaded_mtime"] = None
    st.session_state["active_cal_popup"] = None


# 4. 엑셀 파싱 엔진
# 4-1. 단가 매핑 파일 파서 (선택 업로드용)
def parse_mapping_excel(file_src):
    try:
        df_temp = pd.read_excel(file_src, header=None)
        header_idx = 0
        for i, row in df_temp.head(15).iterrows():
            row_str = "".join([str(x).upper() for x in row.values])
            if "GH" in row_str or "단가" in row_str or "PRICE" in row_str or "모델" in row_str:
                header_idx = i
                break

        df = pd.read_excel(file_src, header=header_idx)
        df.columns = df.columns.astype(str).str.strip().str.upper()

        gh_col = next((c for c in df.columns if "GH" in c or "SEC" in c or "MATERIAL" in c), None)
        mdl_col = next((c for c in df.columns if "모델" in c or "MODEL" in c), None)
        prc_col = next(
            (c for c in df.columns if "단가" in c or "PRICE" in c or "금액" in c or "USD" in c or "$" in c),
            None,
        )

        mapping = {"gh_map": {}, "suffix_map": {}, "model_map": {}, "gh_to_model": {}}
        if prc_col:
            for _, row in df.iterrows():
                price = extract_num(row[prc_col]) or 0.0
                mdl_val = str(row[mdl_col]).strip() if mdl_col and pd.notna(row[mdl_col]) else ""
                if gh_col and pd.notna(row[gh_col]):
                    gh_raw = str(row[gh_col])
                    gh_k = clean_key(gh_raw)
                    gh_suf = get_gh_suffix(gh_raw)
                    if gh_k:
                        mapping["gh_map"][gh_k] = price
                        mapping["suffix_map"][gh_suf] = price
                        if mdl_val:
                            mapping["gh_to_model"][gh_k] = mdl_val
                            mapping["gh_to_model"][gh_suf] = mdl_val
                if mdl_val:
                    mdl_k = clean_key(mdl_val)
                    if mdl_k:
                        mapping["model_map"][mdl_k] = price
        return mapping
    except Exception:
        return None


# 4-2. 생산 실적 파일 파서 (생산 섹션 전용 분리 · 주차별 순실적 계산 · 완제품재고 정밀 추출)
def parse_production_excel(file_src, target_date_str):
    wb = openpyxl.load_workbook(file_src, data_only=True)
    t_y, t_m, t_d = parse_target_ymd(target_date_str)
    target_date_obj = datetime.date(t_y, t_m, t_d)
    target_iso_week = get_iso_week_label(target_date_obj)

    parsed = {
        "daily_prod_plan": 0,
        "daily_prod_actual": 0,
        "cur_week_label": target_iso_week,
        "cur_week_prod_p": 0,
        "cur_week_prod_a": 0,
        "weekly_labels": [],
        "weekly_plan": [],
        "weekly_actual": [],
        "monthly_plan": 0,
        "monthly_actual": 0,
        "monthly_sheet_total": 0,
        "monthly_cur_w_add": 0,
        "monthly_months": [f"{t_m}월"],
        "inventory_qty": 0,
        "inventory_days": 0.0,
        "daily_models": [],
        "daily_prod_p": [],
        "daily_prod_a": [],
    }

    summary_sheet_name = next((s for s in wb.sheetnames if "요약" in s), wb.sheetnames[0])
    sheet = wb[summary_sheet_name]
    merged_map = build_merged_lookup(sheet)

    # 1) 당일 완제품재고 추출 (하단 '당일 완제품재고' 아래 '수량' / '재고일수' 행 탐색)
    for r in range(1, sheet.max_row + 1):
        row_joined = "".join(norm_txt(safe_cell(sheet, r, c)) for c in range(1, 8))
        if "당일완제품재고" in row_joined:
            for sub_r in range(r + 1, min(sheet.max_row + 1, r + 5)):
                lbl_a = norm_txt(safe_cell(sheet, sub_r, 1))
                if "수량" in lbl_a:
                    for c_try in range(2, 6):
                        v_q = extract_num(safe_cell(sheet, sub_r, c_try))
                        if v_q is not None and v_q > 0:
                            parsed["inventory_qty"] = int(round(v_q))
                            break
                elif "재고일수" in lbl_a or "일수" in lbl_a:
                    for c_try in range(2, 6):
                        v_d = extract_num(safe_cell(sheet, sub_r, c_try))
                        if v_d is not None and v_d > 0:
                            parsed["inventory_days"] = round(float(v_d), 1)
                            break
            break

    # 2) 요약 테이블 시작 행('구분') 탐색
    summary_start = 6
    for r in range(1, 20):
        if "구분" in "".join(norm_txt(safe_cell(sheet, r, c)) for c in range(1, 8)):
            summary_start = r
            break

    # 3) 헤더 행에서 일자별 컬럼, 주차 합계('W45 합계' 등), 월합계('11월 합계') 위치 파악
    col_map = {}
    weekly_sum_cols = {}
    daily_date_cols = {}  # col_idx -> datetime.date

    for r in range(summary_start, summary_start + 3):
        run_yr = t_y
        prev_m = None
        for c in range(1, sheet.max_column + 1):
            val = safe_cell(sheet, r, c)
            if val is None:
                continue
            if is_match_date(val, target_date_str) and "target" not in col_map:
                col_map["target"] = c

            vs = norm_txt(val)
            if "W" in vs and "합계" in vs:
                m_w = re.search(r"(W\d+)", vs)
                if m_w and m_w.group(1) not in weekly_sum_cols:
                    weekly_sum_cols[m_w.group(1)] = c
            elif "월합계" in vs and "월합계" not in col_map:
                col_map["월합계"] = c
                m_m = re.search(r"(\d+월)", vs)
                if m_m:
                    parsed["monthly_months"] = [m_m.group(1)]
            else:
                dt_cand, run_yr, prev_m = parse_excel_date_header(val, run_yr, prev_m)
                if dt_cand is not None and c not in daily_date_cols:
                    daily_date_cols[c] = dt_cand

    # 4) 오직 '생산' 블록('11월 생산 계획 대비 실적')의 '계획' / '실적' 행만 정확히 추출 (출하/매출 블록 합산 방지!)
    prod_plan_row = None
    prod_act_row = None

    for r in range(summary_start + 1, min(sheet.max_row + 1, summary_start + 35)):
        grp_a = norm_txt(merged_cell(sheet, merged_map, r, 1))
        lbl_b = norm_txt(safe_cell(sheet, r, 2))

        if any(stop_k in (grp_a + lbl_b) for stop_k in ("당일완제품재고", "주요모델별")):
            break

        # '출하' 또는 '매출' 섹션에 진입하면 생산 행 탐색 종료
        if ("출하" in grp_a or "매출" in grp_a) and "생산" not in grp_a:
            if prod_plan_row is not None:
                break
            continue

        if "생산" in grp_a or prod_plan_row is None:
            if lbl_b == "계획" and prod_plan_row is None:
                prod_plan_row = r
            elif lbl_b == "실적" and prod_act_row is None:
                prod_act_row = r

    def get_k_scaled_val(r, c):
        if not r or not c:
            return 0
        v = extract_num(merged_cell(sheet, merged_map, r, c))
        if v is None or v == 0:
            return 0
        # 요약 시트 단위가 K(예: 172 -> 172,000, 4,477 -> 4,477,000)이므로 100,000 미만은 x1,000 변환
        if 0 < abs(v) < 100_000:
            v *= 1000.0
        return int(round(v))

    parsed["daily_prod_plan"] = get_k_scaled_val(prod_plan_row, col_map.get("target"))
    parsed["daily_prod_actual"] = get_k_scaled_val(prod_act_row, col_map.get("target"))

    # 5) 주차별 계획 및 순수 주간 실적 산출 ('Wxx 합계' 실적 셀이 누계로 병합되어 있으므로 일별 실적 합산 우선)
    sorted_weeks = sorted(weekly_sum_cols.items(), key=lambda x: x[1])
    labels, wp, wa = [], [], []
    prev_bound_col = 3  # C열(10월 이월) 이후부터 스캔
    prev_cum_act = 0
    matched_week_lbl = target_iso_week

    for w_lbl, w_col in sorted_weeks:
        w_plan_val = get_k_scaled_val(prod_plan_row, w_col)
        w_sheet_act = get_k_scaled_val(prod_act_row, w_col)

        # prev_bound_col + 1 ~ w_col - 1 구간의 일별 컬럼 중 기준일 이하 실적 합산
        day_cols_in_span = [
            c for c in range(prev_bound_col + 1, w_col)
            if (c in daily_date_cols and daily_date_cols[c] <= target_date_obj)
        ]
        all_days_in_span = [daily_date_cols[c] for c in range(prev_bound_col + 1, w_col) if c in daily_date_cols]

        if col_map.get("target") and prev_bound_col < col_map["target"] <= w_col:
            matched_week_lbl = w_lbl

        if day_cols_in_span:
            pure_w_act = sum(get_k_scaled_val(prod_act_row, c) for c in day_cols_in_span)
            # 첫 주차(W45)에 시트상 누계가 일별 합보다 크면(이월 등) 보정
            if not labels and w_sheet_act > pure_w_act and all_days_in_span and max(all_days_in_span) <= target_date_obj:
                pure_w_act = w_sheet_act
        elif all_days_in_span and min(all_days_in_span) > target_date_obj:
            pure_w_act = 0
        else:
            # 일별 날짜 헤더가 없는 경우 주차 합계(누계) 차분으로 계산
            pure_w_act = max(0, w_sheet_act - prev_cum_act) if w_sheet_act >= prev_cum_act else w_sheet_act

        if w_sheet_act > prev_cum_act:
            prev_cum_act = w_sheet_act

        labels.append(w_lbl)
        wp.append(w_plan_val)
        wa.append(pure_w_act)
        prev_bound_col = w_col

    fallback_weeks = get_month_iso_weeks(t_y, t_m)
    parsed["weekly_labels"] = labels if labels else fallback_weeks
    parsed["weekly_plan"] = wp if wp else [0] * len(parsed["weekly_labels"])
    parsed["weekly_actual"] = wa if wa else [0] * len(parsed["weekly_labels"])

    week_plan_map = dict(zip(parsed["weekly_labels"], parsed["weekly_plan"]))
    week_act_map = dict(zip(parsed["weekly_labels"], parsed["weekly_actual"]))

    parsed["cur_week_label"] = matched_week_lbl
    parsed["cur_week_prod_p"] = week_plan_map.get(matched_week_lbl, 0)
    parsed["cur_week_prod_a"] = week_act_map.get(matched_week_lbl, 0)

    # 6) 월간 생산 계획 & 실적 산출
    m_plan = get_k_scaled_val(prod_plan_row, col_map.get("월합계"))
    m_sheet_act = get_k_scaled_val(prod_act_row, col_map.get("월합계"))

    # 기준일까지의 전체 일별 실적 합계 (당주 진행분 포함)
    valid_month_day_cols = [
        c for c, dt in daily_date_cols.items()
        if dt.year == t_y and dt.month == t_m and dt <= target_date_obj
    ]
    sum_daily_month_act = sum(get_k_scaled_val(prod_act_row, c) for c in valid_month_day_cols)
    sum_weekly_act = sum(wa)

    # 시트의 '11월 합계' 수식이 완료된 주차(예: W47=3,260K)까지만 잡혀 있고 당주(W48=281K)가 빠진 경우 자동 합산
    best_m_act = max(m_sheet_act, sum_weekly_act, sum_daily_month_act)
    parsed["monthly_plan"] = m_plan
    parsed["monthly_actual"] = best_m_act
    parsed["monthly_sheet_total"] = m_sheet_act
    parsed["monthly_cur_w_add"] = max(0, best_m_act - m_sheet_act) if m_sheet_act > 0 else 0

    # 7) 모델별 당일 생산실적 시트 파싱
    model_sheet = next((s for s in wb.sheetnames if "모델" in s), None)
    if model_sheet:
        ms = wb[model_sheet]
        hr, pr, ar = -1, -1, -1
        for r in range(1, 12):
            if "주요모델별" in "".join(norm_txt(safe_cell(ms, r, c)) for c in range(1, 6)):
                hr = r
                break
        if hr != -1:
            models, c_idx = [], []
            for c in range(2, ms.max_column + 1):
                mn = str(safe_cell(ms, hr, c) or "").strip()
                if mn and mn.upper() != "TOTAL" and "None" not in mn and "합계" not in mn:
                    models.append(mn)
                    c_idx.append(c)
            for r in range(hr + 1, hr + 18):
                rt = "".join(norm_txt(safe_cell(ms, r, c)) for c in range(1, 4))
                if "당일계획" in rt and pr == -1:
                    pr = r
                if "당일실적" in rt and ar == -1:
                    ar = r
            if pr != -1 and ar != -1:
                for c in c_idx:
                    parsed["daily_prod_p"].append(int(round(extract_num(safe_cell(ms, pr, c)) or 0)))
                    parsed["daily_prod_a"].append(int(round(extract_num(safe_cell(ms, ar, c)) or 0)))
                parsed["daily_models"] = models
    return parsed


# 4-3. 통합 출하·매출 파일 — (A) '월별요약' 시트 파서 (숨김열 제외 및 미병합 '~기준일' 헤더 완벽 처리)
def parse_monthly_summary_sheet(ws):
    merged_map = build_merged_lookup(ws)
    max_r = min(300, ws.max_row)
    max_c = min(60, ws.max_column)

    # 숨김 처리된 보조열(예: H, K, N, Q열 USD 보조금액 등) 판별
    hidden_cols = set()
    for c in range(1, max_c + 1):
        col_let = get_column_letter(c)
        if col_let in ws.column_dimensions and getattr(ws.column_dimensions[col_let], "hidden", False):
            hidden_cols.add(c)

    visible_cols = [c for c in range(1, max_c + 1) if c not in hidden_cols]

    detected_bands = []
    skip_until_r = 0

    for r in range(1, max_r):
        if r < skip_until_r:
            continue
        qty_cols, amt_cols = [], []
        for c in visible_cols:
            sub_raw = norm_txt(ws.cell(r + 1, c).value)
            if "수량" in sub_raw or sub_raw == "QTY":
                qty_cols.append(c)
            elif any(k in sub_raw for k in ("금액", "매출", "AMT")):
                amt_cols.append(c)

        if len(qty_cols) >= 2 and len(amt_cols) >= 2:
            first_m_c = min(qty_cols + amt_cols)
            last_m_c = min(max_c, max(qty_cols + amt_cols) + 5)

            col_map = {}
            month_label = ""
            ref_date_label = ""

            for c in visible_cols:
                if c >= first_m_c:
                    break
                lbl = norm_txt(merged_cell(ws, merged_map, r, c)) + norm_txt(merged_cell(ws, merged_map, r + 1, c))
                if "구분" in lbl:
                    col_map.setdefault("gubun", c)
                if any(k in lbl for k in ("과제", "PROJECT", "PRJ")):
                    col_map.setdefault("project", c)
                if any(k in lbl for k in ("모델", "MODEL", "품명")):
                    col_map.setdefault("model", c)

            # 상단 헤더가 미병합('I3=목표' | 'J3=~11/25', 'L3=실적' | 'M3=~11/25')이어도 이전 섹션명을 그대로 유지!
            carried_top = ""
            section_keywords = ("계획", "목표", "실적", "대비", "진도", "달성", "차이")

            for c in visible_cols:
                if c < first_m_c or c > last_m_c:
                    continue
                raw_top = str(merged_cell(ws, merged_map, r, c) or "").strip()
                raw_top_norm = norm_txt(raw_top)
                sub_clean = norm_txt(ws.cell(r + 1, c).value) or norm_txt(merged_cell(ws, merged_map, r + 1, c))

                m_mon = re.search(r"(\d{1,2})\s*월", raw_top)
                if m_mon and not month_label:
                    month_label = f"{int(m_mon.group(1))}월"
                m_ref = re.search(r"~\s*(\d{1,2}\s*[-/.]\s*\d{1,2})", raw_top)
                if m_ref and not ref_date_label:
                    ref_date_label = re.sub(r"\s+", "", m_ref.group(1))

                if any(kw in raw_top_norm for kw in section_keywords):
                    carried_top = raw_top_norm

                top_clean = carried_top

                if "계획" in top_clean and "대비" not in top_clean and "진도" not in top_clean:
                    if "수량" in sub_clean or sub_clean == "QTY":
                        col_map.setdefault("plan_qty", c)
                    elif any(k in sub_clean for k in ("금액", "매출", "AMT")):
                        col_map.setdefault("plan_amt", c)
                elif "대비" in top_clean or (top_clean == "차이" and ("수량" in sub_clean or "금액" in sub_clean)):
                    if "수량" in sub_clean or sub_clean == "QTY":
                        col_map.setdefault("diff_qty", c)
                    elif any(k in sub_clean for k in ("금액", "매출", "AMT")):
                        col_map.setdefault("diff_amt", c)
                elif "목표" in top_clean and not any(k in top_clean for k in ("진도", "달성", "대비")):
                    if "수량" in sub_clean or sub_clean == "QTY":
                        col_map.setdefault("target_qty", c)
                    elif any(k in sub_clean for k in ("금액", "매출", "AMT")):
                        col_map.setdefault("target_amt", c)
                elif "진도" in top_clean or "달성" in top_clean:
                    if "목표" in sub_clean or "계획" in sub_clean:
                        col_map.setdefault("rate_target", c)
                    elif "실적" in sub_clean:
                        col_map.setdefault("rate_act", c)
                elif "실적" in top_clean and not any(k in top_clean for k in ("진도", "달성")):
                    if "수량" in sub_clean or sub_clean == "QTY":
                        col_map.setdefault("act_qty", c)
                    elif any(k in sub_clean for k in ("금액", "매출", "AMT")):
                        col_map.setdefault("act_amt", c)
                elif ("차이" in raw_top_norm or "차이" in sub_clean) and not any(
                    k in sub_clean for k in ("수량", "금액", "매출", "QTY", "AMT")
                ) and "대비" not in raw_top_norm:
                    col_map.setdefault("rate_diff", c)

            # 가시 열(visible_cols) 순서 기반 안전 폴백
            if "act_qty" not in col_map and len(qty_cols) >= 3:
                col_map.setdefault("plan_qty", qty_cols[0])
                col_map.setdefault("target_qty", qty_cols[1])
                col_map.setdefault("act_qty", qty_cols[2])
                if len(qty_cols) >= 4:
                    col_map.setdefault("diff_qty", qty_cols[3])

            if "act_amt" not in col_map and len(amt_cols) >= 3:
                col_map.setdefault("plan_amt", amt_cols[0])
                col_map.setdefault("target_amt", amt_cols[1])
                col_map.setdefault("act_amt", amt_cols[2])
                if len(amt_cols) >= 4:
                    col_map.setdefault("diff_amt", amt_cols[3])

            table_kind = "detail" if ("project" in col_map or "model" in col_map) else "summary"
            detected_bands.append(
                {
                    "hdr_r": r,
                    "kind": table_kind,
                    "col_map": col_map,
                    "first_m_c": first_m_c,
                    "month_label": month_label,
                    "ref_date_label": ref_date_label,
                }
            )
            skip_until_r = r + 2

    if not detected_bands:
        return None

    # 상단 요약표에 수식 누락으로 미매핑된 컬럼이 있으면 하단 상세표(병합 헤더가 명확한 표) 컬럼 위치로 교차 보정
    detail_band = next((b for b in detected_bands if b["kind"] == "detail"), None)
    if detail_band:
        for b in detected_bands:
            if b["kind"] == "summary":
                for m_key in (
                    "plan_qty", "plan_amt", "target_qty", "target_amt",
                    "act_qty", "act_amt", "diff_qty", "diff_amt",
                    "rate_target", "rate_act", "rate_diff",
                ):
                    if m_key not in b["col_map"] and m_key in detail_band["col_map"]:
                        b["col_map"][m_key] = detail_band["col_map"][m_key]

    header_rows_set = set()
    for b in detected_bands:
        header_rows_set.add(b["hdr_r"])
        header_rows_set.add(b["hdr_r"] + 1)

    sample_qtys, sample_amts, sample_pcts = [], [], []
    for idx, b in enumerate(detected_bands):
        start_r = b["hdr_r"] + 2
        end_r = (detected_bands[idx + 1]["hdr_r"] - 1) if (idx + 1 < len(detected_bands)) else max_r
        cmap = b["col_map"]
        for r in range(start_r, end_r + 1):
            if r in header_rows_set:
                continue
            for qk in ("plan_qty", "target_qty", "act_qty"):
                ci = cmap.get(qk)
                v = ws.cell(r, ci).value if ci else None
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0:
                    sample_qtys.append(abs(float(v)))
            for ak in ("plan_amt", "target_amt", "act_amt"):
                ci = cmap.get(ak)
                v = ws.cell(r, ci).value if ci else None
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 0:
                    sample_amts.append(abs(float(v)))
            for pk in ("rate_target", "rate_act"):
                ci = cmap.get(pk)
                v = ws.cell(r, ci).value if ci else None
                if isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0:
                    sample_pcts.append(abs(float(v)))

    max_qty = max(sample_qtys) if sample_qtys else 0
    qty_is_k_unit = (0 < max_qty < 50_000)

    max_amt = max(sample_amts) if sample_amts else 0
    if max_amt >= 50_000_000:
        amt_divisor = 100_000_000.0
    elif max_amt >= 50_000:
        amt_divisor = 100_000.0
    elif max_amt >= 5_000:
        amt_divisor = 100.0
    else:
        amt_divisor = 1.0

    med_pct = statistics.median(sample_pcts) if sample_pcts else 0.8
    pct_is_ratio = med_pct <= 5.0

    def read_qty_pieces(r, c, allow_none=False):
        if not c:
            return None if allow_none else 0
        cell = ws.cell(r, c)
        val = cell.value
        if val is None:
            return None if allow_none else 0
        if isinstance(val, str):
            num = extract_num(val)
            if num is None:
                return None if allow_none else 0
            s_up = val.upper()
            if "K" in s_up or "천" in s_up:
                return int(round(num * 1000.0))
            return int(round(num * (1000.0 if qty_is_k_unit else 1.0)))
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            fmt = str(cell.number_format or "")
            if re.search(r"[0#],\s*\"?K", fmt, re.IGNORECASE) or fmt.rstrip('"').endswith(","):
                return int(round(float(val)))
            return int(round(float(val) * (1000.0 if qty_is_k_unit else 1.0)))
        return 0

    def read_amt_ok(r, c, allow_none=False):
        if not c:
            return None if allow_none else 0.0
        cell = ws.cell(r, c)
        val = cell.value
        if val is None:
            return None if allow_none else 0.0
        if isinstance(val, str):
            num = extract_num(val)
            if num is None:
                return None if allow_none else 0.0
            return round(num, 2) if "억" in val else round(num / amt_divisor, 2)
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            return round(float(val) / amt_divisor, 2)
        return 0.0

    def read_pct(r, c, allow_none=False):
        if not c:
            return None if allow_none else 0.0
        cell = ws.cell(r, c)
        val = cell.value
        if val is None:
            return None if allow_none else 0.0
        if isinstance(val, str):
            num = extract_num(val)
            return round(num, 1) if num is not None else (None if allow_none else 0.0)
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            fmt = str(cell.number_format or "")
            if "%" in fmt or pct_is_ratio:
                return round(float(val) * 100.0, 1)
            return round(float(val), 1)
        return 0.0

    def parse_row_metrics(r, cmap):
        p_q = read_qty_pieces(r, cmap.get("plan_qty"))
        p_a = read_amt_ok(r, cmap.get("plan_amt"))
        t_q = read_qty_pieces(r, cmap.get("target_qty"))
        t_a = read_amt_ok(r, cmap.get("target_amt"))
        a_q = read_qty_pieces(r, cmap.get("act_qty"))
        a_a = read_amt_ok(r, cmap.get("act_amt"))

        d_q = read_qty_pieces(r, cmap.get("diff_qty"), allow_none=True)
        if d_q is None:
            d_q = a_q - t_q
        d_a = read_amt_ok(r, cmap.get("diff_amt"), allow_none=True)
        if d_a is None:
            d_a = round(a_a - t_a, 2)
        r_t = read_pct(r, cmap.get("rate_target"), allow_none=True)
        if r_t is None:
            r_t = round(t_q / p_q * 100.0, 1) if p_q > 0 else 0.0
        r_a = read_pct(r, cmap.get("rate_act"), allow_none=True)
        if r_a is None:
            r_a = round(a_q / p_q * 100.0, 1) if p_q > 0 else 0.0
        r_d = read_pct(r, cmap.get("rate_diff"), allow_none=True)
        if r_d is None:
            r_d = round(r_a - r_t, 1)

        return {
            "plan_qty": p_q,
            "plan_amt": p_a,
            "target_qty": t_q,
            "target_amt": t_a,
            "act_qty": a_q,
            "act_amt": a_a,
            "diff_qty": d_q,
            "diff_amt": d_a,
            "rate_target": r_t,
            "rate_act": r_a,
            "rate_diff": r_d,
        }

    month_lbl, ref_date_lbl = "", ""
    summary_rows, detail_rows = [], []

    for idx, b in enumerate(detected_bands):
        month_lbl = month_lbl or b["month_label"]
        ref_date_lbl = ref_date_lbl or b["ref_date_label"]
        cmap = b["col_map"]
        first_m_c = b["first_m_c"]
        start_r = b["hdr_r"] + 2
        end_r = (detected_bands[idx + 1]["hdr_r"] - 1) if (idx + 1 < len(detected_bands)) else ws.max_row

        blank_streak = 0
        if b["kind"] == "summary":
            g_col = cmap.get("gubun", 1)
            for r in range(start_r, end_r + 1):
                lbl = ""
                for c_chk in range(g_col, first_m_c):
                    if c_chk in hidden_cols:
                        continue
                    v_chk = str(ws.cell(r, c_chk).value or merged_cell(ws, merged_map, r, c_chk) or "").strip()
                    if v_chk and norm_txt(v_chk) != "구분":
                        lbl = v_chk
                        break
                if not lbl:
                    blank_streak += 1
                    if blank_streak >= 3:
                        break
                    continue
                blank_streak = 0
                item = {"구분": lbl}
                item.update(parse_row_metrics(r, cmap))
                summary_rows.append(item)
        else:
            curr_gubun = ""
            g_col = cmap.get("gubun", 1)
            p_col = cmap.get("project")
            m_col = cmap.get("model")

            for r in range(start_r, end_r + 1):
                raw_g = str(ws.cell(r, g_col).value or merged_cell(ws, merged_map, r, g_col) or "").strip()
                proj = str(ws.cell(r, p_col).value or "").strip() if p_col else ""
                mdl = str(ws.cell(r, m_col).value or "").strip() if m_col else ""

                if not proj and not mdl:
                    blank_streak += 1
                    if blank_streak >= 5:
                        break
                    continue
                blank_streak = 0

                if raw_g and norm_txt(raw_g) != "구분":
                    curr_gubun = raw_g

                concat_chk = norm_txt(f"{raw_g}{proj}{mdl}")
                if any(k in concat_chk for k in ("합계", "총계", "소계", "TOTAL", "매출예상")):
                    continue

                item = {
                    "구분": curr_gubun or "삼성향",
                    "과제명": proj or mdl,
                    "모델명": mdl or proj,
                }
                item.update(parse_row_metrics(r, cmap))
                detail_rows.append(item)

    if detail_rows and not summary_rows:
        tot_pq = sum(x["plan_qty"] for x in detail_rows)
        tot_pa = round(sum(x["plan_amt"] for x in detail_rows), 2)
        tot_tq = sum(x["target_qty"] for x in detail_rows)
        tot_ta = round(sum(x["target_amt"] for x in detail_rows), 2)
        tot_aq = sum(x["act_qty"] for x in detail_rows)
        tot_aa = round(sum(x["act_amt"] for x in detail_rows), 2)
        rt_calc = round(tot_tq / tot_pq * 100.0, 1) if tot_pq > 0 else 0.0
        ra_calc = round(tot_aq / tot_pq * 100.0, 1) if tot_pq > 0 else 0.0
        summary_rows.append(
            {
                "구분": f"{month_lbl or '당월'} 합계",
                "plan_qty": tot_pq,
                "plan_amt": tot_pa,
                "target_qty": tot_tq,
                "target_amt": tot_ta,
                "act_qty": tot_aq,
                "act_amt": tot_aa,
                "diff_qty": tot_aq - tot_tq,
                "diff_amt": round(tot_aa - tot_ta, 2),
                "rate_target": rt_calc,
                "rate_act": ra_calc,
                "rate_diff": round(ra_calc - rt_calc, 1),
            }
        )

    if not summary_rows and not detail_rows:
        return None

    return {
        "sheet_name": ws.title,
        "month_label": month_lbl or "당월",
        "ref_date_label": ref_date_lbl,
        "summary_rows": summary_rows,
        "detail_rows": detail_rows,
    }


# 4-4. 통합 출하·매출 파일 — (B) '11월 진도현황' 시트 파서 (하단 Total/요약행 이중합산 및 단가 전이 차단)
def parse_daily_progress_sheet(ws, target_date_str, mapping_data, fallback_ex_rate=1330.0):
    merged_map = build_merged_lookup(ws)
    t_y, t_m, t_d = parse_target_ymd(target_date_str)
    target_date_obj = datetime.date(t_y, t_m, t_d)
    target_iso_key = get_iso_week_key(target_date_obj)

    sheet_year = t_y
    detected_ex_rate = None
    for r in range(1, 5):
        for c in range(1, min(20, ws.max_column + 1)):
            raw_v = str(ws.cell(r, c).value or "").strip()
            if not raw_v:
                continue
            m_yr = re.search(r"(\d{2,4})\s*년", raw_v)
            if m_yr:
                yr_n = int(m_yr.group(1))
                sheet_year = (2000 + yr_n) if yr_n < 100 else yr_n
            if "환율" in raw_v:
                m_ex = re.search(r"환율[^\d]*([\d,]+(?:\.\d+)?)", raw_v)
                if m_ex:
                    detected_ex_rate = float(m_ex.group(1).replace(",", ""))
                else:
                    nxt = extract_num(safe_cell(ws, r, c + 1))
                    if nxt and nxt > 500:
                        detected_ex_rate = nxt

    ex_rate = detected_ex_rate or fallback_ex_rate

    header_row = 3
    for r in range(1, 8):
        row_join = "".join(norm_txt(safe_cell(ws, r, c)) for c in range(1, min(20, ws.max_column + 1)))
        if ("SECCODE" in row_join or "SEC-CODE" in row_join or "매출(KRW)" in row_join) and "합계" in row_join:
            header_row = r
            break

    model_cand_cols = []
    sec_col, price_col, gubun_col, usd_col, krw_col, sum_col = None, None, None, None, None, None

    for c in range(1, min(20, ws.max_column + 1)):
        h_txt = norm_txt(safe_cell(ws, header_row, c))
        if not h_txt:
            continue
        if h_txt in ("모델", "모델명", "MODEL"):
            model_cand_cols.append(c)
        elif "SEC" in h_txt or "GH" in h_txt or "품번" in h_txt:
            sec_col = sec_col or c
        elif re.fullmatch(r"[1-4]Q", h_txt) or any(k in h_txt for k in ("단가", "PRICE")):
            price_col = price_col or c
        elif h_txt == "구분":
            gubun_col = c
        elif "USD" in h_txt:
            usd_col = usd_col or c
        elif "KRW" in h_txt or "원화" in h_txt:
            krw_col = krw_col or c
        elif h_txt in ("합계", "누계"):
            sum_col = sum_col or c

    if len(model_cand_cols) >= 2:
        cust_col = model_cand_cols[0]
        model_col = model_cand_cols[-1]
    elif len(model_cand_cols) == 1:
        cust_col = max(1, model_cand_cols[0] - 1)
        model_col = model_cand_cols[0]
    else:
        cust_col, model_col = 2, 3

    sec_col = sec_col or 5
    if price_col is None:
        for c_chk in range(sec_col + 1, (gubun_col or 9)):
            sample_p = [
                extract_num(safe_cell(ws, r_chk, c_chk))
                for r_chk in range(header_row + 1, min(header_row + 40, ws.max_row + 1))
            ]
            valid_p = [x for x in sample_p if x is not None and 0 < x < 500]
            if valid_p:
                price_col = c_chk
                break
    price_col = price_col or 7
    gubun_col = gubun_col or 8
    sum_col = sum_col or 12

    date_cols = []
    run_yr = sheet_year
    prev_m = None
    for c in range(sum_col + 1, ws.max_column + 1):
        dt_obj, run_yr, prev_m = parse_excel_date_header(safe_cell(ws, header_row, c), run_yr, prev_m)
        if dt_obj is not None:
            date_cols.append((c, dt_obj, dt_obj.strftime("%Y-%m-%d"), get_iso_week_key(dt_obj), get_iso_week_label(dt_obj)))

    if not date_cols:
        return None

    if detected_ex_rate is None and usd_col and krw_col:
        for r in range(header_row + 1, ws.max_row + 1):
            u_v = extract_num(safe_cell(ws, r, usd_col))
            k_v = extract_num(safe_cell(ws, r, krw_col))
            if u_v and k_v and u_v > 1000 and k_v > 1_000_000:
                ratio = k_v / u_v
                if 900 <= ratio <= 2000:
                    ex_rate = round(ratio, 2)
                    detected_ex_rate = ex_rate
                    break

    plan_by_date = {ymd: {"qty": 0, "rev": 0.0, "items": []} for _, _, ymd, _, _ in date_cols}
    actual_by_date = {ymd: {"qty": 0, "rev": 0.0, "items": []} for _, _, ymd, _, _ in date_cols}
    bottom_daily_act_rev = {}
    bottom_daily_plan_rev = {}

    weekly_ship_plan_iso, weekly_rev_plan_iso = {}, {}
    weekly_ship_act_iso, weekly_rev_act_iso = {}, {}
    target_month_w_ship_p, target_month_w_rev_p = {}, {}
    target_month_w_ship_a, target_month_w_rev_a = {}, {}

    monthly_ship_p, monthly_rev_p = 0, 0.0
    monthly_ship_a, monthly_rev_a = 0, 0.0

    plan_model_period = {"daily": {}, "weekly": {}, "monthly": {}}
    actual_model_period = {"daily": {}, "weekly": {}, "monthly": {}}

    reached_total_footer = False

    for r in range(header_row + 1, ws.max_row + 1):
        raw_b = str(merged_cell(ws, merged_map, r, cust_col) or "").strip()
        raw_c = str(merged_cell(ws, merged_map, r, model_col) or "").strip()
        raw_sec = str(merged_cell(ws, merged_map, r, sec_col) or "").strip()
        raw_gbn = norm_txt(ws.cell(r, gubun_col).value)
        raw_sum_val = ws.cell(r, sum_col).value
        raw_sum_lbl = norm_txt(raw_sum_val) if isinstance(raw_sum_val, str) else ""

        # 하단 요약행('출하실적 | 금액' 또는 '출하계획 | 금액') 추출 후 모델 집계에서는 완전 제외!
        if raw_sum_lbl in ("수량", "금액"):
            if raw_sum_lbl == "금액" and raw_gbn in ("출하실적", "출하계획"):
                target_dict = bottom_daily_act_rev if raw_gbn == "출하실적" else bottom_daily_plan_rev
                for c_idx, _, ymd, _, _ in date_cols:
                    v_amt = extract_num(ws.cell(r, c_idx).value)
                    if v_amt is not None and v_amt > 0:
                        rev_ok = (v_amt / 100_000_000.0) if v_amt >= 10_000_000 else v_amt
                        target_dict[ymd] = round(rev_ok, 4)
            continue

        row_id_txt = norm_txt(f"{raw_b}{raw_c}{raw_sec}")
        if any(k in row_id_txt for k in ("TOTAL", "G.TOTAL", "합계", "총계", "소계")):
            reached_total_footer = True
            continue

        # 삼성향 Total / G.Total 이후 행이거나 모델명·품번이 모두 비어있는 하단 행은 절대 개별 모델로 읽지 않음!
        if reached_total_footer or (not raw_c and not raw_sec):
            continue

        if raw_gbn not in ("출하계획", "출하실적"):
            continue

        curr_gh = "" if raw_sec.upper() in ("#N/A", "N/A", "NONE") else raw_sec
        mdl_name = raw_c or curr_gh
        if not mdl_name:
            continue

        # 행 자체의 단가만 사용 (이전 모델의 단가를 $0 행으로 전이시키지 않음!)
        p_num = extract_num(merged_cell(ws, merged_map, r, price_col))
        row_usd_price = p_num if (p_num is not None and p_num > 0) else 0.0

        unit_krw = row_usd_price * ex_rate if row_usd_price > 0 else 0.0
        if unit_krw == 0 and krw_col and sum_col:
            tot_k = extract_num(ws.cell(r, krw_col).value)
            tot_q = extract_num(ws.cell(r, sum_col).value)
            if tot_k and tot_q and tot_q > 0:
                unit_krw = tot_k / tot_q

        gh_clean = clean_key(curr_gh)
        gh_suf = get_gh_suffix(curr_gh)

        if unit_krw > 0:
            if gh_clean:
                mapping_data["gh_map"][gh_clean] = unit_krw
                mapping_data["suffix_map"][gh_suf] = unit_krw
            if mdl_name:
                mapping_data["model_map"][clean_key(mdl_name)] = unit_krw
        if gh_clean and mdl_name and mdl_name != "본사샘플":
            mapping_data["gh_to_model"][gh_clean] = mdl_name
            mapping_data["gh_to_model"][gh_suf] = mdl_name

        for c_idx, dt_obj, ymd, w_key, w_lbl in date_cols:
            q_num = extract_num(ws.cell(r, c_idx).value)
            q_val = int(round(q_num)) if q_num else 0
            if q_val == 0:
                continue

            rev_ok = (q_val * unit_krw) / 100_000_000.0
            is_target_m = (dt_obj.year == t_y and dt_obj.month == t_m)

            if raw_gbn == "출하계획":
                plan_by_date[ymd]["qty"] += q_val
                plan_by_date[ymd]["rev"] += rev_ok
                plan_by_date[ymd]["items"].append(
                    {
                        "model": mdl_name,
                        "gh_code": curr_gh or "-",
                        "site": raw_b or "삼성향",
                        "qty": q_val,
                        "price": unit_krw,
                        "rev_ok": rev_ok,
                    }
                )
                weekly_ship_plan_iso[w_key] = weekly_ship_plan_iso.get(w_key, 0) + q_val
                weekly_rev_plan_iso[w_key] = weekly_rev_plan_iso.get(w_key, 0.0) + rev_ok

                if is_target_m:
                    target_month_w_ship_p[w_lbl] = target_month_w_ship_p.get(w_lbl, 0) + q_val
                    target_month_w_rev_p[w_lbl] = target_month_w_rev_p.get(w_lbl, 0.0) + rev_ok
                    monthly_ship_p += q_val
                    monthly_rev_p += rev_ok

                    mm = plan_model_period["monthly"].setdefault(mdl_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    mm["Plan_Qty"] += q_val
                    mm["Plan_Rev"] += rev_ok

                if w_key == target_iso_key:
                    mw = plan_model_period["weekly"].setdefault(mdl_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    mw["Plan_Qty"] += q_val
                    mw["Plan_Rev"] += rev_ok

                if dt_obj == target_date_obj:
                    md = plan_model_period["daily"].setdefault(mdl_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    md["Plan_Qty"] += q_val
                    md["Plan_Rev"] += rev_ok

            elif raw_gbn == "출하실적":
                actual_by_date[ymd]["qty"] += q_val
                actual_by_date[ymd]["rev"] += rev_ok
                actual_by_date[ymd]["items"].append(
                    {
                        "model": mdl_name,
                        "gh_code": curr_gh or "-",
                        "site": raw_b or "삼성향",
                        "qty": q_val,
                        "price": unit_krw,
                        "rev_ok": rev_ok,
                    }
                )

                if dt_obj <= target_date_obj:
                    weekly_ship_act_iso[w_key] = weekly_ship_act_iso.get(w_key, 0) + q_val
                    weekly_rev_act_iso[w_key] = weekly_rev_act_iso.get(w_key, 0.0) + rev_ok

                    if is_target_m:
                        target_month_w_ship_a[w_lbl] = target_month_w_ship_a.get(w_lbl, 0) + q_val
                        target_month_w_rev_a[w_lbl] = target_month_w_rev_a.get(w_lbl, 0.0) + rev_ok
                        monthly_ship_a += q_val
                        monthly_rev_a += rev_ok

                        mm_a = actual_model_period["monthly"].setdefault(mdl_name, {"Actual": 0, "Rev_Act": 0.0})
                        mm_a["Actual"] += q_val
                        mm_a["Rev_Act"] += rev_ok

                    if w_key == target_iso_key:
                        mw_a = actual_model_period["weekly"].setdefault(mdl_name, {"Actual": 0, "Rev_Act": 0.0})
                        mw_a["Actual"] += q_val
                        mw_a["Rev_Act"] += rev_ok

                if dt_obj == target_date_obj:
                    md_a = actual_model_period["daily"].setdefault(mdl_name, {"Actual": 0, "Rev_Act": 0.0})
                    md_a["Actual"] += q_val
                    md_a["Rev_Act"] += rev_ok

    # 하단 노란색 요약행(출하실적 금액 / 출하계획 금액)이 존재하면 일별 공식 매출액 반영
    for ymd, official_rev in bottom_daily_act_rev.items():
        if ymd in actual_by_date and actual_by_date[ymd]["qty"] > 0:
            actual_by_date[ymd]["rev"] = official_rev

    for ymd, official_p_rev in bottom_daily_plan_rev.items():
        if ymd in plan_by_date and plan_by_date[ymd]["qty"] > 0:
            plan_by_date[ymd]["rev"] = official_p_rev

    if bottom_daily_act_rev:
        weekly_rev_act_iso.clear()
        target_month_w_rev_a.clear()
        monthly_rev_a = 0.0
        for _, dt_obj, ymd, w_key, w_lbl in date_cols:
            if dt_obj <= target_date_obj:
                r_v = actual_by_date[ymd]["rev"]
                weekly_rev_act_iso[w_key] = weekly_rev_act_iso.get(w_key, 0.0) + r_v
                if dt_obj.year == t_y and dt_obj.month == t_m:
                    target_month_w_rev_a[w_lbl] = target_month_w_rev_a.get(w_lbl, 0.0) + r_v
                    monthly_rev_a += r_v

    if bottom_daily_plan_rev:
        weekly_rev_plan_iso.clear()
        target_month_w_rev_p.clear()
        monthly_rev_p = 0.0
        for _, dt_obj, ymd, w_key, w_lbl in date_cols:
            r_pv = plan_by_date[ymd]["rev"]
            weekly_rev_plan_iso[w_key] = weekly_rev_plan_iso.get(w_key, 0.0) + r_pv
            if dt_obj.year == t_y and dt_obj.month == t_m:
                target_month_w_rev_p[w_lbl] = target_month_w_rev_p.get(w_lbl, 0.0) + r_pv
                monthly_rev_p += r_pv

    target_ymd = target_date_obj.strftime("%Y-%m-%d")
    daily_p_info = plan_by_date.get(target_ymd, {"qty": 0, "rev": 0.0})
    daily_a_info = actual_by_date.get(target_ymd, {"qty": 0, "rev": 0.0})

    loaded_yms = sorted(list(set(dt.strftime("%Y-%m") for _, dt, _, _, _ in date_cols)))

    weekly_ship_p_chart, weekly_rev_p_chart = {}, {}
    for d in range(1, calendar.monthrange(t_y, t_m)[1] + 1):
        d_tmp = datetime.date(t_y, t_m, d)
        wl = get_iso_week_label(d_tmp)
        wk = get_iso_week_key(d_tmp)
        if wl not in weekly_ship_p_chart:
            weekly_ship_p_chart[wl] = weekly_ship_plan_iso.get(wk, target_month_w_ship_p.get(wl, 0))
            weekly_rev_p_chart[wl] = round(weekly_rev_plan_iso.get(wk, target_month_w_rev_p.get(wl, 0.0)), 2)

    return {
        "sheet_name": ws.title,
        "detected_ex_rate": detected_ex_rate,
        "loaded_yms": loaded_yms,
        "plan_by_date": plan_by_date,
        "actual_by_date": actual_by_date,
        "daily_ship_p": daily_p_info["qty"],
        "daily_rev_p": round(daily_p_info["rev"], 2),
        "daily_ship_a": daily_a_info["qty"],
        "daily_rev_a": round(daily_a_info["rev"], 2),
        "cur_week_ship_p": weekly_ship_plan_iso.get(target_iso_key, 0),
        "cur_week_rev_p": round(weekly_rev_plan_iso.get(target_iso_key, 0.0), 2),
        "cur_week_ship_a": weekly_ship_act_iso.get(target_iso_key, 0),
        "cur_week_rev_a": round(weekly_rev_act_iso.get(target_iso_key, 0.0), 2),
        "weekly_ship_plan_chart": weekly_ship_p_chart,
        "weekly_rev_plan_chart": weekly_rev_p_chart,
        "weekly_ship_act_chart": target_month_w_ship_a,
        "weekly_rev_act_chart": {k: round(v, 2) for k, v in target_month_w_rev_a.items()},
        "monthly_ship_p": monthly_ship_p,
        "monthly_rev_p": round(monthly_rev_p, 2),
        "monthly_ship_a": monthly_ship_a,
        "monthly_rev_a": round(monthly_rev_a, 2),
        "plan_model_period": plan_model_period,
        "actual_model_period": actual_model_period,
    }


# 4-5. 통합 출하·매출 파일 — (C) '매출 참고' (Global DP 현황 / RTF) 시트 파서
def parse_sales_ref_fcst_sheet(ws, target_date_str, mapping_data, exchange_rate=1330.0):
    t_y, t_m, t_d = parse_target_ymd(target_date_str)
    target_date_obj = datetime.date(t_y, t_m, t_d)
    target_iso_key = get_iso_week_key(target_date_obj)
    target_month_key = f"{t_y:04d}-{t_m:02d}"

    # 헤더 행 탐색 (Plant / Material / Project 또는 기존 Main 시트 4행)
    header_row = 4
    plant_col, mat_col, proj_col, gubun_j_col, gubun_aj_col, mibanyoung_col = None, None, None, None, None, None
    price_col = 1

    for r in range(1, 8):
        row_txts = [norm_txt(safe_cell(ws, r, c)) for c in range(1, min(45, ws.max_column + 1))]
        joined = "".join(row_txts)
        if ("PLANT" in joined and "MATERIAL" in joined) or ("SITE" in joined and "GH" in joined):
            header_row = r
            break

    gubun_cols = []
    for c in range(1, min(50, ws.max_column + 1)):
        ht = norm_txt(safe_cell(ws, header_row, c))
        if not ht:
            continue
        if ht in ("PLANT", "SITE", "거점", "출하거점"):
            plant_col = plant_col or c
        elif ht in ("MATERIAL", "GHCODE", "GH-CODE", "SEC-CODE", "품번"):
            mat_col = mat_col or c
        elif ht in ("PROJECT", "MODEL", "모델", "모델명", "과제명"):
            proj_col = proj_col or c
        elif ht in ("단가", "PRICE", "USD"):
            price_col = c
        elif ht == "구분":
            gubun_cols.append(c)
        elif "미반영" in ht:
            mibanyoung_col = c

    # '매출 참고' 양식(D=4, E=5, G=7, J=10, AJ=36, AK=37) vs 기존 'RTF Main'(C=3, D=4, F=6, I=9) 자동 대응
    plant_col = plant_col or 4
    mat_col = mat_col or 5
    proj_col = proj_col or 7
    if gubun_cols:
        gubun_j_col = gubun_cols[0]
        if len(gubun_cols) >= 2:
            gubun_aj_col = gubun_cols[-1]
    else:
        gubun_j_col = 10

    date_scan_start = (mibanyoung_col + 1) if mibanyoung_col else ((gubun_aj_col + 1) if gubun_aj_col else (gubun_j_col + 1))

    raw_date_cols = []
    run_yr = t_y
    prev_m = None
    for c in range(date_scan_start, ws.max_column + 1):
        dt_obj, run_yr, prev_m = parse_excel_date_header(safe_cell(ws, header_row, c), run_yr, prev_m)
        if dt_obj is not None:
            raw_date_cols.append((c, dt_obj))

    if not raw_date_cols:
        return None

    span_multi_year = len(set(d.year for _, d in raw_date_cols)) > 1
    enriched_date_cols = []
    fcst_date_cols = []
    seen_lbls = {}

    for c, d_obj in raw_date_cols:
        base_lbl = d_obj.strftime("%y/%m/%d") if span_multi_year else d_obj.strftime("%m/%d")
        if base_lbl in seen_lbls:
            seen_lbls[base_lbl] += 1
            d_lbl = f"{base_lbl}_{seen_lbls[base_lbl]}"
        else:
            seen_lbls[base_lbl] = 1
            d_lbl = base_lbl
        fcst_date_cols.append(d_lbl)
        enriched_date_cols.append(
            (
                c,
                d_obj,
                get_iso_week_key(d_obj),
                get_iso_week_label(d_obj),
                f"{d_obj.year:04d}-{d_obj.month:02d}",
                d_lbl,
                d_obj.strftime("%Y-%m-%d"),
            )
        )

    gh_map = mapping_data.get("gh_map", {})
    suffix_map = mapping_data.get("suffix_map", {})
    gh_to_model = mapping_data.get("gh_to_model", {})
    model_map = mapping_data.get("model_map", {})

    known_krw = [
        (p if p > 100 else p * exchange_rate)
        for p in list(gh_map.values()) + list(suffix_map.values()) + list(model_map.values())
        if p > 0
    ]
    avg_price_krw = (sum(known_krw) / len(known_krw)) if known_krw else (5.5 * exchange_rate)

    daily_ship_p, daily_rev_p = 0, 0.0
    weekly_ship_by_isokey, weekly_rev_by_isokey = {}, {}
    target_month_weekly_ship, target_month_weekly_rev = {}, {}
    monthly_ship_by_ym, monthly_rev_by_ym = {}, {}
    plan_by_model_period = {"daily": {}, "weekly": {}, "monthly": {}}
    fcst_rows = []

    curr_site, curr_gh, curr_proj, curr_row_usd = "", "", "", 0.0

    for r in range(header_row + 1, ws.max_row + 1):
        v_site = str(safe_cell(ws, r, plant_col) or "").strip()
        v_gh = str(safe_cell(ws, r, mat_col) or "").strip()
        v_proj = str(safe_cell(ws, r, proj_col) or "").strip()
        v_p_a = extract_num(safe_cell(ws, r, price_col))

        if v_site:
            curr_site = v_site
        if v_gh:
            curr_gh = v_gh
        if v_proj:
            curr_proj = v_proj
        if v_p_a is not None and v_p_a > 0:
            curr_row_usd = v_p_a

        gbn_j = norm_txt(safe_cell(ws, r, gubun_j_col))
        gbn_aj = norm_txt(safe_cell(ws, r, gubun_aj_col)) if gubun_aj_col else ""

        # '매출 참고' 시트의 공급(계획) 행 필터링
        is_supply_plan_row = (gbn_j == "공급") or (not gbn_j and gbn_aj == "계획")
        if not is_supply_plan_row:
            continue
        if gbn_aj and gbn_aj not in ("계획", ""):
            continue
        if not curr_site and not curr_gh:
            continue
        if any(k in norm_txt(f"{v_site}{v_gh}{v_proj}{curr_site}{curr_gh}") for k in ("합계", "소계", "총계", "TOTAL", "G.TOTAL")):
            curr_site, curr_gh, curr_proj = "", "", ""
            continue

        gh_clean = clean_key(curr_gh)
        gh_suf = get_gh_suffix(curr_gh)
        model_name = resolve_model_name(curr_gh, curr_proj, gh_to_model)

        # 단가 결정: A열 USD단가 -> 진도현황/매핑표 단가 -> 평균단가
        if curr_row_usd > 0:
            price_krw = curr_row_usd if curr_row_usd > 100 else curr_row_usd * exchange_rate
        else:
            p_lookup = gh_map.get(gh_clean, 0.0) or suffix_map.get(gh_suf, 0.0) or model_map.get(clean_key(model_name), 0.0)
            if p_lookup > 0:
                price_krw = p_lookup if p_lookup > 100 else p_lookup * exchange_rate
            else:
                price_krw = avg_price_krw

        row_fcst = {
            "SITE": curr_site or "-",
            "GH_CODE": curr_gh or "-",
            "과제명": curr_proj or model_name,
            "모델명": model_name,
            "단가_원": price_krw,
            "_by_ymd": {},
        }
        row_fcst_total = 0

        for c_idx, d_obj, w_key, w_lbl, ym_key, d_lbl, ymd_str in enriched_date_cols:
            q_num = extract_num(safe_cell(ws, r, c_idx))
            q_val = int(round(q_num)) if q_num else 0

            row_fcst[d_lbl] = q_val
            if q_val > 0:
                row_fcst["_by_ymd"][ymd_str] = row_fcst["_by_ymd"].get(ymd_str, 0) + q_val
            row_fcst_total += q_val

            if q_val == 0:
                continue

            rev_val = (q_val * price_krw) / 100_000_000.0
            weekly_ship_by_isokey[w_key] = weekly_ship_by_isokey.get(w_key, 0) + q_val
            weekly_rev_by_isokey[w_key] = weekly_rev_by_isokey.get(w_key, 0.0) + rev_val

            if ym_key == target_month_key:
                target_month_weekly_ship[w_lbl] = target_month_weekly_ship.get(w_lbl, 0) + q_val
                target_month_weekly_rev[w_lbl] = target_month_weekly_rev.get(w_lbl, 0.0) + rev_val
                mm = plan_by_model_period["monthly"].setdefault(model_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                mm["Plan_Qty"] += q_val
                mm["Plan_Rev"] += rev_val

            monthly_ship_by_ym[ym_key] = monthly_ship_by_ym.get(ym_key, 0) + q_val
            monthly_rev_by_ym[ym_key] = monthly_rev_by_ym.get(ym_key, 0.0) + rev_val

            if w_key == target_iso_key:
                mw = plan_by_model_period["weekly"].setdefault(model_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                mw["Plan_Qty"] += q_val
                mw["Plan_Rev"] += rev_val

            if d_obj == target_date_obj:
                daily_ship_p += q_val
                daily_rev_p += rev_val
                md = plan_by_model_period["daily"].setdefault(model_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                md["Plan_Qty"] += q_val
                md["Plan_Rev"] += rev_val

        row_fcst["합계"] = row_fcst_total
        if row_fcst_total > 0:
            fcst_rows.append(row_fcst)

    weekly_ship_plan_for_chart, weekly_rev_plan_for_chart = {}, {}
    for d in range(1, calendar.monthrange(t_y, t_m)[1] + 1):
        d_tmp = datetime.date(t_y, t_m, d)
        wl = get_iso_week_label(d_tmp)
        wk = get_iso_week_key(d_tmp)
        if wl not in weekly_ship_plan_for_chart:
            weekly_ship_plan_for_chart[wl] = weekly_ship_by_isokey.get(wk, target_month_weekly_ship.get(wl, 0))
            weekly_rev_plan_for_chart[wl] = round(
                weekly_rev_by_isokey.get(wk, target_month_weekly_rev.get(wl, 0.0)), 2
            )

    return {
        "sheet_name": ws.title,
        "daily_ship_p": daily_ship_p,
        "daily_rev_p": round(daily_rev_p, 2),
        "cur_week_ship_p": weekly_ship_by_isokey.get(target_iso_key, 0),
        "cur_week_rev_p": round(weekly_rev_by_isokey.get(target_iso_key, 0.0), 2),
        "cur_month_ship_p": monthly_ship_by_ym.get(target_month_key, 0),
        "cur_month_rev_p": round(monthly_rev_by_ym.get(target_month_key, 0.0), 2),
        "weekly_ship_plan": weekly_ship_plan_for_chart,
        "weekly_rev_plan": weekly_rev_plan_for_chart,
        "plan_by_model_period": plan_by_model_period,
        "fcst_rows": fcst_rows,
        "fcst_date_cols": fcst_date_cols,
    }


# 4-6. 통합 출하·매출 워크북 일괄 파서 (하나의 엑셀에서 3개 시트 동시 추출)
def parse_integrated_sales_workbook(file_src, target_date_str, mapping_data, ex_rate_input=1330.0):
    wb = openpyxl.load_workbook(file_src, data_only=True)
    sheet_names = wb.sheetnames

    t_y, t_m, _ = parse_target_ymd(target_date_str)

    # 1) 시트 역할 자동 매칭
    summary_ws = None
    progress_ws = None
    ref_ws = None

    # (a) 월별요약 시트 탐색
    for s in sheet_names:
        sn = norm_txt(s)
        if "월별요약" in sn or ("월별" in sn and "요약" in sn):
            summary_ws = wb[s]
            break
    if summary_ws is None:
        for s in sheet_names:
            if "요약" in norm_txt(s):
                summary_ws = wb[s]
                break

    # (b) 진도현황 시트 탐색 (조회 기준월과 일치하는 시트 우선, 없으면 '진도현황' 포함 시트)
    progress_candidates = [s for s in sheet_names if "진도" in norm_txt(s)]
    for s in progress_candidates:
        if f"{t_m}월" in s.replace(" ", ""):
            progress_ws = wb[s]
            break
    if progress_ws is None and progress_candidates:
        progress_ws = wb[progress_candidates[0]]

    # (c) 매출 참고 (Global DP / RTF Main) 시트 탐색
    for s in sheet_names:
        sn = norm_txt(s)
        if "매출참고" in sn or "GLOBALDP" in sn or sn == "MAIN" or "RTF" in sn:
            ref_ws = wb[s]
            break

    # 이름만으로 매칭되지 않은 시트가 있으면 구조 스캔으로 보완
    for s in sheet_names:
        ws = wb[s]
        if summary_ws is None and ws not in (progress_ws, ref_ws):
            chk = parse_monthly_summary_sheet(ws)
            if chk and chk["detail_rows"]:
                summary_ws = ws
        if progress_ws is None and ws not in (summary_ws, ref_ws):
            top_join = "".join(norm_txt(safe_cell(ws, r, c)) for r in range(1, 6) for c in range(1, 16))
            if "출하실적" in top_join or ("SECCODE" in top_join and "합계" in top_join):
                progress_ws = ws
        if ref_ws is None and ws not in (summary_ws, progress_ws):
            top_join = "".join(norm_txt(safe_cell(ws, r, c)) for r in range(1, 7) for c in range(1, 15))
            if ("PLANT" in top_join and "MATERIAL" in top_join) or "입고요청" in top_join:
                ref_ws = ws

    # 파싱 순서: [11월 진도현황] 먼저 실행하여 SEC-CODE -> 모델명/단가 사전 구축 -> [월별요약] -> [매출 참고]
    prog_res = parse_daily_progress_sheet(progress_ws, target_date_str, mapping_data, ex_rate_input) if progress_ws else None
    effective_ex = (prog_res["detected_ex_rate"] if (prog_res and prog_res.get("detected_ex_rate")) else ex_rate_input)

    sum_res = parse_monthly_summary_sheet(summary_ws) if summary_ws else None
    ref_res = parse_sales_ref_fcst_sheet(ref_ws, target_date_str, mapping_data, effective_ex) if ref_ws else None

    return {
        "all_sheets": sheet_names,
        "summary": sum_res,
        "progress": prog_res,
        "sales_ref": ref_res,
        "effective_ex_rate": effective_ex,
    }


def _build_merged_table(plan_dict, act_dict, key_col_name):
    all_keys = sorted(list(set(list(plan_dict.keys()) + list(act_dict.keys()))))
    rows = []
    for k in all_keys:
        p_info = plan_dict.get(k, {"Plan_Qty": 0, "Plan_Rev": 0.0})
        a_info = act_dict.get(k, {"Actual": 0, "Rev_Act": 0.0})
        p_qty = p_info["Plan_Qty"]
        p_rev = round(p_info["Plan_Rev"], 2)
        a_qty = a_info["Actual"]
        a_rev = round(a_info["Rev_Act"], 2)
        if p_qty == 0 and a_qty == 0 and p_rev == 0 and a_rev == 0:
            continue
        qty_rate = round((a_qty / p_qty * 100), 1) if p_qty > 0 else 0.0
        rev_rate = round((a_rev / p_rev * 100), 1) if p_rev > 0 else 0.0
        rows.append(
            {
                key_col_name: k,
                "출하계획(개)": p_qty,
                "출하계획(K)": format_k(p_qty),
                "출하실적(개)": a_qty,
                "출하실적(K)": format_k(a_qty),
                "수량차이(개)": a_qty - p_qty,
                "수량차이(K)": fmt_signed_k(a_qty - p_qty),
                "수량달성률(%)": qty_rate,
                "매출계획(억원)": p_rev,
                "매출실적(억원)": a_rev,
                "매출차이(억원)": round(a_rev - p_rev, 2),
                "매출달성률(%)": rev_rate,
            }
        )
    return rows


# 5. 전체 데이터 파이프라인 실행 함수
def run_dashboard_pipeline(target_date_input, ex_rate_input, uploaded_files):
    os.makedirs(STORE_DIR, exist_ok=True)
    for slot_key, up_file in uploaded_files.items():
        if up_file is not None:
            with open(CACHED_EXCEL_PATHS[slot_key], "wb") as f:
                f.write(up_file.getvalue())

    st.session_state["target_date"] = target_date_input
    st.session_state["ex_rate"] = float(ex_rate_input)

    t_y, t_m, t_d = parse_target_ymd(target_date_input)
    target_date_obj = datetime.date(t_y, t_m, t_d)
    target_iso_week = get_iso_week_label(target_date_obj)
    default_weeks = get_month_iso_weeks(t_y, t_m)

    st.session_state["cur_week_label"] = target_iso_week
    st.session_state["chart_data"]["weekly_weeks"] = default_weeks
    st.session_state["chart_data"]["monthly_months"] = [f"{t_m}월"]

    messages = []

    # 1. 선택 단가 매핑 반영
    if os.path.exists(CACHED_EXCEL_PATHS["map_rev"]):
        nmap = parse_mapping_excel(CACHED_EXCEL_PATHS["map_rev"])
        if nmap:
            st.session_state["mapping_data"].update(nmap)
            messages.append(("success", f"✅ 추가 단가표 반영 완료 (GH-CODE {len(nmap['gh_map'])}개)"))

    # 2. 생산 실적 반영 (현행 유지)
    if os.path.exists(CACHED_EXCEL_PATHS["prod"]):
        try:
            p = parse_production_excel(CACHED_EXCEL_PATHS["prod"], target_date_input)
            w_weeks = list(dict.fromkeys(p["weekly_labels"])) if p["weekly_labels"] else default_weeks
            st.session_state["cur_week_label"] = p["cur_week_label"]
            st.session_state["kpi_data"]["daily"]["prod_p"] = p["daily_prod_plan"]
            st.session_state["kpi_data"]["daily"]["prod_a"] = p["daily_prod_actual"]
            st.session_state["kpi_data"]["weekly"]["prod_p"] = p["cur_week_prod_p"]
            st.session_state["kpi_data"]["weekly"]["prod_a"] = p["cur_week_prod_a"]
            st.session_state["kpi_data"]["monthly"]["prod_p"] = p["monthly_plan"]
            st.session_state["kpi_data"]["monthly"]["prod_a"] = p["monthly_actual"]
            st.session_state["kpi_data"]["monthly"]["prod_sheet_total"] = p.get("monthly_sheet_total", 0)
            st.session_state["kpi_data"]["monthly"]["prod_cur_w_add"] = p.get("monthly_cur_w_add", 0)

            st.session_state["chart_data"]["daily_models"] = p["daily_models"]
            st.session_state["chart_data"]["daily_prod_p"] = p["daily_prod_p"]
            st.session_state["chart_data"]["daily_prod_a"] = p["daily_prod_a"]
            st.session_state["chart_data"]["weekly_weeks"] = w_weeks
            st.session_state["chart_data"]["weekly_prod_p"] = p["weekly_plan"]
            st.session_state["chart_data"]["weekly_prod_a"] = p["weekly_actual"]
            st.session_state["chart_data"]["monthly_months"] = p["monthly_months"]
            st.session_state["chart_data"]["monthly_prod_p"] = [p["monthly_plan"]]
            st.session_state["chart_data"]["monthly_prod_a"] = [p["monthly_actual"]]

            st.session_state["inventory_data"]["qty"] = p["inventory_qty"]
            st.session_state["inventory_data"]["days"] = p["inventory_days"]
            st.session_state["data_loaded"] = True
            messages.append(
                (
                    "success",
                    f"✅ 생산 실적 & 재고 반영 완료 (일간 실적: {format_k(p['daily_prod_actual'])} | 재고: {format_k(p['inventory_qty'])})",
                )
            )
        except Exception as e:
            messages.append(("error", f"생산 실적 파싱 오류: {e}"))

    w_labels = st.session_state["chart_data"].get("weekly_weeks", default_weeks)
    active_week_lbl = st.session_state.get("cur_week_label", target_iso_week)

    # 3. 통합 출하·매출 파일 반영 ('월별요약' + '진도현황' + '매출 참고' 일괄 처리)
    if os.path.exists(CACHED_EXCEL_PATHS["integrated"]):
        try:
            integ = parse_integrated_sales_workbook(
                CACHED_EXCEL_PATHS["integrated"],
                target_date_input,
                st.session_state["mapping_data"],
                float(ex_rate_input),
            )
            sum_res = integ["summary"]
            prog_res = integ["progress"]
            ref_res = integ["sales_ref"]
            st.session_state["detected_ex_rate"] = integ.get("effective_ex_rate")

            sheet_status = {
                "summary": sum_res["sheet_name"] if sum_res else None,
                "progress": prog_res["sheet_name"] if prog_res else None,
                "sales_ref": ref_res["sheet_name"] if ref_res else None,
                "all_sheets": integ["all_sheets"],
            }
            st.session_state["loaded_sheet_info"] = sheet_status

            plan_model_period = {"daily": {}, "weekly": {}, "monthly": {}}
            actual_model_period = {"daily": {}, "weekly": {}, "monthly": {}}

            # (3-A) '11월 진도현황' 반영 -> Daily & Weekly 계획/실적, 일자별 캘린더 실적/계획
            if prog_res:
                st.session_state["plan_by_date"] = prog_res["plan_by_date"]
                st.session_state["actual_by_date"] = prog_res["actual_by_date"]
                st.session_state["actual_months_loaded"] = prog_res["loaded_yms"]

                st.session_state["kpi_data"]["daily"]["ship_p"] = prog_res["daily_ship_p"]
                st.session_state["kpi_data"]["daily"]["rev_p"] = prog_res["daily_rev_p"]
                st.session_state["kpi_data"]["daily"]["ship_a"] = prog_res["daily_ship_a"]
                st.session_state["kpi_data"]["daily"]["rev_a"] = prog_res["daily_rev_a"]

                st.session_state["kpi_data"]["weekly"]["ship_p"] = prog_res["cur_week_ship_p"]
                st.session_state["kpi_data"]["weekly"]["rev_p"] = prog_res["cur_week_rev_p"]
                st.session_state["kpi_data"]["weekly"]["ship_a"] = prog_res["cur_week_ship_a"]
                st.session_state["kpi_data"]["weekly"]["rev_a"] = prog_res["cur_week_rev_a"]
                st.session_state["kpi_data"]["weekly"]["pure_rev_a"] = prog_res["cur_week_rev_a"]
                st.session_state["kpi_data"]["weekly"]["carryover_rev"] = 0.0

                st.session_state["chart_data"]["weekly_ship_p"] = [
                    prog_res["weekly_ship_plan_chart"].get(w, 0) for w in w_labels
                ]
                st.session_state["chart_data"]["weekly_rev_p"] = [
                    prog_res["weekly_rev_plan_chart"].get(w, 0.0) for w in w_labels
                ]
                st.session_state["chart_data"]["weekly_ship_a"] = [
                    prog_res["weekly_ship_act_chart"].get(w, 0) for w in w_labels
                ]
                st.session_state["chart_data"]["weekly_rev_a"] = [
                    prog_res["weekly_rev_act_chart"].get(w, 0.0) for w in w_labels
                ]

                # '월별요약' 시트가 없을 경우의 월간 기본값 세팅
                st.session_state["kpi_data"]["monthly"]["ship_p"] = prog_res["monthly_ship_p"]
                st.session_state["kpi_data"]["monthly"]["rev_p"] = prog_res["monthly_rev_p"]
                st.session_state["kpi_data"]["monthly"]["ship_a"] = prog_res["monthly_ship_a"]
                st.session_state["kpi_data"]["monthly"]["rev_a"] = prog_res["monthly_rev_a"]
                st.session_state["kpi_data"]["monthly"]["pure_rev_a"] = prog_res["monthly_rev_a"]
                st.session_state["chart_data"]["monthly_ship_p"] = [prog_res["monthly_ship_p"]]
                st.session_state["chart_data"]["monthly_ship_a"] = [prog_res["monthly_ship_a"]]
                st.session_state["chart_data"]["monthly_rev_p"] = [prog_res["monthly_rev_p"]]
                st.session_state["chart_data"]["monthly_rev_a"] = [prog_res["monthly_rev_a"]]

                plan_model_period = prog_res["plan_model_period"]
                actual_model_period = prog_res["actual_model_period"]

                messages.append(
                    (
                        "success",
                        f"✅ [{prog_res['sheet_name']}] 반영 완료 — 기준일({t_m}/{t_d}) 출하실적: {format_k(prog_res['daily_ship_a'])} ({prog_res['daily_rev_a']:,.2f}억) | {active_week_lbl} 주간실적: {format_k(prog_res['cur_week_ship_a'])}",
                    )
                )

            # (3-B) '월별요약' 반영 -> 현황 Monthly 총 수량/총 매출 및 1번 메뉴 진도율 리포트
            if sum_res:
                st.session_state["shipment_report"] = sum_res
                if sum_res["summary_rows"]:
                    # '매출 예상' 행 또는 첫 번째 요약행을 공식 월간 총계로 반영
                    main_sum = sum_res["summary_rows"][-1]
                    st.session_state["kpi_data"]["monthly"]["ship_p"] = main_sum["plan_qty"]
                    st.session_state["kpi_data"]["monthly"]["ship_a"] = main_sum["act_qty"]
                    st.session_state["kpi_data"]["monthly"]["rev_p"] = main_sum["plan_amt"]
                    st.session_state["kpi_data"]["monthly"]["rev_a"] = main_sum["act_amt"]
                    st.session_state["kpi_data"]["monthly"]["target_ship"] = main_sum["target_qty"]
                    st.session_state["kpi_data"]["monthly"]["target_rev"] = main_sum["target_amt"]
                    st.session_state["kpi_data"]["monthly"]["pure_rev_a"] = main_sum["act_amt"]
                    st.session_state["kpi_data"]["monthly"]["carryover_rev"] = 0.0

                    st.session_state["chart_data"]["monthly_ship_p"] = [main_sum["plan_qty"]]
                    st.session_state["chart_data"]["monthly_ship_a"] = [main_sum["act_qty"]]
                    st.session_state["chart_data"]["monthly_rev_p"] = [main_sum["plan_amt"]]
                    st.session_state["chart_data"]["monthly_rev_a"] = [main_sum["act_amt"]]

                messages.append(
                    (
                        "success",
                        f"✅ [{sum_res['sheet_name']}] 반영 완료 — 월간 계획 {format_k(st.session_state['kpi_data']['monthly']['ship_p'])} ({st.session_state['kpi_data']['monthly']['rev_p']:,.1f}억) / 실적 {format_k(st.session_state['kpi_data']['monthly']['ship_a'])} ({st.session_state['kpi_data']['monthly']['rev_a']:,.1f}억) | 상세 모델 {len(sum_res['detail_rows'])}건",
                    )
                )

            # (3-C) '매출 참고' 반영 -> 고객 FCST 조회 & 캘린더 공급 일정 (진도현황 계획 부재 시 계획 폴백)
            if ref_res:
                st.session_state["fcst_data"] = {
                    "rows": ref_res.get("fcst_rows", []),
                    "date_cols": ref_res.get("fcst_date_cols", []),
                }
                if st.session_state["kpi_data"]["daily"]["ship_p"] == 0 and ref_res["daily_ship_p"] > 0:
                    st.session_state["kpi_data"]["daily"]["ship_p"] = ref_res["daily_ship_p"]
                    st.session_state["kpi_data"]["daily"]["rev_p"] = ref_res["daily_rev_p"]
                if st.session_state["kpi_data"]["weekly"]["ship_p"] == 0 and ref_res["cur_week_ship_p"] > 0:
                    st.session_state["kpi_data"]["weekly"]["ship_p"] = ref_res["cur_week_ship_p"]
                    st.session_state["kpi_data"]["weekly"]["rev_p"] = ref_res["cur_week_rev_p"]
                if sum(st.session_state["chart_data"]["weekly_ship_p"]) == 0:
                    st.session_state["chart_data"]["weekly_ship_p"] = [
                        ref_res["weekly_ship_plan"].get(w, 0) for w in w_labels
                    ]
                    st.session_state["chart_data"]["weekly_rev_p"] = [
                        ref_res["weekly_rev_plan"].get(w, 0.0) for w in w_labels
                    ]
                if not prog_res:
                    plan_model_period = ref_res["plan_by_model_period"]

                messages.append(
                    (
                        "success",
                        f"✅ [{ref_res['sheet_name']}] FCST 반영 완료 — 공급 계획 품목 {len(ref_res['fcst_rows'])}건 ({len(ref_res['fcst_date_cols'])}개 일자)",
                    )
                )

            if not (sum_res or prog_res or ref_res):
                messages.append(
                    (
                        "error",
                        f"⚠️ 업로드된 파일에서 '월별요약', '진도현황', '매출 참고' 구조를 찾지 못했습니다. (감지된 시트: {', '.join(integ['all_sheets'])})",
                    )
                )
            else:
                st.session_state["data_loaded"] = True

            # 모델별 일간/주간/월간 비교 테이블 생성
            model_period_tables = {}
            for p_key in ["daily", "weekly", "monthly"]:
                model_period_tables[p_key] = _build_merged_table(
                    plan_model_period.get(p_key, {}), actual_model_period.get(p_key, {}), "모델명"
                )
            st.session_state["chart_data"]["ship_by_model_period"] = model_period_tables
            st.session_state["chart_data"]["ship_by_model"] = model_period_tables["daily"]

        except Exception as e:
            messages.append(("error", f"통합 출하·매출 파일 파싱 오류: {e}"))

    st.session_state["last_updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_state_to_disk()
    return messages


# 세션 초기화 및 버전 변경 시 저장된 엑셀로 자동 재집계
def init_and_sync_session():
    defaults = get_default_state()
    if "active_cal_popup" not in st.session_state:
        st.session_state["active_cal_popup"] = None

    disk_mtime = os.path.getmtime(STATE_JSON_PATH) if os.path.exists(STATE_JSON_PATH) else None
    loaded_mtime = st.session_state.get("_loaded_mtime")

    if disk_mtime is not None and (loaded_mtime is None or disk_mtime > loaded_mtime):
        ok, saved_ver = load_state_from_disk()
        if ok:
            if saved_ver != STATE_VERSION and (
                os.path.exists(CACHED_EXCEL_PATHS["integrated"]) or os.path.exists(CACHED_EXCEL_PATHS["prod"])
            ):
                run_dashboard_pipeline(
                    st.session_state.get("target_date", default_target_date),
                    st.session_state.get("ex_rate", 1330.0),
                    {},
                )
            return

    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_and_sync_session()


# 6. UI 컴포넌트 렌더링 함수
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

    sub_html = (
        f"<div style='font-size: 11px; color: #6b7280; font-weight: 700; margin-top: 6px; border-top: 1px dashed #e5e7eb; padding-top: 4px;'>{sub_caption}</div>"
        if sub_caption
        else ""
    )

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


def render_shipment_html_table(rows, is_summary=False, month_lbl="11월", ref_lbl=""):
    ref_suffix = f" ~{ref_lbl}" if (is_summary and ref_lbl) else ""
    left_headers = (
        '<th rowspan="2" style="min-width:140px; border:1px solid #111;">구분</th>'
        if is_summary
        else (
            '<th rowspan="2" style="border:1px solid #111;">구분</th>'
            '<th rowspan="2" style="border:1px solid #111;">과제명</th>'
            '<th rowspan="2" style="border:1px solid #111;">모델명</th>'
        )
    )

    body_html = ""
    for idx, r in enumerate(rows):
        is_est_row = is_summary and ("예상" in r["구분"] or (idx == len(rows) - 1 and len(rows) > 1))
        row_bg = "#dce6f2" if is_est_row else "#ffffff"
        fw = "800" if is_est_row else "600"

        dq_bg = "#f8d7da" if r["diff_qty"] < 0 else row_bg
        dq_col = "#b91c1c" if r["diff_qty"] < 0 else "#111111"
        da_bg = "#f8d7da" if r["diff_amt"] < 0 else row_bg
        da_col = "#b91c1c" if r["diff_amt"] < 0 else "#111111"
        rd_bg = "#f8d7da" if r["rate_diff"] < 0 else row_bg
        rd_col = "#b91c1c" if r["rate_diff"] < 0 else "#111111"

        pct_fmt = "{:,.1f}%" if is_summary else "{:,.0f}%"
        force_dec = not is_summary

        if is_summary:
            left_tds = f"<td style='background:{row_bg}; font-weight:{fw};'>{r['구분']}</td>"
        else:
            left_tds = (
                f"<td style='background:{row_bg};'>{r['구분']}</td>"
                f"<td style='background:{row_bg}; font-weight:700;'>{r['과제명']}</td>"
                f"<td style='background:{row_bg}; color:#374151;'>{r['모델명']}</td>"
            )

        body_html += f"""
        <tr style="font-weight:{fw};">
            {left_tds}
            <td style="background:{row_bg};">{format_k(r['plan_qty'])}</td>
            <td style="background:{row_bg};">{fmt_amt_ok(r['plan_amt'], always_decimal=force_dec)}</td>
            <td style="background:{row_bg};">{format_k(r['target_qty'])}</td>
            <td style="background:{row_bg};">{fmt_amt_ok(r['target_amt'], always_decimal=True)}</td>
            <td style="background:{row_bg}; font-weight:800;">{format_k(r['act_qty'])}</td>
            <td style="background:{row_bg}; font-weight:800;">{fmt_amt_ok(r['act_amt'], always_decimal=force_dec)}</td>
            <td style="background:{dq_bg}; color:{dq_col}; font-weight:800;">{fmt_signed_k(r['diff_qty'])}</td>
            <td style="background:{da_bg}; color:{da_col}; font-weight:800;">{fmt_amt_ok(r['diff_amt'], always_decimal=force_dec)}</td>
            <td style="background:{row_bg};">{pct_fmt.format(r['rate_target'])}</td>
            <td style="background:{row_bg}; font-weight:800;">{pct_fmt.format(r['rate_act'])}</td>
            <td style="background:{rd_bg}; color:{rd_col}; font-weight:800;">{pct_fmt.format(r['rate_diff'])}</td>
        </tr>
        """

    html = f"""
    <div style="overflow-x:auto; margin-bottom:18px; border:2.5px solid #111111; box-shadow:5px 5px 0px #0044ff; border-radius:4px;">
        <table class="ship-rep-table" style="width:100%; border-collapse:collapse; font-size:13px; text-align:center; background:#ffffff;">
            <thead>
                <tr style="background:#24476e; color:#ffffff; font-weight:800;">
                    {left_headers}
                    <th colspan="2" style="padding:8px; border:1px solid #111;">{month_lbl} 계획</th>
                    <th colspan="2" style="padding:8px; border:1px solid #111;">목표{ref_suffix}</th>
                    <th colspan="2" style="padding:8px; border:1px solid #111;">실적{ref_suffix}</th>
                    <th colspan="2" style="padding:8px; border:1px solid #111;">목표 대비 차이</th>
                    <th colspan="2" style="padding:8px; border:1px solid #111;">진도율</th>
                    <th rowspan="2" style="padding:8px; border:1px solid #111;">차이</th>
                </tr>
                <tr style="background:#2d5986; color:#ffffff; font-weight:700; font-size:12px;">
                    <th style="padding:6px; border:1px solid #111;">수량</th><th style="padding:6px; border:1px solid #111;">금액</th>
                    <th style="padding:6px; border:1px solid #111;">수량</th><th style="padding:6px; border:1px solid #111;">금액</th>
                    <th style="padding:6px; border:1px solid #111;">수량</th><th style="padding:6px; border:1px solid #111;">금액</th>
                    <th style="padding:6px; border:1px solid #111;">수량</th><th style="padding:6px; border:1px solid #111;">금액</th>
                    <th style="padding:6px; border:1px solid #111;">목표</th><th style="padding:6px; border:1px solid #111;">실적</th>
                </tr>
            </thead>
            <tbody>
                {body_html}
            </tbody>
        </table>
    </div>
    """
    render_html(html)


def _render_day_detail_body(d_num, m_num, y_num, d_info, act_info, show_actual_mode):
    plan_items = d_info.get("items", [])
    plan_q = d_info.get("total", 0)
    plan_rev = d_info.get("total_rev", 0.0)

    act_items = act_info.get("items", []) if act_info else []
    act_q = act_info.get("qty", 0) if act_info else 0
    act_rev = act_info.get("rev", 0.0) if act_info else 0.0

    if show_actual_mode:
        diff_q = act_q - plan_q
        diff_q_str = f"+{format_k(diff_q)}" if diff_q > 0 else (f"-{format_k(abs(diff_q))}" if diff_q < 0 else "0K")
        diff_rev = act_rev - plan_rev
        diff_rev_str = f"+{diff_rev:,.2f}억" if diff_rev > 0 else (f"-{abs(diff_rev):,.2f}억" if diff_rev < 0 else "0.00억")
        rate_q = (act_q / plan_q * 100.0) if plan_q > 0 else 0.0
        act_models_cnt = len(set(it["model"] for it in act_items))
        plan_models_cnt = len(set(it["model"] for it in plan_items))

        cards_html = f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px;">
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">🚚 출하 수량 (실적 / 계획)</div>
                <div style="font-size: 18px; font-weight: 900; color: #0044ff; margin-top: 3px;">{format_k(act_q)} <span style="font-size:13px; color:#6b7280;">/ {format_k(plan_q)}</span></div>
                <div style="font-size: 11px; font-weight: 800; color: #111111;">계획 대비 {diff_q_str}</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">💰 매출액 (실적 / 계획)</div>
                <div style="font-size: 18px; font-weight: 900; color: #10b981; margin-top: 3px;">{act_rev:,.2f}억 <span style="font-size:13px; color:#6b7280;">/ {plan_rev:,.2f}억</span></div>
                <div style="font-size: 11px; font-weight: 800; color: #111111;">계획 대비 {diff_rev_str}</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">📱 모델 수 (실적 / 계획)</div>
                <div style="font-size: 18px; font-weight: 900; color: #111111; margin-top: 3px;">{act_models_cnt}종 <span style="font-size:13px; color:#6b7280;">/ {plan_models_cnt}종</span></div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">📊 당일 수량 달성률</div>
                <div style="font-size: 20px; font-weight: 900; color: #111111; margin-top: 3px;">{rate_q:,.1f}%</div>
            </div>
        </div>
        """
        render_html(cards_html)

        t_act, t_plan = st.tabs(["✅ 당일 출하 실적 상세", "📋 당일 출하/공급 계획 상세"])
        with t_act:
            if act_items:
                df_act = pd.DataFrame(act_items).sort_values(by="qty", ascending=False)
                df_act["출하수량"] = df_act["qty"].apply(format_k)
                df_act["적용단가(원)"] = df_act["price"].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
                df_act["실적매출(억원)"] = df_act["rev_ok"].apply(lambda x: f"{x:,.2f}억원")
                df_act = df_act.rename(columns={"model": "모델명", "gh_code": "SEC/GH-CODE", "site": "구분/SITE"})
                cols_a = ["모델명", "SEC/GH-CODE", "구분/SITE", "출하수량", "적용단가(원)", "실적매출(억원)"]
                st.dataframe(df_act[cols_a], use_container_width=True, hide_index=True)
                st.download_button(
                    label=f"📥 {y_num}년 {m_num}월 {d_num}일 출하 실적 다운로드 (CSV)",
                    data=convert_df(df_act[cols_a]),
                    file_name=f"actual_shipment_{y_num}_{m_num:02d}_{d_num:02d}.csv",
                    mime="text/csv",
                    key=f"dl_act_pop_{y_num}_{m_num}_{d_num}",
                )
            else:
                st.info("해당 일자에 집계된 출하 실적 내역이 없습니다.")

        with t_plan:
            if plan_items:
                df_plan = pd.DataFrame(plan_items).sort_values(by="qty", ascending=False)
                df_plan["계획수량"] = df_plan["qty"].apply(format_k)
                df_plan["단가(원)"] = df_plan["price"].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
                df_plan["예상매출(억원)"] = df_plan["rev_ok"].apply(lambda x: f"{x:,.2f}억원")
                df_plan = df_plan.rename(columns={"model": "모델명", "gh_code": "SEC/GH-CODE", "site": "구분/SITE"})
                cols_p = ["모델명", "SEC/GH-CODE", "구분/SITE", "계획수량", "단가(원)", "예상매출(억원)"]
                st.dataframe(df_plan[cols_p], use_container_width=True, hide_index=True)
                st.download_button(
                    label=f"📥 {y_num}년 {m_num}월 {d_num}일 공급 계획 다운로드 (CSV)",
                    data=convert_df(df_plan[cols_p]),
                    file_name=f"fcst_supply_{y_num}_{m_num:02d}_{d_num:02d}.csv",
                    mime="text/csv",
                    key=f"dl_plan_pop_{y_num}_{m_num}_{d_num}",
                )
            else:
                st.info("해당 일자에 예정된 공급 계획이 없습니다.")
    else:
        unique_models = sorted(list(set(it["model"] for it in plan_items)))
        cards_html = f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px;">
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">📱 전체 모델 수</div>
                <div style="font-size: 20px; font-weight: 900; color: #111111; margin-top: 3px;">{len(unique_models)}종</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">📦 총 공급 품목(GH)</div>
                <div style="font-size: 20px; font-weight: 900; color: #111111; margin-top: 3px;">{len(plan_items)}건</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">🚚 총 공급 수량</div>
                <div style="font-size: 20px; font-weight: 900; color: #0044ff; margin-top: 3px;">{format_k(plan_q)}</div>
            </div>
            <div style="background-color: #ffffff; border: 2px solid #111111; padding: 10px 12px; box-shadow: 3px 3px 0px #0044ff; border-radius: 4px; text-align: center;">
                <div style="font-size: 12px; font-weight: 800; color: #6b7280;">💰 총 예상 매출액</div>
                <div style="font-size: 20px; font-weight: 900; color: #10b981; margin-top: 3px;">{plan_rev:,.2f}억원</div>
            </div>
        </div>
        """
        render_html(cards_html)
        st.divider()

        if plan_items:
            df_pop = pd.DataFrame(plan_items).sort_values(by="qty", ascending=False)
            df_pop["공급수량"] = df_pop["qty"].apply(format_k)
            df_pop["단가(원)"] = df_pop["price"].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
            df_pop["예상매출(억원)"] = df_pop["rev_ok"].apply(lambda x: f"{x:,.2f}억원")
            df_pop = df_pop.rename(columns={"model": "모델명", "gh_code": "GH-CODE", "site": "출하 SITE"})
            cols = ["모델명", "GH-CODE", "출하 SITE", "공급수량", "단가(원)", "예상매출(억원)"]
            st.dataframe(df_pop[cols], use_container_width=True, hide_index=True)
            st.download_button(
                label=f"📥 {y_num}년 {m_num}월 {d_num}일 공급 상세 다운로드 (CSV)",
                data=convert_df(df_pop[cols]),
                file_name=f"fcst_supply_{y_num}_{m_num:02d}_{d_num:02d}_detail.csv",
                mime="text/csv",
                key=f"dl_popup_{y_num}_{m_num}_{d_num}",
            )
        else:
            st.info("해당 일자에 예정된 공급 일정이 없습니다.")


if hasattr(st, "dialog"):
    @st.dialog("📋 일자별 공급 및 출하 상세 내역", width="large")
    def show_day_detail_dialog(d_num, m_num, y_num, d_info, act_info, show_actual_mode):
        st.markdown(f"#### 📅 {y_num}년 {m_num}월 {d_num}일 상세 내역")
        _render_day_detail_body(d_num, m_num, y_num, d_info, act_info, show_actual_mode)


# 7. CSS 스타일링
render_html(
    """
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
        font-size: 14px;
        margin-bottom: 18px;
        box-shadow: 4px 4px 0px #0044ff;
    }
    .ship-rep-table tbody td {
        padding: 7px 8px;
        border: 1px solid #d1d5db;
    }
    div[data-testid="stDialog"] div[role="dialog"] {
        max-width: 920px !important;
        width: 90% !important;
    }
    .cal-card-wrapper {
        position: relative;
        margin-bottom: 6px;
    }
    .cal-card {
        min-height: 108px;
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
        width: 335px;
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
"""
)

# 8. 사이드바 (기존 순서 유지 및 SITE별 실적 제거)
with st.sidebar:
    menu = st.radio(
        "📂 Folder",
        (
            "🏠 현황",
            "1. 모델별 출하실적",
            "2. 모델별 생산실적",
            "3. 고객 FCST 조회",
            "4. 완제품 재고 조회",
            "5. 캘린더",
            "6. 데이터 관리 (Admin)",
        ),
    )


def create_active_chart(title, categories, plan_data, actual_data, name1="계획", name2="실적", unit="개", highlight_cat=None):
    fig = go.Figure()
    if not categories:
        categories, plan_data, actual_data = ["-"], [0], [0]

    if unit == "억원":
        txt_p = [f"{v:,.2f}" if v > 0 else "" for v in plan_data]
        txt_a = [f"{v:,.2f}" if v > 0 else "" for v in actual_data]
    else:
        txt_p = [format_k(v) if v > 0 else "" for v in plan_data]
        txt_a = [format_k(v) if v > 0 else "" for v in actual_data]

    act_colors = [
        "#1d4ed8" if (highlight_cat and c == highlight_cat) else "#0044ff"
        for c in categories
    ]
    line_widths = [
        2.8 if (highlight_cat and c == highlight_cat) else 1.5
        for c in categories
    ]

    fig.add_trace(
        go.Bar(
            x=categories,
            y=plan_data,
            name=name1,
            marker_color="#d1d5db",
            marker_line=dict(width=line_widths, color="#111111"),
            text=txt_p,
            textposition="outside",
        )
    )
    fig.add_trace(
        go.Bar(
            x=categories,
            y=actual_data,
            name=name2,
            marker_color=act_colors,
            marker_line=dict(width=line_widths, color="#111111"),
            text=txt_a,
            textposition="outside",
        )
    )
    fig.update_layout(
        title=dict(text=f"<b>{title} ({unit})</b>", font=dict(size=14, color="#111111")),
        barmode="group",
        plot_bgcolor="rgba(0,0,0,0)",
        hovermode="x unified",
        margin=dict(l=20, r=20, t=45, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    return fig


# 9. 메인 메뉴 분기 로직
if menu == "6. 데이터 관리 (Admin)":
    st.title("⚙️ 6. DATA MANAGEMENT (Admin)")
    if st.text_input("비밀번호", type="password") == "2250":
        prod_ok = "🟢 등록됨" if os.path.exists(CACHED_EXCEL_PATHS["prod"]) else "⚪ 미등록"
        integ_ok = "🟢 등록됨" if os.path.exists(CACHED_EXCEL_PATHS["integrated"]) else "⚪ 미등록"
        s_info = st.session_state.get("loaded_sheet_info", {})
        sheet_desc = ""
        if s_info:
            sheet_desc = (
                f"  \n📑 **감지된 통합 시트:** 요약=`{s_info.get('summary') or '-'}` | "
                f"일별진도=`{s_info.get('progress') or '-'}` | FCST참고=`{s_info.get('sales_ref') or '-'}`"
            )
        det_ex = st.session_state.get("detected_ex_rate")
        ex_desc = f" (시트 자동감지 환율: {det_ex:,.0f}원/USD)" if det_ex else ""

        last_upd = st.session_state.get("last_updated_at") or "반영 이력 없음"
        st.info(
            f"💾 **서버 보관 상태:** 생산실적 {prod_ok} | 출하·매출 통합파일 {integ_ok}{ex_desc}{sheet_desc}  \n"
            f"🕒 **마지막 반영 시각:** {last_upd} (새로고침 시 자동 유지, 날짜만 변경 후 재반영 가능)"
        )

        c_set1, c_set2 = st.columns(2)
        with c_set1:
            target_date_input = st.text_input(
                "📅 조회 기준 날짜 지정 (예: 11/25 또는 2026-11-25)",
                value=st.session_state.get("target_date", default_target_date),
            )
        with c_set2:
            ex_rate_input = st.number_input(
                "💵 기본 환율 (시트 내 환율 미기재 시 적용, 원/USD)",
                value=float(st.session_state.get("ex_rate", 1330.0)),
                step=10.0,
            )

        col1, col2 = st.columns(2)
        with col1:
            prod_file = st.file_uploader("1. 생산 실적 파일 (생산실적 · 완제품 재고)", type=["xlsx"])
        with col2:
            integ_file = st.file_uploader(
                "2. 출하 · 매출 통합 파일 ('월별요약' + '진도현황' + '매출 참고' 포함)",
                type=["xlsx"],
            )

        with st.expander("➕ 별도 단가 매핑 파일 업로드 (선택 사항)"):
            map_rev_file = st.file_uploader("3. 추가 단가 매핑 파일 (.xlsx)", type=["xlsx"])

        btn_c1, btn_c2 = st.columns([0.75, 0.25])
        with btn_c1:
            if st.button("🚀 데이터 분석 및 서버 저장 반영", type="primary", use_container_width=True):
                with st.spinner("통합 엑셀 시트(월별요약 · 진도현황 · 매출참고) 및 생산 데이터를 분석 중입니다..."):
                    msgs = run_dashboard_pipeline(
                        target_date_input,
                        ex_rate_input,
                        {
                            "prod": prod_file,
                            "integrated": integ_file,
                            "map_rev": map_rev_file,
                        },
                    )
                for level, txt in msgs:
                    if level == "success":
                        st.success(txt)
                    else:
                        st.error(txt)
                if not msgs:
                    st.warning("업로드되거나 서버에 보관된 엑셀 파일이 없습니다. 파일을 업로드해 주세요.")
        with btn_c2:
            if st.button("🗑️ 저장 데이터 전체 초기화", use_container_width=True):
                reset_persisted_store()
                st.success("서버 저장 데이터와 대시보드 상태가 초기화되었습니다.")
                st.rerun()

elif menu == "🏠 현황":
    st.title("📊 실시간 현황 (전체 통합)")

    cur_target_date = st.session_state.get("target_date", default_target_date)
    t_y, t_m, t_d = parse_target_ymd(cur_target_date)
    cur_week_lbl = st.session_state.get("cur_week_label", get_iso_week_label(datetime.date(t_y, t_m, t_d)))
    last_upd = st.session_state.get("last_updated_at")
    upd_badge = f" &nbsp;|&nbsp; 🕒 최근 반영 : <span style='color:#4b5563;'>{last_upd}</span>" if last_upd else ""

    render_html(
        f"""
    <div class="date-badge">
        📅 조회 기준일 : <span style="color: #0044ff;">{t_y}-{t_m:02d}-{t_d:02d} ({t_m}/{t_d})</span>
        &nbsp;|&nbsp; 📊 조회 주차 : <span style="color: #0044ff;">{cur_week_lbl}</span>
        &nbsp;|&nbsp; 🌐 집계 범위 : <span style="color: #059669;">전 거점 통합</span>{upd_badge}
    </div>
    """
    )

    kpi = st.session_state.get("kpi_data", get_default_state()["kpi_data"])
    cdata = st.session_state.get("chart_data", get_default_state()["chart_data"])

    tab_monthly, tab_weekly, tab_daily = st.tabs(["🗓️ Monthly", "📊 Weekly", "📅 Daily"])

    # 1. Monthly 탭 ('월별요약' 시트 총계 연동)
    with tab_monthly:
        m1, m2, m3 = st.columns(3)
        t_ship_goal = kpi["monthly"].get("target_ship", 0)
        t_rev_goal = kpi["monthly"].get("target_rev", 0.0)
        ship_sub = f"기준일 누적 목표: {format_k(t_ship_goal)} (목표 대비 {fmt_signed_k(kpi['monthly']['ship_a'] - t_ship_goal)})" if t_ship_goal > 0 else None
        rev_sub = f"기준일 누적 목표: {t_rev_goal:,.1f}억 (목표 대비 {kpi['monthly']['rev_a'] - t_rev_goal:+,.1f}억)" if t_rev_goal > 0 else None

        p_sheet_tot = kpi["monthly"].get("prod_sheet_total", 0)
        p_cur_add = kpi["monthly"].get("prod_cur_w_add", 0)
        prod_sub = f"시트 월합계 {format_k(p_sheet_tot)} + {cur_week_lbl} 일별실적 {format_k(p_cur_add)}" if p_cur_add > 0 else None
        with m1:
            render_kpi_card(f"{t_m}월 생산 누계 (전체)", "🏭", kpi["monthly"]["prod_a"], kpi["monthly"]["prod_p"], sub_caption=prod_sub)
            render_achievement_gauge("생산", kpi["monthly"]["prod_a"], kpi["monthly"]["prod_p"], unit="K")
        with m2:
            render_kpi_card(f"{t_m}월 출하 누계 (전체)", "🚚", kpi["monthly"]["ship_a"], kpi["monthly"]["ship_p"], sub_caption=ship_sub)
            render_achievement_gauge("출하", kpi["monthly"]["ship_a"], kpi["monthly"]["ship_p"], unit="K")
        with m3:
            render_kpi_card(
                f"{t_m}월 매출 누계 (전체)", "💰", kpi["monthly"]["rev_a"], kpi["monthly"]["rev_p"], is_rev=True, sub_caption=rev_sub
            )
            render_achievement_gauge("매출", kpi["monthly"]["rev_a"], kpi["monthly"]["rev_p"], unit="억원")

        mch1, mch2, mch3 = st.columns(3)
        m_months = cdata.get("monthly_months", [f"{t_m}월"])
        with mch1:
            st.plotly_chart(
                create_active_chart("월별 생산 현황 (전체)", m_months, cdata.get("monthly_prod_p", [0]), cdata.get("monthly_prod_a", [0]), unit="K"),
                use_container_width=True,
            )
        with mch2:
            st.plotly_chart(
                create_active_chart("월별 출하 현황 (전체)", m_months, cdata.get("monthly_ship_p", [0]), cdata.get("monthly_ship_a", [0]), unit="K"),
                use_container_width=True,
            )
        with mch3:
            st.plotly_chart(
                create_active_chart("월별 매출 현황 (전체)", m_months, cdata.get("monthly_rev_p", [0.0]), cdata.get("monthly_rev_a", [0.0]), unit="억원"),
                use_container_width=True,
            )

    # 2. Weekly 탭 (조회 기준일 주차 단일 집계)
    with tab_weekly:
        wk1, wk2, wk3 = st.columns(3)
        with wk1:
            render_kpi_card(f"주간 생산 ({cur_week_lbl} 전체)", "🏭", kpi["weekly"]["prod_a"], kpi["weekly"]["prod_p"])
            render_achievement_gauge("생산", kpi["weekly"]["prod_a"], kpi["weekly"]["prod_p"], unit="K")
        with wk2:
            render_kpi_card(f"주간 출하 ({cur_week_lbl} 전체)", "🚚", kpi["weekly"]["ship_a"], kpi["weekly"]["ship_p"])
            render_achievement_gauge("출하", kpi["weekly"]["ship_a"], kpi["weekly"]["ship_p"], unit="K")
        with wk3:
            render_kpi_card(
                f"주간 매출 ({cur_week_lbl} 전체)",
                "💰",
                kpi["weekly"]["rev_a"],
                kpi["weekly"]["rev_p"],
                is_rev=True,
                sub_caption=f"{cur_week_lbl} 당주 누적 실적 기준",
            )
            render_achievement_gauge("매출", kpi["weekly"]["rev_a"], kpi["weekly"]["rev_p"], unit="억원")

        wch1, wch2, wch3 = st.columns(3)
        w_weeks = cdata.get("weekly_weeks", get_month_iso_weeks(t_y, t_m))
        with wch1:
            st.plotly_chart(
                create_active_chart(
                    "주차별 생산 추이 (전체)",
                    w_weeks,
                    cdata.get("weekly_prod_p", [0] * len(w_weeks)),
                    cdata.get("weekly_prod_a", [0] * len(w_weeks)),
                    unit="K",
                    highlight_cat=cur_week_lbl,
                ),
                use_container_width=True,
            )
        with wch2:
            st.plotly_chart(
                create_active_chart(
                    "주차별 출하 추이 (전체)",
                    w_weeks,
                    cdata.get("weekly_ship_p", [0] * len(w_weeks)),
                    cdata.get("weekly_ship_a", [0] * len(w_weeks)),
                    unit="K",
                    highlight_cat=cur_week_lbl,
                ),
                use_container_width=True,
            )
        with wch3:
            st.plotly_chart(
                create_active_chart(
                    "주차별 매출 추이 (전체)",
                    w_weeks,
                    cdata.get("weekly_rev_p", [0.0] * len(w_weeks)),
                    cdata.get("weekly_rev_a", [0.0] * len(w_weeks)),
                    unit="억원",
                    highlight_cat=cur_week_lbl,
                ),
                use_container_width=True,
            )

    # 3. Daily 탭 ('진도현황' 시트 조회일 당일 계획/실적 연동)
    with tab_daily:
        k1, k2, k3 = st.columns(3)
        with k1:
            render_kpi_card(f"일간 생산 ({t_m}/{t_d})", "🏭", kpi["daily"]["prod_a"], kpi["daily"]["prod_p"])
            render_achievement_gauge("생산", kpi["daily"]["prod_a"], kpi["daily"]["prod_p"], unit="K")
        with k2:
            render_kpi_card(f"일간 출하 ({t_m}/{t_d})", "🚚", kpi["daily"]["ship_a"], kpi["daily"]["ship_p"])
            render_achievement_gauge("출하", kpi["daily"]["ship_a"], kpi["daily"]["ship_p"], unit="K")
        with k3:
            render_kpi_card(f"일간 매출 ({t_m}/{t_d})", "💰", kpi["daily"]["rev_a"], kpi["daily"]["rev_p"], is_rev=True)
            render_achievement_gauge("매출", kpi["daily"]["rev_a"], kpi["daily"]["rev_p"], unit="억원")

        ch1, ch2, ch3 = st.columns(3)
        with ch1:
            st.plotly_chart(
                create_active_chart("일간 생산 현황 (전체)", ["전체"], [kpi["daily"]["prod_p"]], [kpi["daily"]["prod_a"]], unit="K"),
                use_container_width=True,
            )
        with ch2:
            st.plotly_chart(
                create_active_chart("일간 출하 현황 (전체)", ["전체"], [kpi["daily"]["ship_p"]], [kpi["daily"]["ship_a"]], unit="K"),
                use_container_width=True,
            )
        with ch3:
            st.plotly_chart(
                create_active_chart("일간 매출 현황 (전체)", ["전체"], [kpi["daily"]["rev_p"]], [kpi["daily"]["rev_a"]], unit="억원"),
                use_container_width=True,
            )

elif menu == "1. 모델별 출하실적":
    rep = st.session_state.get("shipment_report", {})
    sum_rows = rep.get("summary_rows", [])
    det_rows = rep.get("detail_rows", [])
    m_lbl = rep.get("month_label") or "당월"
    ref_lbl = rep.get("ref_date_label") or ""

    period_dict = st.session_state["chart_data"].get("ship_by_model_period", {})
    has_daily_prog = any(bool(period_dict.get(k)) for k in ("daily", "weekly", "monthly"))

    title_ref = f" (기준: ~{ref_lbl})" if ref_lbl else ""
    st.subheader(f"📦 1. 모델별 출하 및 매출 진도 현황{title_ref}")

    if not sum_rows and not det_rows and not has_daily_prog:
        st.info("💡 **[6. 데이터 관리 (Admin)] → 2. 출하 · 매출 통합 파일**을 업로드 및 반영해 주세요.")
    else:
        main_tab1, main_tab2 = st.tabs(
            [
                f"📊 {m_lbl} 요약 및 과제/모델별 진도율 (월별요약)",
                "📅 일간 · 주간 · 월간 출하 상세 (진도현황)",
            ]
        )

        # 탭 1: '월별요약' 시트 기반 요약표 + 모델별 진도율 표 & 3중 비교 차트
        with main_tab1:
            if sum_rows:
                st.markdown(f"##### 📌 {m_lbl} 전체 출하 · 매출 요약")
                render_shipment_html_table(sum_rows, is_summary=True, month_lbl=m_lbl, ref_lbl=ref_lbl)

            if det_rows:
                df_det = pd.DataFrame(det_rows)
                f_c1, f_c2, f_c3 = st.columns([1.2, 1.8, 1.0])
                with f_c1:
                    gubun_opts = sorted(df_det["구분"].unique().tolist())
                    sel_gubun = st.multiselect("🏢 구분 필터", options=gubun_opts, default=[])
                with f_c2:
                    proj_opts = df_det["과제명"].tolist()
                    sel_proj = st.multiselect("📱 과제명 필터", options=proj_opts, default=[])
                with f_c3:
                    x_axis_col = st.radio("차트 X축 기준", ["과제명", "모델명"], horizontal=True)

                filtered_rows = [
                    r
                    for r in det_rows
                    if (not sel_gubun or r["구분"] in sel_gubun)
                    and (not sel_proj or r["과제명"] in sel_proj)
                ]

                if filtered_rows:
                    x_labels = [r[x_axis_col] for r in filtered_rows]
                    tab_q, tab_a, tab_r = st.tabs(
                        [
                            "🚚 수량 비교 (계획 vs 목표 vs 실적)",
                            "💰 금액 비교 (억원)",
                            "📊 진도율 비교 (%)",
                        ]
                    )

                    with tab_q:
                        fig_q = go.Figure()
                        fig_q.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["plan_qty"] for r in filtered_rows],
                                name=f"{m_lbl} 계획",
                                marker_color="#d1d5db",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[format_k(r["plan_qty"]) if r["plan_qty"] > 0 else "" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_q.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["target_qty"] for r in filtered_rows],
                                name="목표",
                                marker_color="#10b981",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[format_k(r["target_qty"]) if r["target_qty"] > 0 else "" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_q.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["act_qty"] for r in filtered_rows],
                                name="실적",
                                marker_color="#0044ff",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[format_k(r["act_qty"]) if r["act_qty"] > 0 else "" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_q.update_layout(
                            barmode="group",
                            plot_bgcolor="rgba(0,0,0,0)",
                            hovermode="x unified",
                            margin=dict(l=20, r=20, t=35, b=20),
                            legend=dict(orientation="h", y=1.08, x=1, xanchor="right"),
                        )
                        st.plotly_chart(fig_q, use_container_width=True)

                    with tab_a:
                        fig_a = go.Figure()
                        fig_a.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["plan_amt"] for r in filtered_rows],
                                name=f"{m_lbl} 계획(억)",
                                marker_color="#d1d5db",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[f"{r['plan_amt']:,.1f}" if r["plan_amt"] > 0 else "" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_a.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["target_amt"] for r in filtered_rows],
                                name="목표(억)",
                                marker_color="#10b981",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[f"{r['target_amt']:,.1f}" if r["target_amt"] > 0 else "" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_a.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["act_amt"] for r in filtered_rows],
                                name="실적(억)",
                                marker_color="#0044ff",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[f"{r['act_amt']:,.1f}" if r["act_amt"] > 0 else "" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_a.update_layout(
                            barmode="group",
                            plot_bgcolor="rgba(0,0,0,0)",
                            hovermode="x unified",
                            margin=dict(l=20, r=20, t=35, b=20),
                            legend=dict(orientation="h", y=1.08, x=1, xanchor="right"),
                        )
                        st.plotly_chart(fig_a, use_container_width=True)

                    with tab_r:
                        fig_r = go.Figure()
                        fig_r.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["rate_target"] for r in filtered_rows],
                                name="목표 진도율(%)",
                                marker_color="#9ca3af",
                                marker_line=dict(width=1.2, color="#111"),
                                text=[f"{r['rate_target']:,.0f}%" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_r.add_trace(
                            go.Bar(
                                x=x_labels,
                                y=[r["rate_act"] for r in filtered_rows],
                                name="실적 진도율(%)",
                                marker_color=["#ef4444" if r["rate_diff"] < 0 else "#0044ff" for r in filtered_rows],
                                marker_line=dict(width=1.2, color="#111"),
                                text=[f"{r['rate_act']:,.0f}%" for r in filtered_rows],
                                textposition="outside",
                            )
                        )
                        fig_r.update_layout(
                            barmode="group",
                            plot_bgcolor="rgba(0,0,0,0)",
                            hovermode="x unified",
                            margin=dict(l=20, r=20, t=35, b=20),
                            legend=dict(orientation="h", y=1.08, x=1, xanchor="right"),
                        )
                        st.plotly_chart(fig_r, use_container_width=True)

                    hdr_c1, hdr_c2 = st.columns([0.78, 0.22])
                    with hdr_c1:
                        st.markdown(f"##### 📋 과제 / 모델별 상세 출하 실적 ({len(filtered_rows)}개 모델)")
                    with hdr_c2:
                        df_export = pd.DataFrame(filtered_rows).rename(
                            columns={
                                "plan_qty": f"{m_lbl}계획_수량(개)",
                                "plan_amt": f"{m_lbl}계획_금액(억원)",
                                "target_qty": "목표_수량(개)",
                                "target_amt": "목표_금액(억원)",
                                "act_qty": "실적_수량(개)",
                                "act_amt": "실적_금액(억원)",
                                "diff_qty": "목표대비차이_수량(개)",
                                "diff_amt": "목표대비차이_금액(억원)",
                                "rate_target": "목표진도율(%)",
                                "rate_act": "실적진도율(%)",
                                "rate_diff": "진도율차이(%)",
                            }
                        )
                        st.download_button(
                            "📥 상세 표 CSV 다운로드",
                            data=convert_df(df_export),
                            file_name=f"shipment_summary_{m_lbl}.csv",
                            mime="text/csv",
                            use_container_width=True,
                        )
                    render_shipment_html_table(filtered_rows, is_summary=False, month_lbl=m_lbl, ref_lbl=ref_lbl)
            elif not sum_rows:
                st.info("파일 내 '월별요약' 시트 데이터가 없습니다. 우측 탭에서 일별 진도현황을 확인하세요.")

        # 탭 2: '11월 진도현황' 시트 기반 Daily / Weekly / Monthly 모델별 출하계획 vs 출하실적
        with main_tab2:
            if has_daily_prog:
                t_y, t_m, t_d = parse_target_ymd(st.session_state.get("target_date", default_target_date))
                cur_w = st.session_state.get("cur_week_label", get_iso_week_label(datetime.date(t_y, t_m, t_d)))
                p_label_map = {
                    "daily": f"📅 일간 ({t_m}/{t_d} 당일)",
                    "weekly": f"📊 주간 ({cur_w} 해당 주차)",
                    "monthly": f"🗓️ 월간 ({t_m}월 누계)",
                }
                sel_period = st.radio(
                    "조회 집계 범위",
                    options=["daily", "weekly", "monthly"],
                    format_func=lambda x: p_label_map[x],
                    horizontal=True,
                )
                rows_p = period_dict.get(sel_period, [])
                df_ship = pd.DataFrame(rows_p)
                if not df_ship.empty:
                    df_ship = df_ship.sort_values(by=["출하실적(개)", "출하계획(개)"], ascending=False)
                    sub_tq, sub_tr = st.tabs(["🚚 출하 수량 (계획 vs 실적)", "💰 매출액 (계획 vs 실적)"])
                    with sub_tq:
                        st.plotly_chart(
                            create_active_chart(
                                f"모델별 출하 수량 - {p_label_map[sel_period]}",
                                df_ship["모델명"].tolist(),
                                df_ship["출하계획(개)"].tolist(),
                                df_ship["출하실적(개)"].tolist(),
                                unit="K",
                            ),
                            use_container_width=True,
                        )
                    with sub_tr:
                        st.plotly_chart(
                            create_active_chart(
                                f"모델별 매출액 - {p_label_map[sel_period]}",
                                df_ship["모델명"].tolist(),
                                df_ship["매출계획(억원)"].tolist(),
                                df_ship["매출실적(억원)"].tolist(),
                                unit="억원",
                            ),
                            use_container_width=True,
                        )

                    c_d1, c_d2 = st.columns([0.8, 0.2])
                    with c_d2:
                        st.download_button(
                            label="📥 기간별 실적 CSV 다운로드",
                            data=convert_df(df_ship),
                            file_name=f"daily_progress_{sel_period}.csv",
                            mime="text/csv",
                            use_container_width=True,
                        )
                    st.dataframe(
                        df_ship[
                            [
                                "모델명",
                                "출하계획(K)",
                                "출하실적(K)",
                                "수량차이(K)",
                                "수량달성률(%)",
                                "매출계획(억원)",
                                "매출실적(억원)",
                                "매출차이(억원)",
                                "매출달성률(%)",
                            ]
                        ],
                        use_container_width=True,
                        hide_index=True,
                    )
                else:
                    st.info("선택한 기간에 출하 계획 또는 실적 물량이 있는 모델이 없습니다.")
            else:
                st.info("파일 내 '진도현황' 시트 데이터가 없습니다.")

elif menu == "2. 모델별 생산실적":
    st.subheader("🛠️ 2. 모델별 생산실적")
    cdata = st.session_state["chart_data"]
    st.plotly_chart(
        create_active_chart(
            "모델별 생산 현황 (계획 vs 실적)",
            cdata.get("daily_models", []),
            cdata.get("daily_prod_p", []),
            cdata.get("daily_prod_a", []),
            unit="K",
        ),
        use_container_width=True,
    )

    df_prod = pd.DataFrame(
        {
            "모델명": cdata.get("daily_models", []),
            "생산계획": cdata.get("daily_prod_p", []),
            "생산실적": cdata.get("daily_prod_a", []),
        }
    )
    if not df_prod.empty:
        df_prod["생산계획(K)"] = df_prod["생산계획"].apply(format_k)
        df_prod["생산실적(K)"] = df_prod["생산실적"].apply(format_k)
        df_prod["달성률(%)"] = df_prod.apply(
            lambda r: round(r["생산실적"] / r["생산계획"] * 100.0, 1) if r["생산계획"] > 0 else 0.0, axis=1
        )
        df_prod["차질수량(K)"] = (df_prod["생산계획"] - df_prod["생산실적"]).apply(format_k)
        col1, col2 = st.columns([0.8, 0.2])
        with col2:
            st.download_button(
                label="📥 엑셀(CSV) 다운로드",
                data=convert_df(df_prod),
                file_name="model_production.csv",
                mime="text/csv",
            )
        st.dataframe(
            df_prod[["모델명", "생산계획(K)", "생산실적(K)", "달성률(%)", "차질수량(K)"]],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("💡 **[6. 데이터 관리 (Admin)] → 1. 생산 실적 파일**을 업로드해 주세요.")

elif menu == "3. 고객 FCST 조회":
    st.title("📈 3. 고객 FCST 조회 (모델별 공급 일정 — 매출 참고)")
    fcst_data = st.session_state.get("fcst_data", {"rows": [], "date_cols": []})
    fcst_rows = fcst_data.get("rows", [])
    date_cols = fcst_data.get("date_cols", [])

    if not fcst_rows:
        st.info("💡 **[6. 데이터 관리 (Admin)] → 2. 출하 · 매출 통합 파일**('매출 참고' 시트 포함)을 업로드 및 반영해 주세요.")
    else:
        clean_rows = [{k: v for k, v in r.items() if k != "_by_ymd"} for r in fcst_rows]
        df_fcst = pd.DataFrame(clean_rows)

        total_fcst_qty = int(df_fcst["합계"].sum())
        model_count = int(df_fcst["모델명"].nunique())
        site_count = int(df_fcst["SITE"].nunique())
        top_model = df_fcst.groupby("모델명")["합계"].sum().idxmax()
        top_model_qty = int(df_fcst.groupby("모델명")["합계"].sum().max())

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("📦 총 공급 계획 물량", f"{format_k(total_fcst_qty)}")
        with m2:
            st.metric("📱 계획 대상 모델 수", f"{model_count} 개 모델")
        with m3:
            st.metric("🏭 공급 대상 Plant(SITE)", f"{site_count} 개 거점")
        with m4:
            st.metric("👑 최대 공급 모델", f"{top_model}", f"{format_k(top_model_qty)}")

        st.markdown("<br>", unsafe_allow_html=True)
        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        with f1:
            all_models = sorted(df_fcst["모델명"].unique().tolist())
            selected_models = st.multiselect("🔍 모델명 필터 (다중 선택 가능)", options=all_models, default=[])
        with f2:
            all_sites = sorted(df_fcst["SITE"].unique().tolist())
            selected_sites = st.multiselect("🏢 Plant(SITE) 필터", options=all_sites, default=[])
        with f3:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            hide_zero_cols = st.checkbox("공급 물량이 있는 날짜만 표시", value=True)

        df_filtered = df_fcst.copy()
        if selected_models:
            df_filtered = df_filtered[df_filtered["모델명"].isin(selected_models)]
        if selected_sites:
            df_filtered = df_filtered[df_filtered["SITE"].isin(selected_sites)]

        if hide_zero_cols:
            valid_dates = [c for c in date_cols if c in df_filtered.columns and df_filtered[c].sum() > 0]
        else:
            valid_dates = [c for c in date_cols if c in df_filtered.columns]

        tab_by_model, tab_by_detail, tab_charts = st.tabs(
            ["📊 모델별 공급 일정 (통합)", "📋 Material(GH) / Plant별 상세 일정", "📈 공급 계획 추이 차트"]
        )

        with tab_by_model:
            st.markdown("##### 📌 모델별 공급 계획 일정 (일자별 합계, 단위: K)")
            if not df_filtered.empty and valid_dates:
                df_model_pivot = df_filtered.groupby("모델명")[valid_dates].sum()
                df_model_pivot["총 공급계획(개)"] = df_model_pivot.sum(axis=1)
                df_model_pivot = df_model_pivot.sort_values(by="총 공급계획(개)", ascending=False).reset_index()

                df_model_disp = df_model_pivot.copy()
                df_model_disp["총 공급계획"] = df_model_disp["총 공급계획(개)"].apply(format_k)
                for col_d in valid_dates:
                    df_model_disp[col_d] = df_model_disp[col_d].apply(format_k)

                cols_order = ["모델명", "총 공급계획"] + valid_dates
                df_model_disp = df_model_disp[cols_order]

                c_btn1, c_btn2 = st.columns([0.8, 0.2])
                with c_btn2:
                    st.download_button(
                        label="📥 모델별 일정 CSV 다운로드",
                        data=convert_df(df_model_pivot),
                        file_name="fcst_supply_by_model.csv",
                        mime="text/csv",
                    )
                st.dataframe(df_model_disp, use_container_width=True, hide_index=True)
            else:
                st.warning("선택한 조건에 해당하는 공급 일정이 없습니다.")

        with tab_by_detail:
            st.markdown("##### 📌 Material(GH-CODE) 및 Plant(SITE)별 상세 공급 일정 (단위: K)")
            if not df_filtered.empty and valid_dates:
                base_cols = ["모델명", "과제명", "GH_CODE", "SITE", "합계"] if "과제명" in df_filtered.columns else ["모델명", "GH_CODE", "SITE", "합계"]
                detail_cols = base_cols + valid_dates
                df_detail = df_filtered[detail_cols].sort_values(by="합계", ascending=False)

                df_detail_disp = df_detail.copy()
                df_detail_disp["합계"] = df_detail_disp["합계"].apply(format_k)
                for col_d in valid_dates:
                    df_detail_disp[col_d] = df_detail_disp[col_d].apply(format_k)

                c_btn1, c_btn2 = st.columns([0.8, 0.2])
                with c_btn2:
                    st.download_button(
                        label="📥 상세 일정 CSV 다운로드",
                        data=convert_df(df_detail),
                        file_name="fcst_supply_detailed.csv",
                        mime="text/csv",
                    )
                st.dataframe(df_detail_disp, use_container_width=True, hide_index=True)
            else:
                st.warning("선택한 조건에 해당하는 공급 일정이 없습니다.")

        with tab_charts:
            st.markdown("##### 📌 모델별 공급 물량 및 일자별 공급 추이")
            ch_col1, ch_col2 = st.columns(2)

            with ch_col1:
                top10_models = df_filtered.groupby("모델명")["합계"].sum().sort_values(ascending=False).head(10)
                fig_top10 = go.Figure()
                fig_top10.add_trace(
                    go.Bar(
                        x=top10_models.index.tolist(),
                        y=top10_models.values.tolist(),
                        marker_color="#0044ff",
                        marker_line=dict(width=1.5, color="#111111"),
                        text=[format_k(v) for v in top10_models.values],
                        textposition="outside",
                    )
                )
                fig_top10.update_layout(
                    title="<b>TOP 10 모델별 공급 계획 물량 (단위: K)</b>",
                    plot_bgcolor="rgba(0,0,0,0)",
                    hovermode="x unified",
                    margin=dict(l=20, r=20, t=45, b=20),
                )
                st.plotly_chart(fig_top10, use_container_width=True)

            with ch_col2:
                daily_trend = df_filtered[valid_dates].sum() if valid_dates else pd.Series(dtype=float)
                fig_trend = go.Figure()
                fig_trend.add_trace(
                    go.Bar(
                        x=daily_trend.index.tolist(),
                        y=daily_trend.values.tolist(),
                        marker_color="#10b981",
                        marker_line=dict(width=1.5, color="#111111"),
                        text=[format_k(v) if v > 0 else "" for v in daily_trend.values],
                        textposition="outside",
                    )
                )
                fig_trend.update_layout(
                    title="<b>일자별 전체 공급 계획 추이 (단위: K)</b>",
                    plot_bgcolor="rgba(0,0,0,0)",
                    hovermode="x unified",
                    margin=dict(l=20, r=20, t=45, b=20),
                )
                st.plotly_chart(fig_trend, use_container_width=True)

elif menu == "4. 완제품 재고 조회":
    st.subheader("🏢 4. 완제품 재고 조회 (전체 통합)")
    st.metric(
        "📦 당일 완제품 총 재고",
        f"{format_k(st.session_state['inventory_data']['qty'])}",
        f"재고일수: {st.session_state['inventory_data']['days']} 일",
        delta_color="off",
    )

elif menu == "5. 캘린더":
    st.title("🗓️ 5. 월간 공급 계획 및 출하 실적 캘린더")
    fcst_data = st.session_state.get("fcst_data", {"rows": [], "date_cols": []})
    fcst_rows = fcst_data.get("rows", [])
    plan_by_date = st.session_state.get("plan_by_date", {})
    actual_by_date = st.session_state.get("actual_by_date", {})
    loaded_yms = set(st.session_state.get("actual_months_loaded", []))

    if not fcst_rows and not actual_by_date and not plan_by_date:
        st.info("💡 **[6. 데이터 관리 (Admin)] → 2. 출하 · 매출 통합 파일**을 업로드 및 분석해 주세요.")
    else:
        t_y, t_m, t_d = parse_target_ymd(st.session_state.get("target_date", default_target_date))
        target_date_obj = datetime.date(t_y, t_m, t_d)

        ym_set = set()
        for delta_m in (-1, 0, 1, 2):
            total_m = (t_y * 12 + (t_m - 1)) + delta_m
            ym_set.add((total_m // 12, (total_m % 12) + 1))

        for r_item in fcst_rows:
            for ymd_k in r_item.get("_by_ymd", {}).keys():
                try:
                    parts = ymd_k.split("-")
                    ym_set.add((int(parts[0]), int(parts[1])))
                except Exception:
                    pass
        for ymd_k in list(actual_by_date.keys()) + list(plan_by_date.keys()):
            try:
                parts = ymd_k.split("-")
                ym_set.add((int(parts[0]), int(parts[1])))
            except Exception:
                pass

        ym_options = sorted(list(ym_set))
        default_ym = (t_y, t_m) if (t_y, t_m) in ym_options else ym_options[0]

        c_m1, c_m2 = st.columns([1.2, 2.8])
        with c_m1:
            selected_ym = st.selectbox(
                "📅 조회 대상 연월 선택",
                options=ym_options,
                index=ym_options.index(default_ym),
                format_func=lambda x: f"{x[0]}년 {x[1]}월 (기준월)" if x == (t_y, t_m) else f"{x[0]}년 {x[1]}월",
            )

        cur_year, selected_m = selected_ym
        cur_ym_str = f"{cur_year:04d}-{selected_m:02d}"
        has_month_actuals = cur_ym_str in loaded_yms or any(k.startswith(cur_ym_str) for k in actual_by_date.keys())

        cal_instance = calendar.Calendar(firstweekday=6)
        month_matrix = cal_instance.monthdayscalendar(cur_year, selected_m)
        num_days = calendar.monthrange(cur_year, selected_m)[1]

        # FCST('매출 참고') 공급계획 집계 + '진도현황' 출하계획 보완
        fcst_day_map = {d: {"total": 0, "total_rev": 0.0, "items": []} for d in range(1, num_days + 1)}
        for r_item in fcst_rows:
            m_name = r_item["모델명"]
            gh = r_item["GH_CODE"]
            site = r_item["SITE"]
            price = r_item.get("단가_원", 0.0)
            by_ymd = r_item.get("_by_ymd", {})

            for d in range(1, num_days + 1):
                ymd_key = f"{cur_year:04d}-{selected_m:02d}-{d:02d}"
                q = by_ymd.get(ymd_key, 0) if by_ymd else 0
                if q > 0:
                    rev_val = (q * price) / 100_000_000.0 if price > 0 else 0.0
                    fcst_day_map[d]["total"] += q
                    fcst_day_map[d]["total_rev"] += rev_val
                    fcst_day_map[d]["items"].append(
                        {"model": m_name, "gh_code": gh, "site": site, "qty": q, "price": price, "rev_ok": rev_val}
                    )

        has_month_prog_plan = any(
            k.startswith(cur_ym_str) and v.get("qty", 0) > 0 for k, v in plan_by_date.items()
        )
        day_supply_map = {}
        for d in range(1, num_days + 1):
            ymd_key = f"{cur_year:04d}-{selected_m:02d}-{d:02d}"
            prog_p = plan_by_date.get(ymd_key)
            fcst_p = fcst_day_map[d]
            # 해당 월에 '진도현황' 출하계획이 존재하면 월 전체 일자에 동일 기준(진도현황 출하계획)을 일관 적용하고, 없는 월(익월 등)은 '매출 참고' FCST 적용
            if has_month_prog_plan:
                if prog_p and (prog_p["qty"] > 0 or ymd_key in plan_by_date):
                    day_supply_map[d] = {"total": prog_p["qty"], "total_rev": prog_p["rev"], "items": prog_p["items"]}
                else:
                    day_supply_map[d] = fcst_p
            else:
                day_supply_map[d] = fcst_p

        m_total_qty = sum(day_supply_map[d]["total"] for d in day_supply_map)
        m_total_rev = sum(day_supply_map[d]["total_rev"] for d in day_supply_map)
        supply_days_cnt = sum(1 for d in day_supply_map if day_supply_map[d]["total"] > 0)

        m_model_totals = {}
        for d in day_supply_map:
            for it in day_supply_map[d]["items"]:
                m_model_totals[it["model"]] = m_model_totals.get(it["model"], 0) + it["qty"]

        top_m_name = max(m_model_totals, key=m_model_totals.get) if m_model_totals else "-"
        top_m_qty = m_model_totals.get(top_m_name, 0)

        with c_m2:
            render_html(
                f"""
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
            """
            )

        st.markdown("<br>", unsafe_allow_html=True)

        if st.session_state.get("active_cal_popup"):
            p_day, p_month, p_year, p_act_mode = st.session_state["active_cal_popup"]
            if p_year == cur_year and p_month == selected_m:
                p_ymd = f"{p_year:04d}-{p_month:02d}-{p_day:02d}"
                p_info = day_supply_map.get(p_day, {"total": 0, "total_rev": 0.0, "items": []})
                p_act = actual_by_date.get(p_ymd, {"qty": 0, "rev": 0.0, "items": []})
                with st.container():
                    st.markdown(f"#### 📅 {p_year}년 {p_month}월 {p_day}일 상세 내역")
                    _render_day_detail_body(p_day, p_month, p_year, p_info, p_act, p_act_mode)
                    if st.button("✖️ 상세 닫기", key="btn_close_fallback_popup"):
                        st.session_state["active_cal_popup"] = None
                        st.rerun()
                    st.divider()

        days_headers = [
            ("<span style='color: #ef4444;'>일 (SUN)</span>", "#ef4444"),
            ("월 (MON)", "#111111"),
            ("화 (TUE)", "#111111"),
            ("수 (WED)", "#111111"),
            ("목 (THU)", "#111111"),
            ("금 (FRI)", "#111111"),
            ("<span style='color: #0044ff;'>토 (SAT)</span>", "#0044ff"),
        ]

        h_cols = st.columns(7)
        for i, (h_title, _) in enumerate(days_headers):
            with h_cols[i]:
                render_html(
                    f"""
                <div style="background-color: #ffffff; border: 2px solid #111111; padding: 6px; text-align: center; font-weight: 900; font-size: 13px; box-shadow: 2px 2px 0px #111111; margin-bottom: 6px; border-radius: 4px;">
                    {h_title}
                </div>
                """
                )

        for week in month_matrix:
            w_cols = st.columns(7)
            for day_idx, d_num in enumerate(week):
                with w_cols[day_idx]:
                    if d_num == 0:
                        render_html(
                            "<div style='height: 112px; background: rgba(0,0,0,0.02); border: 1px dashed #d1d5db; border-radius: 4px; margin-bottom: 6px;'></div>"
                        )
                    else:
                        cell_date = datetime.date(cur_year, selected_m, d_num)
                        ymd_key = cell_date.strftime("%Y-%m-%d")

                        d_info = day_supply_map[d_num]
                        plan_q = d_info["total"]
                        plan_rev = d_info["total_rev"]

                        is_target = cell_date == target_date_obj
                        is_past_or_target = cell_date <= target_date_obj
                        show_actual_mode = is_past_or_target and has_month_actuals

                        act_day_data = actual_by_date.get(ymd_key, {"qty": 0, "rev": 0.0, "items": []})
                        act_q = act_day_data["qty"]
                        act_rev = act_day_data["rev"]

                        if is_target:
                            border_css = "3px solid #0044ff; box-shadow: 4px 4px 0px #0044ff;"
                            bg_css = "#f0f5ff"
                            badge_target = "<span style='background: #0044ff; color: #ffffff; font-size: 10px; font-weight: 800; padding: 1px 4px; border-radius: 3px; margin-left: 4px;'>기준일</span>"
                        else:
                            border_css = "2px solid #111111; box-shadow: 3px 3px 0px #111111;"
                            bg_css = "#ffffff" if (plan_q > 0 or act_q > 0) else "#fafafa"
                            badge_target = ""

                        if day_idx == 0:
                            day_color = "#ef4444"
                        elif day_idx == 6:
                            day_color = "#0044ff"
                        else:
                            day_color = "#111111"

                        if show_actual_mode:
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

                        hover_source_items = act_day_data["items"] if (show_actual_mode and act_day_data["items"]) else d_info["items"]
                        unique_models_cnt = len(
                            set(it["model"] for it in (d_info["items"] + (act_day_data["items"] if show_actual_mode else [])))
                        )
                        items_sorted = sorted(hover_source_items, key=lambda x: x["qty"], reverse=True)

                        hover_list_html = ""
                        for it in items_sorted[:6]:
                            q_txt = format_k(it["qty"])
                            hover_list_html += (
                                f"<div style='font-size: 11px; color: #111827; padding: 2px 0; border-bottom: 1px dotted #e5e7eb; "
                                f"display: flex; justify-content: space-between;'><span>• {it['model']} ({it['site']})</span>"
                                f"<span><b>{q_txt}</b> / {it['rev_ok']:,.2f}억</span></div>"
                            )
                        if len(items_sorted) > 6:
                            hover_list_html += (
                                f"<div style='font-size: 10px; color: #6b7280; font-weight: 800; text-align: center; margin-top: 4px;'>"
                                f"+ 외 {len(items_sorted) - 6}개 품목 더 있음</div>"
                            )

                        if day_idx <= 1:
                            pos_class = "pos-left"
                        elif day_idx >= 5:
                            pos_class = "pos-right"
                        else:
                            pos_class = "pos-center"

                        if show_actual_mode:
                            hover_hdr_title = f"📅 {selected_m}월 {d_num}일 출하 실적 / 계획 요약"
                            hover_kpi_left = ("총 실적수량 (계획)", f"{format_k(act_q)} ({format_k(plan_q)})")
                            hover_kpi_right = ("총 실적매출 (계획)", f"{act_rev:,.2f}억 ({plan_rev:,.2f}억)")
                        else:
                            hover_hdr_title = f"📅 {selected_m}월 {d_num}일 공급 스케줄 (전체 {unique_models_cnt}종)"
                            hover_kpi_left = ("총 계획수량", format_k(plan_q))
                            hover_kpi_right = ("총 예상매출", f"{plan_rev:,.2f}억원")

                        hover_popup_html = f"""
                        <div class="cal-hover-popup {pos_class}">
                            <div style="font-size: 13px; font-weight: 900; color: #111111; border-bottom: 2px solid #111111; padding-bottom: 4px; margin-bottom: 8px;">
                                {hover_hdr_title}
                            </div>
                            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px; margin-bottom: 8px;">
                                <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 5px; border-radius: 3px; text-align: center;">
                                    <div style="font-size: 10px; font-weight: 800; color: #6b7280;">{hover_kpi_left[0]}</div>
                                    <div style="font-size: 13px; font-weight: 900; color: #0044ff;">{hover_kpi_left[1]}</div>
                                </div>
                                <div style="background: #f9fafb; border: 1.5px solid #111111; padding: 5px; border-radius: 3px; text-align: center;">
                                    <div style="font-size: 10px; font-weight: 800; color: #6b7280;">{hover_kpi_right[0]}</div>
                                    <div style="font-size: 13px; font-weight: 900; color: #10b981;">{hover_kpi_right[1]}</div>
                                </div>
                            </div>
                            <div style="max-height: 140px; overflow-y: auto;">
                                {hover_list_html if hover_list_html else "<div style='font-size: 11px; color: #9ca3af; text-align: center;'>내역 없음</div>"}
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
                        render_html(full_cell_html)

                        btn_label = f"🔍 상세 ({unique_models_cnt}종)" if (plan_q > 0 or act_q > 0) else f"🔍 {d_num}일"
                        if st.button(btn_label, key=f"btn_pop_{cur_year}_{selected_m}_{d_num}", use_container_width=True):
                            if hasattr(st, "dialog"):
                                show_day_detail_dialog(d_num, selected_m, cur_year, d_info, act_day_data, show_actual_mode)
                            else:
                                st.session_state["active_cal_popup"] = (d_num, selected_m, cur_year, show_actual_mode)
                                st.rerun()

        st.divider()

        st.markdown(f"### 📈 {cur_year}년 {selected_m}월 일자별 공급 계획 및 출하 추이")
        trend_days = [f"{selected_m}/{d:02d}" for d in range(1, num_days + 1)]
        trend_plan_qtys = [day_supply_map[d]["total"] for d in range(1, num_days + 1)]
        trend_act_qtys = [
            actual_by_date.get(f"{cur_year:04d}-{selected_m:02d}-{d:02d}", {}).get("qty", 0)
            for d in range(1, num_days + 1)
        ]
        trend_revs = [round(day_supply_map[d]["total_rev"], 2) for d in range(1, num_days + 1)]

        fig_cal_trend = go.Figure()
        fig_cal_trend.add_trace(
            go.Bar(
                x=trend_days,
                y=trend_plan_qtys,
                name="계획수량(K)",
                marker_color="#d1d5db",
                marker_line=dict(width=1.3, color="#111111"),
                text=[format_k(q) if q > 0 else "" for q in trend_plan_qtys],
                textposition="outside",
            )
        )
        if sum(trend_act_qtys) > 0:
            fig_cal_trend.add_trace(
                go.Bar(
                    x=trend_days,
                    y=trend_act_qtys,
                    name="출하실적(K)",
                    marker_color="#0044ff",
                    marker_line=dict(width=1.3, color="#111111"),
                    text=[format_k(q) if q > 0 else "" for q in trend_act_qtys],
                    textposition="outside",
                )
            )
        fig_cal_trend.add_trace(
            go.Scatter(
                x=trend_days,
                y=trend_revs,
                name="계획매출(억원)",
                yaxis="y2",
                mode="lines+markers",
                line=dict(color="#ef4444", width=2.5),
                marker=dict(size=6, color="#ef4444"),
            )
        )
        fig_cal_trend.update_layout(
            title=f"<b>{cur_year}년 {selected_m}월 일자별 계획 vs 실적 수량 (막대) 및 예상 매출액 (꺾은선)</b>",
            barmode="group",
            plot_bgcolor="rgba(0,0,0,0)",
            hovermode="x unified",
            margin=dict(l=20, r=20, t=45, b=20),
            xaxis=dict(tickangle=-45),
            yaxis=dict(title="수량 (개)"),
            yaxis2=dict(title="매출액 (억원)", overlaying="y", side="right"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_cal_trend, use_container_width=True)
