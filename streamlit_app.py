import os
import re
import json
import shutil
import datetime
import calendar
import openpyxl
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# 1. 페이지 설정 및 영구 저장소 경로 설정
st.set_page_config(page_title="마케팅 부문 대시보드", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
STORE_DIR = os.path.join(BASE_DIR, ".dashboard_store")
STATE_JSON_PATH = os.path.join(STORE_DIR, "dashboard_state.json")
CACHED_EXCEL_PATHS = {
    "map_rev": os.path.join(STORE_DIR, "latest_map_rev.xlsx"),
    "prod": os.path.join(STORE_DIR, "latest_prod.xlsx"),
    "ship_actual": os.path.join(STORE_DIR, "latest_ship_actual.xlsx"),
    "ship_plan": os.path.join(STORE_DIR, "latest_ship_plan.xlsx"),
}

today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
default_target_date = f"{yesterday.month}/{yesterday.day}"


def render_html(html_str):
    if hasattr(st, "html"):
        st.html(html_str)
    else:
        st.markdown(html_str, unsafe_allow_html=True)


# 2. 날짜·포맷·유틸리티 함수
def safe_cell(sheet, r, c):
    if r < 1 or c < 1 or r > sheet.max_row or c > sheet.max_column:
        return None
    return sheet.cell(r, c).value


def clean_key(val):
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).upper())


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

    # YYYY-MM-DD, YYYY/MM/DD, YY.MM.DD 형식 지원
    m_full = re.search(r"\b(\d{2,4})\s*[-/.]\s*0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s)
    if m_full:
        y_raw = int(m_full.group(1))
        y = (2000 + y_raw) if y_raw < 100 else y_raw
        dt = clamp_date(y, int(m_full.group(2)), int(m_full.group(3)))
        return dt.year, dt.month, dt.day

    # M/D, M-D, M.D 형식 지원
    m_md = re.search(r"\b0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s)
    if m_md:
        mon = int(m_md.group(1))
        day = int(m_md.group(2))
        y = infer_year_for_month(mon, ref)
        dt = clamp_date(y, mon, day)
        return dt.year, dt.month, dt.day

    # D일 단독 형식 지원
    m_d = re.search(r"(\d{1,2})일?", s)
    if m_d:
        dt = clamp_date(ref.year, ref.month, int(m_d.group(1)))
        return dt.year, dt.month, dt.day

    return ref.year, ref.month, ref.day


def parse_target_day_info(target_str):
    _, m, d = parse_target_ymd(target_str)
    return m, d


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
    if gh_clean in gh_to_model:
        return gh_to_model[gh_clean]
    if gh_suf in gh_to_model:
        return gh_to_model[gh_suf]
    m = re.search(r"CAM_([A-Z0-9]+)", str(fallback_model).upper())
    if m:
        extracted = m.group(1)
        return extracted[:-1] if extracted.endswith("E") else extracted
    clean_m = str(fallback_model)[:-1] if str(fallback_model).endswith("E") else str(fallback_model)
    return clean_m.strip()


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


# 3. 기본 상태 생성 및 서버 영구 저장/복구 로직
def get_default_state():
    t_y, t_m, t_d = parse_target_ymd(default_target_date)
    init_weeks = get_month_iso_weeks(t_y, t_m)
    cur_w = get_iso_week_label(datetime.date(t_y, t_m, t_d))
    return {
        "target_date": default_target_date,
        "ex_rate": 1350.0,
        "cur_week_label": cur_w,
        "last_updated_at": None,
        "data_loaded": False,
        "mapping_data": {"gh_map": {}, "suffix_map": {}, "model_map": {}, "gh_to_model": {}},
        "fcst_data": {"rows": [], "date_cols": []},
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
                "month_carryover_total": 0.0,
            },
            "monthly": {
                "prod_p": 0,
                "prod_a": 0,
                "ship_p": 0,
                "ship_a": 0,
                "rev_p": 0.0,
                "rev_a": 0.0,
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
            "ship_by_site": [],
            "ship_by_model_period": {"daily": [], "weekly": [], "monthly": []},
            "ship_by_site_period": {"daily": [], "weekly": [], "monthly": []},
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
    payload = {k: st.session_state.get(k) for k in PERSIST_KEYS}
    tmp_path = STATE_JSON_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=_json_default)
    os.replace(tmp_path, STATE_JSON_PATH)
    st.session_state["_loaded_mtime"] = os.path.getmtime(STATE_JSON_PATH)


def load_state_from_disk():
    if not os.path.exists(STATE_JSON_PATH):
        return False
    try:
        with open(STATE_JSON_PATH, "r", encoding="utf-8") as f:
            saved = json.load(f)
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
        return True
    except Exception:
        return False


def init_and_sync_session():
    defaults = get_default_state()
    if "active_cal_popup" not in st.session_state:
        st.session_state["active_cal_popup"] = None

    disk_mtime = os.path.getmtime(STATE_JSON_PATH) if os.path.exists(STATE_JSON_PATH) else None
    loaded_mtime = st.session_state.get("_loaded_mtime")

    if disk_mtime is not None and (loaded_mtime is None or disk_mtime > loaded_mtime):
        if load_state_from_disk():
            return

    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_persisted_store():
    if os.path.exists(STORE_DIR):
        shutil.rmtree(STORE_DIR, ignore_errors=True)
    defaults = get_default_state()
    for k, v in defaults.items():
        st.session_state[k] = v
    st.session_state["_loaded_mtime"] = None
    st.session_state["active_cal_popup"] = None


init_and_sync_session()


# 4. UI 컴포넌트 헬퍼 함수
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

        t_act, t_plan = st.tabs(["✅ 출하 실적 상세 내역", "📋 공급 계획(FCST) 상세 내역"])
        with t_act:
            if act_items:
                df_act = pd.DataFrame(act_items).sort_values(by="qty", ascending=False)
                df_act["출하수량"] = df_act["qty"].apply(format_k)
                df_act["적용단가(원)"] = df_act["price"].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
                df_act["실적매출(억원)"] = df_act["rev_ok"].apply(lambda x: f"{x:,.2f}억원")
                df_act = df_act.rename(columns={"model": "모델명", "gh_code": "GH-CODE", "site": "출하 SITE"})
                cols_a = ["모델명", "GH-CODE", "출하 SITE", "출하수량", "적용단가(원)", "실적매출(억원)"]
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
                df_plan["공급수량"] = df_plan["qty"].apply(format_k)
                df_plan["단가(원)"] = df_plan["price"].apply(lambda x: f"{int(round(x)):,}원" if x > 0 else "-")
                df_plan["예상매출(억원)"] = df_plan["rev_ok"].apply(lambda x: f"{x:,.2f}억원")
                df_plan = df_plan.rename(columns={"model": "모델명", "gh_code": "GH-CODE", "site": "출하 SITE"})
                cols_p = ["모델명", "GH-CODE", "출하 SITE", "공급수량", "단가(원)", "예상매출(억원)"]
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


# 5. 엑셀 파싱 엔진 (거점 통합·다년도·기간별 집계 반영)
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

        gh_col = next((c for c in df.columns if "GH" in c), None)
        mdl_col = next((c for c in df.columns if "모델" in c or "MODEL" in c), None)
        prc_col = next(
            (c for c in df.columns if "단가" in c or "PRICE" in c or "금액" in c or "USD" in c or "$" in c),
            None,
        )

        mapping = {"gh_map": {}, "suffix_map": {}, "model_map": {}, "gh_to_model": {}}
        if prc_col:
            for _, row in df.iterrows():
                p_val = str(row[prc_col]).replace("$", "").replace(",", "").strip()
                try:
                    price = float(p_val)
                except Exception:
                    price = 0.0

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
        "monthly_months": [f"{t_m}월"],
        "inventory_qty": 0,
        "inventory_days": 0.0,
        "daily_models": [],
        "daily_prod_p": [],
        "daily_prod_a": [],
    }

    summary_sheet_name = next((s for s in wb.sheetnames if "요약" in s), wb.sheetnames[0])
    sheet = wb[summary_sheet_name]

    # 완제품 재고 추출 (전체/합계 열 우선 선택)
    for r in range(1, 35):
        row_joined = "".join([str(safe_cell(sheet, r, c) or "") for c in range(1, 10)]).replace(" ", "")
        if "당일완제품재고" in row_joined:
            inv_col = None
            for c in range(2, 10):
                hdr = str(safe_cell(sheet, r + 1, c) or "").replace(" ", "").upper()
                if any(k in hdr for k in ["전체", "합계", "총계", "TOTAL"]):
                    inv_col = c
                    break
            try:
                raw_q = safe_cell(sheet, r + 2, inv_col) if inv_col else (safe_cell(sheet, r + 2, 2) or safe_cell(sheet, r + 2, 3))
                raw_d = safe_cell(sheet, r + 3, inv_col) if inv_col else (safe_cell(sheet, r + 3, 2) or safe_cell(sheet, r + 3, 3))
                parsed["inventory_qty"] = int(float(str(raw_q or 0).replace(",", "")))
                parsed["inventory_days"] = float(str(raw_d or 0.0).replace(",", ""))
            except Exception:
                pass
            break

    summary_start = 6
    for r in range(1, 25):
        if "구분" in "".join([str(safe_cell(sheet, r, c) or "") for c in range(1, 10)]).replace(" ", ""):
            summary_start = r
            break

    col_map = {}
    weekly_cols = {}
    w_start_col = None
    w_first_label = get_month_iso_weeks(t_y, t_m)[0]

    for r in range(summary_start, summary_start + 3):
        for c in range(1, sheet.max_column + 1):
            val = safe_cell(sheet, r, c)
            if not val:
                continue
            if is_match_date(val, target_date_str) and "target" not in col_map:
                col_map["target"] = c
            vs = str(val).replace(" ", "").upper()
            if re.fullmatch(r"W\d+", vs) and w_start_col is None:
                w_start_col = c
                w_first_label = vs
            if "W" in vs and "합계" in vs:
                m = re.search(r"(W\d+)", vs)
                if m and m.group(1) not in weekly_cols:
                    weekly_cols[m.group(1)] = c
            if "월합계" in vs and "월합계" not in col_map:
                col_map["월합계"] = c
                m = re.search(r"(\d+월)", vs)
                if m:
                    parsed["monthly_months"] = [m.group(1)]

    # 베트남 / 한국 / 해외 구분 행이 존재하더라도 전체 통합(TOTAL/합계 우선, 없으면 거점 합산)으로 처리
    total_plan_rows, total_act_rows = [], []
    all_plan_rows, all_act_rows = [], []
    curr_group = ""
    empty_streak = 0

    for r in range(summary_start + 1, min(sheet.max_row + 1, summary_start + 35)):
        row_text = "".join([str(safe_cell(sheet, r, c) or "") for c in range(1, 8)]).replace(" ", "")
        if not row_text:
            empty_streak += 1
            if empty_streak >= 3 and all_plan_rows:
                break
            continue
        empty_streak = 0
        if any(stop_k in row_text for stop_k in ["당일완제품재고", "주요모델별"]):
            break

        val_a = str(safe_cell(sheet, r, 1) or "").replace(" ", "").strip()
        val_b = str(safe_cell(sheet, r, 2) or "").replace(" ", "").strip()

        if val_a and val_a not in ("계획", "실적", "달성률", "차이", "차질"):
            curr_group = val_a.upper()

        is_total_grp = any(k in curr_group for k in ["전체", "합계", "총계", "TOTAL", "전사", "통합"])

        if val_a == "계획" or val_b == "계획":
            all_plan_rows.append(r)
            if is_total_grp:
                total_plan_rows.append(r)
        elif val_a == "실적" or val_b == "실적":
            all_act_rows.append(r)
            if is_total_grp:
                total_act_rows.append(r)

    target_plan_rows = total_plan_rows[:1] if total_plan_rows else all_plan_rows
    target_act_rows = total_act_rows[:1] if total_act_rows else all_act_rows

    def get_single_cell_val(r, c):
        if not r or not c:
            return 0
        try:
            v = float(str(safe_cell(sheet, r, c) or 0).replace(",", "").replace("-", "0"))
            if 0 < abs(v) < 10000:
                v *= 1000
            return int(round(v))
        except Exception:
            return 0

    def get_rows_val(rows, c):
        if not rows or not c:
            return 0
        return sum(get_single_cell_val(r, c) for r in rows)

    parsed["daily_prod_plan"] = get_rows_val(target_plan_rows, col_map.get("target"))
    parsed["daily_prod_actual"] = get_rows_val(target_act_rows, col_map.get("target"))

    labels, wp, wa = [], [], []
    seen_labels = set()
    col_bounds_for_week = []

    if w_start_col and weekly_cols:
        min_w_col = min(weekly_cols.values())
        w_first_plan, w_first_act = 0, 0
        for c in range(w_start_col, min_w_col):
            w_first_plan += get_rows_val(target_plan_rows, c)
            w_first_act += get_rows_val(target_act_rows, c)
        labels.append(w_first_label)
        wp.append(w_first_plan)
        wa.append(w_first_act)
        seen_labels.add(w_first_label)
        col_bounds_for_week.append((w_first_label, min_w_col - 1))

    for lbl, c_idx in sorted(weekly_cols.items(), key=lambda x: x[1]):
        if lbl in seen_labels:
            continue
        labels.append(lbl)
        wp.append(get_rows_val(target_plan_rows, c_idx))
        wa.append(get_rows_val(target_act_rows, c_idx))
        seen_labels.add(lbl)
        col_bounds_for_week.append((lbl, c_idx))

    fallback_weeks = get_month_iso_weeks(t_y, t_m)
    parsed["weekly_labels"] = labels if labels else fallback_weeks
    parsed["weekly_plan"] = wp if wp else [0] * len(parsed["weekly_labels"])
    parsed["weekly_actual"] = wa if wa else [0] * len(parsed["weekly_labels"])

    # 조회일 기준 해당 주차 합계만 산출
    week_plan_map = dict(zip(parsed["weekly_labels"], parsed["weekly_plan"]))
    week_act_map = dict(zip(parsed["weekly_labels"], parsed["weekly_actual"]))

    matched_week_lbl = target_iso_week
    if matched_week_lbl not in week_plan_map and col_map.get("target") and col_bounds_for_week:
        t_col = col_map["target"]
        for w_lbl, end_col in col_bounds_for_week:
            if t_col <= end_col:
                matched_week_lbl = w_lbl
                break

    parsed["cur_week_label"] = matched_week_lbl
    parsed["cur_week_prod_p"] = week_plan_map.get(matched_week_lbl, 0)
    parsed["cur_week_prod_a"] = week_act_map.get(matched_week_lbl, 0)

    m_plan = get_rows_val(target_plan_rows, col_map.get("월합계"))
    m_act = get_rows_val(target_act_rows, col_map.get("월합계"))
    if m_act == 0 and sum(wa) > 0:
        m_act = sum(wa)
    parsed["monthly_plan"] = m_plan
    parsed["monthly_actual"] = m_act

    model_sheet = next((s for s in wb.sheetnames if "모델" in s), None)
    if model_sheet:
        ms = wb[model_sheet]
        hr, pr, ar = -1, -1, -1
        for r in range(1, 12):
            if "주요모델별" in "".join([str(safe_cell(ms, r, c) or "") for c in range(1, 6)]).replace(" ", ""):
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
                rt = "".join([str(safe_cell(ms, r, c) or "") for c in range(1, 4)]).replace(" ", "")
                if "당일계획" in rt and pr == -1:
                    pr = r
                if "당일실적" in rt and ar == -1:
                    ar = r
            if pr != -1 and ar != -1:
                for c in c_idx:
                    try:
                        parsed["daily_prod_p"].append(
                            int(float(str(safe_cell(ms, pr, c) or 0).replace(",", "").replace("-", "0")))
                        )
                    except Exception:
                        parsed["daily_prod_p"].append(0)
                    try:
                        parsed["daily_prod_a"].append(
                            int(float(str(safe_cell(ms, ar, c) or 0).replace(",", "").replace("-", "0")))
                        )
                    except Exception:
                        parsed["daily_prod_a"].append(0)
                parsed["daily_models"] = models
    return parsed


def _parse_single_shipment_sheet(sheet, sheet_year, sheet_month, target_date_obj, mapping_data):
    gh_col, model_col, site_col, gubun_col = None, None, None, None
    header_row = 1

    for r in range(1, 6):
        for c in range(1, 15):
            val = str(safe_cell(sheet, r, c) or "").replace(" ", "").upper()
            if any(k in val for k in ["품번", "GH"]) and not gh_col:
                gh_col = c
                header_row = r
            if any(k in val for k in ["모델명", "MODEL"]) and not model_col:
                model_col = c
                header_row = r
            if any(k in val for k in ["업체명", "SITE", "거래선"]) and not site_col:
                site_col = c
                header_row = r
            if any(k in val for k in ["구분", "TYPE"]) and not gubun_col:
                gubun_col = c
                header_row = r

    gh_col, model_col, site_col, gubun_col = gh_col or 1, model_col or 3, site_col or 4, gubun_col or 6

    carryover_col = 6
    for c in range(1, 12):
        v_h1 = str(safe_cell(sheet, 1, c) or "") + str(safe_cell(sheet, 2, c) or "")
        if ("이월" in v_h1 or "전월" in v_h1) and ("KRW" in v_h1 or "원화" in v_h1 or "금액" in v_h1):
            carryover_col = c
            break

    max_valid_day = calendar.monthrange(sheet_year, sheet_month)[1]
    day_map = {}
    for c in range(1, sheet.max_column + 1):
        for r in range(1, header_row + 2):
            v = safe_cell(sheet, r, c)
            if not v:
                continue
            if isinstance(v, (datetime.datetime, datetime.date)):
                if v.month == sheet_month and 1 <= v.day <= max_valid_day and v.day not in day_map:
                    d = v.day
                    day_map[d] = {"qty": c, "krw": c + 1}
                    break
            vs = str(v).replace(" ", "")
            m = re.search(r"^0?(\d{1,2})일?$", vs)
            if m:
                d = int(m.group(1))
                if 1 <= d <= max_valid_day and d not in day_map:
                    qty_c, krw_c = c, c + 1
                    for sub_c in [c, c + 1, c + 2]:
                        sub_v = str(safe_cell(sheet, r + 1, sub_c) or "").replace(" ", "").upper()
                        if "수량" in sub_v:
                            qty_c = sub_c
                        if "KRW" in sub_v or "원화" in sub_v:
                            krw_c = sub_c
                    day_map[d] = {"qty": qty_c, "krw": krw_c}
                    break

    target_iso_key = get_iso_week_key(target_date_obj)
    is_target_month = (sheet_year == target_date_obj.year and sheet_month == target_date_obj.month)

    daily_ship_a, daily_rev_a = 0, 0.0
    weekly_ship_actual, weekly_rev_actual = {}, {}
    weekly_ship_by_isokey, weekly_rev_by_isokey = {}, {}
    monthly_ship_actual, monthly_rev_actual = 0, 0.0

    model_period = {"daily": {}, "weekly": {}, "monthly": {}}
    site_period = {"daily": {}, "weekly": {}, "monthly": {}}

    learned_prices, learned_suffix_prices, learned_model_prices = {}, {}, {}
    total_carryover_krw = 0.0

    date_actual_map = {
        f"{sheet_year:04d}-{sheet_month:02d}-{d:02d}": {"qty": 0, "rev": 0.0, "items": []}
        for d in day_map
    }

    curr_gh, curr_model, curr_site = "", "", ""

    for r in range(header_row + 2, sheet.max_row + 1):
        v_gh = safe_cell(sheet, r, gh_col)
        v_mdl = safe_cell(sheet, r, model_col)
        v_ste = safe_cell(sheet, r, site_col)
        clean_gbn = str(safe_cell(sheet, r, gubun_col) or "").replace(" ", "").strip()

        if v_gh:
            curr_gh = str(v_gh).strip()
        if v_mdl:
            curr_model = str(v_mdl).strip()
        if v_ste:
            curr_site = str(v_ste).strip()

        val_f = str(safe_cell(sheet, r, carryover_col) or 0).replace(",", "").strip()
        try:
            f_krw = float(val_f) if val_f else 0.0
        except Exception:
            f_krw = 0.0

        row_concat = f"{v_gh or ''}{v_mdl or ''}{v_ste or ''}{clean_gbn}"
        is_summary = any(k in row_concat for k in ["합계", "소계", "누계", "지지품", "총계", "총합", "차이"])
        if f_krw > 0 and not is_summary:
            total_carryover_krw += f_krw

        if is_summary:
            continue

        if clean_gbn == "실적":
            base_model = curr_model[:-1] if curr_model.endswith("E") else curr_model
            gh_clean = clean_key(curr_gh)
            gh_suf = get_gh_suffix(curr_gh)

            for d, cols in day_map.items():
                q_raw = str(safe_cell(sheet, r, cols["qty"]) or 0).replace(",", "").replace("-", "0").strip()
                k_raw = str(safe_cell(sheet, r, cols["krw"]) or 0).replace(",", "").replace("-", "0").strip()
                try:
                    q_val = int(float(q_raw)) if q_raw else 0
                except Exception:
                    q_val = 0
                try:
                    k_val = float(k_raw) if k_raw else 0.0
                except Exception:
                    k_val = 0.0

                rev_ok = k_val / 100_000_000.0
                d_obj = datetime.date(sheet_year, sheet_month, d)
                ymd_key = d_obj.strftime("%Y-%m-%d")

                date_actual_map[ymd_key]["qty"] += q_val
                date_actual_map[ymd_key]["rev"] += rev_ok
                if q_val > 0:
                    date_actual_map[ymd_key]["items"].append(
                        {
                            "model": base_model,
                            "gh_code": curr_gh,
                            "site": curr_site,
                            "qty": q_val,
                            "price": (k_val / q_val) if q_val > 0 else 0.0,
                            "rev_ok": rev_ok,
                        }
                    )

                if q_val > 0 and k_val > 0:
                    u_price = k_val / q_val
                    learned_prices[gh_clean] = u_price
                    learned_suffix_prices[gh_suf] = u_price
                    if base_model:
                        learned_model_prices[clean_key(base_model)] = u_price

                w_lbl = get_iso_week_label(d_obj)
                w_key = get_iso_week_key(d_obj)

                if d_obj <= target_date_obj:
                    weekly_ship_by_isokey[w_key] = weekly_ship_by_isokey.get(w_key, 0) + q_val
                    weekly_rev_by_isokey[w_key] = weekly_rev_by_isokey.get(w_key, 0.0) + rev_ok

                    if is_target_month:
                        weekly_ship_actual[w_lbl] = weekly_ship_actual.get(w_lbl, 0) + q_val
                        weekly_rev_actual[w_lbl] = weekly_rev_actual.get(w_lbl, 0.0) + rev_ok
                        monthly_ship_actual += q_val
                        monthly_rev_actual += rev_ok

                        if base_model:
                            m_rec = model_period["monthly"].setdefault(base_model, {"Actual": 0, "Rev_Act": 0.0})
                            m_rec["Actual"] += q_val
                            m_rec["Rev_Act"] += rev_ok
                        if curr_site:
                            s_rec = site_period["monthly"].setdefault(curr_site, {"Actual": 0, "Rev_Act": 0.0})
                            s_rec["Actual"] += q_val
                            s_rec["Rev_Act"] += rev_ok

                    if w_key == target_iso_key:
                        if base_model:
                            mw_rec = model_period["weekly"].setdefault(base_model, {"Actual": 0, "Rev_Act": 0.0})
                            mw_rec["Actual"] += q_val
                            mw_rec["Rev_Act"] += rev_ok
                        if curr_site:
                            sw_rec = site_period["weekly"].setdefault(curr_site, {"Actual": 0, "Rev_Act": 0.0})
                            sw_rec["Actual"] += q_val
                            sw_rec["Rev_Act"] += rev_ok

                if d_obj == target_date_obj:
                    daily_ship_a += q_val
                    daily_rev_a += rev_ok
                    if base_model:
                        md_rec = model_period["daily"].setdefault(base_model, {"Actual": 0, "Rev_Act": 0.0})
                        md_rec["Actual"] += q_val
                        md_rec["Rev_Act"] += rev_ok
                    if curr_site:
                        sd_rec = site_period["daily"].setdefault(curr_site, {"Actual": 0, "Rev_Act": 0.0})
                        sd_rec["Actual"] += q_val
                        sd_rec["Rev_Act"] += rev_ok

    if total_carryover_krw == 0:
        for r in range(max(1, sheet.max_row - 15), sheet.max_row + 1):
            val_f = str(safe_cell(sheet, r, carryover_col) or 0).replace(",", "").strip()
            try:
                v_num = float(val_f)
                if v_num > 100_000_000:
                    total_carryover_krw = max(total_carryover_krw, v_num)
            except Exception:
                pass

    for k, v in learned_prices.items():
        if k not in mapping_data["gh_map"] or mapping_data["gh_map"][k] == 0:
            mapping_data["gh_map"][k] = v
    for k, v in learned_suffix_prices.items():
        if k not in mapping_data["suffix_map"] or mapping_data["suffix_map"][k] == 0:
            mapping_data["suffix_map"][k] = v
    for k, v in learned_model_prices.items():
        if k not in mapping_data["model_map"] or mapping_data["model_map"][k] == 0:
            mapping_data["model_map"][k] = v

    return {
        "sheet_ym": f"{sheet_year:04d}-{sheet_month:02d}",
        "is_target_month": is_target_month,
        "daily_ship_a": daily_ship_a,
        "daily_rev_a": round(daily_rev_a, 2),
        "carryover_rev_ok": round(total_carryover_krw / 100_000_000.0, 2),
        "date_actual_map": date_actual_map,
        "weekly_ship_actual": weekly_ship_actual,
        "weekly_rev_actual": {k: round(v, 2) for k, v in weekly_rev_actual.items()},
        "weekly_ship_by_isokey": weekly_ship_by_isokey,
        "weekly_rev_by_isokey": weekly_rev_by_isokey,
        "monthly_ship_actual": monthly_ship_actual,
        "monthly_rev_actual": round(monthly_rev_actual, 2),
        "model_period": model_period,
        "site_period": site_period,
    }


def parse_new_shipment_actual_excel(file_src, target_date_str, mapping_data):
    wb = openpyxl.load_workbook(file_src, data_only=True)
    mapping_data.setdefault("gh_map", {})
    mapping_data.setdefault("suffix_map", {})
    mapping_data.setdefault("model_map", {})
    mapping_data.setdefault("gh_to_model", {})

    t_y, t_m, t_d = parse_target_ymd(target_date_str)
    target_date_obj = datetime.date(t_y, t_m, t_d)
    target_iso_key = get_iso_week_key(target_date_obj)

    # 기준월 시트 우선 선택 및 통합 내 다른 월 시트도 함께 감지
    candidate_sheets = []
    for s_name in wb.sheetnames:
        m_mon = re.search(r"(\d{1,2})\s*월", s_name)
        m_yr = re.search(r"(20\d{2}|\d{2})\s*년", s_name)
        if m_mon:
            s_m = int(m_mon.group(1))
            if 1 <= s_m <= 12:
                if m_yr:
                    yr_raw = int(m_yr.group(1))
                    s_y = (2000 + yr_raw) if yr_raw < 100 else yr_raw
                else:
                    s_y = infer_year_for_month(s_m, target_date_obj)
                candidate_sheets.append((s_name, s_y, s_m))

    if not candidate_sheets:
        primary_name = next((s for s in wb.sheetnames if any(k in s for k in ["출하", "일별", "실적"])), wb.sheetnames[0])
        candidate_sheets.append((primary_name, t_y, t_m))

    combined_date_actual = {}
    loaded_yms = []
    total_weekly_ship_iso, total_weekly_rev_iso = {}, {}
    merged_weekly_model, merged_weekly_site = {}, {}

    target_res = None
    for s_name, s_y, s_m in candidate_sheets:
        res = _parse_single_shipment_sheet(wb[s_name], s_y, s_m, target_date_obj, mapping_data)
        combined_date_actual.update(res["date_actual_map"])
        loaded_yms.append(res["sheet_ym"])

        for wk, wv in res["weekly_ship_by_isokey"].items():
            total_weekly_ship_iso[wk] = total_weekly_ship_iso.get(wk, 0) + wv
        for wk, wv in res["weekly_rev_by_isokey"].items():
            total_weekly_rev_iso[wk] = total_weekly_rev_iso.get(wk, 0.0) + wv

        for mdl, vals in res["model_period"]["weekly"].items():
            rec = merged_weekly_model.setdefault(mdl, {"Actual": 0, "Rev_Act": 0.0})
            rec["Actual"] += vals["Actual"]
            rec["Rev_Act"] += vals["Rev_Act"]

        for ste, vals in res["site_period"]["weekly"].items():
            rec = merged_weekly_site.setdefault(ste, {"Actual": 0, "Rev_Act": 0.0})
            rec["Actual"] += vals["Actual"]
            rec["Rev_Act"] += vals["Rev_Act"]

        if res["is_target_month"] and target_res is None:
            target_res = res

    if target_res is None:
        empty_period = {"daily": {}, "weekly": merged_weekly_model, "monthly": {}}
        empty_site_p = {"daily": {}, "weekly": merged_weekly_site, "monthly": {}}
        return {
            "daily_ship_a": 0,
            "daily_rev_a": 0.0,
            "cur_week_ship_a": total_weekly_ship_iso.get(target_iso_key, 0),
            "cur_week_pure_rev_a": round(total_weekly_rev_iso.get(target_iso_key, 0.0), 2),
            "carryover_rev_ok": 0.0,
            "date_actual_map": combined_date_actual,
            "loaded_yms": loaded_yms,
            "weekly_ship_actual": {},
            "weekly_rev_actual": {},
            "monthly_ship_actual": 0,
            "monthly_rev_actual": 0.0,
            "actual_by_model_period": empty_period,
            "actual_by_site_period": empty_site_p,
        }

    target_res["model_period"]["weekly"] = merged_weekly_model
    target_res["site_period"]["weekly"] = merged_weekly_site

    return {
        "daily_ship_a": target_res["daily_ship_a"],
        "daily_rev_a": target_res["daily_rev_a"],
        "cur_week_ship_a": total_weekly_ship_iso.get(target_iso_key, 0),
        "cur_week_pure_rev_a": round(total_weekly_rev_iso.get(target_iso_key, 0.0), 2),
        "carryover_rev_ok": target_res["carryover_rev_ok"],
        "date_actual_map": combined_date_actual,
        "loaded_yms": loaded_yms,
        "weekly_ship_actual": target_res["weekly_ship_actual"],
        "weekly_rev_actual": target_res["weekly_rev_actual"],
        "monthly_ship_actual": target_res["monthly_ship_actual"],
        "monthly_rev_actual": target_res["monthly_rev_actual"],
        "actual_by_model_period": target_res["model_period"],
        "actual_by_site_period": target_res["site_period"],
    }


def parse_rtf_main_plan_excel(file_src, target_date_str, mapping_data, exchange_rate=1350.0):
    wb = openpyxl.load_workbook(file_src, data_only=True)
    sheet = wb["Main"] if "Main" in wb.sheetnames else wb.active

    t_y, t_m, t_d = parse_target_ymd(target_date_str)
    target_date_obj = datetime.date(t_y, t_m, t_d)
    target_iso_key = get_iso_week_key(target_date_obj)
    target_month_key = f"{t_y:04d}-{t_m:02d}"

    raw_date_headers = []
    for c in range(1, sheet.max_column + 1):
        v = safe_cell(sheet, 4, c)
        if isinstance(v, (datetime.datetime, datetime.date)):
            raw_date_headers.append((c, "dt", v))
        elif v:
            s_val = str(v).strip().upper()
            if not re.search(r"\bW\d+\b", s_val):
                m_full = re.search(r"\b(\d{2,4})\s*[-/.]\s*0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s_val)
                if m_full:
                    yr_raw = int(m_full.group(1))
                    yr = (2000 + yr_raw) if yr_raw < 100 else yr_raw
                    raw_date_headers.append((c, "ymd", (yr, int(m_full.group(2)), int(m_full.group(3)))))
                else:
                    m_dt = re.search(r"\b0?(\d{1,2})\s*[-/.]\s*0?(\d{1,2})\b", s_val)
                    if m_dt:
                        raw_date_headers.append((c, "md", (int(m_dt.group(1)), int(m_dt.group(2)))))

    # 연말연시(12월 -> 1월) 문자열 날짜 헤더 연도 자동 이월 추적
    date_cols = []
    running_year = None
    prev_month = None

    for c, kind, val in raw_date_headers:
        if kind == "dt":
            y = val.year if val.year >= 2000 else infer_year_for_month(val.month, target_date_obj)
            d_obj = clamp_date(y, val.month, val.day)
            running_year = d_obj.year
            prev_month = d_obj.month
        elif kind == "ymd":
            d_obj = clamp_date(val[0], val[1], val[2])
            running_year = d_obj.year
            prev_month = d_obj.month
        else:
            mon, day = val
            if running_year is None:
                running_year = infer_year_for_month(mon, target_date_obj)
            elif prev_month is not None and mon < prev_month and (prev_month - mon) >= 6:
                running_year += 1
            d_obj = clamp_date(running_year, mon, day)
            prev_month = mon

        date_cols.append((c, d_obj))

    distinct_years = sorted(list(set(d.year for _, d in date_cols)))
    span_multi_year = len(distinct_years) > 1

    enriched_date_cols = []
    fcst_date_cols = []
    seen_col_labels = {}

    for c, d_obj in date_cols:
        base_lbl = d_obj.strftime("%y/%m/%d") if span_multi_year else d_obj.strftime("%m/%d")
        if base_lbl in seen_col_labels:
            seen_col_labels[base_lbl] += 1
            d_lbl = f"{base_lbl}_{seen_col_labels[base_lbl]}"
        else:
            seen_col_labels[base_lbl] = 1
            d_lbl = base_lbl

        fcst_date_cols.append(d_lbl)
        enriched_date_cols.append(
            (
                c,
                d_obj,
                get_iso_week_key(d_obj),
                get_iso_week_label(d_obj),
                f"{d_obj.year:04d}-{d_obj.month:02d}",
                f"{d_obj.month}월",
                d_lbl,
                d_obj.strftime("%Y-%m-%d"),
            )
        )

    site_col, gh_col, gubun_col = 3, 4, 9

    gh_map = mapping_data.get("gh_map", {})
    suffix_map = mapping_data.get("suffix_map", {})
    gh_to_model = mapping_data.get("gh_to_model", {})
    model_map = mapping_data.get("model_map", {})

    known_prices = [p for p in list(gh_map.values()) + list(suffix_map.values()) + list(model_map.values()) if p > 0]
    avg_price = (sum(known_prices) / len(known_prices)) if known_prices else 8800.0

    daily_ship_p, daily_rev_p = 0, 0.0
    weekly_ship_by_isokey, weekly_rev_by_isokey = {}, {}
    target_month_weekly_ship, target_month_weekly_rev = {}, {}
    monthly_ship_by_ym, monthly_rev_by_ym = {}, {}

    plan_by_model_period = {"daily": {}, "weekly": {}, "monthly": {}}
    plan_by_site_period = {"daily": {}, "weekly": {}, "monthly": {}}
    fcst_rows = []

    for r in range(5, sheet.max_row + 1):
        v_gubun = str(safe_cell(sheet, r, gubun_col) or "").replace(" ", "").strip()
        site = str(safe_cell(sheet, r, site_col) or "").strip()
        gh_code = str(safe_cell(sheet, r, gh_col) or "").strip()

        if not v_gubun or v_gubun not in ["공급", "공 급"]:
            continue
        if not site and not gh_code:
            continue
        if any(k in f"{site}{gh_code}".upper() for k in ["합계", "소계", "총계", "TOTAL"]):
            continue

        gh_clean = clean_key(gh_code)
        gh_suf = get_gh_suffix(gh_code)
        proj_val = str(safe_cell(sheet, r, 6) or "").strip()
        model_name = resolve_model_name(gh_code, proj_val, gh_to_model)

        price = gh_map.get(gh_clean, 0.0)
        if price == 0:
            price = suffix_map.get(gh_suf, 0.0)
        if price == 0 and model_name:
            price = model_map.get(clean_key(model_name), 0.0)
        if price == 0:
            for k, v in gh_map.items():
                if len(k) >= 5 and (k in gh_clean or gh_clean in k or gh_suf in k):
                    price = v
                    break
        if price == 0:
            price = avg_price

        price_krw = price if price > 100 else price * exchange_rate

        row_fcst = {
            "SITE": site,
            "GH_CODE": gh_code,
            "모델명": model_name,
            "단가_원": price_krw,
            "_by_ymd": {},
        }
        row_fcst_total = 0

        for c_idx, d_obj, w_key, w_lbl, ym_key, _, d_lbl, ymd_str in enriched_date_cols:
            q_raw = str(safe_cell(sheet, r, c_idx) or 0).replace(",", "").replace("-", "0").strip()
            try:
                q_val = int(float(q_raw)) if q_raw else 0
            except Exception:
                q_val = 0

            row_fcst[d_lbl] = q_val
            if q_val > 0:
                row_fcst["_by_ymd"][ymd_str] = row_fcst["_by_ymd"].get(ymd_str, 0) + q_val
            row_fcst_total += q_val

            rev_val = (q_val * price_krw) / 100_000_000.0

            weekly_ship_by_isokey[w_key] = weekly_ship_by_isokey.get(w_key, 0) + q_val
            weekly_rev_by_isokey[w_key] = weekly_rev_by_isokey.get(w_key, 0.0) + rev_val

            if ym_key == target_month_key:
                target_month_weekly_ship[w_lbl] = target_month_weekly_ship.get(w_lbl, 0) + q_val
                target_month_weekly_rev[w_lbl] = target_month_weekly_rev.get(w_lbl, 0.0) + rev_val

                if model_name:
                    mm = plan_by_model_period["monthly"].setdefault(model_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    mm["Plan_Qty"] += q_val
                    mm["Plan_Rev"] += rev_val
                if site:
                    sm = plan_by_site_period["monthly"].setdefault(site, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    sm["Plan_Qty"] += q_val
                    sm["Plan_Rev"] += rev_val

            monthly_ship_by_ym[ym_key] = monthly_ship_by_ym.get(ym_key, 0) + q_val
            monthly_rev_by_ym[ym_key] = monthly_rev_by_ym.get(ym_key, 0.0) + rev_val

            if w_key == target_iso_key:
                if model_name:
                    mw = plan_by_model_period["weekly"].setdefault(model_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    mw["Plan_Qty"] += q_val
                    mw["Plan_Rev"] += rev_val
                if site:
                    sw = plan_by_site_period["weekly"].setdefault(site, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    sw["Plan_Qty"] += q_val
                    sw["Plan_Rev"] += rev_val

            if d_obj == target_date_obj:
                daily_ship_p += q_val
                daily_rev_p += rev_val
                if model_name:
                    md = plan_by_model_period["daily"].setdefault(model_name, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    md["Plan_Qty"] += q_val
                    md["Plan_Rev"] += rev_val
                if site:
                    sd = plan_by_site_period["daily"].setdefault(site, {"Plan_Qty": 0, "Plan_Rev": 0.0})
                    sd["Plan_Qty"] += q_val
                    sd["Plan_Rev"] += rev_val

        row_fcst["합계"] = row_fcst_total
        if row_fcst_total > 0:
            fcst_rows.append(row_fcst)

    # 당월 차트용 주차 계획: 해당 주차 전체 계획(월 경계 포함 시 ISO 주차 기준 우선)
    month_weeks = get_month_iso_weeks(t_y, t_m)
    weekly_ship_plan_for_chart = {}
    weekly_rev_plan_for_chart = {}
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
        "daily_ship_p": daily_ship_p,
        "daily_rev_p": round(daily_rev_p, 2),
        "cur_week_ship_p": weekly_ship_by_isokey.get(target_iso_key, 0),
        "cur_week_rev_p": round(weekly_rev_by_isokey.get(target_iso_key, 0.0), 2),
        "cur_month_ship_p": monthly_ship_by_ym.get(target_month_key, 0),
        "cur_month_rev_p": round(monthly_rev_by_ym.get(target_month_key, 0.0), 2),
        "weekly_ship_plan": weekly_ship_plan_for_chart,
        "weekly_rev_plan": weekly_rev_plan_for_chart,
        "month_weeks": month_weeks,
        "plan_by_model_period": plan_by_model_period,
        "plan_by_site_period": plan_by_site_period,
        "fcst_rows": fcst_rows,
        "fcst_date_cols": fcst_date_cols,
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
                "수량달성률(%)": qty_rate,
                "수량차질(K)": format_k(p_qty - a_qty),
                "매출계획(억원)": p_rev,
                "매출실적(억원)": a_rev,
                "매출달성률(%)": rev_rate,
            }
        )
    return rows


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

    # 1. 단가 매핑 반영
    if os.path.exists(CACHED_EXCEL_PATHS["map_rev"]):
        nmap = parse_mapping_excel(CACHED_EXCEL_PATHS["map_rev"])
        if nmap:
            st.session_state["mapping_data"].update(nmap)
            messages.append(("success", f"✅ 단가표 반영 완료 (GH-CODE {len(nmap['gh_map'])}개)"))

    # 2. 생산 실적 반영 (주간 KPI는 조회 기준 주차만 반영)
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
                    f"✅ 생산 데이터 반영 완료 (일간 실적: {format_k(p['daily_prod_actual'])} | {p['cur_week_label']} 주간 실적: {format_k(p['cur_week_prod_a'])})",
                )
            )
        except Exception as e:
            messages.append(("error", f"생산 파싱 오류: {e}"))

    w_labels = st.session_state["chart_data"].get("weekly_weeks", default_weeks)
    active_week_lbl = st.session_state.get("cur_week_label", target_iso_week)

    actual_model_period = {"daily": {}, "weekly": {}, "monthly": {}}
    actual_site_period = {"daily": {}, "weekly": {}, "monthly": {}}
    plan_model_period = {"daily": {}, "weekly": {}, "monthly": {}}
    plan_site_period = {"daily": {}, "weekly": {}, "monthly": {}}

    # 3. 출하 실적 반영 (주간 KPI는 조회 기준 주차만 반영, 캘린더용 YYYY-MM-DD 누적 보관)
    if os.path.exists(CACHED_EXCEL_PATHS["ship_actual"]):
        try:
            s_act = parse_new_shipment_actual_excel(
                CACHED_EXCEL_PATHS["ship_actual"], target_date_input, st.session_state["mapping_data"]
            )
            carryover = s_act["carryover_rev_ok"]

            st.session_state["actual_by_date"].update(s_act["date_actual_map"])
            existing_yms = set(st.session_state.get("actual_months_loaded", []))
            existing_yms.update(s_act["loaded_yms"])
            st.session_state["actual_months_loaded"] = sorted(list(existing_yms))

            st.session_state["kpi_data"]["daily"]["ship_a"] = s_act["daily_ship_a"]
            st.session_state["kpi_data"]["daily"]["rev_a"] = s_act["daily_rev_a"]

            w_ship_a = [s_act["weekly_ship_actual"].get(w, 0) for w in w_labels]
            w_rev_a = []
            for idx, w in enumerate(w_labels):
                pure_w = s_act["weekly_rev_actual"].get(w, 0.0)
                w_rev_a.append(round(pure_w + carryover, 2) if idx == 0 else round(pure_w, 2))

            st.session_state["chart_data"]["weekly_ship_a"] = w_ship_a
            st.session_state["chart_data"]["weekly_rev_a"] = w_rev_a

            # 조회일이 월 첫 주차(w_labels[0])일 때만 해당 주차 KPI에 전월 이월 가산
            cur_w_pure_rev = s_act["cur_week_pure_rev_a"]
            cur_w_carryover = carryover if (w_labels and active_week_lbl == w_labels[0]) else 0.0

            st.session_state["kpi_data"]["weekly"]["ship_a"] = s_act["cur_week_ship_a"]
            st.session_state["kpi_data"]["weekly"]["pure_rev_a"] = cur_w_pure_rev
            st.session_state["kpi_data"]["weekly"]["carryover_rev"] = cur_w_carryover
            st.session_state["kpi_data"]["weekly"]["month_carryover_total"] = carryover
            st.session_state["kpi_data"]["weekly"]["rev_a"] = round(cur_w_pure_rev + cur_w_carryover, 2)

            pure_monthly_rev = s_act["monthly_rev_actual"]
            st.session_state["kpi_data"]["monthly"]["ship_a"] = s_act["monthly_ship_actual"]
            st.session_state["kpi_data"]["monthly"]["pure_rev_a"] = pure_monthly_rev
            st.session_state["kpi_data"]["monthly"]["carryover_rev"] = carryover
            st.session_state["kpi_data"]["monthly"]["rev_a"] = round(pure_monthly_rev + carryover, 2)

            st.session_state["chart_data"]["monthly_ship_a"] = [s_act["monthly_ship_actual"]]
            st.session_state["chart_data"]["monthly_rev_a"] = [round(pure_monthly_rev + carryover, 2)]

            actual_model_period = s_act["actual_by_model_period"]
            actual_site_period = s_act["actual_by_site_period"]
            st.session_state["data_loaded"] = True
            messages.append(
                (
                    "success",
                    f"✅ 출하 실적 반영 완료 (일간: {format_k(s_act['daily_ship_a'])} | {active_week_lbl} 주간: {format_k(s_act['cur_week_ship_a'])} | 이월매출: {carryover:,.2f}억원)",
                )
            )
        except Exception as e:
            messages.append(("error", f"출하 실적 파싱 오류: {e}"))

    # 4. 출하 및 매출 계획 반영 (RTF Main)
    if os.path.exists(CACHED_EXCEL_PATHS["ship_plan"]):
        try:
            s_plan = parse_rtf_main_plan_excel(
                CACHED_EXCEL_PATHS["ship_plan"],
                target_date_input,
                st.session_state["mapping_data"],
                float(ex_rate_input),
            )
            st.session_state["kpi_data"]["daily"]["ship_p"] = s_plan["daily_ship_p"]
            st.session_state["kpi_data"]["daily"]["rev_p"] = s_plan["daily_rev_p"]
            st.session_state["kpi_data"]["weekly"]["ship_p"] = s_plan["cur_week_ship_p"]
            st.session_state["kpi_data"]["weekly"]["rev_p"] = s_plan["cur_week_rev_p"]
            st.session_state["kpi_data"]["monthly"]["ship_p"] = s_plan["cur_month_ship_p"]
            st.session_state["kpi_data"]["monthly"]["rev_p"] = s_plan["cur_month_rev_p"]

            st.session_state["chart_data"]["weekly_ship_p"] = [
                s_plan["weekly_ship_plan"].get(w, 0) for w in w_labels
            ]
            st.session_state["chart_data"]["weekly_rev_p"] = [
                s_plan["weekly_rev_plan"].get(w, 0.0) for w in w_labels
            ]
            st.session_state["chart_data"]["monthly_ship_p"] = [s_plan["cur_month_ship_p"]]
            st.session_state["chart_data"]["monthly_rev_p"] = [s_plan["cur_month_rev_p"]]

            st.session_state["fcst_data"] = {
                "rows": s_plan.get("fcst_rows", []),
                "date_cols": s_plan.get("fcst_date_cols", []),
            }
            plan_model_period = s_plan["plan_by_model_period"]
            plan_site_period = s_plan["plan_by_site_period"]
            st.session_state["data_loaded"] = True
            messages.append(
                (
                    "success",
                    f"✅ 출하 & 매출 계획 반영 완료 (일간 계획: {format_k(s_plan['daily_ship_p'])} | {active_week_lbl} 주간 계획: {format_k(s_plan['cur_week_ship_p'])})",
                )
            )
        except Exception as e:
            messages.append(("error", f"출하 계획 파싱 오류: {e}"))

    # 5. 모델별 / SITE별 기간별(일간·주간·월간) 통합 테이블 생성
    model_period_tables = {}
    site_period_tables = {}
    for p_key in ["daily", "weekly", "monthly"]:
        model_period_tables[p_key] = _build_merged_table(
            plan_model_period.get(p_key, {}), actual_model_period.get(p_key, {}), "모델명"
        )
        site_period_tables[p_key] = _build_merged_table(
            plan_site_period.get(p_key, {}), actual_site_period.get(p_key, {}), "SITE"
        )

    st.session_state["chart_data"]["ship_by_model_period"] = model_period_tables
    st.session_state["chart_data"]["ship_by_site_period"] = site_period_tables
    st.session_state["chart_data"]["ship_by_model"] = model_period_tables["daily"]
    st.session_state["chart_data"]["ship_by_site"] = site_period_tables["daily"]

    st.session_state["last_updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_state_to_disk()
    return messages


# 6. CSS 스타일링
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

# 7. 사이드바
with st.sidebar:
    menu = st.radio(
        "📂 Folder",
        (
            "🏠 현황",
            "1. 모델별 출하실적",
            "2. 모델별 생산실적",
            "3. 출하 SITE별 실적",
            "4. 고객 FCST 조회",
            "5. 완제품 재고 조회",
            "6. 캘린더",
            "7. 데이터 관리 (Admin)",
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


# 8. 메인 메뉴 분기 로직
if menu == "7. 데이터 관리 (Admin)":
    st.title("⚙️ DATA MANAGEMENT")
    if st.text_input("비밀번호", type="password") == "2250":
        saved_files_status = []
        for label_kr, slot_k in [
            ("생산실적", "prod"),
            ("출하실적", "ship_actual"),
            ("출하계획(RTF)", "ship_plan"),
            ("단가매핑", "map_rev"),
        ]:
            if os.path.exists(CACHED_EXCEL_PATHS[slot_k]):
                saved_files_status.append(f"🟢 {label_kr}")
            else:
                saved_files_status.append(f"⚪ {label_kr}(미등록)")

        last_upd = st.session_state.get("last_updated_at") or "반영 이력 없음"
        st.info(
            f"💾 **서버 자동 보관 상태:** {' | '.join(saved_files_status)}  \n"
            f"🕒 **마지막 반영 시각:** {last_upd} (새로고침 및 타 PC 접속 시 자동 유지, 변경된 파일이나 기준일만 수정 후 재반영 가능)"
        )

        c_set1, c_set2 = st.columns(2)
        with c_set1:
            target_date_input = st.text_input(
                "📅 조회 기준 날짜 지정 (예: 11/25 또는 2026-11-25)",
                value=st.session_state.get("target_date", default_target_date),
            )
        with c_set2:
            ex_rate_input = st.number_input(
                "💵 적용 환율 (원/USD)",
                value=float(st.session_state.get("ex_rate", 1350.0)),
                step=10.0,
            )

        col1, col2 = st.columns(2)
        with col1:
            prod_file = st.file_uploader("1. 생산 실적 파일", type=["xlsx"])
        with col2:
            ship_actual_file = st.file_uploader("2. 출하 실적 파일 (신규 양식)", type=["xlsx"])

        col3, col4 = st.columns(2)
        with col3:
            ship_plan_file = st.file_uploader("3. 출하 계획 파일 (RTF Main 시트)", type=["xlsx"])
        with col4:
            map_rev_file = st.file_uploader("4. 단가 매핑 파일 (선택)", type=["xlsx"])

        btn_c1, btn_c2 = st.columns([0.75, 0.25])
        with btn_c1:
            if st.button("🚀 데이터 분석 및 서버 저장 반영", type="primary", use_container_width=True):
                with st.spinner("데이터를 통합 분석하고 서버 저장소에 반영하고 있습니다..."):
                    msgs = run_dashboard_pipeline(
                        target_date_input,
                        ex_rate_input,
                        {
                            "prod": prod_file,
                            "ship_actual": ship_actual_file,
                            "ship_plan": ship_plan_file,
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

    # 1. Monthly 탭 (전체 통합)
    with tab_monthly:
        m1, m2, m3 = st.columns(3)
        with m1:
            render_kpi_card(f"{t_m}월 생산 누계 (전체)", "🏭", kpi["monthly"]["prod_a"], kpi["monthly"]["prod_p"])
            render_achievement_gauge("생산", kpi["monthly"]["prod_a"], kpi["monthly"]["prod_p"], unit="K")
        with m2:
            render_kpi_card(f"{t_m}월 출하 누계 (전체)", "🚚", kpi["monthly"]["ship_a"], kpi["monthly"]["ship_p"])
            render_achievement_gauge("출하", kpi["monthly"]["ship_a"], kpi["monthly"]["ship_p"], unit="K")
        with m3:
            rev_sub = f"당월 순실적 {kpi['monthly']['pure_rev_a']:,.2f}억 + 전월 이월 {kpi['monthly']['carryover_rev']:,.2f}억"
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

    # 2. Weekly 탭 (조회일 기준 해당 주차 합계만 집계, 전체 통합)
    with tab_weekly:
        wk1, wk2, wk3 = st.columns(3)
        with wk1:
            render_kpi_card(f"주간 생산 ({cur_week_lbl} 전체)", "🏭", kpi["weekly"]["prod_a"], kpi["weekly"]["prod_p"])
            render_achievement_gauge("생산", kpi["weekly"]["prod_a"], kpi["weekly"]["prod_p"], unit="K")
        with wk2:
            render_kpi_card(f"주간 출하 ({cur_week_lbl} 전체)", "🚚", kpi["weekly"]["ship_a"], kpi["weekly"]["ship_p"])
            render_achievement_gauge("출하", kpi["weekly"]["ship_a"], kpi["weekly"]["ship_p"], unit="K")
        with wk3:
            w_carry = kpi["weekly"].get("carryover_rev", 0.0)
            m_carry_tot = kpi["weekly"].get("month_carryover_total", 0.0)
            if w_carry > 0:
                w_rev_sub = f"{cur_week_lbl} 순실적 {kpi['weekly']['pure_rev_a']:,.2f}억 + 전월 이월 {w_carry:,.2f}억"
            elif m_carry_tot > 0:
                first_w = cdata.get("weekly_weeks", [cur_week_lbl])[0]
                w_rev_sub = f"{cur_week_lbl} 당주 순실적 기준 (전월 이월 {m_carry_tot:,.2f}억은 {first_w} 반영)"
            else:
                w_rev_sub = f"{cur_week_lbl} 당주 실적 기준"
            render_kpi_card(
                f"주간 매출 ({cur_week_lbl} 전체)", "💰", kpi["weekly"]["rev_a"], kpi["weekly"]["rev_p"], is_rev=True, sub_caption=w_rev_sub
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

    # 3. Daily 탭 (전체 통합)
    with tab_daily:
        k1, k2, k3 = st.columns(3)
        with k1:
            render_kpi_card("일간 생산 (전체)", "🏭", kpi["daily"]["prod_a"], kpi["daily"]["prod_p"])
            render_achievement_gauge("생산", kpi["daily"]["prod_a"], kpi["daily"]["prod_p"], unit="K")
        with k2:
            render_kpi_card("일간 출하 (전체)", "🚚", kpi["daily"]["ship_a"], kpi["daily"]["ship_p"])
            render_achievement_gauge("출하", kpi["daily"]["ship_a"], kpi["daily"]["ship_p"], unit="K")
        with k3:
            render_kpi_card("일간 매출 (전체)", "💰", kpi["daily"]["rev_a"], kpi["daily"]["rev_p"], is_rev=True)
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
    st.subheader("📦 1. 모델별 출하 및 매출 현황 (계획 vs 실적)")
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

    period_dict = st.session_state["chart_data"].get("ship_by_model_period", {})
    rows = period_dict.get(sel_period) if period_dict else st.session_state["chart_data"].get("ship_by_model", [])
    df_ship = pd.DataFrame(rows or [])

    if not df_ship.empty:
        df_ship = df_ship.sort_values("출하실적(개)", ascending=False)

        tab_qty, tab_rev = st.tabs(["🚚 출하 수량 현황 (K)", "💰 매출액 현황 (억원)"])
        with tab_qty:
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
        with tab_rev:
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

        col1, col2 = st.columns([0.8, 0.2])
        with col2:
            st.download_button(
                label="📥 엑셀(CSV) 다운로드",
                data=convert_df(df_ship),
                file_name=f"model_shipment_{sel_period}.csv",
                mime="text/csv",
            )
        st.dataframe(
            df_ship[
                [
                    "모델명",
                    "출하계획(K)",
                    "출하실적(K)",
                    "수량달성률(%)",
                    "수량차질(K)",
                    "매출계획(억원)",
                    "매출실적(억원)",
                    "매출달성률(%)",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("선택한 기간의 출하 실적/계획 데이터가 없습니다. [7. 데이터 관리 (Admin)]에서 파일을 업로드해 주세요.")

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

elif menu == "3. 출하 SITE별 실적":
    st.subheader("📍 3. 출하 SITE별 출하 및 매출 현황 (계획 vs 실적)")
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

    period_dict = st.session_state["chart_data"].get("ship_by_site_period", {})
    rows = period_dict.get(sel_period) if period_dict else st.session_state["chart_data"].get("ship_by_site", [])
    df_site = pd.DataFrame(rows or [])

    if not df_site.empty:
        df_site = df_site.sort_values("출하실적(개)", ascending=False)

        tab_s_qty, tab_s_rev = st.tabs(["🚚 SITE별 출하 수량 (K)", "💰 SITE별 매출액 (억원)"])
        with tab_s_qty:
            st.plotly_chart(
                create_active_chart(
                    f"SITE별 출하 수량 - {p_label_map[sel_period]}",
                    df_site["SITE"].tolist(),
                    df_site["출하계획(개)"].tolist(),
                    df_site["출하실적(개)"].tolist(),
                    unit="K",
                ),
                use_container_width=True,
            )
        with tab_s_rev:
            st.plotly_chart(
                create_active_chart(
                    f"SITE별 매출액 - {p_label_map[sel_period]}",
                    df_site["SITE"].tolist(),
                    df_site["매출계획(억원)"].tolist(),
                    df_site["매출실적(억원)"].tolist(),
                    unit="억원",
                ),
                use_container_width=True,
            )

        col1, col2 = st.columns([0.8, 0.2])
        with col2:
            st.download_button(
                label="📥 엑셀(CSV) 다운로드",
                data=convert_df(df_site),
                file_name=f"site_shipment_{sel_period}.csv",
                mime="text/csv",
            )
        st.dataframe(
            df_site[
                [
                    "SITE",
                    "출하계획(K)",
                    "출하실적(K)",
                    "수량달성률(%)",
                    "수량차질(K)",
                    "매출계획(억원)",
                    "매출실적(억원)",
                    "매출달성률(%)",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("선택한 기간의 SITE별 출하 실적/계획 데이터가 없습니다. [7. 데이터 관리 (Admin)]에서 파일을 업로드해 주세요.")

elif menu == "4. 고객 FCST 조회":
    st.title("📈 4. 고객 FCST 조회 (모델별 공급 일정)")
    fcst_data = st.session_state.get("fcst_data", {"rows": [], "date_cols": []})
    fcst_rows = fcst_data.get("rows", [])
    date_cols = fcst_data.get("date_cols", [])

    if not fcst_rows:
        st.info("💡 출하 계획 파일(RTF Main 시트)을 **[7. 데이터 관리 (Admin)]** 메뉴에서 업로드 및 분석해 주세요.")
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
            st.metric("🏭 공급 대상 SITE 수", f"{site_count} 개 거점")
        with m4:
            st.metric("👑 최대 공급 모델", f"{top_model}", f"{format_k(top_model_qty)}")

        st.markdown("<br>", unsafe_allow_html=True)
        f1, f2, f3 = st.columns([1.5, 1.2, 1.3])
        with f1:
            all_models = sorted(df_fcst["모델명"].unique().tolist())
            selected_models = st.multiselect("🔍 모델명 필터 (다중 선택 가능)", options=all_models, default=[])
        with f2:
            all_sites = sorted(df_fcst["SITE"].unique().tolist())
            selected_sites = st.multiselect("🏢 출하 SITE 필터", options=all_sites, default=[])
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
            ["📊 모델별 공급 일정 (통합)", "📋 품목 / SITE별 상세 공급 일정", "📈 공급 계획 추이 차트"]
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
            st.markdown("##### 📌 품목코드(GH-CODE) 및 출하 SITE별 상세 공급 일정 (단위: K)")
            if not df_filtered.empty and valid_dates:
                detail_cols = ["모델명", "GH_CODE", "SITE", "합계"] + valid_dates
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

elif menu == "5. 완제품 재고 조회":
    st.subheader("🏢 5. 완제품 재고 조회 (전체 통합)")
    st.metric(
        "📦 당일 완제품 총 재고",
        f"{format_k(st.session_state['inventory_data']['qty'])}",
        f"재고일수: {st.session_state['inventory_data']['days']} 일",
        delta_color="off",
    )

elif menu == "6. 캘린더":
    st.title("🗓️ 6. 월간 FCST 공급 및 출하 캘린더")
    fcst_data = st.session_state.get("fcst_data", {"rows": [], "date_cols": []})
    fcst_rows = fcst_data.get("rows", [])
    actual_by_date = st.session_state.get("actual_by_date", {})
    loaded_yms = set(st.session_state.get("actual_months_loaded", []))

    if not fcst_rows and not actual_by_date:
        st.info("💡 출하 계획(RTF Main) 또는 출하 실적 파일을 **[7. 데이터 관리 (Admin)]** 메뉴에서 업로드 및 분석해 주세요.")
    else:
        t_y, t_m, t_d = parse_target_ymd(st.session_state.get("target_date", default_target_date))
        target_date_obj = datetime.date(t_y, t_m, t_d)

        # 데이터에 존재하는 모든 연-월 + 기준월 전후 월 자동 구성
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
        for ymd_k in actual_by_date.keys():
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

        day_supply_map = {d: {"total": 0, "total_rev": 0.0, "items": []} for d in range(1, num_days + 1)}

        for r_item in fcst_rows:
            m_name = r_item["모델명"]
            gh = r_item["GH_CODE"]
            site = r_item["SITE"]
            price = r_item.get("단가_원", 0.0)
            by_ymd = r_item.get("_by_ymd", {})

            for d in range(1, num_days + 1):
                ymd_key = f"{cur_year:04d}-{selected_m:02d}-{d:02d}"
                if by_ymd:
                    q = by_ymd.get(ymd_key, 0)
                else:
                    q = r_item.get(f"{selected_m:02d}/{d:02d}", r_item.get(f"{selected_m}/{d}", 0))

                if q > 0:
                    rev_val = (q * price) / 100_000_000.0 if price > 0 else 0.0
                    day_supply_map[d]["total"] += q
                    day_supply_map[d]["total_rev"] += rev_val
                    day_supply_map[d]["items"].append(
                        {
                            "model": m_name,
                            "gh_code": gh,
                            "site": site,
                            "qty": q,
                            "price": price,
                            "rev_ok": rev_val,
                        }
                    )

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

        # 구버전 Streamlit 용 Fallback 상세 컨테이너
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
                        elif is_past_or_target and not has_month_actuals:
                            border_css = "2px solid #111111; box-shadow: 3px 3px 0px #111111;"
                            bg_css = "#ffffff" if plan_q > 0 else "#fafafa"
                            badge_target = "<span style='background: #e5e7eb; color: #4b5563; font-size: 9px; font-weight: 800; padding: 1px 4px; border-radius: 3px; margin-left: 4px;'>실적미등록</span>"
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

        st.markdown(f"### 📈 {cur_year}년 {selected_m}월 일자별 공급 수량 및 예상 매출 추이")
        trend_days = [f"{selected_m}/{d:02d}" for d in range(1, num_days + 1)]
        trend_qtys = [day_supply_map[d]["total"] for d in range(1, num_days + 1)]
        trend_revs = [round(day_supply_map[d]["total_rev"], 2) for d in range(1, num_days + 1)]

        fig_cal_trend = go.Figure()
        fig_cal_trend.add_trace(
            go.Bar(
                x=trend_days,
                y=trend_qtys,
                name="공급계획수량(K)",
                marker_color=[
                    "#0044ff" if datetime.date(cur_year, selected_m, d) == target_date_obj else "#d1d5db"
                    for d in range(1, num_days + 1)
                ],
                marker_line=dict(width=1.5, color="#111111"),
                text=[format_k(q) if q > 0 else "" for q in trend_qtys],
                textposition="outside",
            )
        )
        fig_cal_trend.add_trace(
            go.Scatter(
                x=trend_days,
                y=trend_revs,
                name="예상매출(억원)",
                yaxis="y2",
                mode="lines+markers",
                line=dict(color="#ef4444", width=2.5),
                marker=dict(size=6, color="#ef4444"),
            )
        )
        fig_cal_trend.update_layout(
            title=f"<b>{cur_year}년 {selected_m}월 일자별 공급 계획 수량 (막대) 및 예상 매출액 (꺾은선)</b>",
            plot_bgcolor="rgba(0,0,0,0)",
            hovermode="x unified",
            margin=dict(l=20, r=20, t=45, b=20),
            xaxis=dict(tickangle=-45),
            yaxis=dict(title="공급수량 (개)"),
            yaxis2=dict(title="예상매출 (억원)", overlaying="y", side="right"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_cal_trend, use_container_width=True)
