"""
داشبورد گزارش‌های ماهانه و دوره‌ای کدال
اجرا: python -m streamlit run dashboard_seasonal_monthly.py
"""
import os
import glob
import sqlite3
import re
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from difflib import SequenceMatcher

# مسیر پایه: پوشه‌ای که همین فایل در آن است (روی ویندوز و روی سرور هر دو کار می‌کند)
try:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(_BASE_DIR, "data")

DB_PATH        = os.path.join(DATA_DIR, "codal_revenue.db")
USD_RATE_PATH  = os.path.join(DATA_DIR, "monthly_usd_rate.xlsx")
INCOME_DB_PATH = os.path.join(DATA_DIR, "codal_income.db")
ANNUAL_DB_PATH       = os.path.join(DATA_DIR, "codal_annual.db")
ANNUAL_USD_RATE_PATH = os.path.join(DATA_DIR, "usd_rate_annual.xlsx")
DAILY_DIR            = os.path.join(DATA_DIR, "daily2")   # فایل‌های روزانهٔ معاملات (TSETMC)
STOCKS_PATH          = os.path.join(DATA_DIR, "stck.xlsx")   # لیست نمادها + insCode (id)
SHAREHOLDERS_DIR     = os.path.join(DATA_DIR, "shareholders")   # خروجی shareholders_all.py
ASSEMBLY_DB_PATH     = os.path.join(DATA_DIR, "codal_assembly.db")
FORECAST_DB_PATH     = os.path.join(DATA_DIR, "forecasts.db")
PRODUCT_PL_DB_PATH   = os.path.join(DATA_DIR, "codal_product_pl.db")

# سال مالیِ مجمعِ امسال (مجمع ۱۴۰۵ مربوط به سال مالی منتهی به ۱۴۰۴ است).
# برای شاخص d استفاده می‌شود؛ هر سال این عدد را یک واحد افزایش دهید.
DIVIDEND_FY = 1404

PERSIAN_MONTHS = {
    "01": "فروردین", "02": "اردیبهشت", "03": "خرداد",
    "04": "تیر",     "05": "مرداد",    "06": "شهریور",
    "07": "مهر",     "08": "آبان",     "09": "آذر",
    "10": "دی",      "11": "بهمن",     "12": "اسفند",
}

EN_TO_FA = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
PERSIAN_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
ARABIC_DIGITS  = "٠١٢٣٤٥٦٧٨٩"


def normalize_digits(s):
    if not s:
        return ""
    return s.translate(str.maketrans(PERSIAN_DIGITS + ARABIC_DIGITS, "0123456789" * 2))


def parse_fa_number(v):
    """مقدارِ عددیِ ذخیره‌شده با قالبِ فارسی را به float تبدیل می‌کند.
    در فایل‌های روزانهٔ TSETMC، اعشار گاهی با ممیزِ فارسی «/» نوشته می‌شود
    (مثلِ «11/24» یعنی 11.24) و ارقام ممکن است فارسی/عربی باشند.
    اگر نتواند تبدیل کند، None برمی‌گرداند."""
    if v is None:
        return None
    s = normalize_digits(str(v)).strip()
    if not s or s.lower() in ("nan", "none", "-", "—"):
        return None
    s = s.replace(",", "")          # جداکنندهٔ هزارگان
    s = s.replace("/", ".").replace("٫", ".")   # ممیزِ فارسی → نقطه
    try:
        return float(s)
    except ValueError:
        return None


# انواعِ مختلفِ «الف» (آ أ إ ٱ …) و حروفِ عربی‌نما را یکدست می‌کند تا نامِ نماد در
# منابعِ مختلف یکی شمرده شود — مثلاً «غپآذر» در فایل‌های روزانهٔ D:\bourse\daily2 با
# «غپاذر» در دیتابیس. یعنی «ا» و «آ» (و دیگر شکل‌های الف) معادلِ هم در نظر گرفته می‌شوند.
_ALEF_VARIANTS = "آأإٱٲٳﺁﺂﺃﺄﺇﺈ"


def normalize_symbol(s):
    """نرمال‌سازیِ نامِ نماد فقط برای *تطبیق* (نه نمایش):
       ی/ک عربی → فارسی، همهٔ شکل‌های الف (آ/أ/إ/ٱ/…) → «ا»، و حذفِ نیم‌فاصله،
       کاراکترهای نامرئی و فاصله‌های اضافی."""
    if s is None:
        return ""
    s = str(s)
    s = s.replace('ي', 'ی').replace('ك', 'ک').replace('ى', 'ی')
    for _ch in _ALEF_VARIANTS:          # همهٔ انواع الف → ا  (یعنی  ا ≡ آ)
        s = s.replace(_ch, 'ا')
    for _ch in ('\u200c', '\u200f', '\u200e', '\u200b', '\u00a0'):
        s = s.replace(_ch, ' ')
    return re.sub(r'\s+', ' ', s).strip()


def to_fa_digits(text):
    if text is None:
        return ""
    return str(text).translate(EN_TO_FA)


def fmt_fa(num, decimals=1):
    if num is None or pd.isna(num):
        return "—"
    return f"{num:,.{decimals}f}".translate(EN_TO_FA)


def fmt_pct_fa(num, decimals=1):
    if num is None or pd.isna(num):
        return "—"
    return f"{num:+,.{decimals}f}%".translate(EN_TO_FA)


def format_period(period_end):
    try:
        y, m, _ = period_end.split("/")
        return f"{PERSIAN_MONTHS.get(m, m)} {to_fa_digits(y)}"
    except (ValueError, AttributeError):
        return to_fa_digits(period_end)


def _assign_year_month(df, parser, period_col="period_end", y_col="سال", m_col="ماه_num"):
    """انتساب امنِ ستون‌های سال/ماه از روی period_end.
    مقاوم در برابر دیتافریم خالی و تفاوت نسخه‌های پانداس
    (جایگزین الگوی شکنندهٔ df[[a,b]] = s.apply(lambda: pd.Series(...)))."""
    _pe = df[period_col].map(parser)
    df[y_col] = _pe.map(lambda t: t[0] if isinstance(t, tuple) else None)
    df[m_col] = _pe.map(lambda t: t[1] if isinstance(t, tuple) else None)
    return df


# ─── تنظیمات صفحه ─────────────────────────────────────────────
st.set_page_config(
    page_title="داشبورد کدال",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css" rel="stylesheet">
<style>
    html, body, [class*="css"], .stApp, .main, .block-container,
    .stMarkdown, .stText, .stSelectbox, .stMultiSelect, .stRadio,
    .stDataFrame, .stMetric, .stButton, .stTextInput, button, input, select, textarea,
    [data-testid="stSidebar"], [data-testid="stMetricLabel"], [data-testid="stMetricValue"],
    [data-testid="stMarkdownContainer"] {
        font-family: 'Vazirmatn', Tahoma, sans-serif !important;
        direction: rtl;
    }
    .stMarkdown, .stText, .stSelectbox, .stMultiSelect { text-align: right; }
    .stDataFrame { direction: rtl; }
    [data-testid="stMetricValue"] { direction: ltr; text-align: right; }
    .stRadio > div { direction: rtl; }
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Vazirmatn', Tahoma, sans-serif !important;
        font-weight: 600;
    }
    [data-testid="stDataFrame"] td,
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] .dvn-scroller *,
    [data-testid="stDataFrame"] [class*="cell"] *,
    [data-testid="stDataFrame"] [class*="header"] *,
    [data-testid="stDataFrame"] canvas + div *,
    .glideDataEditor *, .wzrd-header * {
        font-family: 'Vazirmatn', Tahoma, sans-serif !important;
    }
</style>
""", unsafe_allow_html=True)

PLOTLY_FONT       = dict(family="Vazirmatn, Tahoma, sans-serif", size=13)
PLOTLY_FONT_TITLE = dict(family="Vazirmatn, Tahoma, sans-serif", size=18)

def render_table(df, use_container_width=True, row_styles=None, center_values=False,
                 col_styles=None):
    """رندر جدول با فونت وزیری به جای st.dataframe.
    row_styles: dict اختیاری {شماره ردیف: css} برای رنگی/برجسته کردن سطرهای خاص.
    col_styles: dict اختیاری {نام ستون: css} برای رنگی کردنِ کلِ یک ستون (هدر+سلول‌ها).
    center_values: اگر True، ستون‌های داده وسط‌چین و چپ‌به‌راست (LTR) می‌شوند تا
                   اعداد وسط بیفتند و علامت منفی سمت چپِ عدد قرار بگیرد."""
    row_styles = row_styles or {}
    col_styles = col_styles or {}
    TABLE_CSS = (
        "width:100%;border-collapse:collapse;direction:rtl;"
        "font-family:Vazirmatn,Tahoma,sans-serif;font-size:13px;color:#e2e8f0;"
    )
    TH_CSS  = "background:#1e293b;padding:9px 12px;text-align:right;border:1px solid #334155;font-weight:600;"
    TD_CSS  = "padding:8px 12px;text-align:right;border:1px solid #2d3748;"
    TH_CTR  = "background:#1e293b;padding:9px 12px;text-align:center;border:1px solid #334155;font-weight:600;"
    TD_CTR  = "padding:8px 12px;text-align:center;direction:ltr;border:1px solid #2d3748;"
    ODD_BG  = "background:#0f172a;"
    EVEN_BG = "background:#1a2236;"

    cols = list(df.columns)
    html = f'<div style="overflow-x:auto"><table style="{TABLE_CSS}"><thead><tr>'
    for ci, col in enumerate(cols):
        th = (TH_CTR if (center_values and ci > 0) else TH_CSS)
        cs = col_styles.get(col, "")
        html += f'<th style="{th}{cs}">{col}</th>'
    html += '</tr></thead><tbody>'
    for i, (_, row) in enumerate(df.iterrows()):
        bg = ODD_BG if i % 2 == 0 else EVEN_BG
        extra = row_styles.get(i, "")   # css دلخواه برای این ردیف
        html += f'<tr style="{bg}{extra}">'
        for ci, col in enumerate(cols):
            td = (TD_CTR if (center_values and ci > 0) else TD_CSS)
            cs = col_styles.get(col, "")
            val = str(row[col]) if row[col] is not None else ""
            html += f'<td style="{td}{cs}">{val}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    if height is None:
        height = min(100 + len(df) * 30, 800)
    st.write(html, unsafe_allow_html=True)


def render_table_sortable(df, use_container_width=True, height=None):
    """رندر جدولِ مرتب‌پذیرِ HTML با دکمه‌های سرتیتر برای مرتب کردنِ اعدادِ فارسی و
    اعدادِ لاتین سمت راست. عملگرها: ↕ (نامرتب)، ▲/▼ (مرتب) در هدر.
    شماریِ ستون: ستونِ اول = ۱ (سمت راست تا چپ)."""
    height = height or min(100 + len(df) * 30, 800)
    cols = list(df.columns)
    html = f"""<html><head><meta charset="utf-8">
<style>
table {{ border-collapse:collapse; width:100%; font-family:Vazirmatn,Tahoma,sans-serif; }}
th {{ background:#1e293b; padding:9px 12px; text-align:right; border:1px solid #334155;
      font-weight:600; color:#e2e8f0; cursor:pointer; user-select:none; }}
td {{ padding:8px 12px; text-align:right; border:1px solid #2d3748; color:#e2e8f0; }}
tr:nth-child(even) {{ background:#1a2236; }}
tr:nth-child(odd) {{ background:#0f172a; }}
tr:hover {{ background:#1e293b; }}
.ar {{ opacity:.45; font-size:12px; display:inline-block; margin-right:8px; }}
</style></head><body><table id="t"><thead><tr>"""
    for col in cols:
        html += f'<th><span class="ar">⇅</span>{col}</th>'
    html += '</tr></thead><tbody>'
    for _, row in df.iterrows():
        html += '<tr>'
        for col in cols:
            val = str(row[col]) if row[col] is not None else ""
            html += f'<td>{val}</td>'
        html += '</tr>'
    html += f"""</tbody></table>
<script>
var ths=document.querySelectorAll('#t th');
ths.forEach(function(th, ci){{
  th.addEventListener('click', function(){{
    var tb=document.querySelector('#t tbody');
    var d=(th.dataset.sort||'desc')==='asc'?'desc':'asc';
    th.dataset.sort=d;
    var rows=Array.from(tb.querySelectorAll('tr'));
    rows.sort(function(a,b){{
      var x=a.querySelectorAll('td')[ci].textContent.trim();
      var y=b.querySelectorAll('td')[ci].textContent.trim();
      var nx=parseFloat(x), ny=parseFloat(y);
      var r=(!isNaN(nx)&&!isNaN(ny))?(nx-ny):String(x).localeCompare(String(y),'fa');
      return d==='asc'?r:-r;
    }});
    rows.forEach(function(r){{tb.appendChild(r);}});
    var ars=document.querySelectorAll('#t thead .ar');
    ars.forEach(function(s){{s.textContent=' ⇅'; s.style.opacity='.45';}});
    var sp=th.querySelector('.ar'); if(sp){{sp.textContent=d==='asc'?' ▲':' ▼'; sp.style.opacity='1';}}
  }});
}});
</script></body></html>"""
    components.html(html, height=min(height, 760), scrolling=(height > 760))


# ════════════════════════════════════════════════════════════════
# لود دیتا
# ════════════════════════════════════════════════════════════════
@st.cache_data(ttl=60)
def load_main():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT symbol, period_end, report_type,
               total_billion_toman, domestic_sales_btmn,
               export_sales_btmn, sales_return_btmn, discounts_btmn
        FROM monthly_revenue
        ORDER BY symbol, period_end
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=60)
def load_products():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT symbol, period_end, category, product_name, unit,
               production_qty, sales_qty, sales_rate_rial, sales_amount_btmn
        FROM monthly_product_breakdown
        ORDER BY symbol, period_end, category
    """, conn)
    conn.close()
    return df


@st.cache_data(ttl=60)
def load_product_pl():
    """مبلغ فروش/بهای تمام شده/سود ناخالصِ محصولات از codal_product_pl.db (خروجی اکستراکتر)."""
    if not os.path.exists(PRODUCT_PL_DB_PATH):
        return pd.DataFrame()
    try:
        conn = sqlite3.connect(PRODUCT_PL_DB_PATH)
        df = pd.read_sql_query(
            "SELECT symbol, period_end, duration_months, section, product, "
            "is_total, is_estimate, sales, cost, gross, qty_prod, qty_sold, price "
            "FROM product_pl", conn)
        conn.close()
    except Exception:
        return pd.DataFrame()
    return df


# ─── گزارش‌های دوره‌ای (صورت سود و زیان) ──────────────────────────
@st.cache_data(ttl=60)
def load_income_reports():
    """متادیتای گزارش‌های دوره‌ای از codal_income.db"""
    try:
        conn = sqlite3.connect(INCOME_DB_PATH)
        df = pd.read_sql_query("""
            SELECT id, symbol, period_end, duration_months,
                   COALESCE(is_audited, 0)      AS is_audited,
                   COALESCE(is_consolidated, 0) AS is_consolidated,
                   report_type, title, letter_url, sent_date
            FROM reports
            ORDER BY symbol, period_end DESC, duration_months DESC
        """, conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=60)
def load_income_items():
    """اقلام صورت سود و زیان (هر ردیف با عنوان واقعی) از codal_income.db"""
    try:
        conn = sqlite3.connect(INCOME_DB_PATH)
        df = pd.read_sql_query("""
            SELECT report_id, row_order, label, value
            FROM line_items
            ORDER BY report_id, row_order
        """, conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=60)
def load_assembly():
    """سود نقدی/خالص هر سهم (مجمع عادی سالیانه) از codal_assembly.db
       برای محاسبهٔ «درصد توزیع سود» = dps/eps."""
    try:
        conn = sqlite3.connect(ASSEMBLY_DB_PATH)
        df = pd.read_sql_query("""
            SELECT symbol, period_end, assembly_date, dps, eps,
                   title, letter_url, sent_date
            FROM assembly_decisions
            ORDER BY symbol, period_end DESC
        """, conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


# ─── پیش‌بینی سود (ذخیرهٔ دستیِ کاربر در forecasts.db) ─────────────
def _forecast_init_db():
    """ساخت جدول پیش‌بینی در صورت نبود. هر ردیف = یک قلمِ صورت سود و زیان
    برای یک نماد و یک سال مالی."""
    conn = sqlite3.connect(FORECAST_DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS forecasts (
            symbol       TEXT NOT NULL,
            fiscal_year  TEXT NOT NULL,
            row_order    INTEGER NOT NULL,
            label        TEXT NOT NULL,
            value        REAL,
            updated_at   TEXT,
            PRIMARY KEY (symbol, fiscal_year, row_order)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS forecast_notes (
            symbol       TEXT NOT NULL,
            fiscal_year  TEXT NOT NULL,
            note         TEXT,
            updated_at   TEXT,
            PRIMARY KEY (symbol, fiscal_year)
        )
    """)
    conn.commit()
    # مهاجرت: کلیدهای سال مالی که ارقام فارسی/عربی دارند را به انگلیسی یکدست کن
    # (مثلاً «فصلی-1404-Q۴» → «فصلی-1404-Q4») تا با کلیدهای جدید بخوانند.
    try:
        for tbl in ("forecasts", "forecast_notes"):
            rows = conn.execute(f"SELECT DISTINCT fiscal_year FROM {tbl}").fetchall()
            for (fy,) in rows:
                fy_norm = normalize_digits(str(fy))
                if fy_norm != fy:
                    # اگر کلیدِ نرمال‌شده از قبل وجود دارد، نسخهٔ فارسی را حذف کن؛ وگرنه تغییر بده
                    exists = conn.execute(
                        f"SELECT 1 FROM {tbl} WHERE fiscal_year=? LIMIT 1", (fy_norm,)).fetchone()
                    if exists:
                        conn.execute(f"DELETE FROM {tbl} WHERE fiscal_year=?", (fy,))
                    else:
                        conn.execute(f"UPDATE {tbl} SET fiscal_year=? WHERE fiscal_year=?",
                                     (fy_norm, fy))
        conn.commit()
    except Exception:
        pass
    conn.close()


def load_forecast(symbol, fiscal_year):
    """پیش‌بینی ذخیره‌شده را برمی‌گرداند: (df[label,value,row_order], note) یا (خالی, '')"""
    try:
        _forecast_init_db()
        conn = sqlite3.connect(FORECAST_DB_PATH)
        df = pd.read_sql_query(
            "SELECT row_order, label, value FROM forecasts "
            "WHERE symbol=? AND fiscal_year=? ORDER BY row_order",
            conn, params=(symbol, normalize_digits(str(fiscal_year))))
        note_row = conn.execute(
            "SELECT note FROM forecast_notes WHERE symbol=? AND fiscal_year=?",
            (symbol, normalize_digits(str(fiscal_year)))).fetchone()
        conn.close()
        return df, (note_row[0] if note_row else "")
    except Exception:
        return pd.DataFrame(columns=["row_order", "label", "value"]), ""


def save_forecast(symbol, fiscal_year, rows, note=""):
    """rows = لیستی از (label, value). کلِ پیش‌بینیِ این نماد/سال را جایگزین می‌کند."""
    import datetime as _dt
    _forecast_init_db()
    fiscal_year = normalize_digits(str(fiscal_year))   # کلید همیشه با ارقام انگلیسی
    now = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(FORECAST_DB_PATH)
    conn.execute("DELETE FROM forecasts WHERE symbol=? AND fiscal_year=?",
                 (symbol, fiscal_year))
    for i, (label, value) in enumerate(rows):
        conn.execute(
            "INSERT INTO forecasts (symbol, fiscal_year, row_order, label, value, updated_at) "
            "VALUES (?,?,?,?,?,?)",
            (symbol, fiscal_year, i, str(label),
             (float(value) if value is not None and value != "" else None), now))
    conn.execute(
        "INSERT OR REPLACE INTO forecast_notes (symbol, fiscal_year, note, updated_at) "
        "VALUES (?,?,?,?)", (symbol, fiscal_year, note or "", now))
    conn.commit()
    conn.close()


def list_forecast_symbols():
    """فهرستِ همهٔ نمادهایی که پیش‌بینی برای آن‌ها ذخیره شده."""
    try:
        _forecast_init_db()
        conn = sqlite3.connect(FORECAST_DB_PATH)
        df = pd.read_sql_query(
            "SELECT DISTINCT symbol, fiscal_year, updated_at FROM forecasts "
            "ORDER BY symbol, fiscal_year DESC", conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


# ════════════════════════════════════════════════════════════════
# توابع کمکی برای صورت سود و زیان
# ════════════════════════════════════════════════════════════════
def _inc_year_month(period_end):
    """پخش period_end (YYYY/M/D) به (year, month_int)."""
    try:
        y, m, _ = period_end.split("/")
        return int(y), int(m)
    except (ValueError, AttributeError, IndexError):
        return None, None


def _inc_label_key(label):
    """کلیدِ نرمال‌شده‌ای از عنوانِ ردیف برای مقایسهٔ عنوان‌های مختلف."""
    if not label:
        return ""
    return normalize_digits(str(label).lower())


def _inc_build_matrix(reports_df, items_df):
    """ساخت ماتریس متامورفوز سود و زیان: rows = عناوین، cols = گزارش‌ها (بزرگ‌ترین duration اول)
    بازگشت: (labels, parents, matrix{col_idx: {row_idx: val}}, customdata)
    customdata برای Plotly = column titles (period_end + duration_months)."""
    if reports_df.empty or items_df.empty:
        return [], [], {}, []

    # ترتیب گزارش‌ها: بزرگ‌ترین duration/period_end اول
    reports_sorted = reports_df.sort_values(
        ["period_end", "duration_months"], ascending=[False, False]
    ).reset_index(drop=True)

    report_ids = reports_sorted["id"].tolist()
    customdata = [
        f"{row['period_end']} ({row['duration_months']}m)"
        for _, row in reports_sorted.iterrows()
    ]

    # جمع‌آوریِ عناوین و والدین
    labels = []
    label_to_idx = {}

    for rid in report_ids:
        items = items_df[items_df["report_id"] == rid].sort_values("row_order")
        for _, item in items.iterrows():
            lbl = item["label"]
            if lbl not in label_to_idx:
                label_to_idx[lbl] = len(labels)
                labels.append(lbl)

    # ماتریس: matrix[col_idx][row_idx] = value
    matrix = {}
    for col_idx, rid in enumerate(report_ids):
        matrix[col_idx] = {}
        items = items_df[items_df["report_id"] == rid]
        for _, item in items.iterrows():
            lbl = item["label"]
            row_idx = label_to_idx.get(lbl)
            if row_idx is not None:
                matrix[col_idx][row_idx] = item["value"]

    parents = [""] * len(labels)  # والدین خالی (بدون سلسله‌مراتب)
    return labels, parents, matrix, customdata


# ════════════════════════════════════════════════════════════════
# صفحه اصلی
# ════════════════════════════════════════════════════════════════

# لود دیتا
df_main = load_main()
df_products = load_products()
df_product_pl = load_product_pl()
df_income_reports = load_income_reports()
df_income_items = load_income_items()
df_assembly = load_assembly()

has_income = not df_income_reports.empty and not df_income_items.empty
has_annual = not df_income_reports[df_income_reports["duration_months"] == 12].empty if has_income else False
has_assembly = not df_assembly.empty

# لیستِ نمادهای دسترس‌پذیر
symbols_available = sorted(df_main["symbol"].unique().tolist()) if not df_main.empty else []

# منوی نوار کناری
with st.sidebar:
    st.title("🔍 فیلترها")

    # انتخاب نماد
    selected_global = st.selectbox(
        "نماد",
        options=symbols_available,
        key="global_symbol"
    )

    # صفحات
    page = st.radio(
        "صفحه",
        options=[
            "📊 درآمد ماهانه",
            "🏭 محصولات ماهانه",
            "📈 سود و زیان دوره‌ای",
            "📉 سود و زیان محصولات",
            "💰 مالکیت و سهام",
            "🔮 پیش‌بینی سود",
        ],
        key="page_select"
    )

# ════════════════════════════════════════════════════════════════
# صفحه ۱: درآمد ماهانه
# ════════════════════════════════════════════════════════════════
if page == "📊 درآمد ماهانه":
    st.title("درآمد ماهانه")
    sym_monthly = selected_global

    df_sym = df_main[df_main["symbol"] == sym_monthly].sort_values("period_end")

    if not df_sym.empty:
        # چارت
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_sym["period_end"],
            y=df_sym["total_billion_toman"],
            mode="lines+markers",
            name="کل درآمد",
            line=dict(color="#3b82f6", width=2),
            marker=dict(size=6)
        ))
        fig.update_layout(
            title=f"درآمد کل: {sym_monthly}",
            xaxis_title="دوره",
            yaxis_title="میلیارد تومان",
            hovermode="x unified",
            font=PLOTLY_FONT,
            title_font=PLOTLY_FONT_TITLE,
            plot_bgcolor="#0f1419",
            paper_bgcolor="#0f1419",
            font_color="#e2e8f0"
        )
        st.plotly_chart(fig, use_container_width=True)

        # جدول
        df_display = pd.DataFrame({
            "دوره": df_sym["period_end"],
            "کل درآمد (م.ت)": df_sym["total_billion_toman"].map(lambda x: fmt_fa(x, 1)),
            "فروش داخل": df_sym["domestic_sales_btmn"].map(lambda x: fmt_fa(x, 1)),
            "صادرات": df_sym["export_sales_btmn"].map(lambda x: fmt_fa(x, 1)),
            "برگشت": df_sym["sales_return_btmn"].map(lambda x: fmt_fa(x, 1)),
            "تخفیفات": df_sym["discounts_btmn"].map(lambda x: fmt_fa(x, 1)),
        })
        render_table(df_display)
    else:
        st.warning(f"اطلاعات ماهانه برای {sym_monthly} موجود نیست.")


# ════════════════════════════════════════════════════════════════
# صفحه ۲: محصولات ماهانه
# ════════════════════════════════════════════════════════════════
elif page == "🏭 محصولات ماهانه":
    st.title("محصولات ماهانه")
    sym_prod = selected_global

    df_sym_prod = df_products[df_products["symbol"] == sym_prod]

    if not df_sym_prod.empty:
        # انتخاب دوره
        periods = sorted(df_sym_prod["period_end"].unique())
        period_selected = st.selectbox("دوره", periods, key="period_products")

        df_period = df_sym_prod[df_sym_prod["period_end"] == period_selected]

        if not df_period.empty:
            # جدول محصولات
            df_display = pd.DataFrame({
                "دسته": df_period["category"],
                "محصول": df_period["product_name"],
                "واحد": df_period["unit"],
                "تولید": df_period["production_qty"].map(lambda x: fmt_fa(x, 0) if pd.notna(x) else "—"),
                "فروش": df_period["sales_qty"].map(lambda x: fmt_fa(x, 0) if pd.notna(x) else "—"),
                "نرخ (ریال)": df_period["sales_rate_rial"].map(lambda x: fmt_fa(x, 0) if pd.notna(x) else "—"),
                "مبلغ (م.ت)": df_period["sales_amount_btmn"].map(lambda x: fmt_fa(x, 1) if pd.notna(x) else "—"),
            })
            render_table(df_display)
        else:
            st.info("محصولاتی برای این دوره موجود نیست.")
    else:
        st.warning(f"اطلاعات محصول برای {sym_prod} موجود نیست.")


# ════════════════════════════════════════════════════════════════
# صفحه ۳: صورت سود و زیان دوره‌ای
# ════════════════════════════════════════════════════════════════
elif page == "📈 سود و زیان دوره‌ای":
    st.title("صورت سود و زیان دوره‌ای")
    sym_income = selected_global

    if not has_income:
        st.warning("اطلاعات صورت سود و زیان موجود نیست.")
    else:
        df_sym_reports = df_income_reports[df_income_reports["symbol"] == sym_income]

        if not df_sym_reports.empty:
            # انتخاب گزارش
            report_labels = [
                f"{row['period_end']} ({row['duration_months']}ماه) — {row['title'][:40]}"
                for _, row in df_sym_reports.iterrows()
            ]
            report_idx = st.selectbox("گزارش", range(len(report_labels)),
                                     format_func=lambda i: report_labels[i], key="income_report")

            selected_report = df_sym_reports.iloc[report_idx]
            report_id = selected_report["id"]

            df_items = df_income_items[df_income_items["report_id"] == report_id]

            if not df_items.empty:
                df_display = pd.DataFrame({
                    "عنوان": df_items["label"],
                    "مقدار (میلیون ریال)": df_items["value"].map(lambda x: fmt_fa(x/1000000, 1) if pd.notna(x) else "—"),
                })
                render_table(df_display)
            else:
                st.info("اقلام گزارش موجود نیست.")
        else:
            st.warning(f"گزارش سود و زیان برای {sym_income} موجود نیست.")


# ════════════════════════════════════════════════════════════════
# صفحه ۴: محصول P&L
# ════════════════════════════════════════════════════════════════
elif page == "📉 سود و زیان محصولات":
    st.title("سود و زیان محصولات")
    sym_ppl = selected_global

    if not df_product_pl.empty:
        df_sym_ppl = df_product_pl[df_product_pl["symbol"] == sym_ppl]

        if not df_sym_ppl.empty:
            # انتخاب دوره
            periods_ppl = sorted(df_sym_ppl["period_end"].unique())
            period_ppl = st.selectbox("دوره", periods_ppl, key="period_ppl")

            df_period_ppl = df_sym_ppl[df_sym_ppl["period_end"] == period_ppl]

            if not df_period_ppl.empty:
                df_display = pd.DataFrame({
                    "محصول": df_period_ppl["product"],
                    "فروش (م.ت)": df_period_ppl["sales"].map(lambda x: fmt_fa(x, 1) if pd.notna(x) else "—"),
                    "بهای تمام": df_period_ppl["cost"].map(lambda x: fmt_fa(x, 1) if pd.notna(x) else "—"),
                    "سود ناخالص": df_period_ppl["gross"].map(lambda x: fmt_fa(x, 1) if pd.notna(x) else "—"),
                })
                render_table(df_display)
        else:
            st.warning(f"اطلاعات محصول P&L برای {sym_ppl} موجود نیست.")
    else:
        st.warning("دیتابیس محصول P&L خالی است.")


# ════════════════════════════════════════════════════════════════
# صفحه ۵: مالکیت و سهام
# ════════════════════════════════════════════════════════════════
elif page == "💰 مالکیت و سهام":
    st.title("مالکیت و سهام")
    sym_hold = selected_global

    st.info(f"نماد فعال: **{sym_hold}** — اطلاعات مالکیت و سهام در دسترس است.")


# ════════════════════════════════════════════════════════════════
# صفحه ۶: پیش‌بینی سود
# ════════════════════════════════════════════════════════════════
elif page == "🔮 پیش‌بینی سود":
    st.title("پیش‌بینی سود")
    sym_fc = selected_global

    st.caption(f"نماد فعال: **{sym_fc}** — مقادیر به **میلیارد تومان** وارد می‌شوند "
               "(ردیفِ سود هر سهم به ریال).")

    # انتخاب سال مالیِ پیش‌بینی
    _years_have = []
    if has_annual:
        _ra = df_income_reports[df_income_reports["symbol"] == sym_fc]
        _ra = _ra[_ra["duration_months"] == 12]
        for _pe in _ra["period_end"].tolist():
            _yy, _mm = _inc_year_month(_pe)
            if _yy:
                _years_have.append(_yy)

    _last_actual_year = max(_years_have) if _years_have else None
    _default_year = (_last_actual_year + 1) if _last_actual_year else 1405

    c_y1, c_y2 = st.columns([1, 3])
    with c_y1:
        fiscal_year = st.number_input(
            "سال مالی پیش‌بینی:",
            min_value=1390, max_value=1430,
            value=int(_default_year), step=1, key="fc_year"
        )

    fiscal_year = str(int(fiscal_year))

    # لود پیش‌بینیِ ذخیره‌شده
    saved_df, saved_note = load_forecast(sym_fc, fiscal_year)

    # اقلام پیش‌فرض
    base_labels = ["درآمدهای عملیاتی", "بهای تمام شدهٔ درآمدهای عملیاتی",
                   "سود (زیان) ناخالص", "هزینه‌های فروش، اداری و عمومی",
                   "سود (زیان) عملیاتی", "سود (زیان) خالص",
                   "سود (زیان) خالص هر سهم – ریال"]

    if not saved_df.empty:
        seed_vals = dict(zip(saved_df["label"], saved_df["value"]))
        src_note = "از پیش‌بینیِ ذخیره‌شدهٔ قبلی"
    else:
        seed_vals = {}
        src_note = "اقلام پیش‌فرض"

    st.caption(f"منبع اقلام: {src_note}")

    # جدول قابل‌ویرایش
    editor_rows = []
    for lab in base_labels:
        editor_rows.append({
            "شرح": lab,
            f"پیش‌بینی {to_fa_digits(fiscal_year)}": seed_vals.get(lab, None),
        })
    editor_df = pd.DataFrame(editor_rows)

    st.markdown("##### اقلام صورت سود و زیان")
    edited = st.data_editor(
        editor_df,
        key=f"fc_editor_{sym_fc}_{fiscal_year}",
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "شرح": st.column_config.TextColumn("شرح", width="large"),
            f"پیش‌بینی {to_fa_digits(fiscal_year)}": st.column_config.NumberColumn(
                f"پیش‌بینی {to_fa_digits(fiscal_year)}",
                help="میلیارد تومان (EPS: ریال)",
                format="%.2f"
            ),
        },
        hide_index=True,
    )

    note_txt = st.text_area(
        "یادداشت (اختیاری):",
        value=saved_note or "",
        key=f"fc_note_{sym_fc}_{fiscal_year}",
        height=90
    )

    fc_col = f"پیش‌بینی {to_fa_digits(fiscal_year)}"
    c_b1, c_b2, _ = st.columns([1, 1, 3])

    with c_b1:
        if st.button("💾 ذخیرهٔ پیش‌بینی", key="fc_save", type="primary"):
            rows = [(r["شرح"], r[fc_col]) for _, r in edited.iterrows()
                    if str(r.get("شرح", "")).strip() != ""]
            save_forecast(sym_fc, fiscal_year, rows, note_txt)
            st.success(f"پیش‌بینیِ «{sym_fc}» برای سال {to_fa_digits(fiscal_year)} ذخیره شد.")

    with c_b2:
        if st.button("🗑️ پاک‌کردن این پیش‌بینی", key="fc_del"):
            save_forecast(sym_fc, fiscal_year, [], "")
            st.warning("پیش‌بینی پاک شد. (صفحه را رفرش کن)")

    # همهٔ پیش‌بینی‌ها
    st.markdown("---")
    with st.expander("📚 همهٔ پیش‌بینی‌های ذخیره‌شده"):
        all_fc = list_forecast_symbols()
        if all_fc.empty:
            st.caption("هنوز پیش‌بینی‌ای ذخیره نشده.")
        else:
            disp_fc = pd.DataFrame({
                "نماد": all_fc["symbol"].values,
                "سال مالی": [to_fa_digits(y) for y in all_fc["fiscal_year"].values],
                "آخرین به‌روزرسانی": all_fc["updated_at"].values,
            })
            render_table(disp_fc)