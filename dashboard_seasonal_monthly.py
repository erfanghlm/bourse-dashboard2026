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

def _assign_year_month(df, parser, period_col="period_end", y_col="سال", m_col="ماه_num"):
    """انتساب امنِ ستون‌های سال/ماه از روی period_end.
    مقاوم در برابر دیتافریم خالی و تفاوت نسخه‌های پانداس
    (جایگزین الگوی شکنندهٔ df[[a,b]] = s.apply(lambda: pd.Series(...)))."""
    _pe = df[period_col].map(parser)
    df[y_col] = _pe.map(lambda t: t[0] if isinstance(t, tuple) else None)
    df[m_col] = _pe.map(lambda t: t[1] if isinstance(t, tuple) else None)
    return df
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
        for ci, val in enumerate(row):
            td = (TD_CTR if (center_values and ci > 0) else TD_CSS)
            cs = col_styles.get(cols[ci], "")
            html += f'<td style="{td}{cs}">{val}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    st.markdown(html, unsafe_allow_html=True)



COLOR_TOMAN  = "#4DA3FF"
COLOR_DOLLAR = "#10B981"


_SRT_FA = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "0123456789" * 2)


def _srt_key(v, numeric):
    """کلیدِ مرتب‌سازی؛ عددی: پرانتز = منفی، ارقام فارسی→لاتین."""
    txt = re.sub(r"<[^>]+>", "", str(v))
    if not numeric:
        return txt.strip()
    t = txt.translate(_SRT_FA)
    neg = "(" in t and ")" in t
    t = t.replace(",", "").replace("٪", "").replace("%", "")
    t = re.sub(r"[^0-9.]", "", t)
    if t in ("", "."):
        return -1e308
    try:
        x = float(t)
        return -x if neg else x
    except Exception:
        return -1e308


def render_sortable_table(df, numeric_cols=(), green_cols=(), center_from=2, row_height=32,
                          sortkeys=None):
    """جدولِ HTML دارک، قابلِ مرتب‌سازی با کلیک روی ستون و رنگی:
    منفی‌ها (داخل پرانتز) قرمز، ستون‌های green_cols مثبت سبز.
    sortkeys: dict اختیاری {نام ستون: لیستِ کلیدهای عددی هم‌ترازِ ردیف‌ها} برای سورتِ سفارشی
              (مثلاً ستونِ «فصل» بر مبنای زمان)."""
    TABLE = ("width:100%;border-collapse:collapse;direction:rtl;"
             "font-family:Vazirmatn,Tahoma,sans-serif;font-size:13px;color:#e2e8f0;")
    TH = ("background:#1e293b;padding:8px 11px;text-align:right;border:1px solid #334155;"
          "font-weight:600;cursor:pointer;user-select:none;white-space:nowrap;")
    THC = TH.replace("text-align:right", "text-align:center")
    TD = "padding:7px 11px;text-align:right;border:1px solid #2d3748;"
    TDC = "padding:7px 11px;text-align:center;direction:ltr;border:1px solid #2d3748;"
    cols = list(df.columns)
    numeric_cols = set(numeric_cols)
    green_cols = set(green_cols)
    sortkeys = sortkeys or {}
    n = len(df)

    head = ""
    for ci, c in enumerate(cols):
        style = THC if ci >= center_from else TH
        head += (f'<th style="{style}" onclick="srt(this,{ci})">{c}'
                 f'<span class="ar" style="opacity:.45"> ⇅</span></th>')
    body = ""
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#0f172a" if i % 2 == 0 else "#1a2236"
        body += f'<tr style="background:{bg}">'
        for ci, (c, v) in enumerate(zip(cols, row)):
            sval = str(v)
            color = ""
            if sval.startswith("(") and sval.endswith(")"):
                color = "color:#EF4444;"           # منفی → قرمز
            elif c in green_cols and sval not in ("—", "", "۰"):
                color = "color:#10B981;"           # سود/حاشیهٔ مثبت → سبز
            style = (TDC if ci >= center_from else TD) + color
            if c in sortkeys:
                sk = sortkeys[c][i]
            else:
                sk = _srt_key(v, c in numeric_cols)
            body += f'<td style="{style}" data-sort="{sk}">{v}</td>'
        body += "</tr>"

    height = 70 + row_height * max(1, n) + 16
    html = f"""<!doctype html><html><head><meta charset="utf-8">
<link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css" rel="stylesheet">
<style>
  html,body{{margin:0;padding:0;background:transparent;}}
  table{{{TABLE}}}
  th:hover{{background:#243043 !important;}}
  ::-webkit-scrollbar{{height:8px;width:8px;}}
  ::-webkit-scrollbar-thumb{{background:#334155;border-radius:4px;}}
</style></head><body>
<div style="overflow-x:auto"><table id="t"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>
<script>
var dir={{}};
function srt(th,col){{
  var tb=document.getElementById('t').tBodies[0];
  var rows=Array.prototype.slice.call(tb.rows);
  var d=dir[col]==='asc'?'desc':'asc'; dir={{}}; dir[col]=d;
  rows.sort(function(a,b){{
    var x=a.cells[col].getAttribute('data-sort'), y=b.cells[col].getAttribute('data-sort');
    var nx=parseFloat(x), ny=parseFloat(y);
    var r=(!isNaN(nx)&&!isNaN(ny))?(nx-ny):String(x).localeCompare(String(y),'fa');
    return d==='asc'?r:-r;
  }});
  rows.forEach(function(r){{tb.appendChild(r);}});
  var ars=document.querySelectorAll('#t thead .ar');
  ars.forEach(function(s){{s.textContent=' ⇅'; s.style.opacity='.45';}});
  var sp=th.querySelector('.ar'); if(sp){{sp.textContent=d==='asc'?' ▲':' ▼'; sp.style.opacity='1';}}
}}
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
    """نمادهایی که برایشان پیش‌بینی ذخیره شده + آخرین به‌روزرسانی."""
    try:
        _forecast_init_db()
        conn = sqlite3.connect(FORECAST_DB_PATH)
        df = pd.read_sql_query(
            "SELECT symbol, fiscal_year, MAX(updated_at) AS updated_at "
            "FROM forecasts GROUP BY symbol, fiscal_year ORDER BY updated_at DESC", conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


def render_forecast_block(symbol, kind, labels, actual_vals=None,
                          actual_label="آخرین واقعی", unit_note="میلیارد تومان",
                          quarterly=False):
    """یک بلوکِ خودکفا برای نوشتن/ذخیرهٔ پیش‌بینی، زیر جدول‌های گزارش.
    kind: برچسبِ نوع ('فصلی' یا 'سالانه') که در کلیدِ ذخیره‌سازی می‌آید تا با هم قاطی نشوند.
    labels: لیست عنوانِ اقلام (شرح). actual_vals: dict {شرح: مقدار واقعی} برای ستونِ مرجع.
    quarterly: اگر True، یک انتخابگرِ فصل (Q۱..Q۴) اضافه می‌شود و هر فصل جدا ذخیره می‌شود."""
    import streamlit as _st
    actual_vals = actual_vals or {}
    _st.markdown(f"##### ✍️ پیش‌بینی من ({kind})")

    # سالِ مالیِ پیش‌بینی برای این بلوک
    _yk = f"fcq_year_{symbol}_{kind}"
    if quarterly:
        cc1, cc2, _cc3 = _st.columns([1, 1, 2])
        with cc1:
            fy = _st.number_input("سال مالی (شمسی):", min_value=1390, max_value=1430,
                                  value=1405, step=1, key=_yk)
        with cc2:
            q = _st.selectbox("فصل:", ["فصل ۱ (۳ماهه)", "فصل ۲ (۶ماهه)",
                                       "فصل ۳ (۹ماهه)", "فصل ۴ (۱۲ماهه)"],
                              key=f"fcq_q_{symbol}_{kind}")
        q_tag = q.split(" ")[1]                      # «۱».. «۴»
        fyear = f"{kind}-{int(fy)}-Q{q_tag}"         # کلیدِ ذخیره‌سازیِ مجزا برای هر فصل
        _suffix = f"{int(fy)}_{q_tag}"
    else:
        cc1, _cc2 = _st.columns([1, 3])
        with cc1:
            fy = _st.number_input("سال مالی پیش‌بینی (شمسی):", min_value=1390, max_value=1430,
                                  value=1405, step=1, key=_yk)
        fyear = f"{kind}-{int(fy)}"
        _suffix = str(int(fy))

    saved_df, saved_note = load_forecast(symbol, fyear)
    if saved_df is not None and not saved_df.empty:
        base_labels = saved_df["label"].tolist()
        seed = dict(zip(saved_df["label"], saved_df["value"]))
    else:
        base_labels = list(labels) if labels else [
            "درآمدهای عملیاتی", "بهای تمام شدهٔ درآمدهای عملیاتی", "سود (زیان) ناخالص",
            "سود (زیان) عملیاتی", "سود (زیان) خالص", "سود (زیان) خالص هر سهم – ریال"]
        seed = {}

    fc_col = f"پیش‌بینی {to_fa_digits(int(fy))}"

    def _is_eps(lab):
        return "هرسهم" in _inc_label_key(lab)

    rows = []
    for lab in base_labels:
        _raw = actual_vals.get(lab, None)
        # تبدیل مقدارِ مرجع از «میلیون ریال» به «میلیارد تومان» (به‌جز ردیفِ EPS که ریال است)
        if _raw is not None and not pd.isna(_raw) and not _is_eps(lab):
            _ref = _raw / 10000.0
        else:
            _ref = _raw
        rows.append({"شرح": lab, actual_label: _ref, fc_col: seed.get(lab, None)})
    edf = pd.DataFrame(rows)

    edited = _st.data_editor(
        edf, key=f"fcq_editor_{symbol}_{kind}_{_suffix}",
        use_container_width=True, num_rows="dynamic", hide_index=True,
        column_config={
            "شرح": _st.column_config.TextColumn("شرح", width="large"),
            actual_label: _st.column_config.NumberColumn(
                actual_label, disabled=True, format="%.2f",
                help="آخرین مقدار واقعی (میلیارد تومان؛ ردیفِ EPS به ریال) — فقط مرجع"),
            fc_col: _st.column_config.NumberColumn(
                fc_col, format="%.2f", help=f"مقدار پیش‌بینیِ خودت ({unit_note}؛ ردیفِ EPS به ریال)"),
        },
    )
    note_txt = _st.text_area("📝 یادداشت:", value=saved_note or "",
                             key=f"fcq_note_{symbol}_{kind}_{_suffix}", height=90)

    b1, b2, _b3 = _st.columns([1, 1, 3])
    with b1:
        if _st.button("💾 ذخیرهٔ پیش‌بینی", key=f"fcq_save_{symbol}_{kind}_{_suffix}",
                     type="primary"):
            out = [(r["شرح"], r[fc_col]) for _, r in edited.iterrows()
                   if str(r.get("شرح", "")).strip() != ""]
            save_forecast(symbol, fyear, out, note_txt)
            _st.success(f"پیش‌بینیِ «{symbol}» ({kind} {to_fa_digits(int(fy))}) ذخیره شد.")
    with b2:
        if _st.button("🗑️ پاک‌کردن", key=f"fcq_del_{symbol}_{kind}_{_suffix}"):
            save_forecast(symbol, fyear, [], "")
            _st.warning("پاک شد. (صفحه را رفرش کن)")


def _inc_label_key(s):
    """کلید مقایسه برچسب اقلام (مطابق منطق استخراج‌کننده برای ادغام ردیف‌های مشابه)."""
    if not s:
        return ""
    s = str(s)
    for ch in ('\u200c', '\u200f', '\u200e', '\u200b', '\u00a0'):
        s = s.replace(ch, ' ')
    s = s.replace('ي', 'ی').replace('ك', 'ک').replace('ى', 'ی')
    s = re.sub(r'\s*\(\d+\)\s*$', '', s)            # حذف شماره تکراری مثل (2)
    s = re.sub(r'[\s،,()\u200c\-–—]+', '', s)        # حذف فاصله و علائم
    return s


def _inc_year_month(period_end):
    """'1403/06/31' → (سال, ماه) عددی یا (None, None)"""
    try:
        y, m, _ = normalize_digits(period_end).split("/")
        return int(y), int(m)
    except Exception:
        return None, None


def _inc_fy_key(period_end, duration):
    """کلید سال مالی: ماه قبل از شروع دوره — برای گروه‌بندی گزارش‌های هم‌سال‌مالی."""
    y, m = _inc_year_month(period_end)
    if y is None or not duration:
        return None
    return (y * 12 + m) - int(duration)


def _inc_back_ym(period_end, months):
    """کلیدهای 'YYYY/MM' برای `months` ماهِ منتهی به period_end (شامل خودِ ماهِ پایان)."""
    y, m = _inc_year_month(period_end)
    if y is None or not months:
        return []
    base = y * 12 + (m - 1)          # اندیس صفرمبنای ماه پایان
    keys = []
    for i in range(int(months)):
        t = base - i
        yy, mm = t // 12, t % 12 + 1
        keys.append(f"{yy}/{mm:02d}")
    return keys


def _inc_avg_rate(period_end, months, usd_rates):
    """میانگین نرخ دلار (تومان/دلار) روی بازهٔ `months` ماهِ منتهی به period_end.
    اگر هیچ نرخی در بازه نبود → None؛ اگر months خالی بود → نرخ همان ماهِ پایان."""
    if not months:
        r = usd_rates.get(_ym_key(period_end))
        return r if (r and r > 0) else None
    rs = [usd_rates[k] for k in _inc_back_ym(period_end, months)
          if usd_rates.get(k) and usd_rates[k] > 0]
    return (sum(rs) / len(rs)) if rs else None


# ─── گزارش‌های سالانه (codal_annual.db) — هم‌اسکیمای دیتابیس دوره‌ای ──
@st.cache_data(ttl=60)
def load_annual_reports():
    try:
        conn = sqlite3.connect(ANNUAL_DB_PATH)
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
def load_annual_items():
    try:
        conn = sqlite3.connect(ANNUAL_DB_PATH)
        df = pd.read_sql_query("""
            SELECT report_id, row_order, label, value
            FROM line_items ORDER BY report_id, row_order
        """, conn)
        conn.close()
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=3600)
def load_annual_usd_rates():
    """نرخ دلار سالانه از usd_rate_annual.xlsx. ستون۰=سال، ستون۱=نرخ به تومان.
    خروجی: dict {سال(int) → نرخ تومان/دلار(float)}."""
    try:
        df = pd.read_excel(ANNUAL_USD_RATE_PATH, header=None)
        rates = {}
        for _, row in df.iterrows():
            try:
                yr = int(normalize_digits(str(row[0])).strip())
                rt = float(normalize_digits(str(row[1])).replace(",", "").strip())
            except Exception:
                continue
            if yr > 1000 and rt > 0:
                rates[yr] = rt          # تومان به ازای هر دلار (مستقیم)
        return rates
    except Exception:
        return {}


@st.cache_data(ttl=300)
def load_latest_daily():
    """آخرین فایل اکسل روزانهٔ معاملات در D:/bourse/daily2 را می‌خواند.
    خروجی: (df, نام‌فایل) یا (DataFrame خالی, None)."""
    try:
        files = glob.glob(os.path.join(DAILY_DIR, "*.xlsx"))
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if not files:
            return pd.DataFrame(), None
        latest = max(files, key=os.path.getmtime)
        # key (همان insCode) عددِ ۱۷ رقمی است؛ حتماً به‌صورت رشته خوانده شود
        # وگرنه به‌خاطر محدودیتِ دقتِ float رقم‌های آخرش خراب می‌شود (…562 → …560)
        return pd.read_excel(latest, dtype={"key": str}), os.path.basename(latest)
    except Exception:
        return pd.DataFrame(), None


@st.cache_data(ttl=300)
def load_inscode_map():
    """نگاشتِ نامِ‌نرمال‌شدهٔ نماد → insCode (رشته) از stck.xlsx.
    این فایل id را به‌صورت رشتهٔ سالم نگه می‌دارد، پس مرجعِ مطمئنِ insCode است."""
    try:
        df = pd.read_excel(STOCKS_PATH, dtype={"id": str})
        m = {}
        for _, r in df.iterrows():
            nm = normalize_symbol(r.get("name", ""))
            ic = str(r.get("id", "")).split(".")[0].strip()
            if nm and ic.isdigit() and len(ic) >= 10:
                m[nm] = ic
        return m
    except Exception:
        return {}


@st.cache_data(ttl=60)
def load_latest_shareholders():
    """آخرین فایل اکسل سهامداران در D:/bourse/shareholders را می‌خواند.
    ستون‌های موردانتظار: نماد، نام شرکت، نام سهامدار، نماد سهامدار، تعداد سهم، درصد، تاریخ، key.
    خروجی: (df, نام‌فایل) یا (DataFrame خالی, None)."""
    try:
        files = glob.glob(os.path.join(SHAREHOLDERS_DIR, "*.xlsx"))
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if not files:
            return pd.DataFrame(), None
        latest = max(files, key=os.path.getmtime)
        df = pd.read_excel(latest, dtype={"key": str})
        if "نماد سهامدار" not in df.columns:
            df["نماد سهامدار"] = ""
        df["نماد سهامدار"] = df["نماد سهامدار"].fillna("")
        return df, os.path.basename(latest)
    except Exception:
        return pd.DataFrame(), None


def _inc_build_matrix(reps_sorted, items_df):
    """از روی گزارش‌های مرتب‌شدهٔ یک نماد و اقلام، ماتریس مشترک می‌سازد.
    خروجی: (master_labels, master_keys, col_meta, matrix{ci:[val...]})."""
    rep_ids = reps_sorted["id"].tolist()
    items_sub = items_df[items_df["report_id"].isin(rep_ids)]
    by_rep = {}
    for rid, g in items_sub.groupby("report_id"):
        gg = g.sort_values("row_order")
        by_rep[rid] = list(zip(gg["label"].tolist(), gg["value"].tolist()))
    col_meta  = [rep for _, rep in reps_sorted.iterrows()]
    col_items = [by_rep.get(rep["id"], []) for rep in col_meta]

    master_labels, master_keys = [], []
    for items in col_items:
        for label, _v in items:
            k = _inc_label_key(label)
            if not k:
                continue
            if not any(mk == k or SequenceMatcher(None, mk, k).ratio() >= 0.9 for mk in master_keys):
                master_labels.append(label)
                master_keys.append(k)

    def _val_for(items, mk):
        for label, value in items:
            kk = _inc_label_key(label)
            if kk == mk or SequenceMatcher(None, kk, mk).ratio() >= 0.9:
                return value
        return None

    matrix = {ci: [_val_for(it, mk) for mk in master_keys] for ci, it in enumerate(col_items)}
    return master_labels, master_keys, col_meta, matrix


# ─── تشخیص ردیف‌های کلیدی (برای هایلایت) ─────────────────────────
def _stmt_is_gross(k):    # سود (زیان) ناخالص
    return ("ناخالص" in k) and ("هرسهم" not in k)


def _stmt_is_net(k):      # سود (زیان) خالص (نه ناخالص، نه هر سهم)
    return ("خالص" in k) and ("ناخالص" not in k) and ("هرسهم" not in k)


def _stmt_find(keys, pred):
    for i, k in enumerate(keys):
        if pred(k):
            return i
    return None


def _stmt_key_indices(master_keys):
    """اندیس ردیف‌های درآمد عملیاتی، سود ناخالص، و سود خالص عملیات در حال تداوم."""
    rev = _stmt_find(master_keys, lambda k: ("درآمدعملیاتی" in k or "درآمدهایعملیاتی" in k) and "هرسهم" not in k)
    if rev is None:
        rev = _stmt_find(master_keys, lambda k: k.startswith("درآمد") and "هرسهم" not in k)
    gross = _stmt_find(master_keys, _stmt_is_gross)
    # سود خالص «عملیات در حال تداوم» — «متوقف‌شده» هرگز هایلایت نمی‌شود
    netc = _stmt_find(master_keys, lambda k: _stmt_is_net(k) and ("حالتداوم" in k))
    if netc is None:
        netc = _stmt_find(master_keys, lambda k: _stmt_is_net(k) and ("متوقف" not in k) and ("حالتداوم" not in k))
    return rev, gross, netc


@st.cache_data(ttl=60, show_spinner=False)
def _avg_net_margin(sym, years=5):
    """میانگین حاشیه سود خالص = سود خالص ÷ درآمد عملیاتی، روی آخرین `years` گزارش
    سالانهٔ ۱۲ماهه (از codal_income.db). کسر برمی‌گرداند (مثلاً ۰٫۱۸)؛ اگر داده نبود None."""
    try:
        reps = df_income_rep[(df_income_rep["symbol"] == sym) &
                             (df_income_rep["duration_months"] == 12)].copy()
        if reps.empty:
            return None
        reps = reps.sort_values("period_end", ascending=False).head(years)
        ml, mk, cmeta, mtx = _inc_build_matrix(reps, df_income_items)
        rev_i, _g, netc_i = _stmt_key_indices(mk)
        if rev_i is None or netc_i is None:
            return None
        margins = []
        for ci in range(len(cmeta)):
            rev = mtx[ci][rev_i]
            net = mtx[ci][netc_i]
            if rev and net is not None and rev != 0:
                margins.append(net / rev)
        return (sum(margins) / len(margins)) if margins else None
    except Exception:
        return None


@st.cache_data(ttl=60, show_spinner=False)
def _annual_net_musd_by_year(sym):
    """نگاشت {سال(int) → سود خالص دلاری (میلیون دلار)} از گزارش‌های سالانهٔ ۱۲ماهه.
    منبع: اول codal_annual.db؛ برای سال‌های ۱۴۰۲ به بعد که در آن نباشند، از
    گزارشِ ۱۲ماههٔ codal_income.db (همان منطقِ جدولِ «صورت سود و زیان سالانه»).
    «سود خالص» = سود خالص عملیات در حال تداوم. هر سال با نرخ دلار همان سال.
    فقط چند گزارشِ اخیر پردازش می‌شود (برای جلوگیری از ساخت ماتریسِ سنگین O(n²)
    روی کلِ تاریخچه؛ شاخص‌های c و d حداکثر به ~۶ سال اخیر نیاز دارند).
    (کش‌شده تا در حلقهٔ فیلتر برای هر نماد فقط یک‌بار ماتریس ساخته شود.)"""
    if not has_annual_usd:
        return {}
    try:
        reps, items, _fb = annual_reports_with_income_fallback(sym)
        if reps is None or reps.empty:
            return {}
        reps = reps[reps["duration_months"] == 12].copy()
        if reps.empty:
            return {}
        # فقط ۸ گزارشِ اخیر (پوششِ ≥۶ سال مالیِ متمایز حتی با وجود اصلاحیه‌ها)
        reps = reps.sort_values("period_end", ascending=False).head(8)
        _ml, mk, cmeta, mtx = _inc_build_matrix(reps, items)
        _rev_i, _gross_i, netc_i = _stmt_key_indices(mk)
        if netc_i is None:
            return {}
        out = {}
        for ci in range(len(cmeta)):
            net = mtx[ci][netc_i]
            if net is None or pd.isna(net):
                continue
            yy, _mm = _inc_year_month(cmeta[ci]["period_end"])
            rate = annual_usd_rates.get(yy)
            if not rate or rate <= 0:
                continue
            if yy not in out:                     # جدیدترین/حسابرسی‌شده اول می‌ماند
                out[yy] = (net / 10000.0) * 1000.0 / rate
        return out
    except Exception:
        return {}


def _top_net_income_musd(sym, years=5, top_n=2):
    """میانگینِ «top_n بزرگ‌ترین سود خالصِ دلاری» از میان آخرین `years` سالِ مالیِ
    ۱۲ماهه (codal_annual.db). خروجی به میلیون دلار یا None."""
    by_year = _annual_net_musd_by_year(sym)
    if not by_year:
        return None
    recent_years = sorted(by_year.keys(), reverse=True)[:max(1, years)]
    vals = sorted([by_year[y] for y in recent_years], reverse=True)
    if not vals:
        return None
    top = vals[:max(1, top_n)]
    return sum(top) / len(top)


def _avg_payout_ratio(sym):
    """میانگین نسبت تقسیم سود (dps ÷ eps) روی همهٔ سال‌های موجود در codal_assembly.db.
    هر سال جداگانه سقف ۱۰۰٪ می‌خورد (مثلاً ۱۲۰٪ → ۱۰۰٪) و سپس میانگین گرفته می‌شود.
    کسر برمی‌گرداند (مثلاً ۰٫۸ یعنی ۸۰٪)؛ اگر داده نبود None."""
    try:
        adf = df_assembly[df_assembly["symbol"] == sym]
    except Exception:
        return None
    if adf is None or adf.empty:
        return None
    vals = []
    for _, r in adf.iterrows():
        e, d = r.get("eps"), r.get("dps")
        if e is not None and pd.notna(e) and e > 0 and d is not None and pd.notna(d):
            vals.append(min(d / e, 1.0))          # سقف ۱۰۰٪ روی هر سال
    return (sum(vals) / len(vals)) if vals else None


@st.cache_data(ttl=60, show_spinner=False)
def _load_annual_forecast_nets(fiscal_year):
    """{symbol → سود خالص دلاری پیش‌بینی (م.دلار)} برای یک سال مالی.
    همهٔ پیش‌بینی‌های «سالانه-{year}» را یک‌بار می‌خواند (به‌جای باز کردن دیتابیس
    برای هر نماد در حلقهٔ فیلتر)."""
    if not has_annual_usd:
        return {}
    rate = annual_usd_rates.get(int(fiscal_year))
    if not rate or rate <= 0:
        return {}
    try:
        _forecast_init_db()
        conn = sqlite3.connect(FORECAST_DB_PATH)
        fdf = pd.read_sql_query(
            "SELECT symbol, row_order, label, value FROM forecasts "
            "WHERE fiscal_year=? ORDER BY symbol, row_order",
            conn, params=(f"سالانه-{normalize_digits(str(fiscal_year))}",))
        conn.close()
    except Exception:
        return {}
    out = {}
    if fdf.empty:
        return out
    for sym, g in fdf.groupby("symbol"):
        gg = g.sort_values("row_order")
        keys = [_inc_label_key(l) for l in gg["label"].tolist()]
        _rev_i, _gross_i, netc_i = _stmt_key_indices(keys)
        if netc_i is None:
            continue
        net = gg.iloc[netc_i]["value"]
        if net is None or pd.isna(net):
            continue
        out[sym] = (net / 10000.0) * 1000.0 / rate
    return out


def _current_year_dividend_musd(sym):
    """دیویدندِ دلاریِ سالِ مالیِ DIVIDEND_FY که امسال (مجمعِ DIVIDEND_FY+۱) توزیع می‌شود.
    = سودِ دلاریِ سال DIVIDEND_FY × میانگین نسبت تقسیم سود.
    منبعِ سود به‌ترتیبِ اولویت:
      ۱) گزارش سالانهٔ ۱۲ماههٔ همان سال (اگر منتشر شده)
      ۲) پیش‌بینیِ ذخیره‌شدهٔ همان سال
      ۳) سود دلاریِ سالِ قبل (DIVIDEND_FY−۱)
    خروجی: (dividend_musd یا None، منبع به‌صورت متن)."""
    by_year = _annual_net_musd_by_year(sym)
    net_musd = by_year.get(DIVIDEND_FY)
    src = "گزارش ۱۲ماهه"
    if net_musd is None:
        net_musd = _load_annual_forecast_nets(DIVIDEND_FY).get(sym)
        src = "پیش‌بینی"
    if net_musd is None:
        net_musd = by_year.get(DIVIDEND_FY - 1)
        src = f"سود دلاری {DIVIDEND_FY - 1}"
    if net_musd is None:
        return None, None
    payout = _avg_payout_ratio(sym)
    if payout is None:
        return None, None
    div = max(0.0, net_musd) * payout          # زیان → دیویدند صفر
    return div, src



# رنگ‌های هایلایت (دو رنگِ متفاوت): سود ناخالص = آبی، سود خالص = سبز
HL_GROSS = "background:#10233f;color:#7cc4ff;font-weight:700;"
HL_NET   = "background:#0f2e1a;color:#86efac;font-weight:700;"


def _stmt_hidden_row(k):
    """ردیف‌هایی که در جدول نمایش داده نمی‌شوند: هر سهم، و سرمایه."""
    return ("هرسهم" in k) or (k == "سرمایه") or k.startswith("سرمایه‌")


def render_statement(master_labels, master_keys, columns, rev_idx, gross_idx, netc_idx,
                     to_base, fmt_amount, fmt_pct,
                     view_mode, show_pct_rows, hide_small, extra_cols=None):
    """رندر یک صورت سود و زیان (دوره‌ای یا سالانه) با امکانات مشترک.
    columns: لیست dict با کلیدهای header، pe، months، rev(خام)، vals(لیست هم‌ترازِ master).
    to_base(value, period_end, months) → مقدار در واحد نمایش (تومان/دلار).
    extra_cols: ستون‌های اضافی (پیش‌بینی) که مقادیرشان از قبل در واحد نمایش‌اند؛
        هر کدام dict با header، disp_vals(لیست هم‌ترازِ master، واحدِ نمایش)، rev_disp(درآمدِ نمایش)."""
    extra_cols = extra_cols or []
    # رنگِ متمایز برای ستون‌های پیش‌بینی (هدر+سلول‌ها)
    FC_TINT = "background:#2e2640;color:#c4b5fd;"
    _col_styles = {xc["header"]: FC_TINT for xc in extra_cols}

    def _pct(ri, col):
        v, rv = col["vals"][ri], col["rev"]
        return (v / rv * 100.0) if (v is not None and rv not in (None, 0)) else None

    def _pct_x(ri, xc):
        v, rv = xc["disp_vals"][ri], xc.get("rev_disp")
        return (v / rv * 100.0) if (v is not None and rv not in (None, 0)) else None

    def _amt_actual(ri, col):
        return fmt_amount(to_base(col["vals"][ri], col["pe"], col["months"]))

    def _amt_extra(ri, xc):
        v = xc["disp_vals"][ri]
        return fmt_amount(v)

    rows, styles = [], {}

    if view_mode != "کامل":
        # خلاصه: درآمد عملیاتی (مقدار) + سود ناخالص (مقدار+٪) + سود خالص تداوم (مقدار+٪)
        seq = []
        if rev_idx is not None:
            seq.append((rev_idx, None))
        if gross_idx is not None:
            seq.append((gross_idx, HL_GROSS))
        if netc_idx is not None:
            seq.append((netc_idx, HL_NET))
        for ridx, hl in seq:
            row = {"شرح": master_labels[ridx]}
            for xc in extra_cols:                       # پیش‌بینی‌ها سمت راست (کنار شرح)
                row[xc["header"]] = _amt_extra(ridx, xc)
            for col in columns:
                row[col["header"]] = _amt_actual(ridx, col)
            rows.append(row)
            if hl:
                styles[len(rows) - 1] = hl
            if ridx != rev_idx:               # ردیف درصد فقط برای ناخالص و خالص
                rowp = {"شرح": ""}
                for xc in extra_cols:
                    rowp[xc["header"]] = fmt_pct(_pct_x(ridx, xc))
                for col in columns:
                    rowp[col["header"]] = fmt_pct(_pct(ridx, col))
                rows.append(rowp)
                if hl:
                    styles[len(rows) - 1] = hl
        if not rows:
            st.warning("ردیف‌های «درآمد عملیاتی / سود ناخالص / سود خالص عملیات در حال تداوم» پیدا نشد.")
            return
        render_table(pd.DataFrame(rows), row_styles=styles, center_values=True, col_styles=_col_styles)
        return

    # کامل
    for ri, lab in enumerate(master_labels):
        k = master_keys[ri]
        if _stmt_hidden_row(k):
            continue
        if hide_small and ri not in (rev_idx, gross_idx, netc_idx):
            ps = [abs(p) for p in (_pct(ri, c) for c in columns) if p is not None]
            if ps and max(ps) < 1.0:           # همهٔ مقادیر زیر ۱٪ درآمد → پنهان
                continue
        hl = HL_GROSS if ri == gross_idx else (HL_NET if ri == netc_idx else None)
        row = {"شرح": lab}
        for xc in extra_cols:
            row[xc["header"]] = _amt_extra(ri, xc)
        for col in columns:
            row[col["header"]] = _amt_actual(ri, col)
        rows.append(row)
        if hl:
            styles[len(rows) - 1] = hl
        if show_pct_rows:
            rowp = {"شرح": ""}
            for xc in extra_cols:
                rowp[xc["header"]] = fmt_pct(_pct_x(ri, xc))
            for col in columns:
                rowp[col["header"]] = fmt_pct(_pct(ri, col))
            rows.append(rowp)
            if hl:
                styles[len(rows) - 1] = hl
    if not rows:
        st.info("ردیفی برای نمایش نماند (شاید فیلتر «سطرهای کم‌اهمیت» را خاموش کنید).")
        return
    render_table(pd.DataFrame(rows), row_styles=styles, center_values=True, col_styles=_col_styles)


def _next_annual_period(last_year):
    """دورهٔ سالانهٔ بعدی: (سال+۱)."""
    return (int(last_year) + 1, None)


def _next_quarter_period(last_year, last_q):
    """فصلِ بعدی: q<4 → همان سال q+1 ، q==4 → سالِ بعد q1."""
    if last_q is None or last_q >= 4:
        return (int(last_year) + 1, 1)
    return (int(last_year), int(last_q) + 1)


def _fill_derived(vals):
    """اقلامِ مشتقِ خالی را از روی فرمول پر می‌کند (فقط وقتی کاربر خالی گذاشته):
       سود ناخالص   = درآمد عملیاتی − بهای تمام شده
       سود عملیاتی  = سود ناخالص − هزینه‌های فروش/اداری/عمومی − کاهش ارزش + سایر درآمد − سایر هزینه
       سود خالص     = سود عملیاتی − هزینه‌های مالی + سایر درآمد/هزینهٔ غیرعملیاتی − مالیات
    قرارداد علامت‌ها مثل صورت مالی است: هزینه‌ها منفی وارد می‌شوند.
    vals: dict {برچسب: مقدار}. خروجی: همان dict با اقلامِ خالیِ پرشده."""
    out = dict(vals)

    def num(v):
        return None if (v is None or (isinstance(v, float) and pd.isna(v)) or v == "") else float(v)

    def find(*needles, neg_ok=True):
        """مقدارِ اولین برچسبی که همهٔ کلیدواژه‌ها را دارد."""
        for lab, v in out.items():
            k = _inc_label_key(lab)
            if all(_inc_label_key(n) in k for n in needles):
                return num(v)
        return None

    def find_label(*needles):
        for lab in out:
            k = _inc_label_key(lab)
            if all(_inc_label_key(n) in k for n in needles):
                return lab
        return None

    def is_blank(lab):
        return lab is not None and num(out.get(lab)) is None

    rev = find("درآمد", "عملیاتی")
    if rev is None:
        rev = find("درآمدعملیاتی") or find("درآمد")
    cogs = find("بهای", "تمام", "شده")

    # سود ناخالص
    lbl_gross = find_label("ناخالص")
    if is_blank(lbl_gross) and rev is not None and cogs is not None:
        out[lbl_gross] = rev + cogs        # cogs منفی است

    gross = find("ناخالص")
    sga = find("فروش", "اداری") or find("اداری", "عمومی")
    impair = find("کاهش", "ارزش")
    oth_inc = find("سایر", "درآمد")
    oth_exp = find("سایر", "هزینه")

    # سود عملیاتی
    lbl_op = find_label("عملیاتی")
    # دقت: «درآمدهای عملیاتی» هم «عملیاتی» دارد؛ برچسبِ سود عملیاتی باید «سود» داشته باشد
    lbl_op = find_label("سود", "عملیاتی")
    if is_blank(lbl_op) and gross is not None:
        op = gross
        for x in (sga, impair, oth_inc, oth_exp):
            if x is not None:
                op += x
        out[lbl_op] = op

    op_profit = find("سود", "عملیاتی")
    fin_cost = find("هزینه", "مالی")
    noniop = find("غیرعملیاتی")
    # مالیات: ردیفی که «مالیات» دارد ولی «سود/قبل» ندارد (تا با «...قبل از مالیات» اشتباه نشود)
    tax = None
    for lab, v in out.items():
        kk = _inc_label_key(lab)
        if "مالیات" in kk and "سود" not in kk and "قبل" not in kk:
            tax = num(v)
            break

    # سود عملیات در حال تداوم قبل از مالیات = سود عملیاتی + هزینه مالی + سایر غیرعملیاتی
    lbl_pretax = None
    for lab in out:
        kk = _inc_label_key(lab)
        if ("قبلازمالیات" in kk or ("قبل" in kk and "مالیات" in kk)) and "خالص" not in kk:
            lbl_pretax = lab
            break
    if is_blank(lbl_pretax) and op_profit is not None:
        pretax = op_profit
        for x in (fin_cost, noniop):
            if x is not None:
                pretax += x
        out[lbl_pretax] = pretax

    pretax_val = None
    if lbl_pretax is not None:
        pretax_val = num(out.get(lbl_pretax))

    # سود خالص (مراقب: «خالص» زیرمجموعهٔ «ناخالص» است، پس ناخالص را رد کن)
    lbl_net = None
    for lab in out:
        kk = _inc_label_key(lab)
        if "خالص" in kk and "ناخالص" not in kk and "سود" in kk:
            lbl_net = lab
            break
    if is_blank(lbl_net):
        if pretax_val is not None:
            # از قبل‌از‌مالیات: فقط مالیات کم می‌شود
            net = pretax_val + (tax if tax is not None else 0.0)
            out[lbl_net] = net
        elif op_profit is not None:
            # اگر ردیفِ قبل‌از‌مالیات وجود نداشت، مستقیم از عملیاتی
            net = op_profit
            for x in (fin_cost, noniop, tax):
                if x is not None:
                    net += x
            out[lbl_net] = net

    return out


def render_statement_editable(symbol, prefix, master_labels, master_keys, columns,
                              rev_idx, gross_idx, netc_idx,
                              to_base, fmt_amount, fmt_pct, view_mode, show_pct_rows, hide_small,
                              latest_year, latest_q=None, quarter_labels=None):
    """صورت سود و زیانِ رنگی (مثل قبل) + ستون‌های پیش‌بینی.
      • حالت پیش‌فرض: جدولِ HTMLِ رنگی (واقعی‌ها + پیش‌بینی‌های ذخیره‌شده) با همان هایلایت/فونت.
      • دکمهٔ ➕ یک دورهٔ بعدی اضافه می‌کند، ➖ آخری را حذف.
      • دکمهٔ ✏️ حالتِ ویرایش (data_editor) را باز می‌کند؛ بعد از 💾 ذخیره، دوباره رنگی می‌شود.
    prefix: 'فصلی' یا 'سالانه'."""
    quarter_labels = quarter_labels or {1: "فصل اول", 2: "فصل دوم", 3: "فصل سوم", 4: "فصل چهارم"}
    is_q = (prefix == "فصلی")

    vis_idx = [ri for ri in range(len(master_labels)) if not _stmt_hidden_row(master_keys[ri])]
    vis_labels_plain = [master_labels[ri] for ri in vis_idx]

    # ── سشن: فهرست دوره‌های پیش‌بینی + حالت ویرایش ──
    sk = f"fc_cols_{symbol}_{prefix}"
    ek = f"fc_edit_{symbol}_{prefix}"
    if sk not in st.session_state:
        existing = list_forecast_symbols()
        found = []
        if not existing.empty:
            for _, r in existing.iterrows():
                if r["symbol"] == symbol and str(r["fiscal_year"]).startswith(prefix + "-"):
                    found.append(str(r["fiscal_year"]))
        st.session_state[sk] = sorted(found)
    if ek not in st.session_state:
        st.session_state[ek] = False

    def _fc_col_title(fkey):
        parts = fkey.split("-")
        yy = to_fa_digits(parts[1])
        if is_q and len(parts) > 2:
            q = int(parts[2][1:])
            return f"🔮 {quarter_labels.get(q, parts[2])} {yy}"
        return f"🔮 پیش‌بینی {yy}"

    # ── دکمه‌ها ──
    b1, b2, b3, _ = st.columns([1.4, 1, 1, 3])
    with b1:
        if st.button("➕ ستون پیش‌بینی دورهٔ بعد", key=f"fc_add_{symbol}_{prefix}"):
            cur = st.session_state[sk]
            if cur:
                _last = cur[-1].split("-")
                ly = int(_last[1]); lq = int(_last[2][1:]) if (is_q and len(_last) > 2) else None
            else:
                ly, lq = int(latest_year) if latest_year else 1404, latest_q
            if is_q:
                ny, nq = _next_quarter_period(ly, lq)
                key = normalize_digits(f"{prefix}-{ny}-Q{nq}")   # ارقام همیشه انگلیسی
            else:
                ny, _ = _next_annual_period(ly)
                key = normalize_digits(f"{prefix}-{ny}")
            if key not in cur:
                st.session_state[sk] = cur + [key]
            st.session_state[ek] = True       # افزودن ستون → مستقیم برو حالت ویرایش
            st.rerun()
    with b2:
        if st.button("➖ حذف آخرین", key=f"fc_pop_{symbol}_{prefix}") and st.session_state[sk]:
            _rm = st.session_state[sk][-1]
            save_forecast(symbol, _rm, [], "")
            st.session_state[sk] = st.session_state[sk][:-1]
            st.rerun()
    with b3:
        _edit_lbl = "✅ پایان ویرایش" if st.session_state[ek] else "✏️ ویرایش پیش‌بینی"
        if st.button(_edit_lbl, key=f"fc_edit_btn_{symbol}_{prefix}"):
            st.session_state[ek] = not st.session_state[ek]
            st.rerun()

    fc_keys = st.session_state[sk]

    # مقادیر ذخیره‌شدهٔ پیش‌بینی (واحدِ نمایش = میلیارد تومان)
    fc_titles, fc_saved, fc_data = [], {}, {}
    for fkey in reversed(fc_keys):           # جدیدتر اول (کنار شرح)
        title = _fc_col_title(fkey)
        fc_titles.append(title)
        sdf, _n = load_forecast(symbol, fkey)
        smap = dict(zip(sdf["label"], sdf["value"])) if (sdf is not None and not sdf.empty) else {}
        fc_data[title] = [smap.get(lab, None) for lab in vis_labels_plain]
        fc_saved[title] = fkey

    # ════════ حالت ویرایش: data_editor سادهٔ قابل‌تایپ ════════
    if st.session_state[ek] and fc_titles:
        def _fmt_num(v):
            if v is None or pd.isna(v): return "—"
            dec = 0 if abs(v) >= 100 else (1 if abs(v) >= 1 else 2)
            s = f"{abs(v):,.{dec}f}".translate(EN_TO_FA)
            return f"({s})" if v < 0 else s
        data = {"شرح": vis_labels_plain}
        for t in fc_titles:
            data[t] = fc_data[t]
        for col in columns:
            data[col["header"]] = [_fmt_num(to_base(col["vals"][ri], col["pe"], col["months"])) for ri in vis_idx]
        colcfg = {"شرح": st.column_config.TextColumn("شرح", width="large", disabled=True)}
        for t in fc_titles:
            colcfg[t] = st.column_config.NumberColumn(t, format="%.2f",
                                                      help="پیش‌بینیِ خودت (میلیارد تومان)")
        for col in columns:
            colcfg[col["header"]] = st.column_config.TextColumn(col["header"], disabled=True)
        edited = st.data_editor(
            pd.DataFrame(data), key=f"fc_editor_{symbol}_{prefix}_{len(fc_titles)}",
            use_container_width=True, hide_index=True, column_config=colcfg)
        cnote = st.text_area("📝 یادداشت پیش‌بینی:", key=f"fc_note_{symbol}_{prefix}", height=80)
        _auto = st.checkbox("محاسبهٔ خودکارِ اقلامِ خالی (سود ناخالص/عملیاتی/خالص)",
                            value=True, key=f"fc_auto_{symbol}_{prefix}")
        if st.button("💾 ذخیره و نمایش رنگی", key=f"fc_save_{symbol}_{prefix}", type="primary"):
            for t in fc_titles:
                vals = {lab: edited.iloc[i][t] for i, lab in enumerate(vis_labels_plain)}
                if _auto:
                    vals = _fill_derived(vals)
                out = [(lab, vals[lab]) for lab in vis_labels_plain]
                save_forecast(symbol, fc_saved[t], out, cnote)
            st.session_state[ek] = False
            st.success("ذخیره شد.")
            st.rerun()
        st.caption("اعداد پیش‌بینی را وارد کن، بعد «💾 ذخیره و نمایش رنگی» را بزن. "
                   "اقلامِ خالیِ سود ناخالص/عملیاتی/خالص خودکار حساب می‌شوند.")
        return

    # ════════ حالت نمایش: همان جدولِ رنگیِ HTML با ستون‌های پیش‌بینی ════════
    # extra_cols: مقادیرِ پیش‌بینی از قبل در واحدِ نمایش‌اند → disp_vals هم‌ترازِ master
    extra_cols = []
    for fkey in reversed(fc_keys):
        title = _fc_col_title(fkey)
        smap_full = dict(zip(vis_labels_plain, fc_data[title]))
        disp_vals = [smap_full.get(master_labels[ri]) for ri in range(len(master_labels))]
        rev_disp = disp_vals[rev_idx] if (rev_idx is not None) else None
        extra_cols.append({"header": title, "disp_vals": disp_vals, "rev_disp": rev_disp})

    render_statement(master_labels, master_keys, columns, rev_idx, gross_idx, netc_idx,
                     to_base, fmt_amount, fmt_pct, view_mode, show_pct_rows, hide_small,
                     extra_cols=extra_cols)
    if fc_titles:
        st.caption("ستون‌های 🔮 پیش‌بینی هم رنگی نمایش داده شده‌اند. برای تغییر، «✏️ ویرایش پیش‌بینی» را بزن.")
    else:
        st.caption("برای افزودن پیش‌بینی، دکمهٔ «➕ ستون پیش‌بینی دورهٔ بعد» را بزن.")


@st.cache_data(ttl=3600)
def load_usd_rates():
    """
    'دوره یک ماهه منتهی به 1403/06/31'
    خروجی: dict با کلید period_end (مثل '1403/06/31') و مقدار نرخ دلار به تومان
    """
    try:
        df = pd.read_excel(USD_RATE_PATH, header=None)
        rates = {}
        for _, row in df.iterrows():
            rate_rial = row[0]
            label     = str(row[1]) if len(row) > 1 else ""
            label_norm = normalize_digits(label)
            m = re.search(r'(\d{4})/(\d{2})/\d{2}', label_norm)
            if m and pd.notna(rate_rial):
                ym_key = f"{m.group(1)}/{m.group(2)}"    # فقط سال/ماه
                # نرخ به ریاله → تقسیم بر ۱۰ میشه تومان، تقسیم بر ۱۰۰۰۰ میشه هزار تومان
                # ما میلیارد تومان داریم، پس برای تبدیل:
                # میلیارد_تومان / (نرخ_تومان / ۱۰۰۰) = میلیون_دلار
                # نرخ_دلار_به_تومان = rate_rial / 10
                rates[ym_key] = float(rate_rial)  # نرخ دلار به تومان (مستقیم)
        return rates
    except FileNotFoundError:
        st.warning(f"فایل {USD_RATE_PATH} پیدا نشد. نمودار دلاری غیرفعاله.")
        return {}
    except Exception as e:
        st.warning(f"خطا در خواندن نرخ دلار: {e}")
        return {}


def btmn_to_musd(btmn_series, period_series, usd_rates):
    """
    تبدیل میلیارد تومان به میلیون دلار
    نرخ دلار به تومان در usd_rates است
    میلیارد_تومان × ۱۰۰۰_میلیون_تومان / نرخ_تومان_هر_دلار = دلار
    → تقسیم بر ۱_000_000 میشه میلیون دلار
    → یعنی: btmn × 1000 / rate_toman = میلیون_دلار
    """
    result = []
    for btmn, period in zip(btmn_series, period_series):
        rate = usd_rates.get("/".join(period.split("/")[:2]) if period else period)
        if rate and rate > 0 and pd.notna(btmn):
            result.append(btmn * 1000.0 / rate)
        else:
            result.append(None)
    return result

def _ym_key(period):
    """'1403/12/29' -> '1403/12': extract year/month key for usd_rates lookup"""
    try:
        return "/".join(period.split("/")[:2])
    except Exception:
        return period


def smart_unit(values, show_usd):
    """
    بر اساس بزرگ‌ترین مقدار، بهترین واحد رو انتخاب می‌کنه تا اعداد بین ۱ تا ۱۰۰۰ باشن.
    ورودی: لیست مقادیر (میلیارد تومان یا میلیون دلار)، خروجی: (scale, label)
    """
    vals = [v for v in values if v is not None and not pd.isna(v) and v > 0]
    if not vals:
        if show_usd:
            return 1.0, "م.دلار"
        return 1.0, "م.تومان"
    mx = max(vals)
    if show_usd:
        # ورودی: میلیون دلار
        if mx < 1:          return 1000.0,  "هزار دلار"    # ×1000 → هزار دلار
        elif mx < 1000:     return 1.0,     "م.دلار"
        else:               return 0.001,   "م.م.دلار"
    else:
        # ورودی: میلیارد تومان
        if mx < 1:          return 1000.0,  "م.تومان"      # ×1000 → میلیون تومان
        elif mx < 1000:     return 1.0,     "م.تومان"
        else:               return 0.001,   "تریلیون تومان"


def apply_smart_unit(series, show_usd):
    """یه series میلیارد‌تومان/میلیون‌دلار رو با واحد بهینه برمی‌گردونه."""
    vals = list(series)
    scale, label = smart_unit(vals, show_usd)
    scaled = [v * scale if v is not None and not pd.isna(v) else None for v in vals]
    return scaled, label


def smart_rate_unit(rate_values_in_unit, is_usd):
    """
    برای نرخ فروش: اعداد رو با واحد بهینه scale می‌کنه
    ورودی دلار: اگه < ۱ → سنت (×۱۰۰)، اگه < ۱۰۰۰ → دلار، اگه بزرگ‌تر → هزار دلار
    ورودی تومان (م.تومان): اگه < ۱ → هزار تومان، اگه < ۱۰۰۰ → م.تومان، بزرگ‌تر → م.م.تومان
    """
    vals = [v for v in rate_values_in_unit if v is not None and not pd.isna(v) and v > 0]
    if not vals:
        return 1.0, "دلار" if is_usd else "م.تومان"
    mx = max(vals)
    if is_usd:
        if mx < 0.01:    return 100_000.0, "سنت (×۱۰۰۰)"
        elif mx < 0.1:   return 1000.0,   "هزار‌سنت (×۱۰)"
        elif mx < 1:     return 100.0,    "سنت"
        elif mx < 1000:  return 1.0,      "دلار"
        else:            return 0.001,    "هزار دلار"
    else:
        # ورودی به میلیون تومانه (sales_rate_rial / 10_000_000)
        if mx < 0.001:   return 1_000_000.0, "تومان"
        elif mx < 1:     return 1000.0,      "هزار تومان"
        elif mx < 1000:  return 1.0,         "م.تومان"
        else:            return 0.001,       "م.م.تومان"




try:
    df_main = load_main()
    df_prod = load_products()
    usd_rates = load_usd_rates()
except Exception as e:
    st.error(f"خطا در خواندن دیتابیس: {e}")
    st.stop()

if df_main.empty:
    st.warning("دیتابیس خالیه. اول اسکریپت اصلی رو اجرا کن.")
    st.stop()

has_usd = len(usd_rates) > 0

# گزارش‌های دوره‌ای (اختیاری — اگر codal_income.db موجود باشد)
df_income_rep   = load_income_reports()
df_income_items = load_income_items()
has_income      = not df_income_rep.empty

# گزارش‌های سالانه (اختیاری — codal_annual.db + نرخ دلار سالانه)
df_annual_rep    = load_annual_reports()
df_annual_items  = load_annual_items()
has_annual       = not df_annual_rep.empty
annual_usd_rates = load_annual_usd_rates()
has_annual_usd   = len(annual_usd_rates) > 0

# ─── ادغامِ گزارش‌های سالانه با جایگزینِ ۱۲ماههٔ دوره‌ای ────────────────
# برای سال‌های ۱۴۰۲ به بعد، اگر «صورت سالانه» در codal_annual.db نبود، از
# گزارشِ ۱۲ماههٔ codal_income.db به‌عنوان جایگزین استفاده می‌شود
# (اول سالانه، اگر نبود دوره‌ایِ ۱۲ماهه).
INCOME_FALLBACK_FROM_YEAR = 1402

_ANNUAL_REP_COLS = ["id", "symbol", "period_end", "duration_months", "is_audited",
                    "is_consolidated", "report_type", "title", "letter_url", "sent_date"]


def annual_reports_with_income_fallback(sym):
    """گزارش‌های سالانهٔ یک نماد را برمی‌گرداند و خلأِ سال‌های ۱۴۰۲ به بعد را
    با گزارش‌های ۱۲ماههٔ codal_income.db پر می‌کند.

    خروجی: (reps_combined, items_combined, fallback_years)
        reps_combined  : DataFrame هم‌اسکیمای df_annual_rep (id ممکن است آفست‌شده باشد)
        items_combined : DataFrame هم‌اسکیمای df_annual_items (report_id منطبق با reps)
        fallback_years : set[int] سال‌هایی که داده‌شان از گزارشِ ۱۲ماهه آمده است.
    """
    # --- پایهٔ سالانه ---
    if has_annual:
        reps_a  = df_annual_rep[df_annual_rep["symbol"] == sym].copy()
        items_a = df_annual_items.copy()
    else:
        reps_a  = pd.DataFrame(columns=_ANNUAL_REP_COLS)
        items_a = pd.DataFrame(columns=["report_id", "row_order", "label", "value"])

    fallback_years = set()
    if not has_income:
        return reps_a, items_a, fallback_years

    # سال‌هایی که قبلاً از دیتابیس سالانه پوشش داده شده‌اند
    years_have = set()
    for _pe in reps_a["period_end"].tolist():
        _yy, _ = _inc_year_month(_pe)
        if _yy is not None:
            years_have.add(_yy)

    # گزارش‌های ۱۲ماههٔ این نماد (جدیدتر و حسابرسی‌شده در اولویت)
    inc12 = df_income_rep[(df_income_rep["symbol"] == sym) &
                          (df_income_rep["duration_months"] == 12)].copy()
    if inc12.empty:
        return reps_a, items_a, fallback_years
    inc12 = inc12.sort_values(["period_end", "is_audited", "is_consolidated"],
                              ascending=[False, False, False])

    # برای هر سالِ ۱۴۰۲+ که در سالانه نیست، یک گزارشِ ۱۲ماهه انتخاب کن
    pick_ids, seen_years = [], set()
    for _, r in inc12.iterrows():
        _yy, _ = _inc_year_month(r["period_end"])
        if _yy is None or _yy < INCOME_FALLBACK_FROM_YEAR:
            continue
        if _yy in years_have or _yy in seen_years:
            continue
        seen_years.add(_yy)
        pick_ids.append(r["id"])
        fallback_years.add(_yy)

    if not pick_ids:
        return reps_a, items_a, fallback_years

    # آفستِ امن تا id گزارش‌های دو دیتابیس با هم برخورد نکنند
    _max_a = int(df_annual_rep["id"].max()) if (has_annual and not df_annual_rep.empty) else 0
    _max_i = int(df_income_rep["id"].max()) if not df_income_rep.empty else 0
    offset = max(_max_a, _max_i) + 1

    inc_pick = inc12[inc12["id"].isin(pick_ids)].copy()
    inc_pick["id"] = inc_pick["id"] + offset
    inc_pick = inc_pick[_ANNUAL_REP_COLS]

    inc_items = df_income_items[df_income_items["report_id"].isin(pick_ids)].copy()
    inc_items["report_id"] = inc_items["report_id"] + offset

    reps_combined  = pd.concat([reps_a, inc_pick], ignore_index=True)
    items_combined = pd.concat([items_a, inc_items], ignore_index=True)
    return reps_combined, items_combined, fallback_years

# مجمع عادی سالیانه (اختیاری — codal_assembly.db) برای درصد توزیع سود
df_assembly      = load_assembly()


# ════════════════════════════════════════════════════════════════
# سود مجمعِ زیرمجموعه‌ها (مثلِ داشبورد هلدینگ‌ها) — منبع: codal_annual.db + codal_assembly.db
# ════════════════════════════════════════════════════════════════
DIVIDEND_FY_DEFAULT = "1404"
DIVINPUTS_PATH = os.path.join(DATA_DIR, "dividend_inputs.xlsx")
_DIV_COLS = ["هلدینگ", "سال مالی", "نماد", "درصد تقسیم سود", "سود نماد"]


def _div_num(v):
    try:
        s = str(v).replace(",", "").translate(
            str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "0123456789" * 2)).strip()
        return float(s) if s not in ("", "nan", "None") else None
    except Exception:
        return None


def _div_net_from_group(grp):
    """سود خالصِ نهایی (آخرین ردیفِ «خالص» که «ناخالص/هرسهم» نیست)."""
    val = None
    g = grp.sort_values("row_order") if "row_order" in grp.columns else grp
    for _, it in g.iterrows():
        k = re.sub(r"\s", "", normalize_symbol(it.get("label")))
        if "خالص" in k and "ناخالص" not in k and "هرسهم" not in k:
            v = _div_num(it.get("value"))
            if v is not None:
                val = v
    return val


def _div_profit_from_db(dbpath, fy):
    if not os.path.exists(dbpath):
        return {}
    try:
        conn = sqlite3.connect(dbpath)
        reps = pd.read_sql_query(
            "SELECT id, symbol, period_end, COALESCE(is_audited,0) aud, "
            "COALESCE(is_consolidated,0) cons, sent_date FROM reports "
            "WHERE duration_months=12 AND period_end LIKE ?", conn, params=(f"{fy}/%",))
        if reps.empty:
            conn.close()
            return {}
        reps["sent_date"] = reps["sent_date"].astype(str)
        reps = reps.sort_values(["cons", "aud", "sent_date"], ascending=[True, False, False])
        best = reps.drop_duplicates("symbol", keep="first")
        ids = best["id"].tolist()
        qm = ",".join("?" * len(ids))
        items = pd.read_sql_query(
            f"SELECT report_id, row_order, label, value FROM line_items WHERE report_id IN ({qm})",
            conn, params=ids)
        conn.close()
    except Exception:
        return {}
    id2sym = dict(zip(best["id"], best["symbol"]))
    out = {}
    for rid, grp in items.groupby("report_id"):
        v = _div_net_from_group(grp)
        if v is not None:
            out[normalize_symbol(id2sym.get(rid, ""))] = v
    return out


@st.cache_data(ttl=120)
def _div_load_profit(fy):
    """{نمادِ‌نرمال: سود خالصِ سالانهٔ ۱۲ماهه (میلیون ریال)} — اول codal_annual.db، سپس codal_income.db."""
    res = _div_profit_from_db(ANNUAL_DB_PATH, fy)
    for s, v in _div_profit_from_db(INCOME_DB_PATH, fy).items():
        res.setdefault(s, v)
    return res


@st.cache_data(ttl=120)
def _div_load_payout(fy):
    """{نمادِ‌نرمال: نسبتِ تقسیمِ سود (dps/eps، سقف ۱)} — همان سال، وگرنه میانگینِ سال‌های موجود."""
    if not os.path.exists(ASSEMBLY_DB_PATH):
        return {}
    try:
        conn = sqlite3.connect(ASSEMBLY_DB_PATH)
        df = pd.read_sql_query("SELECT symbol, period_end, dps, eps FROM assembly_decisions", conn)
        conn.close()
    except Exception:
        return {}
    fy_val, all_vals = {}, {}
    for _, r in df.iterrows():
        e, d = _div_num(r.get("eps")), _div_num(r.get("dps"))
        if not (e and e > 0 and d is not None):
            continue
        ratio = min(d / e, 1.0)
        sym = normalize_symbol(r.get("symbol"))
        all_vals.setdefault(sym, []).append(ratio)
        if str(r.get("period_end") or "").startswith(fy + "/"):
            fy_val[sym] = ratio
    return {k: fy_val.get(k, sum(v) / len(v)) for k, v in all_vals.items()}


def _div_load_inputs():
    if not os.path.exists(DIVINPUTS_PATH):
        return pd.DataFrame(columns=_DIV_COLS)
    try:
        df = pd.read_excel(DIVINPUTS_PATH, dtype=str)
        for c in _DIV_COLS:
            if c not in df.columns:
                df[c] = None
        return df
    except Exception:
        return pd.DataFrame(columns=_DIV_COLS)


def _div_inputs_map(holding, fy):
    df = _div_load_inputs()
    out = {}
    if df.empty:
        return out
    for _, r in df.iterrows():
        if (normalize_symbol(r.get("هلدینگ")) == normalize_symbol(holding)
                and str(r.get("سال مالی")).strip() == str(fy).strip()):
            out[normalize_symbol(r.get("نماد"))] = (_div_num(r.get("درصد تقسیم سود")),
                                                    _div_num(r.get("سود نماد")))
    return out


def _div_save_inputs(holding, fy, rows):
    df = _div_load_inputs()
    if not df.empty:
        keep = ~((df["هلدینگ"].map(normalize_symbol) == normalize_symbol(holding)) &
                 (df["سال مالی"].astype(str).str.strip() == str(fy).strip()))
        df = df[keep]
    new = []
    for sym, payout, profit in rows:
        sym = str(sym or "").strip()
        if not sym:
            continue
        if (payout is None or payout == "") and (profit is None or profit == ""):
            continue
        new.append({"هلدینگ": holding, "سال مالی": str(fy), "نماد": sym,
                    "درصد تقسیم سود": payout, "سود نماد": profit})
    out = pd.concat([df, pd.DataFrame(new, columns=_DIV_COLS)], ignore_index=True) if new else df
    try:
        os.makedirs(os.path.dirname(DIVINPUTS_PATH) or ".", exist_ok=True)
    except Exception:
        pass
    out.to_excel(DIVINPUTS_PATH, index=False)


# ════════════════════════════════════════════════════════════════
# سایدبار
# ════════════════════════════════════════════════════════════════
st.sidebar.title("📊 داشبورد کدال")
page = st.sidebar.radio(
    "صفحه",
    ["📈 روند یک نماد", "🔬 محصولات", "📑 گزارش‌های دوره‌ای", "🔍 فیلتر نمادها",
     "👥 سهامداران و زیرمجموعه‌ها", "🔮 پیش‌بینی سود"]
)

st.sidebar.markdown("---")

# انتخاب واحد نمایش (سراسری)
if has_usd:
    currency_mode = st.sidebar.radio(
        "واحد نمایش",
        ["تومان (میلیارد)", "دلار (میلیون)"],
        index=0,
    )
    show_usd = currency_mode == "دلار (میلیون)"
else:
    show_usd = False

st.sidebar.markdown("---")

# ─── انتخاب نماد (مشترک بین همه صفحات) ───────────────────────
_all_syms_sidebar = sorted(df_main["symbol"].unique().tolist())
# نماد فعال — یه selectbox با key="sym" در هر صفحه نشون داده میشه
if "sym" not in st.session_state:
    st.session_state["sym"] = _all_syms_sidebar[0]
# اگه نماد جدیدی از جدول فیلتر انتخاب شده
if "_pending_sym" in st.session_state and st.session_state["_pending_sym"] in _all_syms_sidebar:
    _sym_index = _all_syms_sidebar.index(st.session_state["_pending_sym"])
    del st.session_state["_pending_sym"]
else:
    _sym_index = _all_syms_sidebar.index(st.session_state["sym"]) if "sym" in st.session_state and st.session_state["sym"] in _all_syms_sidebar else 0

selected_global = st.sidebar.selectbox("نماد:", _all_syms_sidebar, index=_sym_index, key="sym")

# ─── ارزش بازار نماد فعال (از آخرین فایل روزانهٔ TSETMC) ──────
_daily_df, _daily_fname = load_latest_daily()
if (not _daily_df.empty and "نام" in _daily_df.columns
        and "ارزش بازار" in _daily_df.columns and usd_rates):
    def _norm_sym(s):
        return normalize_symbol(s)
    _m = _daily_df[_daily_df["نام"].apply(_norm_sym) == _norm_sym(selected_global)]
    _mv_rial = None
    if not _m.empty:
        try:
            _mv_rial = float(str(_m.iloc[0]["ارزش بازار"]).replace(",", "").strip())
        except Exception:
            _mv_rial = None
    if _mv_rial and _mv_rial > 0:
        _mv_btmn = _mv_rial / 1e10                       # ریال → میلیارد تومان
        _last_key = max(usd_rates)                       # آخرین ماهِ نرخ دلار
        _last_rate = usd_rates[_last_key]
        _mv_musd = (_mv_btmn * 1000.0 / _last_rate) if (_last_rate and _last_rate > 0) else None
        _btmn_s = f"{_mv_btmn:,.0f}".translate(EN_TO_FA)
        _musd_s = f"{_mv_musd:,.0f}".translate(EN_TO_FA) if _mv_musd else "—"
        _rate_s = f"{_last_rate:,.0f}".translate(EN_TO_FA)

        # ── P/E از همان فایلِ روزانهٔ TSETMC (آخرین روزی که داده داریم) ──
        # نامِ ستون‌ها را با تحملِ تفاوتِ فاصله/حروف پیدا می‌کنیم.
        def _find_col(cols, *cands):
            norm = {str(c).strip().lower().replace(" ", ""): c for c in cols}
            for cand in cands:
                key = cand.strip().lower().replace(" ", "")
                if key in norm:
                    return norm[key]
            return None

        _pe_html = ""
        _pe_col = _find_col(_daily_df.columns, "p/e", "pe", "p / e")
        if _pe_col is not None:
            _pe_val = parse_fa_number(_m.iloc[0][_pe_col])
            # P/E منفی یا صفر بی‌معناست (شرکتِ زیان‌ده) — به‌جای عدد، تیره نشان بده
            if _pe_val is not None and _pe_val > 0:
                _pe_s = f"{_pe_val:,.2f}".translate(EN_TO_FA)
                _peg_col = _find_col(_daily_df.columns, "p/e گروه", "pe گروه", "p/e_group")
                _peg_val = parse_fa_number(_m.iloc[0][_peg_col]) if _peg_col is not None else None
                _peg_txt = ""
                if _peg_val is not None and _peg_val > 0:
                    _peg_s = f"{_peg_val:,.2f}".translate(EN_TO_FA)
                    _peg_txt = f'<span style="color:#64748b;font-size:11px"> · گروه {_peg_s}</span>'
                _pe_html = (
                    f'<div style="font-size:13px;font-weight:700;color:#fbbf24;margin-top:6px">'
                    f'P/E: {_pe_s}{_peg_txt}</div>'
                )

        st.sidebar.markdown(
            '<div style="font-family:Vazirmatn,Tahoma,sans-serif;direction:rtl;'
            'background:#1e293b;border:1px solid #334155;border-radius:10px;'
            'padding:10px 12px;margin-top:8px">'
            f'<div style="font-size:12px;color:#94a3b8;margin-bottom:4px">ارزش بازار {selected_global}</div>'
            f'<div style="font-size:16px;font-weight:700;color:#e2e8f0">{_btmn_s} میلیارد تومان</div>'
            f'<div style="font-size:14px;font-weight:700;color:#4DA3FF;margin-top:2px">{_musd_s} میلیون دلار</div>'
            f'{_pe_html}'
            f'<div style="font-size:10px;color:#64748b;margin-top:5px">نرخ دلار {_last_key}: {_rate_s} تومان</div>'
            '</div>', unsafe_allow_html=True,
        )

# ─── لینک‌های نماد: TSETMC و کدال (باز شدن در تب جدید) ──────
def _norm_sym2(s):
    return normalize_symbol(s)


def _clean_inscode(v):
    """insCode معتبر = رشته‌ای فقط از رقم و با طولِ کافی. اگر اعشاری/علمی/خراب بود None."""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    s = s.split(".")[0]                 # حذفِ بخش اعشاری اگر مثل 5.6e16 خوانده شده بود نتیجه خراب است
    # حالتِ علمی (e/E) یا حضورِ غیر رقم = نامعتبر (دقتِ float از بین رفته)
    if not s.isdigit() or len(s) < 10:
        return None
    return s


_inscode = None
# منبع اولِ مرجع: stck.xlsx (id سالم و رشته‌ای)
_icmap = load_inscode_map()
if _icmap:
    _inscode = _icmap.get(_norm_sym2(selected_global))
# پشتیبان ۱: فایل روزانه
if not _inscode and not _daily_df.empty and "key" in _daily_df.columns and "نام" in _daily_df.columns:
    _mm = _daily_df[_daily_df["نام"].apply(_norm_sym2) == _norm_sym2(selected_global)]
    if not _mm.empty:
        _inscode = _clean_inscode(_mm.iloc[0]["key"])
# پشتیبان ۲: فایل سهامداران
if not _inscode:
    try:
        _sh_for_link, _ = load_latest_shareholders()
        if (not _sh_for_link.empty and "key" in _sh_for_link.columns
                and "نماد" in _sh_for_link.columns):
            _ml = _sh_for_link[_sh_for_link["نماد"].apply(_norm_sym2) == _norm_sym2(selected_global)]
            if not _ml.empty:
                _inscode = _clean_inscode(_ml.iloc[0]["key"])
    except Exception:
        pass

# لینکِ TSETMC (نیازمند insCode) + لینکِ کدال (با خود نماد)
_codal_url = f"https://www.codal.ir/ReportList.aspx?search&Symbol={selected_global}"
_link_css = ('display:block;text-align:center;font-family:Vazirmatn,Tahoma,sans-serif;'
             'direction:rtl;text-decoration:none;border-radius:10px;padding:8px 12px;'
             'margin-top:8px;font-size:13px;font-weight:600;')
if _inscode:
    _ts_url = f"https://www.tsetmc.com/instInfo/{_inscode}"
    st.sidebar.markdown(
        f'<a href="{_ts_url}" target="_blank" rel="noopener noreferrer" '
        f'style="{_link_css}background:#0e2a47;border:1px solid #1d4ed8;color:#93c5fd">'
        f'🔗 TSETMC</a>',
        unsafe_allow_html=True,
    )
st.sidebar.markdown(
    f'<a href="{_codal_url}" target="_blank" rel="noopener noreferrer" '
    f'style="{_link_css}background:#1c2e1c;border:1px solid #15803d;color:#86efac">'
    f'📄 کدال</a>',
    unsafe_allow_html=True,
)

st.sidebar.markdown("---")

def get_value_col_label(show_usd):
    if show_usd:
        return "میلیون دلار", COLOR_DOLLAR
    return "میلیارد تومان", COLOR_TOMAN


def add_usd_col(df, btmn_col="total_billion_toman"):
    """ستون دلاری رو به دیتافریم اضافه می‌کنه"""
    if has_usd:
        df = df.copy()
        df["total_musd"] = btmn_to_musd(df[btmn_col], df["period_end"], usd_rates)
    return df


def scale_for_plot(values, show_usd):
    """مقادیر رو با واحد بهینه برای نمودار scale می‌کنه، (scaled_list, label) برمی‌گردونه"""
    return apply_smart_unit(pd.Series(values), show_usd)


# ════════════════════════════════════════════════════════════════
# صفحه ۱: روند یک نماد
# ════════════════════════════════════════════════════════════════
if page == "📈 روند یک نماد":
    st.title("روند یک نماد")

    selected = selected_global

    sub_all = df_main[df_main["symbol"] == selected].sort_values("period_end").copy()
    sub_all = add_usd_col(sub_all)

    # فیلتر: تنها ردیف‌های «جمع» یا «جمع درآمدهای عملیاتی» را نگه‌دار
    # اگر یک period چند ردیف داشت (مثلاً هر دوی «جمع» و «جمع درآمدهای عملیاتی»)،
    # ترجیح به «جمع درآمدهای عملیاتی» است؛ اگر نبود، «جمع» را برمی‌گیریم.
    if not sub_all.empty and "report_type" in sub_all.columns:
        # برای هر دوره یک ردیف نگه‌دار: اولویت «درآمدهای عملیاتی» > «جمع» > آخرین ردیف.
        # روشِ بُرداری بدونِ groupby.apply تا ستونِ گروه‌بندی (period_end) حفظ شود
        # (پانداس ۳ ستونِ گروه‌بندی را داخلِ apply حذف می‌کند و باعث KeyError می‌شد).
        _norm = sub_all["report_type"].fillna("").str.strip().str.lower()
        _prio = pd.Series(0, index=sub_all.index)
        _prio[_norm.str.contains(r"جمع", na=False)] = 1
        _prio[_norm.str.contains(r"درآمدهای.*عملیاتی|عملیاتی.*درآمدهای", regex=True, na=False)] = 2
        sub_all = sub_all.assign(_prio=_prio, _ord=range(len(sub_all)))
        sub_all = (sub_all.sort_values(["period_end", "_prio", "_ord"])
                          .groupby("period_end", as_index=False, group_keys=False)
                          .tail(1)
                          .drop(columns=["_prio", "_ord"], errors="ignore")
                          .sort_values("period_end")
                          .reset_index(drop=True))

    def extract_year_month(period):
        try:
            y, m, _ = period.split("/")
            return int(y), int(m)
        except Exception:
            return None, None

    _assign_year_month(sub_all, extract_year_month)
    sub_all["ماه_label"] = sub_all["ماه_num"].apply(
        lambda m: PERSIAN_MONTHS.get(f"{int(m):02d}", str(m)) if pd.notna(m) else ""
    )
    sub_all["دوره"] = sub_all["period_end"].apply(format_period)

    y_col, y_label, bar_color_single = (
        ("total_musd", "میلیون دلار", COLOR_DOLLAR) if show_usd
        else ("total_billion_toman", "میلیارد تومان", COLOR_TOMAN)
    )
    decimals = 2 if show_usd else 1

    # ─── میانگین ۱۲ ماه گذشته برای فیلتر ──────────────────────
    periods_sorted_desc = list(sub_all["period_end"].sort_values(ascending=False))
    last_12_periods = periods_sorted_desc[1:13]  # بدون آخرین ماه
    avg_12 = sub_all[sub_all["period_end"].isin(last_12_periods)][y_col].mean() if last_12_periods else 0

    # ─── متریک‌ها ──────────────────────────────────────────────
    st.markdown("---")

    tabs = st.tabs(["مبلغ فروش"])
    YEAR_COLORS = ["#A78BFA", "#4DA3FF", "#10B981", "#F59E0B", "#EF4444", "#EC4899"]

    # ════ تب مبلغ فروش ════════════════════════════════════════
    with tabs[0]:
        ctrl_r1c1, ctrl_r1c2 = st.columns([3, 1])
        with ctrl_r1c1:
            chart_mode = st.radio(
                "نوع نمایش", ["مقایسه‌ای", "تجمعی", "فصلی", "روند"],
                horizontal=True, key="chart_mode_tab0",
            )

        ccol1, ccol2, ccol3, ccol4 = st.columns(4)
        with ccol1:
            show_amounts = st.checkbox("نمایش مبالغ", value=True, key="cb_amounts")
        with ccol2:
            show_growth  = st.checkbox("درصد رشد",    value=False, key="cb_growth")
        with ccol3:
            show_avg     = st.checkbox("میانگین هر سال", value=False, key="cb_avg")
        with ccol4:
            show_cum     = st.checkbox("از ابتدای سال مالی", value=False, key="cb_cum")

        sub_filtered = sub_all.copy()

        # محاسبه scale بهینه برای نمودار
        _raw_vals = sub_filtered[y_col].dropna().tolist()
        _scale, _scaled_lbl = smart_unit(_raw_vals, show_usd and has_usd)
        sub_filtered = sub_filtered.copy()
        sub_filtered["y_scaled"] = sub_filtered[y_col] * _scale
        y_plot = "y_scaled"
        y_plot_label = _scaled_lbl

        years_available = sorted(sub_filtered["سال"].dropna().unique().astype(int), reverse=True)
        fig = go.Figure()

        if chart_mode == "مقایسه‌ای":
            month_order = list(PERSIAN_MONTHS.values())
            for i, yr in enumerate(sorted(years_available)):
                yr_data = sub_filtered[sub_filtered["سال"] == yr].copy()
                if show_cum:
                    yr_data["y_scaled"] = yr_data["y_scaled"].cumsum()
                fig.add_trace(go.Bar(
                    x=yr_data["ماه_label"], y=yr_data["y_scaled"],
                    name=to_fa_digits(str(yr)),
                    marker_color=YEAR_COLORS[i % len(YEAR_COLORS)],
                    hovertemplate=f"<b>{to_fa_digits(str(yr))}</b> | %{{x}}<br>%{{y:,.2f}} {y_plot_label}<extra></extra>",
                ))
            fig.update_layout(barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=month_order, tickfont=PLOTLY_FONT))

        elif chart_mode == "تجمعی":
            month_order = list(PERSIAN_MONTHS.values())
            for i, yr in enumerate(sorted(years_available)):
                yr_data = sub_filtered[sub_filtered["سال"] == yr].copy()
                yr_data["cumval"] = yr_data["y_scaled"].cumsum()
                fig.add_trace(go.Bar(
                    x=yr_data["ماه_label"], y=yr_data["cumval"],
                    name=to_fa_digits(str(yr)),
                    marker_color=YEAR_COLORS[i % len(YEAR_COLORS)],
                    hovertemplate=f"<b>{to_fa_digits(str(yr))}</b> | %{{x}}<br>%{{y:,.2f}} {y_plot_label} (تجمعی)<extra></extra>",
                ))
            fig.update_layout(barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=month_order, tickfont=PLOTLY_FONT))

        elif chart_mode == "فصلی":
            SEASON_LABELS = ["بهار", "تابستان", "پاییز", "زمستان"]
            sf = sub_filtered.copy()
            sf["فصل_idx"] = sf["ماه_num"].apply(
                lambda m: (int(m) - 1) // 3 if pd.notna(m) else None)
            for i, yr in enumerate(sorted(years_available)):
                yd = sf[sf["سال"] == yr]
                sums = yd.groupby("فصل_idx")["y_scaled"].sum()
                ys = [sums.get(s, 0) for s in range(4)]
                fig.add_trace(go.Bar(
                    x=SEASON_LABELS, y=ys, name=to_fa_digits(str(yr)),
                    marker_color=YEAR_COLORS[i % len(YEAR_COLORS)],
                    hovertemplate=f"<b>{to_fa_digits(str(yr))}</b> | %{{x}}<br>%{{y:,.2f}} {y_plot_label} (جمع فصل)<extra></extra>",
                ))
            fig.update_layout(barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=SEASON_LABELS, tickfont=PLOTLY_FONT))

        else:  # روند
            sub_trend = sub_filtered.copy()
            plot_col = "y_scaled"
            if show_cum:
                sub_trend["cum_by_year"] = sub_trend.groupby("سال")["y_scaled"].cumsum()
                plot_col = "cum_by_year"
            fig.add_trace(go.Bar(
                x=sub_trend["دوره"], y=sub_trend[plot_col],
                marker_color=bar_color_single, name=y_plot_label,
                hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.2f}} {y_plot_label}<extra></extra>",
            ))
            if show_avg:
                avg_line = sub_trend.groupby("سال")[plot_col].transform("mean")
                fig.add_trace(go.Scatter(
                    x=sub_trend["دوره"], y=avg_line, mode="lines",
                    name="میانگین سالانه",
                    line=dict(color="#F59E0B", width=2, dash="dash"),
                    hovertemplate=f"<b>%{{x}}</b><br>میانگین: %{{y:,.2f}} {y_plot_label}<extra></extra>",
                ))
            if show_growth:
                pct = sub_trend[plot_col].pct_change() * 100
                fig.add_trace(go.Scatter(
                    x=sub_trend["دوره"], y=pct, mode="lines+markers",
                    name="درصد رشد", line=dict(color="#D946EF", width=2), yaxis="y2",
                    hovertemplate="<b>%{x}</b><br>رشد: %{y:+.1f}%<extra></extra>",
                ))
                fig.update_layout(yaxis2=dict(
                    title=dict(text="درصد رشد", font=PLOTLY_FONT),
                    tickfont=PLOTLY_FONT, overlaying="y", side="right", showgrid=False,
                ))
            fig.update_layout(xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT))

        _avg_lines = []

        if show_avg and chart_mode in ("مقایسه‌ای", "تجمعی"):
            for i, yr in enumerate(sorted(years_available)):
                yr_data = sub_filtered[sub_filtered["سال"] == yr].copy()
                val_col_use = "y_scaled"
                if chart_mode == "تجمعی":
                    yr_data["cumval"] = yr_data["y_scaled"].cumsum()
                    val_col_use = "cumval"
                avg_val = yr_data[val_col_use].mean()
                clr = YEAR_COLORS[i % len(YEAR_COLORS)]
                fig.add_hline(y=avg_val, line_dash="dash", line_color=clr, opacity=0.6)
                _avg_lines.append((f"میانگین {to_fa_digits(str(yr))}: {fmt_fa(avg_val, 2)}", clr))

        if avg_12 and avg_12 > 0 and chart_mode != "فصلی":
            _avg12_scaled = avg_12 * _scale
            fig.add_hline(y=_avg12_scaled, line_dash="dot", line_color="#94A3B8", opacity=0.8)
            _avg_lines.append((f"م.۱۲ماه: {fmt_fa(_avg12_scaled, 2)}", "#94A3B8"))

        # نمایش در گوشه پایین راست — هر خط یه annotation جداگانه با فاصله ثابت
        if _avg_lines:
            LINE_H = 0.055  # فاصله عمودی بین خطوط (paper units)
            _x0 = 1.01
            _y_start = 0.0 + LINE_H * (len(_avg_lines) - 1)
            for _j, (_txt, _clr) in enumerate(_avg_lines):
                fig.add_annotation(
                    x=_x0, y=_y_start - _j * LINE_H,
                    xref="paper", yref="paper",
                    text=_txt,
                    showarrow=False, xanchor="left", yanchor="middle",
                    font=dict(family="Vazirmatn, Tahoma, sans-serif", size=11, color=_clr),
                )

        _pfx = ("تجمعی " if (chart_mode == "تجمعی" or show_cum)
                else ("فصلی " if chart_mode == "فصلی" else ""))
        unit_title = f"{_pfx}مبلغ فروش ({y_label})"
        fig.update_layout(
            title=dict(text=f"{selected} — {unit_title}", font=PLOTLY_FONT_TITLE),
            yaxis=dict(title=dict(text=y_plot_label if "y_plot_label" in dir() else y_label, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
            font=PLOTLY_FONT, template="plotly_dark", height=480,
            hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
            bargap=0.15, bargroupgap=0.05,
            margin=dict(r=160),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ─── نمودار درصد تقسیم سود (DPS ÷ EPS) ────────────────────
    _adf = df_assembly[df_assembly["symbol"] == selected].copy() if not df_assembly.empty else pd.DataFrame()
    if _adf.empty:
        st.markdown("#### نمودار درصد تقسیم سود")
        st.caption("برای این نماد اطلاعات مجمع عادی سالیانه (سود نقدی) موجود نیست.")
    else:
        def _payout(r):
            e, d = r.get("eps"), r.get("dps")
            if e is None or pd.isna(e) or e <= 0 or d is None or pd.isna(d):
                return None
            return d / e * 100.0

        _adf["payout"] = _adf.apply(_payout, axis=1)
        # سال مالی = بخشِ سالِ period_end ؛ صعودی (قدیمی‌ها سمت چپ)
        _adf["fy"] = _adf["period_end"].apply(
            lambda p: int(normalize_digits(p).split("/")[0]) if p else None)
        _pay = _adf.dropna(subset=["payout", "fy"]).sort_values("fy")

        if _pay.empty:
            st.markdown("#### نمودار درصد تقسیم سود")
            st.caption("داده‌ی معتبری برای محاسبهٔ درصد تقسیم سود نبود.")
        else:
            xs   = [to_fa_digits(str(int(y))) for y in _pay["fy"]]
            ys   = _pay["payout"].tolist()
            # تاریخ مجمع هر سال (هم‌ترتیب با میله‌ها) برای نمایش در tooltip
            _asm_dates = []
            for _, _rr in _pay.iterrows():
                _ad = _rr.get("assembly_date")
                _ad = normalize_digits(str(_ad)).strip() if (_ad is not None and not pd.isna(_ad)) else ""
                _asm_dates.append(to_fa_digits(_ad) if (_ad and _ad.lower() not in ("nan", "none", "")) else "—")
            n    = len(ys)
            BLUE, PURPLE = "#7C9CF5", "#B57BEC"   # آبیِ معمول، بنفشِ سال آخر
            colors = [BLUE] * (n - 1) + [PURPLE]   # آخرین (جدیدترین) سال بنفش

            fig_pay = go.Figure(go.Bar(
                x=xs, y=ys, marker_color=colors,
                width=0.62,
                customdata=_asm_dates,
                hovertemplate="<b>سال مالی %{x}</b><br>درصد تقسیم سود: %{y:.0f}٪<br>تاریخ مجمع: %{customdata}<extra></extra>",
            ))
            # برچسبِ «٪ XX» داخلِ حبابِ تیره، وسطِ هر میله
            for xi, yv in zip(xs, ys):
                fig_pay.add_annotation(
                    x=xi, y=max(yv / 2, 6),
                    text=f"٪ {to_fa_digits(f'{yv:.0f}')}",
                    showarrow=False,
                    font=dict(family="Vazirmatn, Tahoma, sans-serif", size=14, color="#ffffff"),
                    bgcolor="rgba(15,18,28,0.92)", bordercolor="#0b0e16",
                    borderwidth=1, borderpad=5,
                )
            # خطِ میانگینِ تقسیم سود
            _avg_pay = sum(ys) / len(ys)
            fig_pay.add_hline(
                y=_avg_pay, line_dash="dash", line_color="#F59E0B", line_width=2, opacity=0.9,
                annotation_text=f"میانگین: ٪ {to_fa_digits(f'{_avg_pay:.0f}')}",
                annotation_position="top left",
                annotation_font=dict(family="Vazirmatn, Tahoma, sans-serif", size=12, color="#F59E0B"),
            )
            _ymax = max(100, (max(ys) // 10 + 1) * 10)
            fig_pay.update_layout(
                title=dict(text="نمودار درصد تقسیم سود", font=PLOTLY_FONT_TITLE, x=0.5, xanchor="center"),
                template="plotly_dark", font=PLOTLY_FONT, height=430,
                margin=dict(t=70, r=20, l=20, b=40), bargap=0.35,
                xaxis=dict(tickfont=PLOTLY_FONT, showgrid=False),
                yaxis=dict(range=[0, _ymax], dtick=10, tickfont=PLOTLY_FONT,
                           ticksuffix="", gridcolor="#243049"),
                hoverlabel=dict(font=PLOTLY_FONT), showlegend=False,
            )
            st.plotly_chart(fig_pay, use_container_width=True)


    # ─── جدول کامل ─────────────────────────────────────────────
    with st.expander("نمایش جدول کامل"):
        display = sub_all[["دوره", "period_end", "total_billion_toman",
                        "domestic_sales_btmn", "export_sales_btmn",
                        "sales_return_btmn", "discounts_btmn"]].copy()
        if has_usd:
            display["total_musd"]    = btmn_to_musd(display["total_billion_toman"], display["period_end"], usd_rates)
            display["domestic_musd"] = btmn_to_musd(display["domestic_sales_btmn"], display["period_end"], usd_rates)
            display["export_musd"]   = btmn_to_musd(display["export_sales_btmn"],   display["period_end"], usd_rates)
        display = display.drop(columns=["period_end"])
        for col in ["total_billion_toman", "domestic_sales_btmn",
                    "export_sales_btmn", "sales_return_btmn", "discounts_btmn"]:
            display[col] = display[col].apply(lambda x: fmt_fa(x, 1))
        rename_map = {
            "total_billion_toman": "جمع کل (م.تومان)",
            "domestic_sales_btmn": "فروش داخلی (م.تومان)",
            "export_sales_btmn":   "فروش صادراتی (م.تومان)",
            "sales_return_btmn":   "برگشت فروش",
            "discounts_btmn":      "تخفیفات",
        }
        if has_usd:
            for col in ["total_musd", "domestic_musd", "export_musd"]:
                display[col] = display[col].apply(lambda x: fmt_fa(x, 2) if x is not None else "—")
            rename_map.update({
                "total_musd":    "جمع کل (م.دلار)",
                "domestic_musd": "فروش داخلی (م.دلار)",
                "export_musd":   "فروش صادراتی (م.دلار)",
            })
        display = display.rename(columns=rename_map)
        render_table(display)







# ════════════════════════════════════════════════════════════════
# صفحه ۴: تحلیل محصول
# ════════════════════════════════════════════════════════════════
elif page == "🔬 محصولات":
    st.title("محصولات")

    if df_prod.empty:
        st.warning("هیچ دیتای محصولی در دیتابیس نیست.")
        st.stop()

    DONUT_COLORS = [
        "#4DA3FF","#F59E0B","#10B981","#D946EF",
        "#EF4444","#8B5CF6","#14B8A6","#F97316",
        "#EC4899","#6366F1","#84CC16","#06B6D4",
    ]

    # ─── نماد — یک بار برای همه ──────────────────────────────
    prod_symbols = sorted(df_prod["symbol"].unique())
    selected_sym = selected_global if selected_global in prod_symbols else prod_symbols[0]

    sym_products = df_prod[df_prod["symbol"] == selected_sym].copy()
    sym_main_pa  = df_main[df_main["symbol"] == selected_sym].copy()
    all_periods_prod = sorted(sym_products["period_end"].unique(), reverse=True)
    last_12_prod     = all_periods_prod[1:13]

    # ─── ادغام ───────────────────────────────────────────────
    merge_cats = st.toggle("ادغام داخلی + صادراتی", value=False, key="pa_merge")

    # ─── slider فیلتر مشترک (برای هر دو بخش ترکیب و کارت‌ها) ──
    st.markdown('''<style>
    div[data-testid="stSlider"] { direction: ltr !important; }
    div[data-testid="stSlider"] > label { direction: rtl !important; display: block; }
    </style>''', unsafe_allow_html=True)
    _shared_pct = st.slider("حداقل سهم % از کل فروش", 0, 30, 3, key="shared_min_pct")
    if merge_cats:
        def weighted_avg_rate(group):
            total_qty = group["sales_qty"].sum()
            if total_qty > 0:
                return (group["sales_rate_rial"] * group["sales_qty"]).sum() / total_qty
            valid = group["sales_rate_rial"][group["sales_rate_rial"] > 0]
            return valid.mean() if not valid.empty else 0
        sym_agg = sym_products.groupby(["period_end","product_name"]).apply(
            lambda g: pd.Series({
                "sales_amount_btmn": g["sales_amount_btmn"].sum(),
                "sales_qty":         g["sales_qty"].sum(),
                "production_qty":    g["production_qty"].sum(),
                "sales_rate_rial":   weighted_avg_rate(g),
                "unit":              g["unit"].iloc[0],
            })
        ).reset_index()
        sym_agg["category"] = "ترکیبی"
        sym_for_cards = sym_agg
    else:
        sym_for_cards = sym_products.copy()
        sym_for_cards["product_name"] = sym_for_cards["product_name"] + " — " + sym_for_cards["category"]

    # sym_prod با label برای بخش ترکیب
    sym_prod = sym_products.copy()
    sym_prod["label"] = sym_prod["category"] + " - " + sym_prod["product_name"]

    use_usd_mix = show_usd and has_usd
    unit_lbl_mix = "م.دلار" if use_usd_mix else "م.تومان"
    dec_mix = 2 if use_usd_mix else 1

    def convert_amt_mix(btmn, period):
        if use_usd_mix and has_usd:
            rate = usd_rates.get(_ym_key(period))
            if rate and rate > 0:
                return btmn * 1000.0 / rate
        return btmn

    # ════════════════════════════════════════════════════════
    # بخش ۱: ترکیب محصولات (stacked bar + دونات + جدول)
    # ════════════════════════════════════════════════════════
    all_periods_mix = sorted(sym_prod["period_end"].unique(), reverse=True)
    if not all_periods_mix:
        st.warning("داده‌ای یافت نشد.")
        st.stop()

    # slider فیلتر
    last_12_mix  = all_periods_mix[0:12]
    label_sum_12 = (sym_prod[sym_prod["period_end"].isin(last_12_mix)]
                    .groupby("label")["sales_amount_btmn"].sum())
    total_12_mix = label_sum_12.sum()
    label_pct_12 = (label_sum_12 / total_12_mix * 100) if total_12_mix > 0 else label_sum_12 * 0
    _max_mix = max(1, min(int(label_pct_12.max()) if not label_pct_12.empty else 50, 50))

    mix_min_pct = st.session_state.get("shared_min_pct", 3)

    top_labels = [l for l in label_pct_12.sort_values(ascending=False).index
                  if label_pct_12.get(l, 0) >= mix_min_pct]
    if not top_labels:
        top_labels = label_pct_12.nlargest(8).index.tolist()

    label_color_mix = {lbl: DONUT_COLORS[i % len(DONUT_COLORS)] for i, lbl in enumerate(top_labels)}

    # کنترل‌های ترکیب
    ctrl1, ctrl2, ctrl3 = st.columns([2, 2, 2])
    with ctrl1:
        sel_period_mix = st.selectbox("دوره:", all_periods_mix, format_func=format_period, key="mix_period")
    with ctrl2:
        compare_mode_mix = st.selectbox(
            "مقایسه با:",
            ["میانگین ۱۲ ماه گذشته", "ماه مشابه سال گذشته", "ماه قبل", "بدون مقایسه"],
            key="mix_compare",
        )
    with ctrl3:
        pie_mode = st.radio("دونات:", ["ماه آخر", "میانگین ۱۲ ماه"], horizontal=True, key="mix_pie_mode")

    # توابع کمکی ترکیب
    def find_single_compare_mix(period, mode, available):
        try:
            y, m, _ = period.split("/"); y, m = int(y), int(m)
        except: return None
        if mode == "ماه مشابه سال گذشته":
            target = f"{y-1}/{m:02d}"
        elif mode == "ماه قبل":
            pm, py = (m-1, y) if m > 1 else (12, y-1)
            target = f"{py}/{pm:02d}"
        else: return None
        return next((p for p in available if p.startswith(target)), None)

    def build_mix_period_data(period):
        sub = sym_prod[sym_prod["period_end"] == period].copy()
        if sub.empty: return pd.DataFrame()
        rows = [{"label": r["label"], "amount_btmn": r["sales_amount_btmn"] or 0}
                for _, r in sub.iterrows()]
        df_o = pd.DataFrame(rows)
        df_o["amount"] = df_o.apply(lambda r: convert_amt_mix(r["amount_btmn"], period), axis=1)
        total_pos = df_o[df_o["amount"] > 0]["amount"].sum()
        df_o["pct"] = df_o["amount"].apply(lambda x: x / total_pos * 100 if total_pos > 0 else 0)
        return df_o[["label","amount","pct"]]

    def build_mix_avg12(period, available):
        idx = next((i for i, p in enumerate(available) if p == period), None)
        if idx is None: return pd.DataFrame()
        past_12 = available[idx+1:idx+13]
        frames = []
        for p in past_12:
            d = build_mix_period_data(p)
            if not d.empty: frames.append(d.set_index("label")["amount"])
        if not frames: return pd.DataFrame()
        combined = pd.concat(frames, axis=1).fillna(0)
        avg = combined.mean(axis=1).reset_index(); avg.columns = ["label","amount"]
        total_pos = avg[avg["amount"] > 0]["amount"].sum()
        avg["pct"] = avg["amount"].apply(lambda x: x / total_pos * 100 if total_pos > 0 else 0)
        return avg

    df_cur_mix = build_mix_period_data(sel_period_mix)
    if compare_mode_mix == "میانگین ۱۲ ماه گذشته":
        df_comp_mix  = build_mix_avg12(sel_period_mix, all_periods_mix)
        comp_lbl_mix = "میانگین ۱۲ ماه گذشته"
    elif compare_mode_mix == "بدون مقایسه":
        df_comp_mix  = pd.DataFrame()
        comp_lbl_mix = ""
    else:
        _comp_p      = find_single_compare_mix(sel_period_mix, compare_mode_mix, all_periods_mix)
        df_comp_mix  = build_mix_period_data(_comp_p) if _comp_p else pd.DataFrame()
        comp_lbl_mix = format_period(_comp_p) if _comp_p else ""

    has_comp_mix = not df_comp_mix.empty and comp_lbl_mix

    is_ret = df_cur_mix["label"].str.contains("برگشت از فروش", na=False) if not df_cur_mix.empty else pd.Series(dtype=bool)
    mask_cur = (df_cur_mix["pct"] >= mix_min_pct) | is_ret if not df_cur_mix.empty else pd.Series(dtype=bool)
    df_cur_f  = df_cur_mix[mask_cur].copy() if not df_cur_mix.empty else pd.DataFrame()
    df_comp_f = df_comp_mix[df_comp_mix["label"].isin(df_cur_f["label"].tolist())].copy() if has_comp_mix and not df_cur_f.empty else pd.DataFrame()

    merged_mix = df_cur_f.rename(columns={"amount":"amount_cur","pct":"pct_cur"}).copy() if not df_cur_f.empty else pd.DataFrame()
    if has_comp_mix and not df_comp_f.empty and not merged_mix.empty:
        _comp_r = df_comp_f.rename(columns={"amount":"amount_comp","pct":"pct_comp"})
        merged_mix = merged_mix.merge(_comp_r[["label","amount_comp","pct_comp"]], on="label", how="left").fillna(0)
    elif not merged_mix.empty:
        merged_mix["amount_comp"] = None; merged_mix["pct_comp"] = None

    # ─── layout: stacked bar (چپ) | جدول+دونات (راست) ──────
    col_bar, col_right = st.columns([3, 2])

    with col_bar:
        # stacked bar
        _all_raw_mix = []
        _label_data  = {}
        for _lbl in top_labels:
            _pd2 = sym_prod[sym_prod["label"] == _lbl].sort_values("period_end")
            _raw2 = (btmn_to_musd(_pd2["sales_amount_btmn"], _pd2["period_end"], usd_rates)
                     if use_usd_mix and has_usd else _pd2["sales_amount_btmn"].tolist())
            _label_data[_lbl] = (_pd2, _raw2)
            _all_raw_mix.extend([v for v in _raw2 if v is not None])

        _bar_scale, _bar_lbl = smart_unit(_all_raw_mix, use_usd_mix and has_usd)
        _bar_dec = 2 if _bar_scale != 1.0 else dec_mix

        # ترتیبِ مرجعِ ماه‌ها: از روی period_end (که عددی و قابلِ‌مرتب‌سازی است)
        # تا وقتی دو سری مجموعه‌ماه‌های متفاوت دارند، Plotly برچسب‌های فارسی را
        # الفبایی نچیند و ترتیب ماه‌ها به‌هم نریزد.
        _periods_sorted = sorted(sym_prod["period_end"].unique())   # صعودی: قدیمی→جدید
        _cat_order = [format_period(p) for p in _periods_sorted]

        fig_bar_mix = go.Figure()
        for _lbl in top_labels:
            _pd2, _raw2 = _label_data[_lbl]
            _scaled = [v * _bar_scale if v is not None else 0 for v in _raw2]
            fig_bar_mix.add_trace(go.Bar(
                x=[format_period(p) for p in _pd2["period_end"]],
                y=_scaled, name=_lbl,
                marker_color=label_color_mix[_lbl],
                hovertemplate=f"<b>{_lbl}</b><br>%{{x}}<br>%{{y:,.{_bar_dec}f}} {_bar_lbl}<extra></extra>",
            ))
        fig_bar_mix.update_layout(
            barmode="stack",
            title=dict(text=f"{selected_sym} — ترکیب ماهانه ({_bar_lbl})", font=PLOTLY_FONT_TITLE),
            xaxis=dict(categoryorder="array", categoryarray=_cat_order,
                       tickfont=PLOTLY_FONT),
            yaxis=dict(title=dict(text=_bar_lbl, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
            font=PLOTLY_FONT, template="plotly_dark", height=380,
            hoverlabel=dict(font=PLOTLY_FONT),
            legend=dict(font=PLOTLY_FONT, orientation="v"),
            bargap=0.1,
        )
        st.plotly_chart(fig_bar_mix, use_container_width=True)

    with col_right:
        # دونات
        if pie_mode == "ماه آخر":
            _pie_df = df_cur_f.copy() if not df_cur_f.empty else pd.DataFrame()
            _pie_title = format_period(sel_period_mix)
        else:
            _pie_df = build_mix_avg12(sel_period_mix, all_periods_mix)
            if not _pie_df.empty:
                _is_ret2 = _pie_df["label"].str.contains("برگشت از فروش", na=False)
                _pie_df  = _pie_df[(_pie_df["pct"] >= mix_min_pct) | _is_ret2]
            _pie_title = "میانگین ۱۲ ماه"

        if not _pie_df.empty:
            _pie_lbls = _pie_df.sort_values("pct", ascending=False)["label"].tolist()
            _pie_vals = _pie_df.sort_values("pct", ascending=False)["amount"].tolist()
            _pie_clrs = [label_color_mix.get(l, DONUT_COLORS[i % len(DONUT_COLORS)])
                         for i, l in enumerate(_pie_lbls)]
            _pie_total_f = sum(v for v in _pie_vals if v > 0)

            # دونات کوچک + لیست کنارش
            _dc1, _dc2 = st.columns([1, 1])
            with _dc1:
                fig_pie_m = go.Figure(go.Pie(
                    labels=_pie_lbls, values=_pie_vals, hole=0.55,
                    textinfo="none", marker=dict(colors=_pie_clrs),
                    hovertemplate="<b>%{label}</b><br>%{percent}<extra></extra>",
                    direction="clockwise", sort=True, showlegend=False,
                ))
                fig_pie_m.update_layout(
                    font=PLOTLY_FONT, template="plotly_dark", height=230,
                    margin=dict(t=25,b=5,l=5,r=5),
                    title=dict(text=_pie_title, font=dict(family="Vazirmatn,Tahoma,sans-serif",size=11),
                               x=0.5, xanchor="center"),
                )
                st.plotly_chart(fig_pie_m, use_container_width=True)
            with _dc2:
                st.markdown('<div style="margin-top:30px"></div>', unsafe_allow_html=True)
                for _ll, _vv in zip(_pie_lbls, _pie_vals):
                    _pp = _vv / _pie_total_f * 100 if _pie_total_f > 0 else 0
                    _cc = label_color_mix.get(_ll, "#4DA3FF")
                    _sh = _ll[:18] + "…" if len(_ll) > 18 else _ll
                    st.markdown(
                        f'''<div style="display:flex;align-items:center;gap:5px;margin-bottom:6px;
                            direction:rtl;font-family:Vazirmatn,Tahoma,sans-serif">
                          <span style="background:{_cc};color:#fff;border-radius:3px;
                                padding:1px 5px;font-size:10px;font-weight:700">{_pp:.0f}٪</span>
                          <span style="font-size:10px;color:#e2e8f0">{_sh}</span>
                        </div>''', unsafe_allow_html=True,
                    )

        # جدول مقایسه
        if not merged_mix.empty:
            st.markdown("**مقایسه ترکیب:**")
            _h_cur  = format_period(sel_period_mix)
            _cv     = merged_mix["amount_cur"].dropna().tolist()
            _sc_c, _lb_c = smart_unit(_cv, use_usd_mix)
            _dc_c = 2 if _sc_c != 1.0 else (2 if use_usd_mix else 1)
            if has_comp_mix:
                _ccv = merged_mix["amount_comp"].dropna().tolist()
                _sc_k, _lb_k = smart_unit(_ccv, use_usd_mix)
                _dc_k = 2 if _sc_k != 1.0 else (2 if use_usd_mix else 1)
                _rows_t = []
                for _, _r in merged_mix.iterrows():
                    _rows_t.append({
                        "محصول":                              _r["label"],
                        f"مبلغ ({_lb_c})":                   fmt_fa((_r["amount_cur"] or 0) * _sc_c, _dc_c),
                        f"% | {_h_cur}":                     fmt_fa(_r["pct_cur"], 0) + "٪",
                        f"مبلغ ({_lb_k})":                   fmt_fa((_r.get("amount_comp") or 0) * _sc_k, _dc_k),
                        f"% | {comp_lbl_mix}":               fmt_fa(_r.get("pct_comp", 0), 0) + "٪",
                    })
            else:
                _rows_t = []
                for _, _r in merged_mix.iterrows():
                    _rows_t.append({
                        "محصول":        _r["label"],
                        f"مبلغ ({_lb_c})": fmt_fa((_r["amount_cur"] or 0) * _sc_c, _dc_c),
                        "% از کل":     fmt_fa(_r["pct_cur"], 0) + "٪",
                    })
            render_table(pd.DataFrame(_rows_t))

    # ════════════════════════════════════════════════════════
    # بخش ۲: کارت‌های محصول + نمودارهای جزئیات
    # ════════════════════════════════════════════════════════
    st.markdown("---")

    item_sum12  = (sym_for_cards[sym_for_cards["period_end"].isin(last_12_prod)]
                   .groupby("product_name")["sales_amount_btmn"].sum())
    item_avg12  = (sym_for_cards[sym_for_cards["period_end"].isin(last_12_prod)]
                   .groupby("product_name")["sales_amount_btmn"].mean())
    total_sum12 = float(item_sum12[item_sum12 > 0].sum())
    item_pct12  = (item_sum12 / total_sum12 * 100).fillna(0) if total_sum12 > 0 else item_sum12 * 0

    max_pct_c   = float(item_pct12.max()) if not item_pct12.empty else 100.0
    max_slider_c = min(round(max_pct_c, 1), 99.0) if max_pct_c > 0 else 1.0



    min_pct_prod = st.session_state.get("shared_min_pct", 3)

    all_items      = sorted(sym_for_cards["product_name"].unique())
    filtered_items = [p for p in all_items if item_pct12.get(p, 0) >= min_pct_prod]
    if not filtered_items:
        filtered_items = all_items

    # ─── نرخ موزون محصول ─────────────────────────────────────
    def calc_weighted_rate(prod_name):
        _sub12 = sym_for_cards[
            (sym_for_cards["product_name"] == prod_name) &
            (sym_for_cards["period_end"].isin(last_12_prod))
        ].copy()
        if _sub12.empty: return None
        _sub12 = _sub12[(_sub12["sales_rate_rial"] > 0) & (_sub12["sales_amount_btmn"] > 0)]
        if _sub12.empty: return None
        total_amt = _sub12["sales_amount_btmn"].sum()
        if total_amt <= 0: return None
        weighted = (_sub12["sales_rate_rial"] / 10_000_000 * _sub12["sales_amount_btmn"]).sum() / total_amt
        return weighted
    weighted_rate_12 = {p: calc_weighted_rate(p) for p in all_items}

    def _default_prod(items):
        return max(items, key=lambda p: float(item_pct12.get(p, 0) or 0)) if items else None

    _prev_sym_pa = st.session_state.get("pa_last_sym", None)
    if _prev_sym_pa != selected_sym:
        st.session_state.pa_selected = _default_prod(filtered_items)
        st.session_state["pa_last_sym"] = selected_sym
    elif st.session_state.get("pa_selected") not in filtered_items:
        st.session_state.pa_selected = _default_prod(filtered_items)

    def get_cat_info(name):
        if merge_cats:
            orig = sym_products[sym_products["product_name"] == name]
            dom  = orig[orig["category"].str.contains("داخلی", na=False)]["period_end"].nunique()
            exp  = orig[orig["category"].str.contains("صادرات", na=False)]["period_end"].nunique()
            parts_info = []
            if dom: parts_info.append(f"{dom} ماه داخلی")
            if exp: parts_info.append(f"{exp} ماه صادراتی")
            return " | ".join(parts_info)
        else:
            parts = name.rsplit(" — ", 1)
            if len(parts) == 2:
                orig = sym_products[(sym_products["product_name"] == parts[0]) &
                                    (sym_products["category"] == parts[1])]
            else:
                orig = sym_products[sym_products["product_name"] == name]
            return f"{orig['period_end'].nunique()} ماه"

    n_cols = 4
    rows_c = [filtered_items[i:i+n_cols] for i in range(0, len(filtered_items), n_cols)]
    for row_items in rows_c:
        cols_card = st.columns(n_cols)
        for j, prod in enumerate(row_items):
            pct    = item_pct12.get(prod, 0)
            avg    = item_avg12.get(prod, 0)
            info   = get_cat_info(prod)
            is_sel = (st.session_state.pa_selected == prod)
            border = "#4DA3FF" if is_sel else "#334155"
            bg     = "#1e3a5f" if is_sel else "#1e293b"
            tick   = "✓ " if is_sel else ""
            with cols_card[j]:
                card_html = (
                    f'<div style="background:{bg};border-radius:10px;'
                    f'padding:14px 16px;border:2px solid {border};'
                    f'font-family:Vazirmatn,Tahoma,sans-serif;direction:rtl;'
                    f'margin-bottom:2px;min-height:110px">'
                    f'<div style="font-size:13px;font-weight:700;color:#e2e8f0'
                    f';margin-bottom:4px">{tick}{prod}</div>'
                    f'<div style="font-size:11px;color:#94a3b8;margin-bottom:6px">{info}</div>'
                    f'<div style="font-size:24px;font-weight:700;color:#4DA3FF;line-height:1">'
                    f'{pct:.1f}٪</div>'
                    f'<div style="font-size:11px;color:#64748b;margin-top:4px">'
                    f'میانگین مبلغ (۱۲ماهه): {avg:.1f} م.تومان</div>'
                    f'</div>'
                )
                st.markdown(card_html, unsafe_allow_html=True)
                if st.button(" ", key=f"pa_btn_{prod}", use_container_width=True, help=f"انتخاب {prod}"):
                    st.session_state.pa_selected = prod
                    st.rerun()
                st.markdown(
                    f'''<style>
                    div[data-testid="stButton"]:has(button[title="انتخاب {prod}"]) button {{
                        margin-top: -8px; height: 6px !important;
                        min-height: 0 !important; opacity: 0 !important;
                    }}
                    div[data-testid="stButton"]:has(button[title="انتخاب {prod}"]) {{
                        margin-top: -120px; height: 120px;
                    }}
                    </style>''', unsafe_allow_html=True,
                )

    # ─── جزئیات محصول انتخابی ────────────────────────────────
    st.markdown("---")
    selected_prod = st.session_state.pa_selected
    st.subheader(f"جزئیات: {selected_prod}")

    sub_all_cats = sym_for_cards[sym_for_cards["product_name"] == selected_prod].copy()
    sub = sub_all_cats[sub_all_cats["sales_rate_rial"] > 0].sort_values("period_end").copy()
    sub["دوره"] = sub["period_end"].apply(format_period)

    rate_col   = "sales_rate_rial"
    rate_color = COLOR_DOLLAR if use_usd_mix else COLOR_TOMAN
    if use_usd_mix and has_usd:
        def _rate_to_usd(row):
            if not row[rate_col]: return None
            rate = usd_rates.get(_ym_key(row["period_end"]), None)
            return row[rate_col] / (rate * 10) if rate else None
        sub["rate_col_val"] = sub.apply(_rate_to_usd, axis=1)
        rate_label = "دلار"
    else:
        sub["rate_col_val"] = sub[rate_col] / 10_000_000
        rate_label = "م.تومان"

    detail_tabs = st.tabs(["نرخ فروش", "مبلغ فروش", "مقدار تولید", "مقدار فروش", "تولید+فروش"])

    with detail_tabs[0]:
        if sub.empty:
            st.info("داده نرخ موجود نیست.")
        else:
            all_rate_vals = sub["rate_col_val"].dropna().tolist()
            rate_scale, rate_scaled_lbl = smart_rate_unit(all_rate_vals, use_usd_mix and has_usd)
            sub2 = sub.copy()
            sub2["rate_scaled"] = sub2["rate_col_val"].apply(
                lambda x: x * rate_scale if x is not None and not pd.isna(x) else None)
            fig_rate = go.Figure()
            for cat in sorted(sub2["category"].unique()):
                cat_sub = sub2[sub2["category"] == cat]
                c = COLOR_DOLLAR if "صادرات" in cat else rate_color
                fig_rate.add_trace(go.Scatter(
                    x=cat_sub["دوره"], y=cat_sub["rate_scaled"], mode="lines+markers",
                    line=dict(color=c, width=3, shape="spline", smoothing=0.6),
                    marker=dict(size=8, color=c), name=cat,
                    hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.2f}} {rate_scaled_lbl}<extra></extra>",
                ))
            fig_rate.update_layout(
                title=dict(text=f"نرخ فروش {selected_prod}", font=PLOTLY_FONT_TITLE),
                xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT),
                yaxis=dict(title=dict(text=rate_scaled_lbl, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
                font=PLOTLY_FONT, template="plotly_dark", height=400,
                hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
            )
            st.plotly_chart(fig_rate, use_container_width=True)

    with detail_tabs[1]:
        mode_amt = st.radio("نوع نمایش", ["مقایسه‌ای", "مقایسه‌ای تجمعی", "روند"],
                            horizontal=True, key="pa_amt_mode")
        amt_vals = sub_all_cats["sales_amount_btmn"].dropna().tolist()
        amt_scale, amt_lbl = smart_unit(amt_vals, use_usd_mix and has_usd)

        sub_amt = sub_all_cats.sort_values("period_end").copy()
        sub_amt["دوره"] = sub_amt["period_end"].apply(format_period)
        # scale ساده بدون lambda پیچیده
        sub_amt["amt_sc"] = [
            (convert_amt_mix(v, sub_amt.iloc[i]["period_end"]) * amt_scale if v is not None else None)
            for i, v in enumerate(sub_amt["sales_amount_btmn"])
        ]

        def _eym_amt(p):
            try: y,m,_=p.split("/"); return int(y),int(m)
            except: return None,None
        _assign_year_month(sub_amt, _eym_amt)
        sub_amt["ماه_label"] = sub_amt["ماه_num"].apply(
            lambda m: PERSIAN_MONTHS.get(f"{int(m):02d}", str(m)) if pd.notna(m) else "")

        YEAR_COLORS_A = ["#A78BFA","#4DA3FF","#10B981","#F59E0B","#EF4444","#EC4899"]
        fig_amt = go.Figure()

        if mode_amt in ("مقایسه‌ای", "مقایسه‌ای تجمعی"):
            _cum_a = (mode_amt == "مقایسه‌ای تجمعی")
            years_a = sorted(sub_amt["سال"].dropna().unique().astype(int))
            month_order = list(PERSIAN_MONTHS.values())
            for cat in sorted(sub_amt["category"].unique()):
                for yi, yr in enumerate(years_a):
                    d = sub_amt[(sub_amt["سال"]==yr) & (sub_amt["category"]==cat)].sort_values("ماه_num")
                    if d.empty: continue
                    yv = d["amt_sc"].cumsum() if _cum_a else d["amt_sc"]
                    clr = YEAR_COLORS_A[yi % len(YEAR_COLORS_A)]
                    _suf = " (تجمعی)" if _cum_a else ""
                    fig_amt.add_trace(go.Bar(
                        x=d["ماه_label"], y=yv,
                        name=f"{to_fa_digits(str(yr))} {cat}",
                        marker_color=clr,
                        opacity=0.7 if "صادرات" in cat else 1.0,
                        hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.2f}} {amt_lbl}{_suf}<extra></extra>",
                    ))
            fig_amt.update_layout(barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=month_order, tickfont=PLOTLY_FONT))
        else:
            # روند — همه دسته‌ها روی هم stacked
            for cat in sorted(sub_amt["category"].unique()):
                cs = sub_amt[sub_amt["category"]==cat]
                c  = COLOR_DOLLAR if "صادرات" in cat else COLOR_TOMAN
                fig_amt.add_trace(go.Bar(
                    x=cs["دوره"], y=cs["amt_sc"], name=cat, marker_color=c,
                    hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.2f}} {amt_lbl}<extra></extra>",
                ))
            fig_amt.update_layout(barmode="stack",
                xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT))

        fig_amt.update_layout(
            title=dict(text=f"مبلغ فروش {selected_prod} ({amt_lbl})", font=PLOTLY_FONT_TITLE),
            yaxis=dict(title=dict(text=amt_lbl, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
            font=PLOTLY_FONT, template="plotly_dark", height=420,
            hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
            bargap=0.15, bargroupgap=0.05,
        )
        st.plotly_chart(fig_amt, use_container_width=True)

    def make_qty_fig_pa(qty_col, qty_label, mode):
        vol_sub = sub_all_cats.sort_values("period_end").copy()
        vol_sub["دوره"] = vol_sub["period_end"].apply(format_period)
        def _eym(p):
            try: y,m,_ = p.split("/"); return int(y),int(m)
            except: return None,None
        _assign_year_month(vol_sub, _eym)
        vol_sub["ماه_label"] = vol_sub["ماه_num"].apply(
            lambda m: PERSIAN_MONTHS.get(f"{int(m):02d}", str(m)) if pd.notna(m) else "")
        unit_str = sub_all_cats.iloc[0]["unit"] if not sub_all_cats.empty else ""
        fig = go.Figure()
        YEAR_COLORS_Q = ["#A78BFA","#4DA3FF","#10B981","#F59E0B","#EF4444","#EC4899"]
        if mode in ("مقایسه‌ای", "مقایسه‌ای تجمعی"):
            _cum_q = (mode == "مقایسه‌ای تجمعی")
            years_q = sorted(vol_sub["سال"].dropna().unique().astype(int))
            month_order = list(PERSIAN_MONTHS.values())
            for cat in sorted(vol_sub["category"].unique()):
                for yi, yr in enumerate(years_q):
                    d = vol_sub[(vol_sub["سال"]==yr) & (vol_sub["category"]==cat)].sort_values("ماه_num")
                    if d.empty: continue
                    yv = d[qty_col].cumsum() if _cum_q else d[qty_col]
                    clr = YEAR_COLORS_Q[yi % len(YEAR_COLORS_Q)]
                    _suf = " (تجمعی)" if _cum_q else ""
                    fig.add_trace(go.Bar(
                        x=d["ماه_label"], y=yv,
                        name=f"{to_fa_digits(str(yr))} {cat}", marker_color=clr,
                        opacity=0.7 if "صادرات" in cat else 1.0,
                        hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.0f}} {unit_str}{_suf}<extra></extra>",
                    ))
            fig.update_layout(barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=month_order, tickfont=PLOTLY_FONT))
        else:
            vol_g = vol_sub.groupby("period_end").agg(
                qty=(qty_col,"sum"), دوره=("دوره","first")
            ).reset_index().sort_values("period_end")
            color = "#10B981" if qty_col=="sales_qty" else "#F59E0B"
            fig.add_trace(go.Bar(
                x=vol_g["دوره"], y=vol_g["qty"], marker_color=color, name=qty_label,
                hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.0f}} {unit_str}<extra></extra>",
            ))
            fig.update_layout(xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT))
        fig.update_layout(
            title=dict(text=f"{selected_prod} — {qty_label} ({unit_str})", font=PLOTLY_FONT_TITLE),
            yaxis=dict(title=dict(text=unit_str, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
            font=PLOTLY_FONT, template="plotly_dark", height=420,
            hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
            bargap=0.15, bargroupgap=0.05,
        )
        return fig

    with detail_tabs[2]:
        mode_prod = st.radio("نوع نمایش", ["مقایسه‌ای", "مقایسه‌ای تجمعی", "روند"],
                             horizontal=True, key="pa_prod_mode")
        st.plotly_chart(make_qty_fig_pa("production_qty", "مقدار تولید", mode_prod), use_container_width=True)

    with detail_tabs[3]:
        mode_sales = st.radio("نوع نمایش", ["مقایسه‌ای", "مقایسه‌ای تجمعی", "روند"],
                              horizontal=True, key="pa_sales_mode")
        st.plotly_chart(make_qty_fig_pa("sales_qty", "مقدار فروش", mode_sales), use_container_width=True)

    with detail_tabs[4]:
        mode_both = st.radio("نوع نمایش", ["مقایسه‌ای", "مقایسه‌ای تجمعی", "روند"],
                             horizontal=True, key="pa_both_mode")
        vol_sub2 = sub_all_cats.sort_values("period_end").copy()
        vol_sub2["دوره"] = vol_sub2["period_end"].apply(format_period)
        unit_str2 = sub_all_cats.iloc[0]["unit"] if not sub_all_cats.empty else ""
        def _eym2(p):
            try: y,m,_=p.split("/"); return int(y),int(m)
            except: return None,None
        _assign_year_month(vol_sub2, _eym2)
        vol_sub2["ماه_label"] = vol_sub2["ماه_num"].apply(
            lambda m: PERSIAN_MONTHS.get(f"{int(m):02d}", str(m)) if pd.notna(m) else "")
        fig_both = go.Figure()
        YEAR_COLORS_B = ["#A78BFA","#4DA3FF","#10B981","#F59E0B","#EF4444","#EC4899"]
        if mode_both in ("مقایسه‌ای", "مقایسه‌ای تجمعی"):
            _cum_b = (mode_both == "مقایسه‌ای تجمعی")
            years_b = sorted(vol_sub2["سال"].dropna().unique().astype(int))
            month_order = list(PERSIAN_MONTHS.values())
            for yi, yr in enumerate(years_b):
                d = vol_sub2[vol_sub2["سال"]==yr].groupby("period_end").agg(
                    sales_qty=("sales_qty","sum"), production_qty=("production_qty","sum"),
                    ماه_label=("ماه_label","first"), ماه_num=("ماه_num","first")).reset_index().sort_values("ماه_num")
                clr = YEAR_COLORS_B[yi % len(YEAR_COLORS_B)]
                yr_fa = to_fa_digits(str(yr))
                pq = d["production_qty"].cumsum() if _cum_b else d["production_qty"]
                sq = d["sales_qty"].cumsum() if _cum_b else d["sales_qty"]
                _suf = " (تجمعی)" if _cum_b else ""
                fig_both.add_trace(go.Bar(x=d["ماه_label"], y=pq,
                    name=f"{yr_fa} تولید", marker_color=clr, opacity=0.6,
                    hovertemplate=f"<b>%{{x}}</b><br>تولید: %{{y:,.0f}}{_suf}<extra></extra>"))
                fig_both.add_trace(go.Bar(x=d["ماه_label"], y=sq,
                    name=f"{yr_fa} فروش", marker_color=clr,
                    hovertemplate=f"<b>%{{x}}</b><br>فروش: %{{y:,.0f}}{_suf}<extra></extra>"))
            fig_both.update_layout(barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=month_order, tickfont=PLOTLY_FONT))
        else:
            vol_g2 = vol_sub2.groupby("period_end").agg(
                sales_qty=("sales_qty","sum"), production_qty=("production_qty","sum"),
                دوره=("دوره","first")).reset_index().sort_values("period_end")
            fig_both.add_trace(go.Bar(x=vol_g2["دوره"], y=vol_g2["production_qty"],
                marker_color="#F59E0B", name="تعداد تولید",
                hovertemplate="<b>%{x}</b><br>تولید: %{y:,.0f}<extra></extra>"))
            fig_both.add_trace(go.Bar(x=vol_g2["دوره"], y=vol_g2["sales_qty"],
                marker_color="#10B981", name="تعداد فروش",
                hovertemplate="<b>%{x}</b><br>فروش: %{y:,.0f}<extra></extra>"))
            fig_both.update_layout(xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT))
        fig_both.update_layout(
            title=dict(text=f"{selected_prod} — تولید و فروش ({unit_str2})", font=PLOTLY_FONT_TITLE),
            yaxis=dict(title=dict(text=unit_str2, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
            font=PLOTLY_FONT, template="plotly_dark", height=420, barmode="group",
            hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
            bargap=0.15, bargroupgap=0.05,
        )
        st.plotly_chart(fig_both, use_container_width=True)

    # جدول اعداد
    _tbl_btn_lbl = "📋 بستن جدول اعداد" if st.session_state.get("pa_table_open") else "📋 نمایش جدول اعداد"
    if st.button(_tbl_btn_lbl, key="pa_tbl_toggle"):
        st.session_state["pa_table_open"] = not st.session_state.get("pa_table_open", False)

    if st.session_state.get("pa_table_open", False):
        tbl = sub_all_cats.sort_values("period_end", ascending=False).copy()
        tbl["دوره"] = tbl["period_end"].apply(format_period)
        raw_rate_tmn = (tbl["sales_rate_rial"] / 10_000_000).tolist()
        scale_rt, lbl_rt = smart_unit(raw_rate_tmn, False)
        tbl["نرخ_تومان_scaled"] = [v * scale_rt for v in raw_rate_tmn]
        raw_amt_tmn = tbl["sales_amount_btmn"].tolist()
        scale_at, lbl_at = smart_unit(raw_amt_tmn, False)
        tbl["مبلغ_تومان_scaled"] = [v * scale_at if v is not None else None for v in raw_amt_tmn]
        cols_tbl = ["دوره","category","unit","production_qty","sales_qty","نرخ_تومان_scaled","مبلغ_تومان_scaled"]
        rename_tbl = {
            "category":"دسته","unit":"واحد","production_qty":"تولید","sales_qty":"فروش",
            "نرخ_تومان_scaled": f"نرخ ({lbl_rt})", "مبلغ_تومان_scaled": f"مبلغ ({lbl_at})",
        }
        if has_usd:
            raw_rate_usd = [
                (tbl.iloc[i]["sales_rate_rial"] / (usd_rates.get(_ym_key(tbl.iloc[i]["period_end"]),None) * 10))
                if usd_rates.get(_ym_key(tbl.iloc[i]["period_end"])) else None
                for i in range(len(tbl))]
            scale_ru, lbl_ru = smart_unit([v for v in raw_rate_usd if v], True)
            tbl["نرخ_دلار_scaled"] = [v * scale_ru if v is not None else None for v in raw_rate_usd]
            raw_amt_usd = btmn_to_musd(tbl["sales_amount_btmn"], tbl["period_end"], usd_rates)
            scale_au, lbl_au = smart_unit([v for v in raw_amt_usd if v], True)
            tbl["مبلغ_دلار_scaled"] = [v * scale_au if v is not None else None for v in raw_amt_usd]
            cols_tbl += ["نرخ_دلار_scaled","مبلغ_دلار_scaled"]
            rename_tbl.update({"نرخ_دلار_scaled": f"نرخ ({lbl_ru})", "مبلغ_دلار_scaled": f"مبلغ ({lbl_au})"})
        display_tbl = tbl[cols_tbl].copy()
        dec_rt2 = 2 if scale_rt != 1 else 1
        dec_at2 = 2 if scale_at != 1 else 1
        display_tbl["نرخ_تومان_scaled"]  = display_tbl["نرخ_تومان_scaled"].apply(lambda x: fmt_fa(x, dec_rt2))
        display_tbl["مبلغ_تومان_scaled"] = display_tbl["مبلغ_تومان_scaled"].apply(
            lambda x: fmt_fa(x, dec_at2) if x is not None else "—")
        for c in ["production_qty","sales_qty"]:
            display_tbl[c] = display_tbl[c].apply(lambda x: fmt_fa(x, 0))
        if has_usd:
            display_tbl["نرخ_دلار_scaled"]  = display_tbl["نرخ_دلار_scaled"].apply(
                lambda x: fmt_fa(x, 2) if x is not None else "—")
            display_tbl["مبلغ_دلار_scaled"] = display_tbl["مبلغ_دلار_scaled"].apply(
                lambda x: fmt_fa(x, 2) if x is not None else "—")
        display_tbl = display_tbl.rename(columns=rename_tbl)
        render_table(display_tbl)

    # ════════════════════════════════════════════════════════
    # نرخ موزون کل نماد در طول زمان
    # ════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("#### 📈 نرخ موزون فروش نماد")
    _wrate_sym_rows = []
    for _pp in sorted(sym_prod["period_end"].unique()):
        _pp_sub = sym_prod[
            (sym_prod["period_end"] == _pp) &
            (sym_prod["sales_rate_rial"] > 0) &
            (sym_prod["sales_amount_btmn"] > 0)
        ]
        if _pp_sub.empty: continue
        _total_a = _pp_sub["sales_amount_btmn"].sum()
        if _total_a <= 0: continue
        _wr_sym = (_pp_sub["sales_rate_rial"] / 10_000_000 * _pp_sub["sales_amount_btmn"]).sum() / _total_a
        if use_usd_mix and has_usd:
            _r_usd = usd_rates.get(_ym_key(_pp))
            if _r_usd and _r_usd > 0:
                _wr_sym = _wr_sym / _r_usd
        _wrate_sym_rows.append({"period_end": _pp, "wr": _wr_sym})

    if _wrate_sym_rows:
        _wr_sym_df = pd.DataFrame(_wrate_sym_rows).sort_values("period_end")
        _wr_sym_df["دوره"] = _wr_sym_df["period_end"].apply(format_period)
        _wr_sym_vals = _wr_sym_df["wr"].dropna().tolist()
        _wr_sym_scale, _wr_sym_lbl = smart_rate_unit(_wr_sym_vals, use_usd_mix and has_usd)
        _wr_sym_df["wr_sc"] = _wr_sym_df["wr"] * _wr_sym_scale
        _last12_wr = sorted(_wr_sym_df["period_end"].unique(), reverse=True)[:12]
        _avg_wr_sym = _wr_sym_df[_wr_sym_df["period_end"].isin(_last12_wr)]["wr_sc"].mean()
        fig_wr_sym = go.Figure()
        fig_wr_sym.add_trace(go.Scatter(
            x=_wr_sym_df["دوره"], y=_wr_sym_df["wr_sc"],
            mode="lines+markers",
            line=dict(color="#10B981", width=3, shape="spline", smoothing=0.6),
            marker=dict(size=6, color="#10B981"),
            fill="tozeroy", fillcolor="rgba(16,185,129,0.1)",
            hovertemplate=f"<b>%{{x}}</b><br>%{{y:,.2f}} {_wr_sym_lbl}<extra></extra>",
        ))
        fig_wr_sym.add_hline(
            y=_avg_wr_sym, line_dash="dot", line_color="#F59E0B", opacity=0.8,
            annotation_text=f"م.۱۲ماه: {fmt_fa(_avg_wr_sym, 2)} {_wr_sym_lbl}",
            annotation_position="bottom left",
            annotation_font=dict(family="Vazirmatn,Tahoma,sans-serif", size=11, color="#F59E0B"),
        )
        fig_wr_sym.update_layout(
            title=dict(text=f"{selected_sym} — نرخ موزون فروش کل محصولات ({_wr_sym_lbl})", font=PLOTLY_FONT_TITLE),
            xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT),
            yaxis=dict(title=dict(text=_wr_sym_lbl, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
            font=PLOTLY_FONT, template="plotly_dark", height=350,
            margin=dict(r=160), hoverlabel=dict(font=PLOTLY_FONT),
        )
        st.plotly_chart(fig_wr_sym, use_container_width=True)
        st.caption("نرخ موزون = Σ(نرخ محصول × مبلغ فروش) / Σ(مبلغ فروش) برای همه محصولات در هر ماه")

    # ════════════════════════════════════════════════════════
    # مبلغ تولیدات نماد  (همیشه دلاری) — نمایش مقایسه‌ایِ سال‌به‌سال
    #   برای هر دوره: Σ ( مقدار تولید هر محصول × نرخِ دلاریِ مرجعِ آن محصول )
    #
    #   نرخِ دلاریِ مرجع بسته به دکمهٔ کناری:
    #     • حالت میانگین (پیش‌فرض): میانگینِ نرخِ دلاریِ تک‌تکِ ماه‌ها.
    #     • حالت «نرخِ آخرین ماهِ پرفروش»: نرخِ دلاریِ آخرین ماهی (در کلِ تاریخچه)
    #       که در آن حداقل ۲۰٪ از تولیدِ همان ماهِ تولیدی فروخته شده باشد.
    # ════════════════════════════════════════════════════════
    st.markdown("---")
    _pv_h1, _pv_h2 = st.columns([3, 2])
    with _pv_h1:
        st.markdown("#### 🏭 مبلغ تولیدات نماد")
    with _pv_h2:
        _pv_use_lastsale = st.checkbox(
            "نرخِ آخرین ماهِ پرفروش (≥۲۰٪ تولید) به‌جای میانگین",
            value=True, key="pv_rate_mode",
        )

    if not has_usd:
        st.info("نرخ دلار موجود نیست؛ امکان محاسبهٔ دلاریِ مبلغ تولیدات وجود ندارد.")
    else:
        # ─── اطلاعاتِ ماهانهٔ هر محصول (هر محصول = product_name + category) ───
        # نرخِ دلاریِ هر ماه = sales_rate_rial / (نرخ دلارِ همان ماه × ۱۰)
        _prod_months = {}   # (pn,cat) -> list[{period, sales_qty, dollar_rate}] مرتب بر اساس دوره
        _avg_rate    = {}   # (pn,cat) -> میانگینِ نرخِ دلاری (دلار به ازای واحد)
        for (_pn, _cat), _g in sym_products.groupby(["product_name", "category"]):
            _rows_m = []
            _rates  = []
            for _, _row in _g.sort_values("period_end").iterrows():
                _rr = _row["sales_rate_rial"]
                _ur = usd_rates.get(_ym_key(_row["period_end"]))
                _dr = None
                if _rr and _rr > 0 and not pd.isna(_rr) and _ur and _ur > 0:
                    _dr = _rr / (_ur * 10.0)
                    _rates.append(_dr)
                _rows_m.append({
                    "period":      _row["period_end"],
                    "sales_qty":   _row["sales_qty"],
                    "dollar_rate": _dr,
                })
            _prod_months[(_pn, _cat)] = _rows_m
            if _rates:
                _avg_rate[(_pn, _cat)] = sum(_rates) / len(_rates)

        def _ref_rate(_pn, _cat, _period, _prod_qty):
            """نرخِ دلاریِ مرجعِ یک محصول برای ماهِ تولیدیِ مشخص."""
            if not _pv_use_lastsale:
                return _avg_rate.get((_pn, _cat))
            # حالت «آخرین ماهِ پرفروش»: آخرین ماهی (در کلِ تاریخچه، بدون محدودیتِ زمانی)
            # که در آن حداقل ۲۰٪ از تولیدِ همین ماهِ تولیدی فروخته شده باشد.
            if _prod_qty is None or pd.isna(_prod_qty) or _prod_qty <= 0:
                return _avg_rate.get((_pn, _cat))
            _thr  = 0.20 * _prod_qty
            _rows_m = _prod_months.get((_pn, _cat), [])
            _cands = [
                x for x in _rows_m
                if x["sales_qty"] is not None and not pd.isna(x["sales_qty"])
                and x["sales_qty"] >= _thr and x["dollar_rate"] is not None
            ]
            if _cands:
                # آخرین (جدیدترین) ماهِ واجدِ شرط در کلِ تاریخچه
                return max(_cands, key=lambda x: x["period"])["dollar_rate"]
            # بازگشتی: اگر هیچ ماهی به آستانهٔ ۲۰٪ نرسید → میانگینِ نرخِ دلاری
            return _avg_rate.get((_pn, _cat))

        # ─── مبلغ تولیدات هر دوره ───
        _pv_rows = []
        for _pp in sorted(sym_products["period_end"].unique()):
            _pp_sub = sym_products[sym_products["period_end"] == _pp]
            _val = 0.0
            _has_any = False
            for _, _row in _pp_sub.iterrows():
                _pq = _row["production_qty"]
                _rt = _ref_rate(_row["product_name"], _row["category"], _pp, _pq)
                if _rt is None or _pq is None or pd.isna(_pq):
                    continue
                _has_any = True
                _val += _pq * _rt / 1_000_000.0              # میلیون دلار
            if _has_any:
                _pv_rows.append({"period_end": _pp, "pv": _val})

        if _pv_rows:
            _pv_df = pd.DataFrame(_pv_rows).sort_values("period_end")
            _pv_scale, _pv_lbl = smart_unit(_pv_df["pv"].tolist(), True)
            _pv_df["pv_sc"] = _pv_df["pv"] * _pv_scale
            _pv_dec = 2

            # تفکیکِ سال/ماه برای نمایشِ مقایسه‌ای
            def _eym_pv(p):
                try: y, m, _ = p.split("/"); return int(y), int(m)
                except: return None, None
            _assign_year_month(_pv_df, _eym_pv)
            _pv_df["ماه_label"] = _pv_df["ماه_num"].apply(
                lambda m: PERSIAN_MONTHS.get(f"{int(m):02d}", str(m)) if pd.notna(m) else "")

            _PV_YEAR_COLORS = ["#A78BFA", "#4DA3FF", "#10B981", "#F59E0B", "#EF4444", "#EC4899", "#14B8A6"]
            _month_order = list(PERSIAN_MONTHS.values())
            _years_pv = sorted(_pv_df["سال"].dropna().unique().astype(int))

            fig_pv = go.Figure()
            for _yi, _yr in enumerate(_years_pv):
                _d = _pv_df[_pv_df["سال"] == _yr].sort_values("ماه_num")
                if _d.empty:
                    continue
                fig_pv.add_trace(go.Bar(
                    x=_d["ماه_label"], y=_d["pv_sc"],
                    name=to_fa_digits(str(_yr)),
                    marker_color=_PV_YEAR_COLORS[_yi % len(_PV_YEAR_COLORS)],
                    hovertemplate=f"<b>%{{x}} {to_fa_digits(str(_yr))}</b><br>%{{y:,.{_pv_dec}f}} {_pv_lbl}<extra></extra>",
                ))

            _pv_mode_lbl = "نرخِ آخرین ماهِ پرفروش" if _pv_use_lastsale else "میانگینِ نرخ دلاری"
            fig_pv.update_layout(
                title=dict(text=f"{selected_sym} — مبلغ تولیدات ({_pv_lbl}) — {_pv_mode_lbl}", font=PLOTLY_FONT_TITLE),
                barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=_month_order, tickfont=PLOTLY_FONT),
                yaxis=dict(title=dict(text=_pv_lbl, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
                font=PLOTLY_FONT, template="plotly_dark", height=420,
                hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
                bargap=0.15, bargroupgap=0.05,
            )
            st.plotly_chart(fig_pv, use_container_width=True)

            if _pv_use_lastsale:
                st.caption(
                    "مبلغ تولیدات = Σ (مقدار تولید هر محصول × نرخِ دلاریِ مرجع) برای همهٔ محصولات در هر ماه. "
                    "نرخِ مرجعِ هر محصول = نرخِ دلاریِ آخرین ماهی (در کلِ تاریخچه) که در آن حداقل ۲۰٪ از "
                    "تولیدِ همان ماهِ تولیدی فروخته شده باشد؛ اگر همان ماهِ تولید هم واجدِ شرط باشد ممکن است "
                    "نرخِ خودش انتخاب شود. اگر هیچ ماهی به آستانهٔ ۲۰٪ نرسد، میانگینِ نرخِ دلاری جایگزین می‌شود."
                )
            else:
                st.caption(
                    "مبلغ تولیدات = Σ (مقدار تولید هر محصول × میانگینِ نرخِ دلاریِ همان محصول) برای همهٔ محصولات در هر ماه. "
                    "نرخِ دلاریِ هر محصول = میانگینِ نرخِ دلاریِ تک‌تکِ ماه‌ها (هر ماه با نرخ دلارِ همان ماه)."
                )
        else:
            st.info("داده‌ای برای محاسبهٔ مبلغ تولیدات موجود نیست.")

    # ════════════════════════════════════════════════════════
    # بخش: مبلغ فروش / بهای تمام شده / سود ناخالصِ محصولات (فصلی) — از گزارش تفسیری
    # ════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 📊 مبلغ فروش / بهای تمام شده / سود ناخالصِ محصولات (فصلی)")
    st.caption("منبع: «خلاصه اطلاعات گزارش تفسیری» کدال (codal_product_pl.db). "
               "ارقامِ تجمعیِ هر دوره با تفاضل به مقدارِ هر فصل تبدیل می‌شوند. واحد: میلیارد تومان.")

    _ppl = load_product_pl()
    if _ppl is None or _ppl.empty:
        st.info("دادهٔ گزارش تفسیری یافت نشد. ابتدا `codal_product_pl_extractor.py` را اجرا کن "
                f"تا `{PRODUCT_PL_DB_PATH}` ساخته شود.")
    else:
        _ppl = _ppl.copy()
        _ppl["symn"] = _ppl["symbol"].apply(normalize_symbol)
        _sub = _ppl[(_ppl["symn"] == normalize_symbol(selected_sym)) & (_ppl["is_estimate"] == 0)].copy()
        if _sub.empty:
            st.info(f"برای «{selected_sym}» دادهٔ گزارش تفسیری نیست "
                    "(شاید گزارشش هنوز استخراج نشده؛ اکستراکتر را برای این نماد اجرا کن).")
        else:
            _FA_TR = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
            # تجمعیِ هر (بخش، محصول، جمع‌بودن) → {(سال,دوره): dict متریک‌ها}
            cumP, glob_durs = {}, {}
            for _, r in _sub.iterrows():
                yr = str(r["period_end"]).translate(_FA_TR)[:4]
                dur = int(r["duration_months"]) if pd.notna(r["duration_months"]) else None
                if not yr.isdigit() or dur is None:
                    continue
                key = (r["section"], r["product"], int(r["is_total"]))
                cumP.setdefault(key, {})[(yr, dur)] = {
                    "sales": r.get("sales"), "cost": r.get("cost"), "gross": r.get("gross"),
                    "qty_prod": r.get("qty_prod"), "qty_sold": r.get("qty_sold")}
                glob_durs.setdefault(yr, set()).add(dur)

            # ستون‌های فصلیِ سراسری + یک ردیفِ «سالانه» برای هر سال (۱۲ماهه، بدونِ تفاضل)
            qmap = [("اول", 3, None), ("دوم", 6, 3), ("سوم", 9, 6), ("چهارم", 12, 9)]
            quarter_cols = []
            for y in sorted(glob_durs):
                if not y.isdigit() or int(y) < 1402:   # فقط از سال مالی ۱۴۰۲ به بعد
                    continue
                durs = glob_durs[y]
                if any(d in (3, 6, 9) for d in durs):
                    for ql, dc, dp in qmap:
                        if dc in durs:
                            quarter_cols.append((y, ql, dc, dp))
                if 12 in durs:
                    quarter_cols.append((y, "سالانه", 12, None))

            if not quarter_cols:
                st.info("دادهٔ کافی برای جدولِ فصلی نیست.")
            else:
                def _qdiff(cmap, fld, dcur, dprev, year):
                    def g(dur):
                        t = cmap.get((year, dur))
                        if not t:
                            return None
                        v = t.get(fld)
                        return None if (v is None or pd.isna(v)) else v
                    a = g(dcur)
                    if a is None:
                        return None
                    if dprev is None:        # فصلِ اول یا ردیفِ سالانه → بدونِ تفاضل
                        return a
                    b = g(dprev)
                    return None if b is None else a - b   # دورهٔ قبلی نبود → قابلِ‌محاسبه نیست

                def _none(v):
                    return None if (v is None or (isinstance(v, float) and pd.isna(v))) else v

                def _fa_int(v):
                    v = _none(v)
                    if v is None:
                        return "—"
                    s = f"{abs(v):,.0f}".translate(EN_TO_FA)
                    return f"({s})" if v < 0 else s

                sec_order = {"داخلی": 0, "صادراتی": 1, "خدمات": 2, "سایر": 3}
                keys = sorted(cumP.keys(), key=lambda k: (sec_order.get(k[0], 9), k[2], k[1]))

                rows = []
                _qrank = {"اول": 1, "دوم": 2, "سوم": 3, "چهارم": 4, "سالانه": 5}
                _name_sales = {}   # مجموعِ مبلغ فروشِ هر محصول طی دوره‌ها (برای چیدنِ گزینه‌ها)
                for k in keys:
                    sec, prod, tot = k
                    name = prod if (tot or prod in ("جمع", "سایر")) else f"{prod} ({sec})"
                    cmap = cumP[k]
                    for (y, q, dc, dp) in quarter_cols:
                        sales = _none(_qdiff(cmap, "sales", dc, dp, y))
                        cost = _none(_qdiff(cmap, "cost", dc, dp, y))
                        gross = _none(_qdiff(cmap, "gross", dc, dp, y))
                        qprod = _none(_qdiff(cmap, "qty_prod", dc, dp, y))
                        qsold = _none(_qdiff(cmap, "qty_sold", dc, dp, y))
                        if all(x is None for x in (sales, cost, gross, qprod, qsold)):
                            continue
                        if sales is not None:
                            _name_sales[name] = _name_sales.get(name, 0.0) + sales
                        rate = (sales * 1e6 / qsold) if (sales is not None and qsold) else None
                        margin = (gross / sales * 100) if (sales and gross is not None and sales != 0) else None
                        season = f"{q} {y}".translate(EN_TO_FA)
                        rows.append({
                            "محصول": name,
                            "فصل": season,
                            "_skey": int(y) * 10 + _qrank.get(q, 0),   # کلیدِ زمانیِ فصل
                            "مقدار تولید": _fa_int(qprod),
                            "مقدار فروش": _fa_int(qsold),
                            "نرخ فروش (ریال)": _fa_int(round(rate) if rate is not None else None),
                            "مبلغ فروش (م.ت)": _fa_int(round(sales / 1e4) if sales is not None else None),
                            "مبلغ بهای تمام شده (م.ت)": _fa_int(round(cost / 1e4) if cost is not None else None),
                            "سود ناخالص (م.ت)": _fa_int(round(gross / 1e4) if gross is not None else None),
                            "حاشیه سود ناخالص": (f"{margin:.1f}".translate(EN_TO_FA) + "٪") if margin is not None else "—",
                        })

                _df_pl = pd.DataFrame(rows)
                # گزینه‌های باکس: بر اساسِ بیشترین مبلغ فروشِ مجموع طی دوره‌ها (نزولی)؛
                # «جمع فروش داخلی» و «جمع فروش صادراتی» از گزینه‌ها حذف می‌شوند
                _excl = {"جمع فروش داخلی", "جمع فروش صادراتی", "جمع"}
                _names = []
                for _p in _df_pl["محصول"].tolist():
                    if _p not in _names and _p not in _excl:
                        _names.append(_p)
                _names.sort(key=lambda nm: _name_sales.get(nm, 0.0), reverse=True)
                _prod_opts = ["همه"] + _names
                _pick = st.selectbox("محصول:", _prod_opts, key="ppl_prod_pick")
                if _pick != "همه":
                    _df_pl = _df_pl[_df_pl["محصول"] == _pick].reset_index(drop=True)

                _skeys = _df_pl["_skey"].tolist()
                _df_show = _df_pl.drop(columns=["_skey"])
                st.caption("هر سطر = یک محصول در یک فصل (از سال مالی ۱۴۰۲). مبالغ به میلیارد تومان "
                           "(بهای/منفی قرمز)، نرخ به ریال. روی عنوانِ هر ستون کلیک کن تا مرتب شود "
                           "(ستونِ «فصل» بر مبنای زمان مرتب می‌شود).")
                render_sortable_table(
                    _df_show,
                    numeric_cols={"مقدار تولید", "مقدار فروش", "نرخ فروش (ریال)", "مبلغ فروش (م.ت)",
                                  "مبلغ بهای تمام شده (م.ت)", "سود ناخالص (م.ت)", "حاشیه سود ناخالص"},
                    green_cols={"سود ناخالص (م.ت)", "حاشیه سود ناخالص"},
                    center_from=2,
                    sortkeys={"فصل": _skeys})



# ════════════════════════════════════════════════════════════════
# صفحه: گزارش‌های دوره‌ای (صورت سود و زیان میان‌دوره‌ای و سالانه)
# ════════════════════════════════════════════════════════════════
elif page == "📑 گزارش‌های دوره‌ای":
    st.title("گزارش‌های دوره‌ای — صورت سود و زیان")

    if not has_income:
        st.warning(
            "دیتابیس گزارش‌های دوره‌ای پیدا نشد.\n\n"
            f"مسیری که داشبورد بررسی کرد: `{INCOME_DB_PATH}`\n\n"
            "این فایل باید در فولدر `D:\\bourse\\codal seasonal` ساخته شود. "
            "ابتدا اسکریپت `codal_income_extractor.py` را در همان فولدر اجرا کنید "
            "تا `codal_income.db` ساخته و پر شود، سپس این صفحه فعال می‌شود."
        )
        st.stop()

    # ─── نماد از نوار حاشیه‌ای (مشترک با بقیه صفحات) ───────────────
    inc_syms = sorted(df_income_rep["symbol"].unique().tolist())
    inc_sym = selected_global
    if inc_sym not in inc_syms:
        st.warning(
            f"برای نماد «{inc_sym}» گزارش دوره‌ای موجود نیست.\n\n"
            "از منوی «نماد» در نوار سمت راست، یکی از نمادهای دارای گزارش دوره‌ای را انتخاب کنید."
        )
        with st.expander(f"نمادهای دارای گزارش دوره‌ای ({to_fa_digits(len(inc_syms))} نماد)"):
            st.write("، ".join(inc_syms))
        st.stop()

    reps = df_income_rep[df_income_rep["symbol"] == inc_sym].copy()
    reps = reps.sort_values(["period_end", "duration_months"], ascending=[False, False])

    # ─── آماده‌سازی اقلام هر گزارش ───────────────────────────────
    rep_ids = reps["id"].tolist()
    items_sym = df_income_items[df_income_items["report_id"].isin(rep_ids)]
    items_by_rep = {}
    for rid, g in items_sym.groupby("report_id"):
        gg = g.sort_values("row_order")
        items_by_rep[rid] = list(zip(gg["label"].tolist(), gg["value"].tolist()))

    col_meta  = [rep for _, rep in reps.iterrows()]
    col_items = [items_by_rep.get(rep["id"], []) for rep in col_meta]

    if not any(col_items):
        st.warning("اقلامی برای این نماد در دیتابیس نیست.")
        st.stop()

    # ─── ادغام هوشمند برچسب‌ها در یک ستون «شرح» مشترک ────────────
    master_labels, master_keys = [], []
    for items in col_items:
        for label, _v in items:
            k = _inc_label_key(label)
            if not k:
                continue
            if not any(mk == k or SequenceMatcher(None, mk, k).ratio() >= 0.9 for mk in master_keys):
                master_labels.append(label)
                master_keys.append(k)

    def _val_for(items, mk):
        for label, value in items:
            k = _inc_label_key(label)
            if k == mk or SequenceMatcher(None, k, mk).ratio() >= 0.9:
                return value
        return None

    matrix = {ci: [_val_for(items, mk) for mk in master_keys]
              for ci, items in enumerate(col_items)}

    # ردیف درآمد عملیاتی (مخرج محاسبه درصدها)
    rev_idx = None
    for _i, mk in enumerate(master_keys):
        if ("درآمدعملیاتی" in mk or "درآمدهایعملیاتی" in mk) and "هرسهم" not in mk:
            rev_idx = _i
            break
    if rev_idx is None:
        for _i, mk in enumerate(master_keys):
            if mk.startswith("درآمد") and "هرسهم" not in mk:
                rev_idx = _i
                break

    def _col_header(rep):
        dur = rep["duration_months"]
        dur_s = f"{to_fa_digits(int(dur))} ماهه" if pd.notna(dur) and dur else "؟ ماهه"
        aud_s = "ح.ش" if rep["is_audited"] == 1 else "ح.ن"
        cons_s = " تلفیقی" if rep["is_consolidated"] == 1 else ""
        return f"{dur_s} {format_period(rep['period_end'])}{cons_s} ({aud_s})"

    # ════════════════════════════════════════════════════════════
    # توابع واحد و تبدیل ارز (هماهنگ با کلید تومان/دلار نوار حاشیه‌ای)
    # ════════════════════════════════════════════════════════════
    # مقادیر خام صورت سود و زیان معمولاً «میلیون ریال» هستند (سود هر سهم: ریال).
    def _inc_to_base(value, period_end, is_per_share=False, months=None):
        """تبدیل مقدار خام به واحد پایه: میلیارد تومان / میلیون دلار (یا تومان‌هر‌سهم / دلار‌هر‌سهم).
        برای حالت دلاری، نرخ = «میانگین نرخ دلار روی `months` ماهِ منتهی به period_end»."""
        if value is None or pd.isna(value):
            return None
        if is_per_share:
            toman_ps = value / 10.0                         # ریال → تومان هر سهم
            if show_usd and has_usd:
                rate = _inc_avg_rate(period_end, months, usd_rates)
                return (toman_ps / rate) if (rate and rate > 0) else None
            return toman_ps
        btmn = value / 10000.0                               # میلیون ریال → میلیارد تومان
        if show_usd and has_usd:
            rate = _inc_avg_rate(period_end, months, usd_rates)
            return (btmn * 1000.0 / rate) if (rate and rate > 0) else None
        return btmn

    def _inc_scale(vals, usd):
        v = [abs(x) for x in vals if x is not None and not pd.isna(x) and abs(x) > 0]
        mx = max(v, default=0)
        if usd:
            return (0.001, "میلیارد دلار") if mx >= 1000 else (1.0, "میلیون دلار")
        return (0.001, "هزار میلیارد تومان") if mx >= 1000 else (1.0, "میلیارد تومان")

    def _ps_unit():
        return "دلار/سهم" if (show_usd and has_usd) else "تومان/سهم"

    cur_note = "دلار" if (show_usd and has_usd) else "تومان"
    st.caption(f"واحد نمایش: **{cur_note}** (از نوار سمت راست قابل تغییر است). "
               "نماد فعال هم از همان نوار انتخاب می‌شود.")
    if show_usd and has_usd:
        st.caption("توجه: تبدیل دلاری هر دوره با **میانگین نرخ دلارِ همان بازه** انجام می‌شود "
                   "(مثلاً دورهٔ ۶ماهه = میانگین ۶ ماهِ منتهی به پایان دوره، و هر فصل = میانگین ۳ ماهِ آن فصل).")

    # ─── کارت‌های کلیدی آخرین گزارش ──────────────────────────────
    def _find_item(items, subs, exclude=()):
        for label, value in items:
            k = _inc_label_key(label)
            if any(s in k for s in subs) and not any(e in k for e in exclude):
                return label, value
        return None, None

    latest = col_meta[0]
    latest_items = col_items[0]
    latest_pe = latest["period_end"]
    latest_months = int(latest["duration_months"]) if pd.notna(latest["duration_months"]) else None
    st.caption(f"آخرین گزارش: {_col_header(latest)}")

    KPI_DEFS = [
        ("درآمد عملیاتی", ["درآمدعملیاتی", "درآمدهایعملیاتی", "درآمدحاصلازارائه"], (), False),
        ("سود ناخالص",   ["سودناخالص"], ("هرسهم",), False),
        ("سود عملیاتی",  ["سودعملیاتی"], ("هرسهم",), False),
        ("سود خالص",     ["سودخالص", "سودزیانخالص"], ("هرسهم",), False),
        ("سود هر سهم",   ["هرسهم"], (), True),
    ]
    kpi_cols = st.columns(len(KPI_DEFS))
    for _i, (kname, ksubs, kexcl, is_eps) in enumerate(KPI_DEFS):
        _lbl, _val = _find_item(latest_items, ksubs, kexcl)
        with kpi_cols[_i]:
            base = _inc_to_base(_val, latest_pe, is_eps, months=latest_months)
            if base is None:
                st.metric(kname, "—")
            elif is_eps:
                st.metric(kname, f"{fmt_fa(base, 0)} {_ps_unit()}")
            else:
                sc, ul = _inc_scale([base], show_usd and has_usd)
                st.metric(kname, f"{fmt_fa(base * sc, 1)}", help=ul)
                st.caption(ul)

    st.markdown("---")

    # ════════════════════════════════════════════════════════════
    # نمودار روند یک قلم — با مقایسه سال‌به‌سال فصلی (مثل صفحه ماهیانه)
    # ════════════════════════════════════════════════════════════
    st.markdown("#### 📈 نمودار روند / مقایسه فصلی")

    def _default_metric_idx():
        for subs in (["سودخالص", "سودزیانخالص"], ["درآمدعملیاتی", "درآمدهایعملیاتی"]):
            for i, mk in enumerate(master_keys):
                if any(s in mk for s in subs) and "هرسهم" not in mk:
                    return i
        return 0

    cm1, cm2 = st.columns([3, 3])
    with cm1:
        metric_label = st.selectbox("قلم:", master_labels, index=_default_metric_idx(), key="inc_metric")
    metric_idx = master_labels.index(metric_label)
    metric_key = master_keys[metric_idx]
    is_per_share = "هرسهم" in metric_key
    with cm2:
        trend_mode = st.radio(
            "نوع نمایش:",
            ["مقایسه‌ای سال‌به‌سال (فصلی)", "روند زمانی (فصلی)", "تجمعی (طبق گزارش)"],
            key="inc_trend_mode",
        )

    Q_LABELS = {1: "فصل اول", 2: "فصل دوم", 3: "فصل سوم", 4: "فصل چهارم"}
    Q_ORDER = ["فصل اول", "فصل دوم", "فصل سوم", "فصل چهارم"]
    YEAR_COLORS_INC = ["#10B981", "#4DA3FF", "#A78BFA", "#F59E0B", "#EF4444", "#EC4899"]

    def _quarter_points(row_idx):
        """ارقام فصلی (تفاضلی) یک ردیف: لیست dict با fy(سال مالی)، q(شماره فصل)، pe، val(خام)."""
        fy_map = {}
        for ci, rep in enumerate(col_meta):
            d = rep["duration_months"]
            if pd.isna(d) or not d:
                continue
            v = matrix[ci][row_idx]
            if v is None:
                continue
            fyk = _inc_fy_key(rep["period_end"], d)
            if fyk is None:
                continue
            fy_map.setdefault(fyk, {})[int(d)] = (rep["period_end"], v)
        out = []
        for fyk, dmap in fy_map.items():
            fy_year = (fyk + 12 - 1) // 12          # سالِ پایانِ سال مالی
            for d, (pe, v) in dmap.items():
                prev = dmap.get(d - 3)
                qv = (v - prev[1]) if prev else v
                q = max(1, min(4, int(round(d / 3))))
                out.append({"fy": fy_year, "q": q, "pe": pe, "val": qv})
        return out

    fig_inc = go.Figure()

    if trend_mode == "تجمعی (طبق گزارش)":
        by_dur = {}
        for ci, rep in enumerate(col_meta):
            d = rep["duration_months"]
            if pd.isna(d):
                continue
            v = matrix[ci][metric_idx]
            if v is None:
                continue
            by_dur.setdefault(int(d), []).append((rep["period_end"], v))
        # واحد و scale یکدست
        all_base = []
        for d, lst in by_dur.items():
            for pe, v in lst:
                b = _inc_to_base(v, pe, is_per_share, months=d)
                if b is not None:
                    all_base.append(b)
        if is_per_share:
            sc, unit_lbl = 1.0, _ps_unit()
        else:
            sc, unit_lbl = _inc_scale(all_base, show_usd and has_usd)
        DUR_LABELS = {3: "۳ ماهه", 6: "۶ ماهه", 9: "۹ ماهه", 12: "۱۲ ماهه"}
        for di, d in enumerate(sorted(by_dur.keys())):
            pts = sorted(by_dur[d], key=lambda t: normalize_digits(t[0]))
            xs = [format_period(p) for p, _ in pts]
            ys = [(_inc_to_base(v, p, is_per_share, months=d) or 0) * sc for p, v in pts]
            fig_inc.add_trace(go.Bar(
                x=xs, y=ys, name=DUR_LABELS.get(d, f"{to_fa_digits(d)} ماهه"),
                marker_color=YEAR_COLORS_INC[di % len(YEAR_COLORS_INC)],
                hovertemplate="<b>%{x}</b><br>%{y:,.2f} " + unit_lbl + "<extra></extra>",
            ))
        fig_inc.update_layout(barmode="group", xaxis=dict(autorange="reversed", tickfont=PLOTLY_FONT))
        _cap = "ارقام «تجمعی» از ابتدای سال مالی؛ هر مدت دوره جداگانه نمایش داده می‌شود."

    else:
        pts = _quarter_points(metric_idx)
        base_pts = [(p["fy"], p["q"], p["pe"], _inc_to_base(p["val"], p["pe"], is_per_share, months=3)) for p in pts]
        if is_per_share:
            sc, unit_lbl = 1.0, _ps_unit()
        else:
            sc, unit_lbl = _inc_scale([b for *_, b in base_pts], show_usd and has_usd)

        if trend_mode == "مقایسه‌ای سال‌به‌سال (فصلی)":
            years = sorted({fy for fy, *_ in base_pts})
            for yi, yr in enumerate(years):
                yv = {1: None, 2: None, 3: None, 4: None}
                for fy, q, pe, b in base_pts:
                    if fy == yr and b is not None:
                        yv[q] = b * sc
                fig_inc.add_trace(go.Bar(
                    x=Q_ORDER, y=[yv[1], yv[2], yv[3], yv[4]],
                    name=to_fa_digits(str(yr)),
                    marker_color=YEAR_COLORS_INC[yi % len(YEAR_COLORS_INC)],
                    hovertemplate="<b>" + to_fa_digits(str(yr)) + "</b> | %{x}<br>%{y:,.2f} " + unit_lbl + "<extra></extra>",
                ))
            fig_inc.update_layout(
                barmode="group",
                xaxis=dict(categoryorder="array", categoryarray=Q_ORDER, tickfont=PLOTLY_FONT),
            )
            _cap = "هر فصل = تفاضل دوره‌های تجمعی هم‌سال‌مالی (۶ماهه − ۳ماهه و …)؛ مقایسه سال‌به‌سال."
        else:  # روند زمانی فصلی
            ser = sorted(base_pts, key=lambda t: normalize_digits(t[2]))
            xs = [f"{Q_LABELS[q]} {to_fa_digits(str(fy))}" for fy, q, pe, b in ser]
            ys = [(b or 0) * sc for fy, q, pe, b in ser]
            fig_inc.add_trace(go.Bar(
                x=xs, y=ys, name=metric_label, marker_color="#4DA3FF",
                hovertemplate="<b>%{x}</b><br>%{y:,.2f} " + unit_lbl + "<extra></extra>",
            ))
            fig_inc.update_layout(xaxis=dict(tickfont=PLOTLY_FONT))
            _cap = "ارقام فصلی (تفاضلی) به ترتیب زمان."

    fig_inc.update_layout(
        title=dict(text=f"{inc_sym} — {metric_label} ({unit_lbl})", font=PLOTLY_FONT_TITLE),
        yaxis=dict(title=dict(text=unit_lbl, font=PLOTLY_FONT), tickfont=PLOTLY_FONT),
        font=PLOTLY_FONT, template="plotly_dark", height=440,
        hoverlabel=dict(font=PLOTLY_FONT), legend=dict(font=PLOTLY_FONT),
        bargap=0.15, bargroupgap=0.05,
    )
    st.plotly_chart(fig_inc, use_container_width=True)
    st.caption(_cap)

    st.markdown("---")

    # ════════════════════════════════════════════════════════════
    # جدول صورت سود و زیان + درصد از درآمد عملیاتی
    # ════════════════════════════════════════════════════════════
    st.markdown("#### 🧾 جدول صورت سود و زیان دوره‌ای + درصد از درآمد")

    c1, c2 = st.columns(2)
    with c1:
        view_mode = st.radio("نمای جدول:", ["خلاصه (٪ سود ناخالص و خالص)", "کامل"],
                             index=0, key="inc_view_mode")
    with c2:
        tbl_kind = st.radio("اعداد جدول:", ["تجمعی (دوره‌ای)", "فصلی (تفاضلی)"],
                            index=1, key="inc_tbl_kind")
    if view_mode == "کامل":
        c3, c4 = st.columns(2)
        show_pct_rows = c3.toggle("ردیف ٪ زیر هر قلم", value=True, key="inc_pct")
        hide_small = c4.toggle("پنهان‌کردن سطرهای کم‌اهمیت (همه < ۱٪ درآمد)", value=True, key="inc_hide_small")
    else:
        show_pct_rows, hide_small = False, False

    _, gross_idx, netc_idx = _stmt_key_indices(master_keys)

    def _fmt_amount(base_val, per_share=False):
        if base_val is None or pd.isna(base_val):
            return "—"
        dec = 0 if abs(base_val) >= 100 else (1 if abs(base_val) >= 1 else 2)
        s = f"{abs(base_val):,.{dec}f}".translate(EN_TO_FA)
        return f"({s})" if base_val < 0 else s

    def _fmt_pct(v):
        if v is None or pd.isna(v):
            return "—"
        s = f"{abs(v):,.1f}".translate(EN_TO_FA)
        return f"-{s}٪" if v < 0 else f"{s}٪"

    def _to_base(value, period_end, months):
        return _inc_to_base(value, period_end, False, months=months)

    # ─── ساخت ستون‌ها (جدیدتر سمت راست = اول لیست) با vals هم‌ترازِ master ──
    columns = []
    if tbl_kind == "تجمعی (دوره‌ای)":
        for ci, rep in enumerate(col_meta):       # col_meta از قبل: جدیدتر اول
            mo = int(rep["duration_months"]) if pd.notna(rep["duration_months"]) else None
            columns.append({
                "header": _col_header(rep), "pe": rep["period_end"], "months": mo,
                "rev": (matrix[ci][rev_idx] if rev_idx is not None else None),
                "vals": matrix[ci],
            })
    else:
        if rev_idx is None:
            st.warning("ردیف «درآمد عملیاتی» برای حالت فصلی/درصد پیدا نشد.")
        else:
            _qcache = {}
            def _qmap(ri):
                if ri not in _qcache:
                    m = {}
                    for p in _quarter_points(ri):
                        m[(p["fy"], p["q"])] = (p["pe"], p["val"])
                    _qcache[ri] = m
                return _qcache[ri]
            rev_map = _qmap(rev_idx)
            for (fy, q) in sorted(rev_map.keys(), reverse=True):   # جدیدتر سمت راست (RTL)
                pe = rev_map[(fy, q)][0]
                vals = []
                for ri in range(len(master_keys)):
                    t = _qmap(ri).get((fy, q))
                    vals.append(t[1] if t else None)
                columns.append({
                    "header": f"{Q_LABELS[q]} {to_fa_digits(str(fy))}",
                    "pe": pe, "months": 3, "rev": rev_map[(fy, q)][1], "vals": vals,
                })

    if not columns:
        if tbl_kind != "تجمعی (دوره‌ای)" and rev_idx is not None:
            st.warning("داده فصلی کافی برای ساخت جدول نیست.")
    else:
        # جدیدترین فصلِ واقعی برای محاسبهٔ دورهٔ بعدی
        _lat_y, _lat_q = None, None
        if tbl_kind != "تجمعی (دوره‌ای)" and columns:
            _hdr0 = columns[0]["header"]   # مثل «فصل اول ۱۴۰۴»
            _ym = re.search(r'(\d{4})', normalize_digits(_hdr0))
            _lat_y = int(_ym.group(1)) if _ym else None
            for _qn, _ql in {1: "اول", 2: "دوم", 3: "سوم", 4: "چهارم"}.items():
                if _ql in _hdr0:
                    _lat_q = _qn
                    break
        else:
            _y0, _m0 = _inc_year_month(col_meta[0]["period_end"]) if col_meta else (None, None)
            _lat_y, _lat_q = _y0, (4 if (_m0 and _m0 >= 12) else None)
        render_statement_editable(
            inc_sym, "فصلی", master_labels, master_keys, columns,
            rev_idx, gross_idx, netc_idx, _to_base,
            _fmt_amount, _fmt_pct, view_mode, show_pct_rows, hide_small,
            latest_year=_lat_y, latest_q=_lat_q,
            quarter_labels={1: "فصل اول", 2: "فصل دوم", 3: "فصل سوم", 4: "فصل چهارم"})

        # ── ذخیرهٔ ارقامِ فصلیِ واقعی (تفاضلی) به تفکیکِ (سال، فصل) برای ساختِ سالانه ──
        # {سال: {فصل: {برچسب: مقدارِ خام میلیون‌ریال}}}
        _q_actual = {}
        for _ri in range(len(master_labels)):
            for _pt in _quarter_points(_ri):
                _q_actual.setdefault(_pt["fy"], {}).setdefault(_pt["q"], {})[master_labels[_ri]] = _pt["val"]
        st.session_state[f"q_actual_{inc_sym}"] = _q_actual
        st.session_state[f"q_labels_{inc_sym}"] = list(master_labels)

    # ════════════════════════════════════════════════════════════
    # جدول گزارش‌های سالانه (codal_annual.db) — جدا، زیر جدول دوره‌ای
    # ════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("### 📅 صورت سود و زیان سالانه")

    # گزارش‌های سالانه + جایگزینِ ۱۲ماههٔ codal_income.db برای سال‌های ۱۴۰۲ به بعد
    reps_a, items_a, _fb_years = annual_reports_with_income_fallback(inc_sym)

    if reps_a.empty:
        if not has_annual and not has_income:
            st.info(
                "دیتابیس گزارش‌های سالانه پیدا نشد.\n\n"
                f"مسیر بررسی‌شده: `{ANNUAL_DB_PATH}`\n\n"
                "ابتدا `codal_annual_extractor.py` را در فولدر `D:\\bourse\\codal seasonal` اجرا کنید."
            )
        else:
            st.info(f"گزارش سالانه یا ۱۲ماهه‌ای برای نماد «{inc_sym}» یافت نشد.")
    else:
        reps_a = reps_a.sort_values(["period_end", "duration_months"], ascending=[False, False])
        ml_a, mk_a, cmeta_a, mtx_a = _inc_build_matrix(reps_a, items_a)

        if not ml_a:
            st.info("اقلامی برای گزارش سالانهٔ این نماد ثبت نشده.")
        else:
            rev_a, gross_a, netc_a = _stmt_key_indices(mk_a)

            ca1, ca2 = st.columns(2)
            with ca1:
                view_mode_a = st.radio("نمای جدول سالانه:", ["کامل", "خلاصه (٪ سود ناخالص و خالص)"],
                                       index=0, key="ann_view_mode")
            view_mode_a = "خلاصه" if view_mode_a.startswith("خلاصه") else "کامل"
            if view_mode_a == "کامل":
                with ca2:
                    cc1, cc2 = st.columns(2)
                    show_pct_a = cc1.toggle("ردیف ٪ زیر هر قلم", value=True, key="ann_pct")
                    hide_small_a = cc2.toggle("پنهان‌کردن سطرهای < ۱٪", value=True, key="ann_hide_small")
            else:
                show_pct_a, hide_small_a = False, False

            # نرخ دلار سالانه (تومان) بر اساس سالِ پایان دوره
            def _annual_to_base(value, period_end, months=None):
                if value is None or pd.isna(value):
                    return None
                btmn = value / 10000.0                       # میلیون ریال → میلیارد تومان
                if show_usd and has_annual_usd:
                    yy, _mm = _inc_year_month(period_end)
                    rate = annual_usd_rates.get(yy)          # تومان به ازای هر دلار
                    return (btmn * 1000.0 / rate) if (rate and rate > 0) else None
                return btmn

            def _ann_header(rep):
                yy, _mm = _inc_year_month(rep["period_end"])
                aud = "ح.ش" if rep["is_audited"] == 1 else "ح.ن"
                cons = " ت" if rep["is_consolidated"] == 1 else ""
                star = "٭" if (yy in _fb_years) else ""   # ٭ = از گزارشِ ۱۲ماههٔ دوره‌ای
                return f"{to_fa_digits(yy)}{cons} ({aud}){star}"

            columns_a = []
            for ci, rep in enumerate(cmeta_a):              # cmeta_a: جدیدتر اول → راست
                columns_a.append({
                    "header": _ann_header(rep), "pe": rep["period_end"],
                    "months": 12, "rev": (mtx_a[ci][rev_a] if rev_a is not None else None),
                    "vals": mtx_a[ci],
                })

            _ua = ("میلیون دلار" if (show_usd and has_annual_usd) else "میلیارد تومان")
            if show_usd and not has_annual_usd:
                st.warning(f"فایل نرخ دلار سالانه پیدا نشد (`{ANNUAL_USD_RATE_PATH}`)؛ مقادیر به تومان نمایش داده می‌شوند.")
            st.caption(f"ستون‌ها = سال مالی (جدیدتر سمت راست). ارقام به **{_ua}**؛ "
                       "تبدیل دلاری با نرخ سالانهٔ همان سال (به تومان) انجام می‌شود.")
            if _fb_years:
                _fb_txt = "، ".join(to_fa_digits(y) for y in sorted(_fb_years))
                st.caption(f"٭ سال‌های {_fb_txt} در دیتابیس سالانه نبودند و از گزارشِ "
                           "**۱۲ماههٔ دوره‌ای** (`codal_income.db`) برداشته شده‌اند.")

            _lat_ya, _ = _inc_year_month(cmeta_a[0]["period_end"]) if cmeta_a else (None, None)

            # ── ساخت خودکارِ سالِ پیش‌بینی از جمعِ چهار فصل (واقعی + پیش‌بینیِ فصلی) ──
            with st.expander("🧮 ساخت سال از روی فصل‌ها (واقعی + پیش‌بینیِ فصلی)"):
                st.caption("هر قلمِ سالانه = جمعِ همان قلم در چهار فصلِ سال. فصل‌هایی که گزارش واقعی "
                           "دارند از داده‌های واقعی، و فصل‌هایی که پیش‌بینی کرده‌ای از پیش‌بینیِ فصلی برداشته می‌شوند.")
                _q_act = st.session_state.get(f"q_actual_{inc_sym}", {})
                _q_labs = st.session_state.get(f"q_labels_{inc_sym}", ml_a)
                _years_av = sorted(_q_act.keys()) if _q_act else []
                _by1, _by2 = st.columns([1, 2])
                with _by1:
                    _build_year = st.number_input(
                        "سال مالی هدف:", min_value=1390, max_value=1430,
                        value=int((_years_av[-1] + 1) if _years_av else ((_lat_ya or 1403) + 1)),
                        step=1, key=f"ann_build_year_{inc_sym}")
                _build_year = int(_build_year)

                # پیش‌نمایش: منبعِ هر فصل (پیش‌بینی اولویت دارد؛ وگرنه واقعی)
                _src = {}
                for _q in (1, 2, 3, 4):
                    _sdf, _ = load_forecast(inc_sym, f"فصلی-{_build_year}-Q{_q}")
                    if _sdf is not None and not _sdf.empty and _sdf["value"].notna().any():
                        _src[_q] = "پیش‌بینی"
                    elif _build_year in _q_act and _q in _q_act[_build_year]:
                        _src[_q] = "واقعی"
                    else:
                        _src[_q] = "—"
                st.caption("منبع هر فصل:  " + "   ".join(
                    f"فصل {to_fa_digits(q)}: {_src[q]}" for q in (1, 2, 3, 4)))

                # ── تشخیص: چه پیش‌بینی‌هایی در DB ذخیره شده‌اند؟ ──
                _allfc = list_forecast_symbols()
                _mine = []
                if _allfc is not None and not _allfc.empty:
                    _mine = [str(r["fiscal_year"]) for _, r in _allfc.iterrows()
                             if r["symbol"] == inc_sym]
                with st.expander("🔎 بررسی فنی (اگر پیش‌بینی در جمع نمی‌آید این را باز کن)"):
                    st.write("کلیدهای پیش‌بینیِ ذخیره‌شده برای این نماد:",
                             _mine if _mine else "(هیچ پیش‌بینی‌ای ذخیره نشده)")
                    st.write(f"بخشِ سالانه دنبالِ این کلیدها می‌گردد: "
                             f"`فصلی-{_build_year}-Q1..Q4`")
                    for _q in (1, 2, 3, 4):
                        _sd, _ = load_forecast(inc_sym, f"فصلی-{_build_year}-Q{_q}")
                        if _sd is not None and not _sd.empty:
                            _rv = next((v for l, v in zip(_sd["label"], _sd["value"])
                                        if "درآمد" in _inc_label_key(l) and "عملیاتی" in _inc_label_key(l)), None)
                            st.write(f"  فصل {to_fa_digits(_q)} (`فصلی-{_build_year}-Q{_q}`): "
                                     f"{len(_sd)} قلم، درآمد={_rv}")
                        else:
                            st.write(f"  فصل {to_fa_digits(_q)} (`فصلی-{_build_year}-Q{_q}`): خالی")
                    _qa = _q_act.get(_build_year, {})
                    st.write(f"فصل‌های واقعیِ موجود برای سال {to_fa_digits(_build_year)}: "
                             f"{sorted(_qa.keys()) if _qa else '(هیچ)'}")

                if st.button("🧮 محاسبه و ذخیره به‌عنوان پیش‌بینیِ سالانه",
                             key=f"ann_build_btn_{inc_sym}", type="primary"):
                    # ۱) مقادیرِ هر فصل را جمع‌آوری کن (همه به میلیارد تومان)، روی اتحادِ همهٔ برچسب‌ها
                    q_vals = {1: {}, 2: {}, 3: {}, 4: {}}
                    q_src = {}
                    for _q in (1, 2, 3, 4):
                        _sdf, _ = load_forecast(inc_sym, f"فصلی-{_build_year}-Q{_q}")
                        has_fc = (_sdf is not None and not _sdf.empty and
                                  _sdf["value"].notna().any())
                        if has_fc:
                            # این فصل پیش‌بینی دارد → پیش‌بینی اولویت دارد (میلیارد تومان)
                            for lab, v in zip(_sdf["label"], _sdf["value"]):
                                if v is not None and not pd.isna(v):
                                    q_vals[_q][lab] = float(v)
                            q_src[_q] = "پیش‌بینی"
                        elif _build_year in _q_act and _q in _q_act[_build_year]:
                            # وگرنه واقعی (خام میلیون‌ریال → میلیارد تومان)
                            for lab, raw in _q_act[_build_year][_q].items():
                                if raw is not None and not pd.isna(raw):
                                    q_vals[_q][lab] = raw / 10000.0
                            q_src[_q] = "واقعی"
                        else:
                            q_src[_q] = "—"

                    # ۲) اتحادِ برچسب‌ها و جمعِ چهار فصل
                    all_labels = []
                    for _q in (1, 2, 3, 4):
                        for lab in q_vals[_q]:
                            if lab not in all_labels:
                                all_labels.append(lab)
                    summed = {}
                    for lab in all_labels:
                        tot, any_q = 0.0, False
                        for _q in (1, 2, 3, 4):
                            if lab in q_vals[_q]:
                                tot += q_vals[_q][lab]; any_q = True
                        summed[lab] = tot if any_q else None

                    # ۳) نگاشتِ برچسب‌های فصلی → برچسب‌های جدولِ سالانه (بر اساس کلیدِ نرمال)
                    ann_by_key = {_inc_label_key(l): l for l in ml_a}
                    out_map = {}
                    for lab, val in summed.items():
                        ann_lab = ann_by_key.get(_inc_label_key(lab), lab)  # اگر نخورد، همان برچسب فصلی
                        out_map[ann_lab] = val
                    # خروجی را بر مبنای ترتیبِ اقلام سالانه بساز (و اقلامِ اضافیِ فصلی را هم نگه‌دار)
                    ordered = [l for l in ml_a] + [l for l in out_map if l not in ml_a]
                    out = [(l, out_map.get(l)) for l in ordered]

                    save_forecast(inc_sym, f"سالانه-{_build_year}", out, "ساخته‌شده از جمع فصل‌ها")
                    _sk_a = f"fc_cols_{inc_sym}_سالانه"
                    _cur = st.session_state.get(_sk_a, [])
                    _key = f"سالانه-{_build_year}"
                    if _key not in _cur:
                        st.session_state[_sk_a] = sorted(_cur + [_key])
                    # خلاصهٔ درآمدِ هر فصل برای کنترل
                    _rev_dbg = []
                    for _q in (1, 2, 3, 4):
                        _rv = next((v for l, v in q_vals[_q].items()
                                    if "درآمد" in _inc_label_key(l) and "عملیاتی" in _inc_label_key(l)), None)
                        _rev_dbg.append(f"ف{to_fa_digits(_q)}: {fmt_fa(_rv,0) if _rv else '—'}")
                    st.success(f"سال {to_fa_digits(_build_year)} ساخته شد. درآمدِ فصل‌ها — " +
                               " | ".join(_rev_dbg))
                    st.rerun()

            render_statement_editable(
                inc_sym, "سالانه", ml_a, mk_a, columns_a,
                rev_a, gross_a, netc_a, _annual_to_base,
                _fmt_amount, _fmt_pct, view_mode_a, show_pct_a, hide_small_a,
                latest_year=_lat_ya, latest_q=None)


elif page == "🔍 فیلتر نمادها":
    st.title("فیلتر و رتبه‌بندی نمادها")
    # بررسی query param از کلیک جدول
    _qp = st.query_params.get("sym", "")
    if _qp and _qp in _all_syms_sidebar:
        st.query_params.clear()
        # widget key رو نمیشه مستقیم ست کرد
        # اما index رو میشه از طریق default value کنترل کرد
        # این کار میکنه چون قبل از render sidebar هستیم — نه
        # تنها راه: از selectbox بدون key استفاده کن در sidebar
        st.session_state["_pending_sym"] = _qp
        st.rerun()

    if df_main.empty:
        st.warning("داده‌ای موجود نیست.")
        st.stop()

    all_syms = sorted(df_main["symbol"].unique())

    # ─── نقشهٔ ارزش بازار (میلیون دلار) از فایل روزانهٔ TSETMC ──
    def _norm_sym2(s):
        return normalize_symbol(s)

    mcap_musd_map = {}
    if (not _daily_df.empty and "نام" in _daily_df.columns
            and "ارزش بازار" in _daily_df.columns and usd_rates):
        _last_rate_f = usd_rates[max(usd_rates)]
        for _, _rr in _daily_df.iterrows():
            try:
                _mv = float(str(_rr["ارزش بازار"]).replace(",", "").strip())
            except Exception:
                continue
            if _mv > 0 and _last_rate_f and _last_rate_f > 0:
                mcap_musd_map[_norm_sym2(_rr["نام"])] = (_mv / 1e10) * 1000.0 / _last_rate_f

    # ─── کمک‌تابع: مبلغ تولید یک ماه (میلیون دلار) با نرخِ «آخرین ماهِ پرفروش» ──
    #   برای هر محصول (product_name+category): آستانه = ۲۰٪ تولیدِ همان ماهِ هدف؛
    #   نرخِ مرجع = نرخِ دلاریِ آخرین ماهی (در کلِ تاریخچه) که فروشش ≥ آستانه بوده؛
    #   اگر هیچ ماهی به آستانه نرسد → میانگینِ نرخِ دلاریِ آن محصول.
    _prod_by_sym = {s: g for s, g in df_prod.groupby("symbol")} if not df_prod.empty else {}

    def _prod_value_lastsale(sym_prod_df, target_period):
        if sym_prod_df is None or sym_prod_df.empty or not target_period:
            return None
        total = 0.0
        has_any = False
        for (_pn, _cat), _g in sym_prod_df.groupby(["product_name", "category"]):
            _tp = _g[_g["period_end"] == target_period]
            if _tp.empty:
                continue
            _pq = _tp["production_qty"].sum()
            if _pq is None or pd.isna(_pq) or _pq <= 0:
                continue
            _thr = 0.20 * _pq
            _cands = []      # (period, نرخ دلاری) ماه‌هایی که ≥۲۰٪ فروخته‌اند
            _all_dr = []     # برای میانگینِ بازگشتی
            for _, _r in _g.iterrows():
                _rr = _r["sales_rate_rial"]
                _ur = usd_rates.get(_ym_key(_r["period_end"]))
                _dr = (_rr / (_ur * 10.0)) if (_rr and _rr > 0 and not pd.isna(_rr)
                                               and _ur and _ur > 0) else None
                if _dr is not None:
                    _all_dr.append(_dr)
                _sq = _r["sales_qty"]
                if _sq is not None and not pd.isna(_sq) and _sq >= _thr and _dr is not None:
                    _cands.append((_r["period_end"], _dr))
            if _cands:
                _ref = max(_cands, key=lambda x: x[0])[1]     # آخرین ماهِ واجدِ شرط
            elif _all_dr:
                _ref = sum(_all_dr) / len(_all_dr)
            else:
                _ref = None
            if _ref is None:
                continue
            has_any = True
            total += _pq * _ref / 1_000_000.0
        return total if has_any else None

    # ─── انتخابِ شاخص‌ها برای محاسبه (هرکدام جدا) ─────────────────────────
    # هر شاخص محاسبهٔ جداگانه و نسبتاً سنگینی دارد؛ به‌صورت پیش‌فرض فقط «شاخص a»
    # محاسبه می‌شود تا پنجرهٔ فیلتر سریع باز شود. بقیه را با تیک‌زدن اضافه کن.
    st.markdown("**شاخص‌های موردنظر برای محاسبه:**")
    _ic1, _ic2, _ic3, _ic4 = st.columns(4)
    with _ic1:
        want_a = st.checkbox("شاخص a", value=True,  key="idx_want_a")
    with _ic2:
        want_b = st.checkbox("شاخص b", value=False, key="idx_want_b")
    with _ic3:
        want_c = st.checkbox("شاخص c", value=False, key="idx_want_c")
    with _ic4:
        want_d = st.checkbox("شاخص d", value=False, key="idx_want_d")
    _need_ab = want_a or want_b
    _need_cd = want_c or want_d
    if not (want_a or want_b or want_c or want_d):
        st.caption("هیچ شاخصی انتخاب نشده — جدول بدونِ ستونِ شاخص نمایش داده می‌شود.")
    else:
        st.caption("هر شاخص جداگانه محاسبه می‌شود؛ هرچه کمتر تیک بزنی، پنجره سریع‌تر باز می‌شود.")

    # ─── محاسبه معیارهای هر نماد ──────────────────────────────
    rows_filter = []
    for sym in all_syms:
        sym_df = df_main[df_main["symbol"] == sym].sort_values("period_end", ascending=False)
        if sym_df.empty:
            continue

        all_p = sym_df["period_end"].tolist()

        # آخرین ماه
        last_val = sym_df.iloc[0]["total_billion_toman"] if not sym_df.empty else None
        last_period = sym_df.iloc[0]["period_end"] if not sym_df.empty else ""

        # میانگین ۱۲ ماه گذشته (بدون ماه آخر)
        past_12 = sym_df.iloc[1:13]["total_billion_toman"].dropna()
        avg_12 = past_12.mean() if not past_12.empty else None

        # میانگین ۳ ماه آخر (شامل ماه آخر)
        last_3 = sym_df.iloc[0:3]["total_billion_toman"].dropna()
        avg_3 = last_3.mean() if not last_3.empty else None

        # نسبت: ماه آخر / میانگین ۱۲ ماه
        ratio_last_12 = (last_val / avg_12) if (last_val and avg_12 and avg_12 > 0) else None

        # نسبت: میانگین ۳ ماه / میانگین ۱۲ ماه
        ratio_3_12 = (avg_3 / avg_12) if (avg_3 and avg_12 and avg_12 > 0) else None

        # دلاری اگه داره
        if has_usd:
            last_musd  = btmn_to_musd(pd.Series([last_val or 0]), pd.Series([last_period]), usd_rates)[0]
            avg12_musd = (btmn_to_musd(sym_df.iloc[1:13]["total_billion_toman"].fillna(0),
                                        sym_df.iloc[1:13]["period_end"], usd_rates))
            avg12_musd_val = float(pd.Series(avg12_musd).mean()) if avg12_musd else None
        else:
            last_musd = None; avg12_musd_val = None

        # ─── شاخص a ────────────────────────────────────────────
        # A = (درآمد دلاری ماه آخر ÷ درآمد دلاری مشابه پارسال)
        #     × میانگین درآمد دلاری ۱۲ماههٔ ۵ سال اخیر
        #     × میانگین حاشیه سود ۱۲ماههٔ ۵ سال اخیر
        #     ÷ ارزش بازار امروز (م.دلار)
        #     × میانگین نسبت تقسیم سود (dps÷eps)
        idx_a = None
        idx_b = None
        if has_usd and _need_ab:
            # درآمد دلاری مشابه پارسال (همان ماه، سالِ قبل)
            _ly, _lm = _inc_year_month(last_period)
            r_yoy_musd = None
            if _ly and _lm:
                _prev = sym_df[sym_df["period_end"].apply(
                    lambda p: _inc_year_month(p) == (_ly - 1, _lm))]
                if not _prev.empty:
                    r_yoy_musd = btmn_to_musd(pd.Series([_prev.iloc[0]["total_billion_toman"]]),
                                              pd.Series([_prev.iloc[0]["period_end"]]), usd_rates)[0]

            # میانگین درآمد دلاری ۱۲ماههٔ ۵ سال اخیر (میانگین ماهانهٔ ۶۰ ماه × ۱۲)
            _rev60 = sym_df.iloc[0:60]
            _m60 = btmn_to_musd(_rev60["total_billion_toman"], _rev60["period_end"], usd_rates)
            _m60 = [v for v in _m60 if v is not None]
            avg_ann_musd_5y = (sum(_m60) / len(_m60)) * 12 if _m60 else None

            _margin = _avg_net_margin(sym)          # کسر
            _payout = _avg_payout_ratio(sym)         # کسر
            _mcap   = mcap_musd_map.get(_norm_sym2(sym))   # میلیون دلار

            if (want_a and last_musd and r_yoy_musd and r_yoy_musd > 0 and avg_ann_musd_5y is not None
                    and _margin is not None and _mcap and _mcap > 0 and _payout is not None):
                idx_a = ((last_musd / r_yoy_musd) * avg_ann_musd_5y * _margin
                         / _mcap * _payout)

            # ─── شاخص b ────────────────────────────────────────────
            # دقیقاً مثل a، اما فاکتورِ اولِ «رشد درآمد دلاری ماه آخر نسبت به پارسال»
            # با «رشد مبلغ تولید (به نرخِ آخرین ماهِ پرفروش) ماه آخر نسبت به مشابه پارسال»
            # جایگزین می‌شود. سه فاکتورِ بعدی (میانگین سود دلاری ۵ سال، حاشیه سود،
            # بازده/تقسیم‌سود نسبت به ارزش بازار) همان a است.
            _spb = _prod_by_sym.get(sym)
            if want_b and _spb is not None and not _spb.empty:
                _pb_last = max(_spb["period_end"])
                _pby, _pbm = _inc_year_month(_pb_last)
                _pv_last = _prod_value_lastsale(_spb, _pb_last)
                _pv_yoy  = None
                if _pby and _pbm:
                    _prev_ps = [p for p in _spb["period_end"].unique()
                                if _inc_year_month(p) == (_pby - 1, _pbm)]
                    if _prev_ps:
                        _pv_yoy = _prod_value_lastsale(_spb, _prev_ps[0])
                if (_pv_last is not None and _pv_yoy and _pv_yoy > 0
                        and avg_ann_musd_5y is not None and _margin is not None
                        and _mcap and _mcap > 0 and _payout is not None):
                    idx_b = ((_pv_last / _pv_yoy) * avg_ann_musd_5y * _margin
                             / _mcap * _payout)

        # ─── شاخص c و d ────────────────────────────────────────
        # c: میانگینِ ۲ سود خالصِ دلاریِ برتر (۵ سال اخیر) ÷ ارزش بازار فعلی.
        # d: مثل c، اما از مخرج «دیویدندِ دلاریِ امسال» (سود سال DIVIDEND_FY ×
        #    میانگین تقسیم سود) کم می‌شود. منبعِ سود: گزارش → پیش‌بینی → سال قبل.
        idx_c = None
        idx_d = None
        _net_top2 = None
        _div_val = 0.0
        _d_src = None
        _d_fallback = False
        try:
            _net_top2 = (_top_net_income_musd(sym, years=5, top_n=2)   # میلیون دلار
                         if _need_cd else None)
            _mcap_c = mcap_musd_map.get(_norm_sym2(sym))             # میلیون دلار
            if _net_top2 is not None and _mcap_c and _mcap_c > 0:
                if want_c:
                    idx_c = _net_top2 / _mcap_c
                if want_d:
                    _div_musd, _d_src = _current_year_dividend_musd(sym)
                    _div_val = _div_musd if (_div_musd is not None) else 0.0
                    _denom_d = _mcap_c - _div_val
                    if _denom_d > 0:
                        idx_d = _net_top2 / _denom_d
                        # علامت‌گذاری: سودِ سال DIVIDEND_FY از گزارش واقعی ۱۲ماهه نیست
                        # (پیش‌بینی، سود سال قبل، یا نامشخص) ⇒ گزارش سالانهٔ ۱۴۰۴ ندارد.
                        _d_fallback = (_d_src != "گزارش ۱۲ماهه")
        except Exception:
            idx_d = None

        rows_filter.append({
            "نماد":          sym,
            "آخرین دوره":    format_period(last_period),
            "last_btmn":     last_val,
            "avg12_btmn":    avg_12,
            "avg3_btmn":     avg_3,
            "ratio_last_12": ratio_last_12,
            "ratio_3_12":    ratio_3_12,
            "last_musd":     last_musd,
            "avg12_musd":    avg12_musd_val,
            "idx_a":         idx_a,
            "idx_b":         idx_b,
            "idx_c":         idx_c,
            "idx_d":         idx_d,
            "net_top2_musd": _net_top2,
            "div_musd":      _div_val,
            "d_fallback":    _d_fallback,
            "d_src":         (_d_src or ""),
        })

    df_filter = pd.DataFrame(rows_filter)
    if df_filter.empty:
        st.warning("داده‌ای موجود نیست.")
        st.stop()

    # مرتب‌سازی پیش‌فرض: نسبت ماه آخر به میانگین ۱۲ماه نزولی
    df_sorted = df_filter.sort_values("ratio_last_12", ascending=False, na_position="last")

    # ─── ساخت جدول نمایش ──────────────────────────────────────
    # فیلتر نمادها: همیشه م.تومان (scale=1)
    _sc_last,  _lb_last  = 1.0, "م.تومان"
    _sc_avg12, _lb_avg12 = 1.0, "م.تومان"
    _sc_avg3,  _lb_avg3  = 1.0, "م.تومان"

    def _fmt_ratio(v):
        if v is None or pd.isna(v): return "—"
        return f"{v:.2f}×"

    rows_disp = []
    for _, _r in df_sorted.iterrows():
        _lv = (_r["last_musd"] if show_usd and _r["last_musd"] is not None else _r["last_btmn"])
        _av = (_r["avg12_musd"] if show_usd and _r["avg12_musd"] is not None else _r["avg12_btmn"])
        rows_disp.append({
            "نماد":                              _r["نماد"],
            "آخرین دوره":                        _r["آخرین دوره"],
            f"فروش آخرین ماه ({_lb_last})":     fmt_fa((_r["last_btmn"] or 0) * _sc_last, 2) if _r["last_btmn"] else "—",
            f"میانگین ۱۲ ماه ({_lb_avg12})":    fmt_fa((_r["avg12_btmn"] or 0) * _sc_avg12, 2) if _r["avg12_btmn"] else "—",
            f"میانگین ۳ ماه ({_lb_avg3})":      fmt_fa((_r["avg3_btmn"] or 0) * _sc_avg3, 2) if _r["avg3_btmn"] else "—",
            "ماه آخر / م۱۲":                     _fmt_ratio(_r["ratio_last_12"]),
            "م۳ / م۱۲":                          _fmt_ratio(_r["ratio_3_12"]),
        })

    # ─── جدول sortable با JavaScript ──────────────────────────
    # ساخت HTML جدول با قابلیت کلیک روی هدر
    _tbl_rows = []
    for _, _r in df_sorted.iterrows():
        _ia = _r["idx_a"]
        _ib = _r["idx_b"]
        _ic = _r["idx_c"]
        _id = _r["idx_d"]
        # سلول شاخص d: اگر سودِ ۱۴۰۴ از گزارش واقعی نبود (پیش‌بینی/سال قبل) با ⚠ و رنگ زرد
        if _id is not None and not pd.isna(_id):
            _d_txt = fmt_fa(_id, 4)
            if _r.get("d_fallback"):
                _src_lbl = _r.get("d_src") or "برآوردی"
                _d_cell = (f'<span style="background:#3a2a0a;color:#fbbf24;'
                           f'padding:1px 6px;border-radius:5px;font-weight:700;" '
                           f'title="گزارش سالانهٔ {to_fa_digits(str(DIVIDEND_FY))} ندارد؛ '
                           f'سود از: {_src_lbl}">{_d_txt} ⚠</span>')
            else:
                _d_cell = _d_txt
        else:
            _d_cell = "—"
        _tbl_rows.append({
            "نماد":         _r["نماد"],
            "آخرین دوره":   _r["آخرین دوره"],
            "فروش آخرین":   fmt_fa((_r["last_btmn"] or 0) * _sc_last, 2) if _r["last_btmn"] else "—",
            "م.۱۲ماه":      fmt_fa((_r["avg12_btmn"] or 0) * _sc_avg12, 2) if _r["avg12_btmn"] else "—",
            "م.۳ماه":       fmt_fa((_r["avg3_btmn"] or 0) * _sc_avg3, 2) if _r["avg3_btmn"] else "—",
            "آخر/م۱۲":      _fmt_ratio(_r["ratio_last_12"]),
            "م۳/م۱۲":       _fmt_ratio(_r["ratio_3_12"]),
            "شاخص a":       (fmt_fa(_ia, 4) if (_ia is not None and not pd.isna(_ia)) else "—"),
            "شاخص b":       (fmt_fa(_ib, 4) if (_ib is not None and not pd.isna(_ib)) else "—"),
            "شاخص c":       (fmt_fa(_ic, 4) if (_ic is not None and not pd.isna(_ic)) else "—"),
            "شاخص d":       _d_cell,
            # مقادیر عددی برای sort
            "_last":        _r["last_btmn"] or 0,
            "_avg12":       _r["avg12_btmn"] or 0,
            "_avg3":        _r["avg3_btmn"] or 0,
            "_r1":          _r["ratio_last_12"] or 0,
            "_r3":          _r["ratio_3_12"] or 0,
            "_a":           (_ia if (_ia is not None and not pd.isna(_ia)) else 0),
            "_b":           (_ib if (_ib is not None and not pd.isna(_ib)) else 0),
            "_c":           (_ic if (_ic is not None and not pd.isna(_ic)) else 0),
            "_d":           (_id if (_id is not None and not pd.isna(_id)) else 0),
        })

    _display_cols = ["نماد","آخرین دوره",f"فروش آخرین ({_lb_last})",f"م.۱۲ماه ({_lb_avg12})",f"م.۳ماه ({_lb_avg3})","آخر/م۱۲","م۳/م۱۲"]
    _numeric_cols = ["_last","_avg12","_avg3","_r1","_r3"]

    import json as _json
    _data_js = _json.dumps([
        {
            "نماد": r["نماد"],
            "آخرین دوره": r["آخرین دوره"],
            f"فروش آخرین ({_lb_last})": r["فروش آخرین"],
            f"م.۱۲ماه ({_lb_avg12})": r["م.۱۲ماه"],
            f"م.۳ماه ({_lb_avg3})": r["م.۳ماه"],
            "آخر/م۱۲": r["آخر/م۱۲"],
            "م۳/م۱۲": r["م۳/م۱۲"],
            "شاخص a": r["شاخص a"],
            "شاخص b": r["شاخص b"],
            "شاخص c": r["شاخص c"],
            "شاخص d": r["شاخص d"],
            "_last": r["_last"],
            "_avg12": r["_avg12"],
            "_avg3": r["_avg3"],
            "_r1": r["_r1"],
            "_r3": r["_r3"],
            "_a": r["_a"],
            "_b": r["_b"],
            "_c": r["_c"],
            "_d": r["_d"],
        }
        for r in _tbl_rows
    ], ensure_ascii=False)

    _idx_disp, _idx_sort = [], []
    if want_a:
        _idx_disp.append("شاخص a"); _idx_sort.append("_a")
    if want_b:
        _idx_disp.append("شاخص b"); _idx_sort.append("_b")
    if want_c:
        _idx_disp.append("شاخص c"); _idx_sort.append("_c")
    if want_d:
        _idx_disp.append("شاخص d"); _idx_sort.append("_d")
    _numeric_keys = _json.dumps(["_last","_avg12","_avg3","_r1","_r3"] + _idx_sort, ensure_ascii=False)
    _display_keys = _json.dumps(["نماد","آخرین دوره",f"فروش آخرین ({_lb_last})",f"م.۱۲ماه ({_lb_avg12})",f"م.۳ماه ({_lb_avg3})","آخر/م۱۲","م۳/م۱۲"] + _idx_disp, ensure_ascii=False)
    _sort_keys    = _json.dumps(["نماد","آخرین دوره","_last","_avg12","_avg3","_r1","_r3"] + _idx_sort, ensure_ascii=False)
    # ستونِ مرتب‌سازیِ پیش‌فرض: اولین شاخصِ انتخاب‌شده، وگرنه «ماه آخر / م۱۲»
    _default_sort_js = _json.dumps(_idx_sort[0] if _idx_sort else "_r1")

    st.components.v1.html(f"""
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="utf-8">
<link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css" rel="stylesheet">
<style>
  body {{ background:transparent; margin:0; font-family:Vazirmatn,Tahoma,sans-serif; direction:rtl; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; color:#e2e8f0; }}
  th {{ background:#1e293b; padding:9px 10px; text-align:right; border:1px solid #334155;
        cursor:pointer; user-select:none; white-space:nowrap; }}
  th:hover {{ background:#273549; }}
  th.sorted-asc::after {{ content:" ↑"; color:#4DA3FF; }}
  th.sorted-desc::after {{ content:" ↓"; color:#4DA3FF; }}
  td {{ padding:8px 10px; border:1px solid #2d3748; }}
  tr:nth-child(odd)  {{ background:#0f172a; }}
  tr:nth-child(even) {{ background:#1a2236; }}
  td.sym-cell {{ cursor:pointer; color:#4DA3FF; font-weight:700; }}
  td.sym-cell:hover {{ text-decoration:underline; color:#7CC3FF; }}
</style>
</head>
<body>
<table id="ftbl"></table>
<script>
const data = {_data_js};
const dispKeys = {_display_keys};
const sortKeys = {_sort_keys};
const numKeys  = {_numeric_keys};
let sortCol = {_default_sort_js}, sortAsc = false;

function numVal(v) {{
  if (typeof v === "number") return v;
  const s = String(v).replace(/[^0-9.\x2d]/g,"");
  return parseFloat(s) || 0;
}}

function render() {{
  const sorted = [...data].sort((a,b) => {{
    const av = numKeys.includes(sortCol) ? (a[sortCol]||0) : String(a[sortCol]||"");
    const bv = numKeys.includes(sortCol) ? (b[sortCol]||0) : String(b[sortCol]||"");
    return sortAsc ? (av>bv?1:av<bv?-1:0) : (av<bv?1:av>bv?-1:0);
  }});

  const tbl = document.getElementById("ftbl");
  // هدر
  const hdr = "<thead><tr>" + dispKeys.map((k,i) => {{
    const sk = sortKeys[i];
    const cls = sk===sortCol ? (sortAsc?"sorted-asc":"sorted-desc") : "";
    return `<th class="${{cls}}" data-sk="${{sk}}">${{k}}</th>`;
  }}).join("") + "</tr></thead>";

  // ردیف‌ها
  const body = "<tbody>" + sorted.map(r => {{
    return "<tr>" + dispKeys.map((k,i) => {{
      if (i===0) return `<td class="sym-cell" onclick="pickSym('${{r["نماد"]}}')">${{r[k]}}</td>`;
      return `<td>${{r[k]}}</td>`;
    }}).join("") + "</tr>";
  }}).join("") + "</tbody>";

  tbl.innerHTML = hdr + body;
  tbl.querySelectorAll("th").forEach(th => {{
    th.onclick = () => {{
      const sk = th.dataset.sk;
      if (sk===sortCol) sortAsc=!sortAsc;
      else {{ sortCol=sk; sortAsc=false; }}
      render();
    }};
  }});
}}

function pickSym(sym) {{
  window.parent.location.href = window.parent.location.pathname + '?sym=' + encodeURIComponent(sym);
}}

render();
</script>

</body>
</html>
""", height=min(60 + len(_tbl_rows)*38, 800), scrolling=True)

    # ─── توضیح فرمول شاخص a ───────────────────────────────────
    st.markdown(
        """
**فرمول شاخص a:**

$$a = \\frac{\\text{درآمد دلاری ماه آخر}}{\\text{درآمد دلاری مشابه پارسال}}\\;\\times\\;\\overline{\\text{درآمد دلاری ۱۲ماهه (۵ سال اخیر)}}\\;\\times\\;\\overline{\\text{حاشیه سود ۱۲ماهه (۵ سال اخیر)}}\\;\\times\\;\\frac{\\overline{\\text{نسبت تقسیم سود}}}{\\text{ارزش بازار امروز (م.دلار)}}$$

به‌بیان ساده: (رشد دلاری ماه آخر نسبت به پارسال) × میانگین سود دلاری سالانهٔ ۵ سال × بازده نسبت به ارزش بازار × میانگین درصد تقسیم سود.

---

**فرمول شاخص b:**

$$b = \\frac{\\text{مبلغ تولید ماه آخر (نرخِ آخرین ماهِ پرفروش)}}{\\text{مبلغ تولید مشابه پارسال (همان روش)}}\\;\\times\\;\\overline{\\text{درآمد دلاری ۱۲ماهه (۵ سال اخیر)}}\\;\\times\\;\\overline{\\text{حاشیه سود ۱۲ماهه (۵ سال اخیر)}}\\;\\times\\;\\frac{\\overline{\\text{نسبت تقسیم سود}}}{\\text{ارزش بازار امروز (م.دلار)}}$$

شاخص b دقیقاً مثل a است، فقط فاکتورِ اول (رشد دلاریِ درآمدِ ماه آخر نسبت به پارسال) با **رشدِ مبلغِ تولیدِ ماه آخر نسبت به مشابهِ پارسال** جایگزین شده، که در آن «مبلغ تولید» = Σ (مقدار تولید هر محصول × نرخِ دلاریِ آخرین ماهی که ≥۲۰٪ تولیدِ همان ماه در آن فروخته شده) است.

---

**فرمول شاخص c:**

$$c = \\frac{\\text{میانگینِ ۲ سود خالصِ دلاریِ برتر (از ۵ سال اخیر)}}{\\text{ارزش بازار فعلی (م.دلار)}}$$

از میان **۵ گزارش سالانهٔ ۱۲ماههٔ اخیر**، سود خالص (عملیات در حال تداوم) هر سال با **نرخ دلار همان سال** به میلیون دلار تبدیل می‌شود؛ سپس **۲ مقدار بزرگ‌تر** انتخاب و میانگین گرفته می‌شود، و بر **ارزش بازار فعلی شرکت (میلیون دلار)** تقسیم می‌شود. این شاخص مستقل از حالت نمایش (تومان/دلار) است و نوعی «بازده سود دلاریِ سال‌های اوج نسبت به ارزش بازار» را نشان می‌دهد.

---

**فرمول شاخص d:**

$$d = \\frac{\\text{میانگینِ ۲ سود خالصِ دلاریِ برتر (از ۵ سال اخیر)}}{\\text{ارزش بازار فعلی (م.دلار)} \\;-\\; \\text{دیویدندِ دلاریِ امسال}}$$

دقیقاً مثل شاخص c، اما از **مخرج** (ارزش بازار) مقدار **دیویدندِ دلاریِ امسال** کم می‌شود، چون این سود به‌زودی در مجمعِ امسال از شرکت خارج می‌شود:

$$\\text{دیویدندِ دلاریِ امسال} = \\text{سود خالصِ دلاریِ سال مالی } %d \\;\\times\\; \\overline{\\text{نسبت تقسیم سود}}$$

منبعِ «سود خالصِ دلاریِ سال %d» به‌ترتیب اولویت: (۱) گزارش سالانهٔ ۱۲ماههٔ همان سال اگر منتشر شده، (۲) در غیر این صورت **پیش‌بینیِ ذخیره‌شده**، (۳) و اگر هیچ‌کدام نبود، **سود دلاری سال %d**. اگر شرکت زیان‌ده باشد، دیویدند صفر در نظر گرفته می‌شود.

> ⚠ سلول‌های شاخص d که با **رنگ زرد و علامت ⚠** مشخص شده‌اند یعنی **گزارش سالانهٔ %d هنوز منتشر نشده** و سودِ مجمعِ امسال از پیش‌بینی یا سود سال قبل **برآورد** شده است (با نگه‌داشتن نشانگر روی سلول، منبع نمایش داده می‌شود).
        """ % (DIVIDEND_FY, DIVIDEND_FY, DIVIDEND_FY - 1, DIVIDEND_FY)
    )
    st.caption(
        "اجزا: «درآمد دلاری ماه آخر» و «مشابه پارسال» از فروش ماهانه (تبدیل با نرخ دلار همان ماه) • "
        "«میانگین درآمد دلاری ۱۲ماههٔ ۵ سال» = میانگین درآمد دلاری ماهانهٔ ۶۰ ماه × ۱۲ • "
        "«حاشیه سود» = سود خالص ÷ درآمد عملیاتی (میانگین ۵ گزارش سالانهٔ ۱۲ماهه) • "
        "«ارزش بازار» از آخرین فایل روزانهٔ TSETMC به میلیون دلار • "
        "«نسبت تقسیم سود» = میانگین dps÷eps از مجمع. "
        "در شاخص b، «مبلغ تولید ماه آخر / مشابه پارسال» جایگزینِ «درآمد دلاری ماه آخر / مشابه پارسال» می‌شود؛ "
        "نرخِ هر محصول = نرخِ دلاریِ آخرین ماهی (کلِ تاریخچه) که فروشش ≥۲۰٪ تولیدِ آن ماه بوده (وگرنه میانگینِ نرخِ دلاری). "
        "اگر هر جزء برای نمادی موجود نباشد، شاخص «—» نمایش داده می‌شود. "
        "تبدیل دلاری فقط در حالت نمایش دلاری معنا دارد."
    )


    st.info(f"نماد فعال: **{selected_global}** — برای تغییر از منوی سمت راست یا کلیک روی نماد در جدول استفاده کنید")

# ════════════════════════════════════════════════════════════════
# صفحه ۵: سهامداران و زیرمجموعه‌ها
# ════════════════════════════════════════════════════════════════
elif page == "👥 سهامداران و زیرمجموعه‌ها":
    st.title("سهامداران و زیرمجموعه‌ها")

    _sh_df, _sh_fname = load_latest_shareholders()

    if _sh_df.empty:
        st.warning(
            f"فایل سهامداران پیدا نشد. اسکریپت `shareholders_all.py` را اجرا کن تا "
            f"خروجی در پوشهٔ `{SHAREHOLDERS_DIR}` ساخته شود."
        )
    else:
        st.caption(f"منبع: «{_sh_fname}» — مجموع {to_fa_digits(len(_sh_df))} ردیف، "
                   f"{to_fa_digits(_sh_df['نماد'].nunique())} نماد")

        def _norm_sym_sh(s):
            return normalize_symbol(s)

        sym_active = selected_global
        sym_active_n = _norm_sym_sh(sym_active)

        def _fmt_int_fa(v):
            try:
                return f"{float(str(v).replace(',', '')):,.0f}".translate(EN_TO_FA)
            except Exception:
                return "—"

        def _fmt_pct(v):
            try:
                return f"{float(v):,.3f}".translate(EN_TO_FA)
            except Exception:
                return to_fa_digits(v)

        # نگاشتِ ارزش بازارِ هر نماد (میلیارد تومان) از آخرین فایل روزانهٔ TSETMC
        # «ارزش بازار» در فایل به ریال است → ÷ ۱۰٬۰۰۰٬۰۰۰٬۰۰۰ = میلیارد تومان (۱ م.ت = ۱۰ میلیارد ریال)
        _mcap_bt = {}
        if (not _daily_df.empty and "نام" in _daily_df.columns
                and "ارزش بازار" in _daily_df.columns):
            for _, _r in _daily_df.iterrows():
                try:
                    _v = float(str(_r["ارزش بازار"]).replace(",", "").strip())
                    _mcap_bt[_norm_sym_sh(_r["نام"])] = _v / 1e10
                except Exception:
                    pass

        def _mcap_of(sym_n):
            """ارزش بازارِ کلِ یک نماد به میلیارد تومان (یا None)."""
            return _mcap_bt.get(sym_n)

        tab_owners, tab_subs = st.tabs(
            ["🏛️ سهامداران این نماد", "🧩 زیرمجموعه‌ها (این نماد سهامدارِ چه نمادهایی است)"]
        )

        # ─── تب ۱: سهامداران نمادِ فعال ─────────────────────────
        with tab_owners:
            st.subheader(f"سهامداران «{sym_active}»")
            owners = _sh_df[_sh_df["نماد"].apply(_norm_sym_sh) == sym_active_n].copy()
            if owners.empty:
                st.info("برای این نماد سهامداری در فایل ثبت نشده است.")
            else:
                # مرتب‌سازی نزولی بر اساس درصد (یا تعداد سهم)
                sort_col = "درصد" if "درصد" in owners.columns else "تعداد سهم"
                try:
                    owners = owners.sort_values(sort_col, ascending=False)
                except Exception:
                    pass

                bourse_owners = (owners["نماد سهامدار"].astype(str).str.strip() != "").sum()
                c1, c2 = st.columns(2)
                c1.metric("تعداد سهامداران عمده", to_fa_digits(len(owners)))
                c2.metric("سهامدارانِ بورسی", to_fa_digits(int(bourse_owners)))

                disp = pd.DataFrame({
                    "نام سهامدار":  owners["نام سهامدار"].astype(str).values,
                    "نماد سهامدار": owners["نماد سهامدار"].astype(str).values,
                    "تعداد سهم":    [_fmt_int_fa(v) for v in owners["تعداد سهم"].values],
                    "درصد":         [_fmt_pct(v) for v in owners["درصد"].values],
                })
                # ردیف‌هایی که سهامدارشان خودش بورسی است را برجسته کن
                _rs = {i: "color:#4DA3FF;font-weight:600;"
                       for i, v in enumerate(owners["نماد سهامدار"].astype(str).str.strip().values) if v}
                render_table(disp, row_styles=_rs)

        # ─── تب ۲: زیرمجموعه‌ها (نمادِ فعال در نقش سهامدار) ──────
        with tab_subs:
            st.subheader(f"«{sym_active}» سهامدارِ چه نمادهایی است؟")
            subs = _sh_df[_sh_df["نماد سهامدار"].apply(_norm_sym_sh) == sym_active_n].copy()
            if subs.empty:
                st.info(
                    "این نماد به‌عنوان «سهامدار» هیچ نماد دیگری در فایل دیده نشد. "
                    "(یا واقعاً سهامدارِ نماد بورسیِ دیگری نیست، یا نامش هنوز به نماد تطبیق نخورده.)"
                )
            else:
                sort_col = "درصد" if "درصد" in subs.columns else "تعداد سهم"
                try:
                    subs = subs.sort_values(sort_col, ascending=False)
                except Exception:
                    pass
                st.metric("تعداد نمادهای زیرمجموعه/تحت‌مالکیت", to_fa_digits(subs["نماد"].nunique()))

                disp2 = pd.DataFrame({
                    "نماد":      subs["نماد"].astype(str).values,
                    "نام شرکت":  subs["نام شرکت"].astype(str).values if "نام شرکت" in subs.columns else "",
                    "تعداد سهم": [_fmt_int_fa(v) for v in subs["تعداد سهم"].values],
                    "درصد":      [_fmt_pct(v) for v in subs["درصد"].values],
                })
                render_table(disp2)

                # ─── 💰 سود مجمعِ زیرمجموعه‌ها (تخمینِ سودِ رسیده به نمادِ فعال) ───
                st.markdown("---")
                st.markdown(f"#### 💰 سود مجمعِ زیرمجموعه‌ها (تخمینِ سودِ رسیده به «{sym_active}»)")
                _dfy = st.text_input("سال مالیِ گزارش:", value=DIVIDEND_FY_DEFAULT, key="subs_div_fy").strip()
                _pay_m = _div_load_payout(_dfy)
                _pro_m = _div_load_profit(_dfy)
                _sav_m = _div_inputs_map(sym_active, _dfy)
                _subs_d = subs.drop_duplicates("نماد").copy()

                def _pf_pay(s):
                    sv = _sav_m.get(normalize_symbol(s))
                    if sv and sv[0] is not None:
                        return round(sv[0], 1)
                    v = _pay_m.get(normalize_symbol(s))
                    return round(v * 100, 1) if v is not None else None

                def _pf_pro(s):
                    sv = _sav_m.get(normalize_symbol(s))
                    if sv and sv[1] is not None:
                        return round(sv[1])
                    v = _pro_m.get(normalize_symbol(s))
                    return round(v / 1e4) if v is not None else None    # میلیون ریال → میلیارد تومان

                _base = pd.DataFrame({
                    "نماد": _subs_d["نماد"].astype(str).values,
                    "نام شرکت": (_subs_d["نام شرکت"].astype(str).values
                                 if "نام شرکت" in _subs_d.columns else ""),
                    "درصد مالکیت": [round(_div_num(v), 2) if _div_num(v) is not None else None
                                    for v in _subs_d["درصد"].values],
                    "درصد تقسیم سود": [_pf_pay(s) for s in _subs_d["نماد"]],
                    "سود نماد (م.ت)": [_pf_pro(s) for s in _subs_d["نماد"]],
                })
                try:
                    _ded = st.data_editor(
                        _base, use_container_width=True, hide_index=True, key="subs_div_editor",
                        column_config={
                            "نماد": st.column_config.TextColumn("نماد", disabled=True),
                            "نام شرکت": st.column_config.TextColumn("نام شرکت", disabled=True),
                            "درصد مالکیت": st.column_config.NumberColumn("درصد مالکیت", disabled=True, format="%.2f"),
                            "درصد تقسیم سود": st.column_config.NumberColumn(
                                "درصد تقسیم سود", format="%.1f", help="dps/eps — پیش‌پرشده، قابلِ ویرایش"),
                            "سود نماد (م.ت)": st.column_config.NumberColumn(
                                "سود نماد (م.ت)", format="%.0f",
                                help="سود خالصِ سالانهٔ نماد به میلیارد تومان — پیش‌پرشده، قابلِ ویرایش"),
                        },
                    )
                    if st.button("💾 ذخیرهٔ اعداد (تقسیم سود و سود نماد)", key="subs_div_save"):
                        _div_save_inputs(sym_active, _dfy,
                                         zip(_ded["نماد"], _ded["درصد تقسیم سود"], _ded["سود نماد (م.ت)"]))
                        st.success(f"ذخیره شد در {DIVINPUTS_PATH} — دفعهٔ بعد همین اعداد می‌آیند.")
                        try:
                            st.rerun()
                        except AttributeError:
                            st.experimental_rerun()
                except Exception:
                    _ded = _base
                    st.warning("ویرایشِ درون‌برنامه‌ای در دسترس نیست؛ Streamlit را به‌روزرسانی کن "
                               "(pip install -U streamlit).")

                def _to_h(r):
                    try:
                        return (float(r["سود نماد (م.ت)"]) * float(r["درصد تقسیم سود"]) / 100.0
                                * float(r["درصد مالکیت"]) / 100.0)
                    except Exception:
                        return None
                _ded = _ded.copy()
                _ded["_d"] = _ded.apply(_to_h, axis=1)
                _tot = _ded["_d"].dropna().sum()
                st.metric(f"جمعِ سودِ مجمعِ رسیده به «{sym_active}» (م.ت)", _fmt_int_fa(_tot))
                _show = pd.DataFrame({
                    "نماد": _ded["نماد"],
                    "درصد مالکیت": _ded["درصد مالکیت"].map(
                        lambda v: (f"{v:,.2f}".translate(EN_TO_FA)) if pd.notna(v) else "—"),
                    "درصد تقسیم سود": _ded["درصد تقسیم سود"].map(
                        lambda v: (f"{v:,.1f}".translate(EN_TO_FA)) if pd.notna(v) else "—"),
                    "سود نماد (م.ت)": _ded["سود نماد (م.ت)"].map(
                        lambda v: _fmt_int_fa(v) if pd.notna(v) else "—"),
                    "سود به هلدینگ (م.ت)": _ded["_d"].map(
                        lambda v: _fmt_int_fa(v) if pd.notna(v) else "—"),
                })
                _show = pd.concat([_show, pd.DataFrame([{
                    "نماد": "جمع", "درصد مالکیت": "", "درصد تقسیم سود": "",
                    "سود نماد (م.ت)": "", "سود به هلدینگ (م.ت)": _fmt_int_fa(_tot),
                }])], ignore_index=True)
                render_table(_show)
                st.caption("سود به هلدینگ = سود نماد × درصد تقسیم سود × درصد مالکیت. "
                           "«درصد تقسیم سود» از codal_assembly.db و «سود نماد» از سود خالصِ سالانهٔ "
                           "۱۲ماههٔ codal_annual.db پیش‌پر می‌شوند؛ هر تغییری بدهی و «ذخیره» بزنی، "
                           "دفعهٔ بعد همان اعداد می‌آیند.")

                # ─── نمودار تو‌در‌توی زیرمجموعه‌ها (مالکیت غیرمستقیم) ───
                st.markdown("---")
                st.markdown("#### 🌳 درختِ زیرمجموعه‌ها (شامل مالکیتِ غیرمستقیم)")
                st.caption("هر شاخه = نمادی که والدش سهامدار آن است. درصدِ کنارِ هر گره، "
                           "مالکیتِ غیرمستقیمِ نمادِ فعال از آن شرکت است "
                           "(حاصل‌ضربِ درصدها در طولِ مسیر). مثال: دپارس ← دفارا ← دکپسول.")

                # نگاشتِ سریع: هر نماد → لیستِ (نمادِ زیرمجموعه، درصد، نام شرکت)
                _sh_df["_p"] = _sh_df["نماد سهامدار"].apply(_norm_sym_sh)
                _sh_df["_c"] = _sh_df["نماد"].apply(_norm_sym_sh)
                _name_of = {}
                for _, r in _sh_df.iterrows():
                    nm = str(r.get("نام شرکت", "") or "")
                    if r["_c"] and r["_c"] not in _name_of:
                        _name_of[r["_c"]] = nm

                def _children_of(parent_n):
                    """زیرمجموعه‌های مستقیم: (نمادِ نمایش، نماد نرمال، درصد)"""
                    out = []
                    sub = _sh_df[_sh_df["_p"] == parent_n]
                    for _, r in sub.iterrows():
                        try:
                            pct = float(r["درصد"])
                        except Exception:
                            pct = None
                        out.append((str(r["نماد"]), r["_c"], pct))
                    return out

                # ساخت گره‌های نمودار با پیمایشِ بازگشتیِ امن (جلوگیری از حلقه)
                ids, labels, parents, values, customdata, node_syms = [], [], [], [], [], []
                MAX_DEPTH = 6

                def _walk(node_n, node_disp, parent_id, acc_pct, path, depth):
                    node_id = (parent_id + " / " if parent_id else "") + node_disp
                    ids.append(node_id)
                    labels.append(node_disp)
                    parents.append(parent_id)
                    node_syms.append(node_n)
                    # مقدارِ نمودار = درصدِ غیرمستقیم (برای ریشه ۱۰۰)
                    values.append(acc_pct if acc_pct is not None else 0.0)
                    customdata.append(acc_pct)
                    if depth >= MAX_DEPTH:
                        return
                    for cdisp, cn, cpct in _children_of(node_n):
                        if not cn or cn in path:          # جلوگیری از حلقه و خودارجاعی
                            continue
                        child_acc = None
                        if acc_pct is not None and cpct is not None:
                            child_acc = acc_pct * cpct / 100.0
                        _walk(cn, cdisp, node_id, child_acc, path | {cn}, depth + 1)

                # ریشه = نمادِ فعال با ۱۰۰٪
                _walk(sym_active_n, sym_active, "", 100.0, {sym_active_n}, 0)

                if len(ids) <= 1:
                    st.info("زیرمجموعه‌ای برای ترسیمِ درخت پیدا نشد.")
                else:
                    # متنِ هر گره: درصد + ارزشِ سهم (م.ت)
                    _txt, _hover = [], []
                    for nsym, v in zip(node_syms, customdata):
                        _mc = _mcap_of(nsym)
                        _sv = (_mc * v / 100.0) if (_mc is not None and v is not None) else None
                        _ptxt = (f"{v:.2f}%".translate(EN_TO_FA) if v is not None else "")
                        _vtxt = (f"{_fmt_int_fa(_sv)} م.ت" if _sv is not None else "")
                        _txt.append((_ptxt + (("<br>" + _vtxt) if _vtxt else "")))
                        _hover.append(
                            f"مالکیت: {_ptxt or '—'}<br>ارزش سهم: {_vtxt or '—'}"
                            + (f"<br>ارزش بازار کل: {_fmt_int_fa(_mc)} م.ت" if _mc is not None else ""))
                    fig_tree = go.Figure(go.Icicle(
                        ids=ids, labels=labels, parents=parents,
                        text=_txt, textinfo="label+text",
                        customdata=_hover,
                        tiling=dict(orientation="h"),
                        hovertemplate="<b>%{label}</b><br>%{customdata}<extra></extra>",
                        marker=dict(colorscale="Tealgrn"),
                    ))
                    fig_tree.update_layout(
                        height=max(320, 60 + 46 * max(2, len(ids) - 1)),
                        margin=dict(t=20, l=10, r=10, b=10),
                        font=dict(family="Vazirmatn,Tahoma,sans-serif", size=13),
                        paper_bgcolor="rgba(0,0,0,0)",
                    )
                    st.plotly_chart(fig_tree, use_container_width=True)

                    # جدولِ مالکیتِ غیرمستقیم (همهٔ گره‌ها به‌جز ریشه)
                    rows_ind = []
                    _total_share_val = 0.0
                    _have_val = False
                    for nid, lab, par, nsym, acc in zip(ids, labels, parents, node_syms, customdata):
                        if not par:
                            continue
                        depth = nid.count(" / ")
                        kind = "مستقیم" if depth == 1 else f"غیرمستقیم (سطح {to_fa_digits(depth)})"
                        mcap = _mcap_of(nsym)                       # ارزش بازارِ کلِ نماد (م.ت)
                        share_val = (mcap * acc / 100.0) if (mcap is not None and acc is not None) else None
                        if share_val is not None:
                            _total_share_val += share_val
                            _have_val = True
                        rows_ind.append({
                            "نماد": lab,
                            "مسیر": par.replace(" / ", " ← ") + " ← " + lab,
                            "نوع": kind,
                            "٪ مالکیت": (f"{acc:.3f}".translate(EN_TO_FA) if acc is not None else "—"),
                            "ارزش بازار کل (م.ت)": (_fmt_int_fa(mcap) if mcap is not None else "—"),
                            "ارزش سهم (م.ت)": (_fmt_int_fa(share_val) if share_val is not None else "—"),
                        })
                    if rows_ind:
                        st.markdown("##### جدولِ مالکیت (مستقیم و غیرمستقیم)")
                        if _have_val:
                            st.metric("جمعِ ارزشِ سهمِ زیرمجموعه‌ها (م.ت)", _fmt_int_fa(_total_share_val))
                        render_table(pd.DataFrame(rows_ind))
                        st.caption("«ارزش سهم» = درصدِ مالکیت × ارزش بازارِ کلِ آن نماد. "
                                   "مثال: ۳۰٪ از یک شرکتِ ۸٬۰۰۰ میلیارد تومانی = ۲٬۴۰۰ میلیارد تومان. "
                                   "ارزش بازارها از آخرین فایل روزانهٔ TSETMC است؛ نمادهایی که در آن فایل نباشند «—» می‌شوند.")

        st.info(f"نماد فعال: **{sym_active}** — برای تغییر، از منوی «نماد» در نوار سمت راست استفاده کن.")


# ════════════════════════════════════════════════════════════════
# صفحه ۶: پیش‌بینی سود (اقلام تفصیلی صورت سود و زیان — ذخیره در forecasts.db)
# ════════════════════════════════════════════════════════════════
elif page == "🔮 پیش‌بینی سود":
    st.title("پیش‌بینی سود (صورت سود و زیان)")
    sym_fc = selected_global
    st.caption(f"نماد فعال: **{sym_fc}** — مقادیر به **میلیارد تومان** وارد می‌شوند "
               "(ردیفِ سود هر سهم به ریال).")

    # ─── انتخاب سال مالیِ پیش‌بینی ───────────────────────────
    # سال‌های موجود در گزارش‌های سالانهٔ این نماد را پیدا کن تا سالِ بعدی را پیشنهاد بدهیم
    _years_have = []
    if has_annual:
        _ra = df_annual_rep[df_annual_rep["symbol"] == sym_fc]
        for _pe in _ra["period_end"].tolist():
            _yy, _mm = _inc_year_month(_pe)
            if _yy:
                _years_have.append(_yy)
    _last_actual_year = max(_years_have) if _years_have else None
    _default_year = (_last_actual_year + 1) if _last_actual_year else 1405

    c_y1, c_y2 = st.columns([1, 3])
    with c_y1:
        fiscal_year = st.number_input(
            "سال مالی پیش‌بینی (شمسی):", min_value=1390, max_value=1430,
            value=int(_default_year), step=1, key="fc_year")
    fiscal_year = str(int(fiscal_year))

    # ─── منبعِ اقلام: پیش‌بینی ذخیره‌شده، وگرنه آخرین صورت سالانه ───
    saved_df, saved_note = load_forecast(sym_fc, fiscal_year)

    # آخرین صورت سالانهٔ واقعی برای نمایشِ کنار دستی و seed اولیه
    actual_labels, actual_vals = [], {}
    if has_annual:
        reps_fc = df_annual_rep[df_annual_rep["symbol"] == sym_fc].copy()
        reps_fc = reps_fc.sort_values(["period_end", "duration_months"], ascending=[False, False])
        ml_fc, mk_fc, cmeta_fc, mtx_fc = _inc_build_matrix(reps_fc, df_annual_items)
        if ml_fc:
            actual_labels = ml_fc
            # ستونِ جدیدترین گزارش = اندیس ۰ — تبدیل م.ریال به میلیارد تومان (به‌جز EPS)
            for ri, lab in enumerate(ml_fc):
                _rawv = mtx_fc[0][ri] if mtx_fc.get(0) else None
                if _rawv is not None and not pd.isna(_rawv) and "هرسهم" not in _inc_label_key(lab):
                    actual_vals[lab] = _rawv / 10000.0
                else:
                    actual_vals[lab] = _rawv

    if saved_df is not None and not saved_df.empty:
        base_labels = saved_df["label"].tolist()
        seed_vals = dict(zip(saved_df["label"], saved_df["value"]))
        src_note = "از پیش‌بینیِ ذخیره‌شدهٔ قبلی"
    elif actual_labels:
        base_labels = actual_labels
        seed_vals = {}      # خالی شروع شود تا کاربر خودش پر کند
        src_note = "ساختار اقلام از آخرین صورت سالانهٔ واقعی"
    else:
        base_labels = ["درآمدهای عملیاتی", "بهای تمام شدهٔ درآمدهای عملیاتی",
                       "سود (زیان) ناخالص", "هزینه‌های فروش، اداری و عمومی",
                       "سود (زیان) عملیاتی", "سود (زیان) خالص",
                       "سود (زیان) خالص هر سهم – ریال", "سرمایه"]
        seed_vals = {}
        src_note = "اقلام پیش‌فرض (گزارش سالانه‌ای برای این نماد یافت نشد)"

    st.caption(f"منبع اقلام: {src_note}")

    # ─── جدول قابل‌ویرایش ────────────────────────────────────
    _last_actual_label = (f"واقعیِ {to_fa_digits(_last_actual_year)}"
                          if _last_actual_year else "آخرین واقعی")
    editor_rows = []
    for lab in base_labels:
        editor_rows.append({
            "شرح": lab,
            _last_actual_label: actual_vals.get(lab, None),
            f"پیش‌بینی {to_fa_digits(fiscal_year)}": seed_vals.get(lab, None),
        })
    editor_df = pd.DataFrame(editor_rows)

    st.markdown("##### اقلام صورت سود و زیان")
    edited = st.data_editor(
        editor_df,
        key=f"fc_editor_{sym_fc}_{fiscal_year}",
        use_container_width=True,
        num_rows="dynamic",     # امکان افزودن/حذف ردیف
        column_config={
            "شرح": st.column_config.TextColumn("شرح", width="large"),
            _last_actual_label: st.column_config.NumberColumn(
                _last_actual_label, help="آخرین مقدار واقعی (میلیارد تومان؛ EPS به ریال) — فقط مرجع",
                disabled=True, format="%.2f"),
            f"پیش‌بینی {to_fa_digits(fiscal_year)}": st.column_config.NumberColumn(
                f"پیش‌بینی {to_fa_digits(fiscal_year)}",
                help="مقدار پیش‌بینیِ خودت (میلیارد تومان؛ ردیفِ EPS به ریال)",
                format="%.2f"),
        },
        hide_index=True,
    )

    note_txt = st.text_area("یادداشت (اختیاری):", value=saved_note or "",
                            key=f"fc_note_{sym_fc}_{fiscal_year}", height=90)

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
            save_forecast(sym_fc, fiscal_year, [], "")   # جایگزینی با خالی = پاک‌سازی
            st.warning("پیش‌بینیِ این نماد/سال پاک شد. (صفحه را رفرش کن)")

    # ─── محاسبهٔ خودکارِ چند نسبت از روی پیش‌بینی ─────────────
    def _find_val(rows_df, *needles):
        for _, r in rows_df.iterrows():
            k = _inc_label_key(r["شرح"])
            if all(_inc_label_key(n) in k for n in needles):
                v = r[fc_col]
                if v is not None and not pd.isna(v):
                    return float(v)
        return None

    rev_f  = _find_val(edited, "درآمد") or _find_val(edited, "درآمدعملیاتی")
    net_f  = _find_val(edited, "سود", "خالص")
    eps_f  = _find_val(edited, "هرسهم")
    if rev_f or net_f or eps_f:
        st.markdown("---")
        st.markdown("##### چکیدهٔ پیش‌بینی")
        m1, m2, m3 = st.columns(3)
        m1.metric("درآمد عملیاتی (م.ت)", fmt_fa(rev_f, 1) if rev_f else "—")
        m2.metric("سود خالص (م.ت)", fmt_fa(net_f, 1) if net_f else "—")
        _margin = (net_f / rev_f * 100) if (rev_f and net_f) else None
        m3.metric("حاشیه سود خالص", (f"{_margin:,.1f}%".translate(EN_TO_FA)) if _margin is not None else "—")
        if eps_f:
            st.caption(f"EPS پیش‌بینی: **{fmt_fa(eps_f, 0)}** ریال")

    # ─── فهرستِ همهٔ پیش‌بینی‌های ذخیره‌شده ───────────────────
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