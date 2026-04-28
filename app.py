import streamlit as st
import pandas as pd
import openpyxl
import io, re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="전표 변환기", layout="wide")
st.title("📊 증권사 전표 자동 변환기")

# ─────────────────────────────
# 기본 유틸
# ─────────────────────────────
def clean(v):
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except:
        pass
    return str(v).strip()

def to_int(v):
    try:
        return int(float(str(v).replace(',', '').replace(' ', '')))
    except:
        return 0

def parse_date(v):
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except:
        return None, None

def extract_stock_name(name):
    name = str(name).replace(" ", "").strip()
    if "#" in name:
        return name.split("#")[-1]
    return name

def normalize_trade_type(t):
    t = str(t)
    if "매수" in t:
        return "BUY"
    if "매도" in t:
        return "SELL"
    if "이자" in t:
        return "INTEREST"
    return None

# ─────────────────────────────
# 거래처 매핑
# ─────────────────────────────
def load_broker_map(file):
    if file is None:
        return {}

    df = pd.read_excel(file)

    return {
        str(name).strip(): (str(code).strip(), str(name).strip())
        for code, name in zip(df.iloc[:, 0], df.iloc[:, 1])
    }

def get_broker_info(stock, broker_map):
    return broker_map.get(stock, ("", "미등록"))

# ─────────────────────────────
# 증권사 감지
# ─────────────────────────────
def detect_broker(df):
    sample = df.astype(str).fillna("").head(10).to_string()

    if "한국투자" in sample:
        return "HANTOO"
    if "키움" in sample:
        return "KIWOOM"
    if "메리츠" in sample:
        return "MERITZ"

    return "GENERIC"

# ─────────────────────────────
# 📌 증권사별 파서 (데이터만 추출)
# ─────────────────────────────

def parse_hantoo(df):
    trades = []

    for _, r in df.iterrows():
        m, d = parse_date(r.get("거래일"))
        if not m:
            continue

        trades.append({
            "date": (m, d),
            "type": clean(r.get("구분")),
            "stock": clean(r.get("종목명")),
            "qty": to_int(r.get("수량")),
            "price": to_int(r.get("단가")),
            "amount": to_int(r.get("거래금액")),
            "fee": to_int(r.get("수수료")),
            "tax": 0
        })

    return trades


def parse_kiwoom(df):
    trades = []

    for _, r in df.iterrows():
        m, d = parse_date(r.get("일자"))
        if not m:
            continue

        trades.append({
            "date": (m, d),
            "type": clean(r.get("구분")),
            "stock": clean(r.get("종목명")),
            "qty": to_int(r.get("수량")),
            "price": to_int(r.get("체결가")),
            "amount": to_int(r.get("거래금액")),
            "fee": to_int(r.get("수수료")),
            "tax": to_int(r.get("세금"))
        })

    return trades


def parse_meritz(df):
    trades = []
    i = 2

    while i < len(df):

        r1 = df.iloc[i]
        r2 = df.iloc[i+1] if i+1 < len(df) else None

        m, d = parse_date(r1.iloc[0])
        if not m:
            i += 1
            continue

        trades.append({
            "date": (m, d),
            "type": clean(r2.iloc[0]) if r2 is not None else "",
            "stock": "",
            "qty": 0,
            "price": 0,
            "amount": to_int(r2.iloc[5]) if r2 is not None else 0,
            "fee": 0,
            "tax": 0
        })

        i += 2

    return trades


def parse_generic(df):
    trades = []

    for _, r in df.iterrows():

        m, d = parse_date(r.get("거래일") or r.get("일자"))
        if not m:
            continue

        trades.append({
            "date": (m, d),
            "type": clean(r.get("구분")),
            "stock": clean(r.get("종목")),
            "qty": to_int(r.get("수량")),
            "price": to_int(r.get("단가")),
            "amount": to_int(r.get("금액")),
            "fee": 0,
            "tax": 0
        })

    return trades

# ─────────────────────────────
# router
# ─────────────────────────────
def parse_router(df):

    broker = detect_broker(df)

    if broker == "HANTOO":
        return parse_hantoo(df)

    if broker == "KIWOOM":
        return parse_kiwoom(df)

    if broker == "MERITZ":
        return parse_meritz(df)

    return parse_generic(df)

# ─────────────────────────────
# 회계 엔진 (단일)
# ─────────────────────────────
def row(m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr):
    return [m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr]


def process_trades(trades, broker_map, broker_code):

    rows = []

    for t in trades:

        m, d = t["date"]
        ttype = normalize_trade_type(t["type"])

        if not ttype:
            continue

        stock = extract_stock_name(t["stock"])
        cp_code, cp_name = get_broker_info(stock, broker_map)

        qty = t["qty"]
        price = t["price"]
        amount = t["amount"]

        # BUY
        if ttype == "BUY":
            rows.append(row(m,d,"차변",10700,"단기매매증권",cp_code,cp_name,stock,qty*price,0))
            rows.append(row(m,d,"대변",12500,"예치금",broker_code,"",stock,0,amount))

        # SELL
        elif ttype == "SELL":
            rows.append(row(m,d,"차변",12500,"예치금",broker_code,"",stock,amount,0))
            rows.append(row(m,d,"대변",10700,"단기매매증권",cp_code,cp_name,stock,0,qty*price))

    return rows

# ─────────────────────────────
# Excel 생성
# ─────────────────────────────
def create_excel(rows):

    wb = openpyxl.Workbook()
    ws = wb.active

    header = [
        "월","일","구분",
        "계정코드","계정명",
        "거래처코드","거래처명",
        "적요","차변","대변"
    ]

    ws.append(header)

    fill = PatternFill(start_color="FFF59D", fill_type="solid")
    for i in range(len(header)):
        ws.cell(1, i+1).fill = fill

    for r in rows:
        ws.append(r)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ─────────────────────────────
# UI
# ─────────────────────────────
broker_file = st.file_uploader("거래처 매핑")
broker_code = st.text_input("증권사 거래처코드")
uploaded = st.file_uploader("엑셀 업로드")

if uploaded and st.button("변환 실행"):

    broker_map = load_broker_map(broker_file)

    xl = pd.ExcelFile(uploaded)

    trades = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        trades += parse_router(df)

    st.write("총 거래:", len(trades))

    rows = process_trades(trades, broker_map, broker_code)

    if not rows:
        st.error("변환 실패")
    else:
        out = create_excel(rows)
        st.success("완료")
        st.download_button("다운로드", data=out, file_name="result.xlsx")
