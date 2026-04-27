import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="전표 변환기", layout="wide")
st.title("📊 증권사 통합 전표 변환기")

# ─────────────────────────────
# 기본 유틸
# ─────────────────────────────
def to_int(v):
    try:
        if pd.isna(v):
            return 0
        if isinstance(v, str):
            v = re.sub(r"[^0-9.-]", "", v)
        return int(float(v))
    except:
        return 0

def clean(v):
    try:
        if pd.isna(v):
            return ""
        return str(v).strip()
    except:
        return ""

def parse_date(v):
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except:
        return None, None

# ─────────────────────────────
# 🔥 거래유형 통합 정규화
# ─────────────────────────────
def normalize_trade_type(v):
    v = str(v)

    if "매도" in v:
        return "SELL"
    if "매수" in v:
        return "BUY"
    if "예탁금" in v:
        return "INTEREST"
    if "입고" in v or "공모주" in v:
        return "StockCredit"
    if "입금" in v:
        return "Credit"
    if "출금" in v:
        return "Debit"

    return None

# ─────────────────────────────
# 종목명 normalize (핵심)
# ─────────────────────────────
def extract_stock_name(v):
    v = str(v).replace(" ", "").strip()
    if "#" in v:
        return v.split("#")[-1]
    return v

# ─────────────────────────────
# 거래처 매핑
# ─────────────────────────────
def load_broker_map(file):
    if file is None:
        return {}

    df = pd.read_excel(file)

    mp = {}
    for code, name in zip(df.iloc[:,0], df.iloc[:,1]):
        key = extract_stock_name(name)
        mp[key] = (str(code).strip(), str(name).strip())

    return mp

def get_broker_info(stock, broker_map):
    key = extract_stock_name(stock)
    return broker_map.get(key, ("", stock))

# ─────────────────────────────
# 컬럼 찾기 (안전)
# ─────────────────────────────
def find_col(df, keys):
    for c in df.columns:
        for k in keys:
            if k in str(c):
                return c
    return None

# ─────────────────────────────
# 🔥 핵심: 증권사 통합 파서
# ─────────────────────────────
def parse_trades(df):
    trades = []

    header = None
    for i in range(min(20, len(df))):
        row = df.iloc[i].astype(str)
        if any("거래일" in x for x in row):
            header = i
            break

    if header is None:
        return []

    df.columns = df.iloc[header]
    df = df.iloc[header+1:].reset_index(drop=True)

    c_date  = find_col(df, ["거래일"])
    c_type  = find_col(df, ["구분","적요","거래종류","내용"])
    c_stock = find_col(df, ["종목"])
    c_qty   = find_col(df, ["수량"])
    c_price = find_col(df, ["단가"])
    c_net   = find_col(df, ["금액","거래대금"])

    for i in range(len(df)):
        try:
            r = df.iloc[i]

            m,d = parse_date(r.get(c_date))
            if not m:
                continue

            ttype = clean(r.get(c_type))
            stock = clean(r.get(c_stock))
            qty   = to_int(r.get(c_qty))
            price = to_int(r.get(c_price))
            net   = to_int(r.get(c_net))

            trades.append({
                "m":m,"d":d,
                "type":ttype,
                "stock":stock,
                "qty":qty,
                "price":price,
                "net":net
            })

        except:
            continue

    return trades

# ─────────────────────────────
# 전표 생성
# ─────────────────────────────
def process(trades, broker_map, broker_code):
    rows = []

    for t in trades:

        ttype = normalize_trade_type(t["type"])
        if not ttype:
            continue

        m,d = t["m"],t["d"]
        stock = extract_stock_name(t["stock"])
        qty = t["qty"]
        price = t["price"]
        net = t["net"]

        cp_code, cp_name = get_broker_info(stock, broker_map)

        memo = f"{stock}"

        if ttype == "SELL":
            rows.append([m,d,"차변",12500,"예치금",broker_code,"",memo,net,0])
            rows.append([m,d,"대변",10700,"주식",cp_code,cp_name,memo,0,qty*price])

        elif ttype == "BUY":
            cost = qty * price
            rows.append([m,d,"차변",10700,"주식",cp_code,cp_name,memo,cost,0])
            rows.append([m,d,"대변",12500,"예치금",broker_code,"",memo,0,cost])

        elif ttype == "INTEREST":
            rows.append([m,d,"차변",12500,"예치금",broker_code,"",memo,net,0])
            rows.append([m,d,"대변",42000,"이자수익","", "",memo,0,net])

    return rows

# ─────────────────────────────
# Excel 출력
# ─────────────────────────────
def to_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active

    header = ["월","일","구분","계정코드","계정명","거래처코드","거래처명","적요","차변","대변"]
    ws.append(header)

    for r in rows:
        ws.append(r)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ─────────────────────────────
# UI
# ─────────────────────────────
broker_file = st.file_uploader("거래처 매핑", type=["xlsx"])
broker_code = st.text_input("증권사 코드")
uploaded = st.file_uploader("엑셀 업로드", type=["xlsx"])

if uploaded:
    if st.button("변환"):

        broker_map = load_broker_map(broker_file)

        tmp = tempfile.NamedTemporaryFile(delete=False)
        tmp.write(uploaded.read())
        tmp.close()

        xl = pd.ExcelFile(tmp.name)

        all_trades = []
        for sh in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sh, header=None)
            all_trades += parse_trades(df)

        rows = process(all_trades, broker_map, broker_code)

        st.success(f"{len(rows)}건 변환 완료")

        out = to_excel(rows)

        st.download_button("다운로드", out, "result.xlsx")
