import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re
from datetime import datetime
from openpyxl.styles import PatternFill

st.set_page_config(page_title="증권사 전표 변환기", layout="wide")
st.title("📊 증권사 전표 통합 변환기")

# ─────────────────────────────
# 기본 유틸
# ─────────────────────────────
def to_int(v):
    try:
        if v is None:
            return 0
        if isinstance(v, str):
            v = re.sub(r"[^0-9.-]", "", v)
        return int(float(v))
    except:
        return 0


def clean(v):
    if v is None:
        return ''
    try:
        if pd.isna(v):
            return ''
    except:
        pass
    return str(v).strip()


# ─────────────────────────────
# 날짜 판별 (핵심 보완)
# ─────────────────────────────
def is_date(val):
    if val is None:
        return False

    if isinstance(val, (pd.Timestamp, datetime)):
        return True

    s = str(val).strip()
    if s in ("", "nan", "NaT"):
        return False

    return bool(re.search(r"\d{4}\D\d{1,2}\D\d{1,2}", s))


def parse_date(v):
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except:
        return None, None


# ─────────────────────────────
# 거래유형
# ─────────────────────────────
def normalize_trade_type(ttype):
    ttype = str(ttype)

    if "매도" in ttype:
        return "SELL"
    elif "매수" in ttype:
        return "BUY"
    elif "예탁금" in ttype or "이용료" in ttype:
        return "INTEREST"
    elif "입고" in ttype:
        return "StockCredit"
    elif "입금" in ttype:
        return "Credit"
    elif "출금" in ttype:
        return "Debit"

    return None


# ─────────────────────────────
# row 생성
# ─────────────────────────────
def row(m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr):
    return [m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr]


# ─────────────────────────────
# 종목 정리
# ─────────────────────────────
def extract_stock_name(name):
    name = str(name).replace(" ", "").strip()
    if "#" in name:
        return name.split("#")[-1]
    return name


# ─────────────────────────────
# 엑셀 컬럼 찾기 (완전 유연)
# ─────────────────────────────
def find_col(df, keys):
    for c in df.columns:
        for k in keys:
            if k in str(c):
                return c
    return None


# ─────────────────────────────
# 핵심 통합 파서 (증권사 공용)
# ─────────────────────────────
def parse_sheet(df):

    header_row = None

    # 1️⃣ 날짜 포함 행 찾기 (전체 행 검사)
    for i in range(len(df)):
        row = df.iloc[i].astype(str)
        if any(is_date(v) for v in row):
            header_row = i
            break

    if header_row is None:
        return []

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    # 2️⃣ 컬럼 자동 매칭
    c_date  = find_col(df, ["거래일", "거래일자", "일자", "날짜"])
    c_type  = find_col(df, ["구분", "적요", "내용", "거래종류"])
    c_stock = find_col(df, ["종목", "종목명"])
    c_qty   = find_col(df, ["수량", "거래수량", "거래좌수"])
    c_price = find_col(df, ["단가", "가격", "기준가"])
    c_net   = find_col(df, ["금액", "거래금액", "거래대금"])
    c_fee   = find_col(df, ["수수료"])
    c_tax   = find_col(df, ["세금", "제세금"])

    trades = []

    i = 0
    while i < len(df):

        try:
            r1 = df.iloc[i]

            # 날짜 아니면 skip
            if not is_date(r1.get(c_date)):
                i += 1
                continue

            # 짝꿍행 체크 (다음 행이 상세일 가능성)
            r2 = df.iloc[i + 1] if i + 1 < len(df) else None
            if r2 is not None and is_date(r2.get(c_date)):
                r2 = None  # 날짜면 같은 거래 아님

            def get(col):
                if col is None:
                    return None
                if r2 is not None:
                    return r2.get(col)
                return r1.get(col)

            m, d = parse_date(r1.get(c_date))

            trade_type = clean(get(c_type))
            stock = clean(get(c_stock))

            trades.append({
                "month": m,
                "day": d,
                "type": trade_type,
                "stock": stock,
                "qty": to_int(get(c_qty)),
                "price": to_int(get(c_price)),
                "net": to_int(get(c_net)),
                "fee": to_int(get(c_fee)),
                "tax": to_int(get(c_tax)),
            })

            i += 2 if r2 is not None else 1

        except:
            i += 1

    return trades


# ─────────────────────────────
# 전표 생성
# ─────────────────────────────
def process(trades, broker_code):

    rows = []

    for t in trades:

        m = t["month"]
        d = t["day"]
        ttype = normalize_trade_type(t["type"])

        if not ttype:
            continue

        stock = extract_stock_name(t["stock"])
        qty = t["qty"]
        price = t["price"]
        net = t["net"]
        fee = t["fee"]

        if ttype == "SELL":
            memo = f"{stock} 매도"
            rows.append(row(m,d,"차변",12500,"예치금",broker_code,"",memo,net,0))
            rows.append(row(m,d,"대변",10700,"증권",broker_code,stock,memo,0,qty*price))

        elif ttype == "BUY":
            cost = qty * price
            memo = f"{stock} 매수"

            rows.append(row(m,d,"차변",10700,"증권",broker_code,stock,memo,cost,0))
            rows.append(row(m,d,"대변",12500,"예치금",broker_code,"",memo,0,cost+fee))

        elif ttype == "INTEREST":
            memo = "이자"
            rows.append(row(m,d,"차변",12500,"예치금",broker_code,"",memo,net,0))
            rows.append(row(m,d,"대변",42000,"이자수익",broker_code,"",memo,0,net))

    return rows


# ─────────────────────────────
# 엑셀 생성
# ─────────────────────────────
def make_excel(rows):

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
broker_code = st.text_input("증권사 코드")
file = st.file_uploader("엑셀 업로드")

if file and st.button("변환"):

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(file.read())
    tmp.close()

    xl = pd.ExcelFile(tmp.name)

    all_trades = []

    for s in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=s, header=None)
        all_trades += parse_sheet(df)

    rows = process(all_trades, broker_code)

    if not rows:
        st.error("0건 (파싱 실패)")
    else:
        out = make_excel(rows)
        st.success(f"{len(rows)}건 변환 완료")

        st.download_button(
            "다운로드",
            data=out,
            file_name="result.xlsx"
        )
