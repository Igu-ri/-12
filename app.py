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
    if pd.isna(v):
        return ''
    return str(v).strip()


# ─────────────────────────────
# 🔥 핵심: 날짜 "인지" 함수 (변환 X)
# ─────────────────────────────
def is_date_like(v):
    if v is None:
        return False

    s = str(v).strip()

    # 완전 숫자/빈값 제외
    if s == "" or s.lower() == "nan":
        return False

    # 날짜 패턴만 체크
    return bool(
        re.match(r"^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}", s) or  # 2026-01-01
        re.match(r"^\d{1,2}[.\-/]\d{1,2}", s) or             # 01-02
        re.match(r"^\d{4}", s)                                # 2026...
    )


# ─────────────────────────────
# 날짜 파싱 (이건 "필요할 때만")
# ─────────────────────────────
def parse_date_safe(v):
    try:
        d = pd.to_datetime(v, errors="coerce")
        if pd.isna(d):
            return None, None
        return d.month, d.day
    except:
        return None, None


# ─────────────────────────────
# 거래유형 정리
# ─────────────────────────────
def normalize_trade_type(t):
    t = str(t)

    if "매도" in t:
        return "SELL"
    if "매수" in t:
        return "BUY"
    if "예탁금" in t:
        return "INTEREST"
    if "입고" in t:
        return "StockCredit"
    if "입금" in t:
        return "Credit"
    if "출금" in t:
        return "Debit"

    return None


# ─────────────────────────────
# 종목명 정리
# ─────────────────────────────
def extract_stock_name(v):
    v = str(v)
    v = v.replace(" ", "")
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

    return {
        extract_stock_name(name): (str(code), str(name))
        for code, name in zip(df.iloc[:, 0], df.iloc[:, 1])
    }


def get_broker_info(stock, broker_map):
    key = extract_stock_name(stock)
    return broker_map.get(key, ("", stock))


# ─────────────────────────────
# 🔥 핵심 파서 (0건 해결 핵심)
# ─────────────────────────────
def parse_trades(df):

    trades = []

    # 1) "날짜처럼 보이는 행" 찾기
    start_idx = None

    for i in range(len(df)):
        row = df.iloc[i].astype(str).values

        if any(is_date_like(v) for v in row):
            start_idx = i
            break

    if start_idx is None:
        return []

    # 2) 데이터 파싱
    i = start_idx

    while i < len(df):

        row = df.iloc[i].astype(str).values

        # 날짜 아닌 행 skip
        if not any(is_date_like(v) for v in row):
            i += 1
            continue

        try:
            date_val = None

            for v in row:
                if is_date_like(v):
                    date_val = v
                    break

            m, d = parse_date_safe(date_val)
            if not m:
                i += 1
                continue

            # 안전 index 접근
            trade_type = clean(row[1]) if len(row) > 1 else ""
            stock = clean(row[2]) if len(row) > 2 else ""
            qty = to_int(row[3]) if len(row) > 3 else 0
            price = to_int(row[4]) if len(row) > 4 else 0
            net = to_int(row[5]) if len(row) > 5 else 0
            fee = to_int(row[6]) if len(row) > 6 else 0
            tax = to_int(row[7]) if len(row) > 7 else 0

            if not trade_type:
                i += 1
                continue

            trades.append({
                "month": m,
                "day": d,
                "type": trade_type,
                "stock": stock,
                "qty": qty,
                "price": price,
                "net": net,
                "fee": fee,
                "tax": tax
            })

        except:
            pass

        i += 1

    return trades


# ─────────────────────────────
# 전표 생성
# ─────────────────────────────
def process_trades(trades, broker_map, broker_code):

    rows = []

    for t in trades:

        m = t["month"]
        d = t["day"]
        ttype = normalize_trade_type(t["type"])

        if not ttype:
            continue

        stock = t["stock"]
        qty = t["qty"]
        price = t["price"]
        net = t["net"]
        fee = t["fee"]

        cp_code, cp_name = get_broker_info(stock, broker_map)
        memo = f"{stock}"

        if ttype == "SELL":
            rows.append([m,d,"차변",12500,"예치금",broker_code,"",memo,net,0])
            rows.append([m,d,"대변",10700,"단기매매증권",cp_code,cp_name,memo,0,qty*price])

        elif ttype == "BUY":
            cost = qty * price
            rows.append([m,d,"차변",10700,"단기매매증권",cp_code,cp_name,memo,cost,0])
            rows.append([m,d,"차변",82800,"수수료",cp_code,cp_name,"수수료",fee,0])
            rows.append([m,d,"대변",12500,"예치금",broker_code,"",memo,0,cost-fee])

        elif ttype == "INTEREST":
            rows.append([m,d,"차변",12500,"예치금",broker_code,"",memo,net,0])
            rows.append([m,d,"대변",42000,"이자수익","","",memo,0,net])

    return rows


# ─────────────────────────────
# 엑셀 생성
# ─────────────────────────────
def create_excel(rows):

    wb = openpyxl.Workbook()
    ws = wb.active

    headers = ["월","일","구분","계정코드","계정","거래처코드","거래처","적요","차변","대변"]
    ws.append(headers)

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
broker_code = st.text_input("증권사 코드")
uploaded = st.file_uploader("엑셀")

if uploaded:

    if st.button("변환"):

        broker_map = load_broker_map(broker_file)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(uploaded.read())
        tmp.close()

        df = pd.read_excel(tmp.name, header=None)

        trades = parse_trades(df)

        st.write("파싱 결과:", len(trades))

        rows = process_trades(trades, broker_map, broker_code)

        if not rows:
            st.error("0건")
        else:
            out = create_excel(rows)

            st.download_button("다운로드", out, "result.xlsx")
