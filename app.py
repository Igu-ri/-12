import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="wide")
st.title("📊 더존 위하고 전표 변환기")

# ─────────────────────────────
# 기본 유틸
# ─────────────────────────────
def to_int(v):
    try:
        if v is None or pd.isna(v):
            return 0
        if isinstance(v, str):
            v = re.sub(r"[^0-9.-]", "", v)
        return int(float(v))
    except:
        return 0


def clean(v):
    if v is None or pd.isna(v):
        return ""
    return str(v).strip()


def parse_date(v):
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except:
        return None, None


# ─────────────────────────────
# 거래유형
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
# 헤더 안정화 (핵심)
# ─────────────────────────────
def merge_multi_header(df, header_row, max_depth=3):

    header_block = df.iloc[header_row:header_row+max_depth]

    header_rows = []

    for i in range(len(header_block)):
        row = header_block.iloc[i].astype(str)

        numeric_ratio = sum(
            v.replace(".", "", 1).isdigit()
            for v in row
            if v not in [None, "", "nan"]
        ) / max(len(row), 1)

        if numeric_ratio > 0.5:
            break

        header_rows.append(row)

    header_df = pd.DataFrame(header_rows).fillna("")

    new_cols = []

    for c in range(df.shape[1]):
        parts = []

        for r in range(len(header_df)):
            v = str(header_df.iloc[r, c]).strip()
            if v and v != "nan":
                parts.append(v)

        new_cols.append("_".join(parts) if parts else f"col_{c}")

    df.columns = new_cols

    return df.iloc[header_row + len(header_rows):].reset_index(drop=True)


# ─────────────────────────────
# 파서
# ─────────────────────────────
def parse_hantoo_sheet(df):

    header_row = None

    # 1️⃣ 헤더 찾기
    for i in range(min(15, len(df))):
        row = df.iloc[i].astype(str)
        if any("거래일" in v for v in row):
            header_row = i
            break

    if header_row is None:
        return []

    # 2️⃣ 단일 헤더 시도
    try:
        df1 = df.copy()
        df1.columns = df1.iloc[header_row]
        df1 = df1.iloc[header_row + 1:].reset_index(drop=True)

        if any("거래일" in str(c) for c in df1.columns):
            df = df1
        else:
            raise Exception()

    except:
        # 3️⃣ 멀티 헤더 fallback (PowerQuery 대체)
        df = merge_multi_header(df, header_row)

    # ─────────────────────────────
    # 컬럼 찾기
    # ─────────────────────────────
    def find(keys):
        for c in df.columns:
            for k in keys:
                if k in str(c):
                    return c
        return None

    c_date  = find(["거래일", "일자", "날짜"])
    c_type  = find(["구분", "적요", "내용"])
    c_stock = find(["종목"])
    c_qty   = find(["수량"])
    c_price = find(["단가"])

    c_amount = find(["정산금액"])
    c_net    = find(["거래금액", "입출금액", "금액"])

    trades = []

    for _, r in df.iterrows():
        try:
            m, d = parse_date(r.get(c_date))
            if not m:
                continue

            ttype_raw = clean(r.get(c_type))
            ttype = normalize_trade_type(ttype_raw)

            if not ttype:
                continue

            qty = to_int(r.get(c_qty))
            price = to_int(r.get(c_price))

            amount = (
                to_int(r.get(c_amount)) or
                to_int(r.get(c_net)) or
                (qty * price if qty and price else 0)
            )

            trades.append({
                "month": m,
                "day": d,
                "type": ttype,
                "stock": clean(r.get(c_stock)),
                "qty": qty,
                "price": price,
                "amount": amount
            })

        except:
            continue

    return trades


# ─────────────────────────────
# 전표 생성
# ─────────────────────────────
def process_trades(trades):

    rows = []

    for t in trades:

        m = t["month"]
        d = t["day"]
        stock = t["stock"]
        qty = t["qty"]
        price = t["price"]
        amount = t["amount"]
        ttype = t["type"]

        if ttype == "SELL":
            memo = f"{stock} 매도"
            rows.append([m,d,"차변",12500,"예치금","","",memo,amount,0])
            rows.append([m,d,"대변",10700,"단기매매증권","",stock,memo,0,amount])

        elif ttype == "BUY":
            memo = f"{stock} 매수"
            cost = qty * price
            rows.append([m,d,"차변",10700,"단기매매증권","",stock,memo,cost,0])
            rows.append([m,d,"대변",12500,"예치금","","",memo,0,cost])

        elif ttype == "INTEREST":
            memo = "예탁금이용료"
            rows.append([m,d,"차변",12500,"예치금","","",memo,amount,0])
            rows.append([m,d,"대변",42000,"이자수익","",stock,memo,0,amount])

        elif ttype == "StockCredit":
            memo = f"{stock} 입고"
            cost = qty * price
            rows.append([m,d,"차변",10700,"단기매매증권","",stock,memo,cost,0])
            rows.append([m,d,"대변",13100,"선급금","",stock,memo,0,cost])

        elif ttype == "Credit":
            memo = "입금"
            rows.append([m,d,"차변",12500,"예치금","",stock,memo,amount,0])
            rows.append([m,d,"대변",12500,"예치금","","미등록",memo,0,amount])

        elif ttype == "Debit":
            memo = "출금"
            rows.append([m,d,"차변",12500,"예치금","","미등록",memo,0,amount])
            rows.append([m,d,"대변",12500,"예치금","",stock,memo,amount,0])

    return rows


# ─────────────────────────────
# Excel 생성
# ─────────────────────────────
def create_excel(rows):

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
uploaded = st.file_uploader("엑셀 업로드", type=["xlsx"])

if uploaded:

    if st.button("변환 실행"):

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(uploaded.read())
        tmp.close()

        xl = pd.ExcelFile(tmp.name)

        all_trades = []

        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            all_trades.extend(parse_hantoo_sheet(df))

        rows = process_trades(all_trades)

        if not rows:
            st.error("❌ 변환 데이터 없음")
        else:
            out = create_excel(rows)
            st.success("완료")
            st.download_button("다운로드", data=out, file_name="result.xlsx")
