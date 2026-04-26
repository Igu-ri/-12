import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="wide")
st.title("📊 더존 위하고 전표 변환기")


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
    elif "예탁금" in ttype:
        return "INTEREST"
    elif "입고" in ttype:
        return "StockCredit"
    elif "입금" in ttype:
        return "Credit"
    elif "출금" in ttype:
        return "Debit"

    return None


# ─────────────────────────────
# row
# ─────────────────────────────
def row(m, d, div, acct_code, acct_name, cp_code, cp_name, memo, dr, cr):
    return [
        m, d, div,
        acct_code,
        acct_name,
        cp_code,
        cp_name,
        memo,
        dr,
        cr
    ]


# ─────────────────────────────
# 파서
# ─────────────────────────────
def parse_hantoo_sheet(df):

    header_row = None

    # 1️⃣ 헤더 찾기
    for i in range(min(15, len(df))):
        row_str = df.iloc[i].astype(str)
        if any("거래일" in str(v) for v in row_str):
            header_row = i
            break

    if header_row is None:
        return []

    # 2️⃣ 단일 헤더 우선
    try:
        df_single = df.copy()
        df_single.columns = df_single.iloc[header_row]
        df_single = df_single.iloc[header_row + 1:].reset_index(drop=True)

        cols = [str(c) for c in df_single.columns]

        if any("거래일" in c for c in cols) and any("종목" in c for c in cols):
            df = df_single
            st.write("✅ 단일 헤더")

        else:
            raise Exception()

    except:
        # 3️⃣ 멀티 헤더
        header_rows = []

        for i in range(header_row, min(header_row + 5, len(df))):
            row_values = df.iloc[i].astype(str)

            numeric_ratio = sum(
                str(v).replace('.', '', 1).isdigit()
                for v in row_values
                if pd.notna(v)
            ) / len(row_values)

            if numeric_ratio > 0.5:
                break

            header_rows.append(df.iloc[i])

        header_rows = pd.DataFrame(header_rows).fillna("")

        new_cols = []
        for col in range(df.shape[1]):
            parts = []
            for r in range(len(header_rows)):
                val = str(header_rows.iloc[r, col]).strip()
                if val and val != "nan":
                    parts.append(val)

            new_cols.append("_".join(parts))

        df.columns = new_cols
        df = df.iloc[header_row + len(header_rows):].reset_index(drop=True)

        st.write("🔁 멀티 헤더")

    st.write("컬럼:", df.columns.tolist())

    # ─────────────────────────────
    # 컬럼 찾기
    # ─────────────────────────────
    def find_col(keys):
        for c in df.columns:
            for k in keys:
                if k in str(c):
                    return c
        return None

    c_date = find_col(["거래일","일자","날짜"])
    c_type = find_col(["구분","적요","내용"])
    c_stock = find_col(["종목"])
    c_qty = find_col(["수량"])
    c_price = find_col(["단가"])

    c_amount = find_col(["정산금액","금액","거래금액","입출금액"])
    c_fee = find_col(["수수료"])
    c_tax = find_col(["세금","제세금"])

    trades = []

    # ─────────────────────────────
    # 데이터 생성
    # ─────────────────────────────
    for _, r in df.iterrows():
        try:
            m, d = parse_date(r.get(c_date))
            if not m:
                continue

            ttype_raw = clean(r.get(c_type))
            ttype = normalize_trade_type(ttype_raw)

            if not ttype:
                continue

            stock = clean(r.get(c_stock))

            qty = to_int(r.get(c_qty))
            price = to_int(r.get(c_price))

            amount = to_int(r.get(c_amount))
            fee = to_int(r.get(c_fee))
            tax = to_int(r.get(c_tax))

            # 🔥 금액 통일
            if amount:
                final_amount = amount
            elif qty and price:
                final_amount = qty * price
            else:
                final_amount = 0

            trades.append({
                "month": m,
                "day": d,
                "type": ttype,
                "stock": stock,
                "qty": qty,
                "price": price,
                "amount": final_amount,
                "fee": fee,
                "tax": tax
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
            rows.append(row(m,d,"차변",12500,"예치금","","",memo,amount,0))
            rows.append(row(m,d,"대변",10700,"단기매매증권","",stock,memo,0,amount))

        elif ttype == "BUY":
            memo = f"{stock} 매수"
            rows.append(row(m,d,"차변",10700,"단기매매증권","",stock,memo,amount,0))
            rows.append(row(m,d,"대변",12500,"예치금","","",memo,0,amount))

        elif ttype == "INTEREST":
            memo = "예탁금이용료"
            rows.append(row(m,d,"차변",12500,"예치금","","",memo,amount,0))
            rows.append(row(m,d,"대변",42000,"이자수익","",stock,memo,0,amount))

        elif ttype == "StockCredit":
            memo = f"{stock} 입고"
            rows.append(row(m,d,"차변",10700,"단기매매증권","",stock,memo,amount,0))
            rows.append(row(m,d,"대변",13100,"선급금","",stock,memo,0,amount))

        elif ttype == "Credit":
            memo = "입금"
            rows.append(row(m,d,"차변",12500,"예치금","",stock,memo,amount,0))
            rows.append(row(m,d,"대변",12500,"예치금","","미등록",memo,0,amount))

        elif ttype == "Debit":
            memo = "출금"
            rows.append(row(m,d,"차변",12500,"예치금","","미등록",memo,0,amount))
            rows.append(row(m,d,"대변",12500,"예치금","",stock,memo,amount,0))

    return rows


# ─────────────────────────────
# 엑셀 생성
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
            trades = parse_hantoo_sheet(df)
            all_trades.extend(trades)

        rows = process_trades(all_trades)

        if not rows:
            st.error("❌ 변환 데이터 없음")
        else:
            out = create_excel(rows)

            st.success("완료")
            st.download_button("다운로드", data=out, file_name="result.xlsx")
