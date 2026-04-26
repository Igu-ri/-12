import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re

st.set_page_config(page_title="더존 전표 변환기", layout="wide")
st.title("📊 더존 위하고 전표 변환기")


# ─────────────────────────────
# 1️⃣ 기본 유틸
# ─────────────────────────────
def to_int(v):
    try:
        if v is None or pd.isna(v):
            return 0
        v = str(v)
        v = re.sub(r"[^0-9.-]", "", v)
        return int(float(v)) if v else 0
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
# 2️⃣ 거래 유형
# ─────────────────────────────
def normalize_type(t):
    t = str(t)

    if "매도" in t:
        return "SELL"
    if "매수" in t:
        return "BUY"
    if "예탁금" in t:
        return "INTEREST"
    if "입고" in t:
        return "STOCK"
    if "입금" in t:
        return "CREDIT"
    if "출금" in t:
        return "DEBIT"

    return None


# ─────────────────────────────
# 3️⃣ 안전 컬럼 탐색 (핵심)
# ─────────────────────────────
def find_col(df, keywords):
    for c in df.columns:
        c_str = str(c).replace(" ", "")
        for k in keywords:
            if k in c_str:
                return c
    return None


# ─────────────────────────────
# 4️⃣ 파서 (엑셀 → 표준 데이터)
# ─────────────────────────────
def parse(df):

    header_row = None

    for i in range(min(15, len(df))):
        row = df.iloc[i].astype(str)
        if any("거래일" in v for v in row):
            header_row = i
            break

    if header_row is None:
        return []

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    c_date  = find_col(df, ["거래일", "일자", "날짜"])
    c_type  = find_col(df, ["구분", "적요", "내용"])
    c_stock = find_col(df, ["종목"])
    c_qty   = find_col(df, ["수량"])
    c_price = find_col(df, ["단가"])

    c_trade  = find_col(df, ["거래금액"])
    c_cash   = find_col(df, ["입출금"])
    c_settle = find_col(df, ["정산"])

    result = []

    for _, r in df.iterrows():

        m, d = parse_date(r.get(c_date))
        if not m:
            continue

        ttype_raw = clean(r.get(c_type))
        ttype = normalize_type(ttype_raw)
        if not ttype:
            continue

        qty = to_int(r.get(c_qty))
        price = to_int(r.get(c_price))

        # 🔥 핵심: 금액 우선순위
        trade_amount  = to_int(r.get(c_trade))
        cash_amount   = to_int(r.get(c_cash))
        settle_amount = to_int(r.get(c_settle))

        amount = (
            trade_amount
            or cash_amount
            or settle_amount
            or (qty * price if qty and price else 0)
        )

        result.append({
            "month": m,
            "day": d,
            "type": ttype,
            "stock": clean(r.get(c_stock)),
            "qty": qty,
            "price": price,
            "amount": amount
        })

    return result


# ─────────────────────────────
# 5️⃣ 전표 생성
# ─────────────────────────────
def journal(data):

    rows = []

    for t in data:

        m = t["month"]
        d = t["day"]
        ttype = t["type"]
        stock = t["stock"]
        qty = t["qty"]
        price = t["price"]
        amount = t["amount"]

        if ttype == "SELL":
            memo = f"{stock} 매도"
            rows.append([m,d,"차변",12500,"예치금","","",memo,amount,0])
            rows.append([m,d,"대변",10700,"증권","","",memo,0,amount])

        elif ttype == "BUY":
            memo = f"{stock} 매수"
            cost = qty * price
            rows.append([m,d,"차변",10700,"증권","","",memo,cost,0])
            rows.append([m,d,"대변",12500,"예치금","","",memo,0,cost])

        elif ttype == "INTEREST":
            memo = "예탁금이용료"
            rows.append([m,d,"차변",12500,"예치금","","",memo,amount,0])
            rows.append([m,d,"대변",42000,"이자수익","","",memo,0,amount])

        elif ttype == "STOCK":
            memo = f"{stock} 입고"
            cost = qty * price
            rows.append([m,d,"차변",10700,"증권","","",memo,cost,0])
            rows.append([m,d,"대변",13100,"선급금","","",memo,0,cost])

        elif ttype == "CREDIT":
            memo = "입금"
            rows.append([m,d,"차변",12500,"예치금","","",memo,amount,0])
            rows.append([m,d,"대변",12500,"예치금","","미등록",memo,0,amount])

        elif ttype == "DEBIT":
            memo = "출금"
            rows.append([m,d,"차변",12500,"예치금","","미등록",memo,0,amount])
            rows.append([m,d,"대변",12500,"예치금","","",memo,amount,0])

    return rows


# ─────────────────────────────
# 6️⃣ 엑셀 생성
# ─────────────────────────────
def export(rows):

    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(["월","일","구분","계정코드","계정명","거래처코드","거래처명","적요","차변","대변"])

    for r in rows:
        ws.append(r)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ─────────────────────────────
# 7️⃣ UI
# ─────────────────────────────
file = st.file_uploader("엑셀 업로드", type=["xlsx"])

if file:

    if st.button("변환"):

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(file.read())
        tmp.close()

        xl = pd.ExcelFile(tmp.name)

        all_data = []

        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            all_data.extend(parse(df))

        rows = journal(all_data)

        if not rows:
            st.error("데이터 없음 (컬럼 매칭 실패)")
        else:
            out = export(rows)
            st.success("완료")
            st.download_button("다운로드", data=out, file_name="result.xlsx")
