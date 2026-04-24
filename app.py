import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="wide")
st.title("📊 더존 위하고 전표 변환기")


# ─────────────────────────────
# 숫자 변환
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
# 🔥 거래유형 정규화 (핵심 추가)
# ─────────────────────────────
def normalize_trade_type(ttype):
    ttype = str(ttype)

    if "매도" in ttype:
        return "SELL"
    elif "매수" in ttype:
        return "BUY"
    return None

# ─────────────────────────────
# row (엑셀 10컬럼 기준)
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
# HANTOO 파서 (자동 컬럼 매칭)
# ─────────────────────────────
def parse_hantoo_sheet(df):
    header_row = None

    for i in range(min(15, len(df))):
        row_str = df.iloc[i].astype(str)
        if any("거래일" in v for v in row_str):
            header_row = i
            break

    if header_row is None:
        return []

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)

    def find_col(keys):
        for c in df.columns:
            for k in keys:
                if k in str(c):
                    return c
        return None

    c_date  = find_col(["거래일","거래일자", "일자", "날짜"])
    c_type  = find_col(["구분", "적요명","내용"])
    c_stock = find_col(["종목","종목명(거래상대명)"])
    c_qty   = find_col(["수량"])
    c_price = find_col(["단가", "가격"])
    c_net   = find_col(["금액"])

    trades = []

    for _, r in df.iterrows():
        try:
            m, d = parse_date(r.get(c_date))
            if not m:
                continue

            trade_type = clean(r.get(c_type))
            stock = clean(r.get(c_stock))

            qty = to_int(r.get(c_qty))
            price = to_int(r.get(c_price))
            net = to_int(r.get(c_net))

            if not trade_type:
                continue

            trades.append({
                "month": m,
                "day": d,
                "type": trade_type,
                "stock": stock,
                "qty": qty,
                "price": price,
                "net": net
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
        net = t["net"]

        # 🔥 정규화 적용
        ttype = normalize_trade_type(t["type"])

        if not ttype:
            continue

        # 매도
        if ttype == "SELL":
            memo = f"{stock}({qty}주*{price})매도"

            rows.append(row(m,d,"차변",12500,"예치금","",stock,memo,net,0))
            rows.append(row(m,d,"대변",10700,"단기매매증권","",stock,memo,0,qty*price))

        # 매수
        elif ttype == "BUY":
            cost = qty * price
            memo = f"{stock}({qty}주*{price})매수"

            rows.append(row(m,d,"차변",10700,"단기매매증권","",stock,memo,cost,0))
            rows.append(row(m,d,"대변",12500,"예치금","",stock,memo,0,cost))

    return rows

# ─────────────────────────────
# Excel 생성
# ─────────────────────────────
def create_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append([
        "1.월","2.일","3.구분",
        "4.계정과목코드","5.계정과목명",
        "6.거래처코드","7.거래처명",
        "8.적요명","9.차변(출금)","10.대변(입금)"
    ])

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
            st.download_button(
                "다운로드",
                data=out,
                file_name="result.xlsx"
            )
