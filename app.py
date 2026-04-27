import streamlit as st
import pandas as pd
import openpyxl
import io, tempfile, re
from openpyxl.styles import PatternFill

st.set_page_config(page_title="더존 전표 변환기", page_icon="📊", layout="wide")
st.title("📊 더존 위하고 전표 변환기")

# ─────────────────────────────
# 🔥 공통 유틸 (데이터 정리)
# ─────────────────────────────

def to_int(v):
    """숫자만 추출해서 int 변환"""
    try:
        if v is None:
            return 0
        if isinstance(v, str):
            v = re.sub(r"[^0-9.-]", "", v)
        return int(float(v))
    except:
        return 0


def clean(v):
    """NaN / 공백 제거"""
    try:
        if pd.isna(v):
            return ''
    except:
        pass
    return str(v).strip() if v is not None else ''


def parse_date(v):
    """월/일 추출"""
    try:
        d = pd.to_datetime(v)
        return d.month, d.day
    except:
        return None, None


# ─────────────────────────────
# 🔥 거래유형 표준화 (핵심)
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
    elif "출금" in ttype and ("이체" in ttype or "은행" in ttype):
        return "Debit"

    return None


# ─────────────────────────────
# 🔥 종목명 정리 (핵심 업그레이드)
# ─────────────────────────────
def extract_stock_name(name):
    """
    예:
    코스닥#메쥬 → 메쥬
    코스피#삼성전자 → 삼성전자
    """
    name = str(name).replace(" ", "").strip()
    return name.split("#")[-1] if "#" in name else name


# ─────────────────────────────
# 🔥 거래처 매핑 (기존 유지 + 강화)
# ─────────────────────────────
def load_broker_map(file):
    """
    거래처 매핑 엑셀:
    A열 = 코드
    B열 = 이름
    """
    if file is None:
        return {}

    df = pd.read_excel(file)

    mapping = {}
    for code, name in zip(df.iloc[:, 0], df.iloc[:, 1]):
        key = extract_stock_name(name)  # 🔥 종목명 기준 통일
        mapping[key] = (str(code).strip(), str(name).strip())

    return mapping


def get_broker_info(stock, broker_map):
    key = extract_stock_name(stock)

    if key in broker_map:
        return broker_map[key]

    return "", stock  # 미매핑


# ─────────────────────────────
# 🔥 엑셀 컬럼 자동 탐지
# ─────────────────────────────
def find_col(df, keys):
    for c in df.columns:
        for k in keys:
            if k in str(c):
                return c
    return None


# ─────────────────────────────
# 🔥 한투/증권사 공통 파서 (핵심)
# ─────────────────────────────
def parse_trades(df):
    header_row = None

    # 1️⃣ 헤더 찾기 (멀티 헤더 대응)
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
    c_type  = find_col(df, ["구분", "적요", "거래종류"])
    c_stock = find_col(df, ["종목", "종목명"])
    c_qty   = find_col(df, ["수량", "거래수량"])
    c_price = find_col(df, ["단가", "가격"])
    c_net   = find_col(df, ["금액", "거래금액", "입출금액"])
    c_fee   = find_col(df, ["수수료"])
    c_tax   = find_col(df, ["세금", "tax"])

    trades = []

    for _, r in df.iterrows():
        try:
            m, d = parse_date(r.get(c_date))
            if not m:
                continue

            trades.append({
                "month": m,
                "day": d,
                "type": clean(r.get(c_type)),
                "stock": clean(r.get(c_stock)),
                "qty": to_int(r.get(c_qty)),
                "price": to_int(r.get(c_price)),
                "net": to_int(r.get(c_net)),
                "fee": to_int(r.get(c_fee)),
                "tax": to_int(r.get(c_tax)),
            })

        except:
            continue

    return trades


# ─────────────────────────────
# 🔥 전표 생성 (핵심 로직)
# ─────────────────────────────
def process_trades(trades, broker_map, broker_code):
    rows = []

    for t in trades:

        m, d = t["month"], t["day"]

        # 🔥 핵심: 거래 타입 표준화
        ttype = normalize_trade_type(t["type"])
        if not ttype:
            continue

        stock = extract_stock_name(t["stock"])
        qty   = t["qty"]
        price = t["price"]
        net   = t["net"]
        fee   = t["fee"]

        cp_code, cp_name = get_broker_info(stock, broker_map)

        # ───────────── SELL ─────────────
        if ttype == "SELL":
            memo = f"{stock} 매도"

            rows.append([m,d,"차변",12500,"예치금",broker_code,"",memo,net,0])
            rows.append([m,d,"대변",10700,"증권",cp_code,cp_name,memo,0,qty*price])

        # ───────────── BUY ─────────────
        elif ttype == "BUY":
            cost = qty * price
            memo = f"{stock} 매수"

            rows.append([m,d,"차변",10700,"증권",cp_code,cp_name,memo,cost,0])
            rows.append([m,d,"차변",82800,"수수료",cp_code,cp_name,"수수료",fee,0])
            rows.append([m,d,"대변",12500,"예치금",broker_code,"",memo,0,cost-fee])

        # ───────────── 이자 ─────────────
        elif ttype == "INTEREST":
            rows.append([m,d,"차변",12500,"예치금",broker_code,"",t["type"],net,0])
            rows.append([m,d,"대변",42000,"이자수익",broker_code,stock,t["type"],0,net])

        # ───────────── 입고 ─────────────
        elif ttype == "StockCredit":
            cost = qty * price
            memo = f"{stock} 입고"

            rows.append([m,d,"차변",10700,"증권",cp_code,cp_name,memo,cost,0])
            rows.append([m,d,"대변",13100,"선급금",cp_code,cp_name,memo,0,cost])

        # ───────────── 입금 ─────────────
        elif ttype == "Credit":
            rows.append([m,d,"차변",12500,"예치금",broker_code,"",t["type"],net,0])
            rows.append([m,d,"대변",12500,"예치금","","미등록",t["type"],0,net])

        # ───────────── 출금 ─────────────
        elif ttype == "Debit":
            rows.append([m,d,"차변",12500,"예치금","","미등록",t["type"],0,net])
            rows.append([m,d,"대변",12500,"예치금",broker_code,"",t["type"],net,0])

    return rows


# ─────────────────────────────
# 🔥 엑셀 생성
# ─────────────────────────────
def create_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active

    headers = [
        "월","일","구분",
        "계정과목코드","계정과목명",
        "거래처코드","거래처명",
        "적요","차변","대변"
    ]

    ws.append(headers)

    yellow = PatternFill(start_color="FFF59D", fill_type="solid")

    for c in ws[1]:
        c.fill = yellow

    for r in rows:
        ws.append(r)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ─────────────────────────────
# 🔥 UI
# ─────────────────────────────
broker_file = st.file_uploader("거래처 매핑", type=["xlsx"])
broker_code = st.text_input("증권사 코드")
uploaded = st.file_uploader("엑셀 업로드", type=["xlsx"])

if uploaded and st.button("변환"):

    broker_map = load_broker_map(broker_file)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(uploaded.read())
    tmp.close()

    xl = pd.ExcelFile(tmp.name)

    all_trades = []
    for s in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=s, header=None)
        all_trades += parse_trades(df)

    rows = process_trades(all_trades, broker_map, broker_code)

    if not rows:
        st.error("데이터 없음")
    else:
        out = create_excel(rows)

        st.success("완료")
        st.download_button("다운로드", data=out, file_name="result.xlsx")
