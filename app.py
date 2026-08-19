import io
import re
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

try:
    import xlrd
except ImportError:
    xlrd = None

st.set_page_config(page_title="행사제안서 자동생성기", layout="wide")

# ---------------------------------------------------------------
# 고정 컬럼 위치 (0-indexed, A=0)
# ---------------------------------------------------------------
COL = {"E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "M": 12, "N": 13}

EVENT_OPTIONS = [
    ("최저가", "최저가"),
    ("상시가", "상시할인가"),
    ("day7", "7일 행사가(폐쇄몰)"),
    ("day3", "3일 행사가"),
    ("공동구매가", "공동구매가"),
    ("기타", "기타(직접입력)"),
]
EVENT_LABEL = dict(EVENT_OPTIONS)
EVENT_KEYS = [k for k, _ in EVENT_OPTIONS]

# 단품 판단에서 "묶음/세트"로 간주할 보조 키워드 (수량 숫자 패턴이 없을 때만 사용)
BUNDLE_HINT_WORDS = ["세트", "SET", "묶음", "구성"]


# ---------------------------------------------------------------
# 유틸
# ---------------------------------------------------------------
def get(row, idx):
    if idx is None:
        return ""
    if idx < len(row):
        v = row[idx]
        return "" if v is None else v
    return ""


def to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("₩", "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def normalize_code(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s == "":
        return ""
    try:
        f = float(s)
        if re.match(r"^-?\d+(\.0+)?$", s):
            return str(int(f))
    except ValueError:
        pass
    return s


def extract_qty(row):
    """행 전체 텍스트에서 수량(개입) 힌트를 찾아 단품 여부를 판단.
    '1개'면 1, '2개입'/'3개' 등이면 그 숫자, 아무 힌트가 없으면 단품(1)으로 간주."""
    text = " ".join(str(c) for c in row if c not in (None, ""))
    nums = [int(m) for m in re.findall(r"(\d+)\s*개", text)]
    if nums:
        return min(nums)
    if any(w in text for w in BUNDLE_HINT_WORDS):
        return 99  # 세트/묶음 힌트만 있고 숫자가 없으면 단품이 아닌 것으로 취급
    return 1


def read_rows(file_bytes, ext):
    """엑셀 파일을 항상 A열=0 기준 절대 위치의 2차원 리스트로 읽는다."""
    if ext == "xlsx":
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        return rows
    elif ext == "xls":
        if xlrd is None:
            raise RuntimeError("xlrd 라이브러리가 설치되어 있지 않습니다 (requirements.txt 확인).")
        wb = xlrd.open_workbook(file_contents=file_bytes)
        sheet = wb.sheet_by_index(0)
        rows = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        return rows
    else:
        raise RuntimeError("지원하지 않는 파일 형식입니다 (.xls 또는 .xlsx만 가능).")


def find_header_row(rows, keywords, min_matches=2):
    best_row, best_map = -1, {}
    for r in range(min(len(rows), 20)):
        row = rows[r]
        col_map = {}
        for ci, cell in enumerate(row):
            v = str(cell).strip() if cell is not None else ""
            for k in keywords:
                if k in v and k not in col_map:
                    col_map[k] = ci
        if len(col_map) >= min_matches:
            return r, col_map
        if len(col_map) > len(best_map):
            best_row, best_map = r, col_map
    return best_row, best_map


def parse_cost_sheet(rows):
    """.xls 원가 파일: 품번, 품명, 원가는 헤더 텍스트로 탐색.
    같은 품번이 여러 줄(묶음 수량별)로 존재하면 단품(1개) 기준 행을 우선 채택."""
    header_row, col_map = find_header_row(rows, ["품번", "품명", "원가"], 2)
    result = {}
    if header_row == -1:
        return result
    c_name = col_map.get("품명")
    c_code = col_map.get("품번")
    c_cost = col_map.get("원가")
    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        if not row:
            continue
        code = normalize_code(get(row, c_code)) if c_code is not None else ""
        if not code:
            continue
        name = str(get(row, c_name)).strip() if c_name is not None else ""
        cost_val = to_num(get(row, c_cost)) if c_cost is not None else None
        qty = extract_qty(row)
        if code not in result or qty < result[code]["_qty"]:
            result[code] = {"품명": name, "원가": cost_val, "_qty": qty}
    return result


def parse_price_sheet(rows):
    """.xlsx 가격표: 상품코드 컬럼은 헤더 탐색, E~J/M/N은 고정 위치.
    같은 상품코드가 여러 줄(묶음 수량별)로 존재하면 단품(1개) 기준 행을 우선 채택."""
    header_row, col_map = find_header_row(rows, ["상품코드", "품번"], 1)
    if header_row == -1:
        header_row = 0
    c_code = col_map.get("상품코드", col_map.get("품번", 0))
    result = {}
    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        if not row:
            continue
        code = normalize_code(get(row, c_code))
        if not code:
            continue
        qty = extract_qty(row)
        if code in result and qty >= result[code]["_qty"]:
            continue  # 이미 더 단품에 가까운 행을 채택한 상태면 건너뜀
        rocket_raw = str(get(row, COL["M"])).strip()
        arrival_raw = str(get(row, COL["N"])).strip()
        result[code] = {
            "정상가": to_num(get(row, COL["E"])),
            "최저가": to_num(get(row, COL["F"])),
            "상시가": to_num(get(row, COL["G"])),
            "day7": to_num(get(row, COL["H"])),
            "day3": to_num(get(row, COL["I"])),
            "공동구매가": to_num(get(row, COL["J"])),
            # 로켓배송: 원본 값(운영/위수탁 등)을 그대로 표기, 공란이면 '-'
            "rocket": rocket_raw if rocket_raw not in ("", "None") else "-",
            "arrival": "도착보장" if arrival_raw not in ("", "None") else "-",
            "_qty": qty,
        }
    return result


def calc_row(r, cost_raw):
    event_price = r["기타금액"] if r["적용타입"] == "기타" else r.get(r["적용타입"])
    price_missing = event_price is None  # 선택한 항목의 금액이 비어있음
    discount_rate = (1 - event_price / r["정상가"]) if r.get("정상가") and event_price is not None else None
    discount_amt = (r["정상가"] - event_price) if r.get("정상가") is not None and event_price is not None else None
    cost_vat_in = cost_raw
    cost_vat_ex = round(cost_raw * 1.1) if cost_raw is not None else None
    fee_dec = (r.get("수수료") or 0) / 100
    supply_vat_in = round(event_price * (1 - fee_dec)) if event_price is not None else None
    supply_vat_ex = round(supply_vat_in * 1.1) if supply_vat_in is not None else None
    margin = (
        supply_vat_ex - cost_vat_ex - (r.get("배송비") or 0)
        if supply_vat_ex is not None and cost_vat_ex is not None
        else None
    )
    margin_rate = (margin / supply_vat_ex) if margin is not None and supply_vat_ex else None
    return {
        "event_price": event_price,
        "price_missing": price_missing,
        "할인율": discount_rate,
        "할인금액": discount_amt,
        "원가(+VAT)": cost_vat_in,
        "원가(-VAT)": cost_vat_ex,
        "공급가(+VAT)": supply_vat_in,
        "공급가(-VAT)": supply_vat_ex,
        "공헌이익": margin,
        "공헌이익률": margin_rate,
    }


def pick_default_type(price):
    """가격이 채워진 항목 중 우선순위대로 기본 적용유형을 고른다."""
    if not price:
        return "day7"
    for k in ["day7", "day3", "최저가", "상시가", "공동구매가"]:
        if price.get(k) is not None:
            return k
    return "day7"


# ---------------------------------------------------------------
# UI
# ---------------------------------------------------------------
st.title("행사제안서 자동생성기")
st.caption("기준 파일 업로드 → 품번 입력 → 수수료/배송비 설정 → 엑셀 다운로드")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 기준 파일 업로드")
    st.caption(".xls = 원가 파일(오클릭)  |  .xlsx = 가격표 파일. 여러 개 업로드 가능합니다.")
    uploaded = st.file_uploader(
        "파일을 드래그하거나 클릭하여 업로드",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )

cost_map, price_map = {}, {}
file_status = []
if uploaded:
    for f in uploaded:
        ext = f.name.lower().rsplit(".", 1)[-1]
        try:
            rows = read_rows(f.getvalue(), ext)
            if ext == "xls":
                m = parse_cost_sheet(rows)
                cost_map.update(m)
                file_status.append((f.name, "원가 파일", len(m)))
            else:
                m = parse_price_sheet(rows)
                price_map.update(m)
                file_status.append((f.name, "가격표 파일", len(m)))
        except Exception as e:
            file_status.append((f.name, f"⚠️ 읽기 실패: {e}", 0))

with col1:
    for name, kind, cnt in file_status:
        st.write(f"- **{name}** — {kind} ({cnt}건 인식)")

with col2:
    st.subheader("2. 품번 입력")
    sku_text = st.text_area(
        "품번을 입력하세요. 엔터(줄바꿈) 또는 콤마로 여러 개 구분",
        height=120,
        placeholder="예) 102649, 102650\n102761",
    )
    lookup = st.button("조회하기", type="primary")

if "rows" not in st.session_state:
    st.session_state.rows = []

if lookup:
    seen, skus = set(), []
    for tok in re.split(r"[\n,，]+", sku_text):
        code = normalize_code(tok)
        if code and code not in seen:
            seen.add(code)
            skus.append(code)
    prev = {r["품번"]: r for r in st.session_state.rows}
    new_rows = []
    for sku in skus:
        cost = cost_map.get(sku)
        price = price_map.get(sku)
        p = prev.get(sku, {})
        new_rows.append(
            {
                "품번": sku,
                "상품코드": sku,
                "품명": cost["품명"] if cost else "",
                "정상가": price["정상가"] if price else None,
                "최저가": price["최저가"] if price else None,
                "상시가": price["상시가"] if price else None,
                "day7": price["day7"] if price else None,
                "day3": price["day3"] if price else None,
                "공동구매가": price["공동구매가"] if price else None,
                "rocket": price["rocket"] if price else "-",
                "arrival": price["arrival"] if price else "-",
                "matched": bool(cost or price),
                "적용타입": p.get("적용타입", pick_default_type(price)),
                "기타금액": p.get("기타금액", None),
                "수수료": p.get("수수료", 0.0),
                "배송비": p.get("배송비", 0.0),
            }
        )
    st.session_state.rows = new_rows

st.subheader("3. 수수료 · 배송비 일괄 적용")
b1, b2, b3 = st.columns([1, 1, 2])
with b1:
    bulk_fee = st.number_input("수수료 일괄(%)", value=0.0, step=0.5)
with b2:
    bulk_ship = st.number_input("배송비 일괄(원)", value=0.0, step=100.0)
with b3:
    st.write("")
    st.write("")
    if st.button("전체 행에 적용"):
        for r in st.session_state.rows:
            r["수수료"] = bulk_fee
            r["배송비"] = bulk_ship

st.subheader("4. 결과표")

if not st.session_state.rows:
    st.info("기준 파일을 업로드하고 품번을 조회하면 결과가 표시됩니다.")
else:
    unmatched = [r["품번"] for r in st.session_state.rows if not r["matched"]]
    if unmatched:
        st.warning("기준 파일에서 매칭되지 않은 품번: " + ", ".join(unmatched))

    edit_df = pd.DataFrame(st.session_state.rows)

    st.caption(
        "※ 적용유형에서 금액이 비어있는 항목을 고르면 아래 결과표에 노란색으로 표시됩니다 "
        "(선택 자체는 가능하지만 금액이 없어 계산에서 제외됩니다)."
    )

    edited = st.data_editor(
        edit_df,
        column_order=[
            "품번", "상품코드", "품명", "정상가", "최저가", "상시가",
            "day7", "day3", "공동구매가", "적용타입", "기타금액", "수수료", "배송비",
            "rocket", "arrival",
        ],
        column_config={
            "품번": "품번",
            "상품코드": "상품코드",
            "품명": "품명",
            "정상가": st.column_config.NumberColumn("정상가(판매가)"),
            "최저가": "최저가",
            "상시가": "상시할인가",
            "day7": "7일 행사가(폐쇄몰)",
            "day3": "3일 행사가",
            "공동구매가": "공동구매가",
            "적용타입": st.column_config.SelectboxColumn("적용유형(선택)", options=EVENT_KEYS, required=True),
            "기타금액": st.column_config.NumberColumn("기타 금액(적용유형=기타일 때)"),
            "수수료": st.column_config.NumberColumn("수수료(%)"),
            "배송비": st.column_config.NumberColumn("배송비"),
            "rocket": st.column_config.TextColumn("로켓배송", disabled=True),
            "arrival": st.column_config.TextColumn("도착배송", disabled=True),
        },
        disabled=["품번", "상품코드", "품명", "정상가", "최저가", "상시가", "day7", "day3", "공동구매가", "rocket", "arrival"],
        hide_index=True,
        use_container_width=True,
        key="editor",
    )
    st.session_state.rows = edited.to_dict("records")

    display_rows = []
    missing_flags = []
    for r in st.session_state.rows:
        cost_raw = cost_map.get(r["품번"], {}).get("원가")
        c = calc_row(r, cost_raw)
        missing_flags.append(c["price_missing"])
        display_rows.append(
            {
                "품번": r["품번"],
                "상품코드": r["상품코드"],
                "품명": r["품명"],
                "정상가(판매가)": r["정상가"],
                "최저가": r["최저가"],
                "상시할인가": r["상시가"],
                "7일 행사가(폐쇄몰)": r["day7"],
                "3일 행사가": r["day3"],
                "공동구매가": r["공동구매가"],
                "적용행사가": c["event_price"] if c["event_price"] is not None else "확인필요",
                "할인율": f"{c['할인율']*100:.1f}%" if c["할인율"] is not None else "-",
                "할인금액": c["할인금액"],
                "원가(+VAT)": c["원가(+VAT)"],
                "원가(-VAT)": c["원가(-VAT)"],
                "수수료(%)": r["수수료"],
                "배송비": r["배송비"],
                "공급가(+VAT)": c["공급가(+VAT)"],
                "공급가(-VAT)": c["공급가(-VAT)"],
                "공헌이익": c["공헌이익"],
                "공헌이익률": f"{c['공헌이익률']*100:.1f}%" if c["공헌이익률"] is not None else "-",
                "로켓배송": r["rocket"],
                "도착배송": r["arrival"],
            }
        )
    result_df = pd.DataFrame(display_rows)

    def highlight_missing(row):
        idx = row.name
        if idx < len(missing_flags) and missing_flags[idx]:
            return ["background-color: #fff3b0"] * len(row)
        return [""] * len(row)

    st.dataframe(result_df.style.apply(highlight_missing, axis=1), use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="행사제안서")
    st.download_button(
        "엑셀로 다운로드",
        data=buf.getvalue(),
        file_name=f"행사제안서_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
