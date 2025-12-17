
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# -------------------------------
# Helper functions
# -------------------------------

COL_START = 1   # A
COL_END = 11    # K

def read_sheet_values_and_fill(file, sheet_name=None):
    """
    Read cell values and whether a cell is 'filled' (has a solid/background fill)
    for columns A..K and all used rows. Returns:
      - values: dict[(row, col)] -> cell value
      - filled: dict[(row, col)] -> bool (True if cell has a fill color)
      - max_row: maximum used row among A..K in this sheet
    """
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    # Determine max_row by scanning A..K for last non-empty cell
    max_row = 0
    for col in range(COL_START, COL_END + 1):
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value not in (None, ""):
                if row > max_row:
                    max_row = row
    # If everything is empty, default to ws.max_row to allow scanning
    if max_row == 0:
        max_row = ws.max_row

    values = {}
    filled = {}
    for row in range(1, max_row + 1):
        for col in range(COL_START, COL_END + 1):
            c = ws.cell(row=row, column=col)
            values[(row, col)] = c.value

            # Determine if the cell has any background fill set
            f = c.fill
            is_filled = False
            if f is not None and getattr(f, "patternType", None):
                # patternType like 'solid' or others means it is intentionally filled
                if f.patternType.lower() != "none":
                    # Some files use theme or indexed colors. We treat non-None pattern as filled.
                    is_filled = True
            filled[(row, col)] = is_filled
    return values, filled, max_row


def is_text(x):
    return isinstance(x, str)


def compare_workbooks(
    file_old, file_new, sheet_old=None, sheet_new=None, exclude_filled_cells=True, trim_spaces=True, case_sensitive=True
):
    """
    Compare two Excel workbooks over columns A..K and return a DataFrame of text changes.
    Only report when both old and new are strings and not equal (after optional trimming/case handling).
    If exclude_filled_cells is True, any cell that is filled in either workbook is ignored.
    """
    old_vals, old_filled, max_row_old = read_sheet_values_and_fill(file_old, sheet_old)
    new_vals, new_filled, max_row_new = read_sheet_values_and_fill(file_new, sheet_new)
    max_row = max(max_row_old, max_row_new)

    records = []
    for row in range(1, max_row + 1):
        for col in range(COL_START, COL_END + 1):
            old_v = old_vals.get((row, col))
            new_v = new_vals.get((row, col))

            if exclude_filled_cells:
                if old_filled.get((row, col), False) or new_filled.get((row, col), False):
                    continue

            if is_text(old_v) and is_text(new_v):
                a = old_v
                b = new_v

                if trim_spaces:
                    a = a.strip()
                    b = b.strip()
                if not case_sensitive:
                    a = a.lower()
                    b = b.lower()

                if a != b:
                    col_letter = get_column_letter(col)
                    # Create a human-readable Korean message
                    msg = f"{col_letter}열 {row}행 '{old_v}'에서 '{new_v}'(으)로 변경됐음"
                    records.append({
                        "열": col_letter,
                        "행": row,
                        "원본값": old_v,
                        "변경값": new_v,
                        "메시지": msg,
                    })

    df = pd.DataFrame.from_records(records, columns=["열", "행", "원본값", "변경값", "메시지"])
    return df


def list_sheets(file):
    wb = load_workbook(file, read_only=True, data_only=True)
    return wb.sheetnames


# -------------------------------
# Streamlit UI
# -------------------------------

st.set_page_config(page_title="엑셀 텍스트 변경 감지기 (A~K열)", layout="wide")

st.title("📘 엑셀 텍스트 변경 감지기 (A~K열)")
st.caption("두 개의 엑셀 파일을 비교해 A열부터 K열까지의 **문자값** 변경만 감지합니다. 색상(채우기)만 바뀐 경우는 제외합니다.")

col_u1, col_u2 = st.columns(2)
with col_u1:
    file_old = st.file_uploader("기준(이전) 엑셀 파일 업로드", type=["xlsx"])
with col_u2:
    file_new = st.file_uploader("비교(이후) 엑셀 파일 업로드", type=["xlsx"])

advanced = st.expander("고급 옵션")
with advanced:
    exclude_filled_cells = st.checkbox("색이 칠해진 셀은 변경 검출에서 제외", value=True)
    trim_spaces = st.checkbox("앞뒤 공백 무시", value=True)
    case_sensitive = st.checkbox("대소문자 구분", value=True)

sheet_old = None
sheet_new = None

if file_old and file_new:
    try:
        sheets_old = list_sheets(file_old)
        sheets_new = list_sheets(file_new)
    except Exception as e:
        st.error(f"시트 목록을 불러오는 데 실패했습니다: {e}")
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        sheet_old = st.selectbox("기준(이전) 시트 선택", options=sheets_old, index=0)
    with c2:
        sheet_new = st.selectbox("비교(이후) 시트 선택", options=sheets_new, index=0)

    if st.button("변경 사항 분석 실행", type="primary"):
        try:
            df = compare_workbooks(
                file_old, file_new,
                sheet_old=sheet_old, sheet_new=sheet_new,
                exclude_filled_cells=exclude_filled_cells,
                trim_spaces=trim_spaces,
                case_sensitive=case_sensitive
            )
            st.success(f"검출된 변경 건수: {len(df)}")
            st.dataframe(df, use_container_width=True, hide_index=True)

            # Provide downloadable CSV and Excel
            if not df.empty:
                csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "결과 CSV 다운로드",
                    data=csv_bytes,
                    file_name="excel_text_changes.csv",
                    mime="text/csv"
                )

                # Excel export
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="text_changes")
                st.download_button(
                    "결과 엑셀(xlsx) 다운로드",
                    data=output.getvalue(),
                    file_name="excel_text_changes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Show example message for clarity
            st.info("예시: **K열 3행 'S'에서 'M'(으)로 변경됐음**")
        except Exception as e:
            st.exception(e)
else:
    st.info("두 개의 xlsx 파일을 업로드하면 비교를 진행할 수 있습니다.")
