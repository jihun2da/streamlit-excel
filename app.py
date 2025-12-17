
import streamlit as st
import pandas as pd
from collections import defaultdict, Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

COL_START = 1   # A
COL_END   = 11  # K
COLS      = [get_column_letter(c) for c in range(COL_START, COL_END+1)]

def normalize_value(v, trim_spaces=True, case_sensitive=True):
    if isinstance(v, str):
        s = v.strip() if trim_spaces else v
        return s if case_sensitive else s.lower()
    return v

def read_sheet_values(file, sheet_name=None, trim_spaces=True, case_sensitive=True):
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    max_row = 0
    for c in range(COL_START, COL_END+1):
        for r in range(1, ws.max_row+1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                max_row = max(max_row, r)
    if max_row == 0:
        max_row = ws.max_row
    rows = []
    for r in range(1, max_row+1):
        orig = {}
        norm = {}
        empty_all = True
        for c in range(COL_START, COL_END+1):
            v = ws.cell(row=r, column=c).value
            col = get_column_letter(c)
            orig[col] = v
            norm[col] = normalize_value(v, trim_spaces, case_sensitive)
            if v not in (None, ""):
                empty_all = False
        if not empty_all:
            rows.append({"_row": r, "orig": orig, "norm": norm})
    return rows

def fill_signature(fill):
    if fill is None:
        return ("none",)
    pt = getattr(fill, "patternType", None)
    if not pt or str(pt).lower() == "none":
        return ("none",)
    def color_tuple(c):
        if c is None:
            return None
        return (
            getattr(c, "type", None),
            getattr(c, "rgb", None),
            getattr(c, "indexed", None),
            getattr(c, "theme", None),
            getattr(c, "tint", None),
        )
    sig = (
        str(pt).lower(),
        color_tuple(getattr(fill, "fgColor", None)),
        color_tuple(getattr(fill, "start_color", None)),
        color_tuple(getattr(fill, "end_color", None)),
    )
    return sig

def read_sheet_fills(file, sheet_name=None):
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    max_row = 0
    for c in range(COL_START, COL_END+1):
        for r in range(1, ws.max_row+1):
            if ws.cell(row=r, column=c).value not in (None, ""):
                max_row = max(max_row, r)
    if max_row == 0:
        max_row = ws.max_row
    fills = {}
    for r in range(1, max_row+1):
        for c in range(COL_START, COL_END+1):
            fills[(r, c)] = fill_signature(ws.cell(row=r, column=c).fill)
    return fills

def row_tuple(norm_row):
    return tuple(norm_row[col] for col in COLS)

def best_pairing(new_rows, old_rows):
    candidates = []
    for i, o in enumerate(old_rows):
        for j, n in enumerate(new_rows):
            eq = sum(1 for col in COLS if o["norm"].get(col) == n["norm"].get(col))
            if eq > 0:
                candidates.append((eq, i, j))
    candidates.sort(reverse=True)
    used_old, used_new = set(), set()
    pairs = []
    for eq, i, j in candidates:
        if i in used_old or j in used_new:
            continue
        pairs.append((i, j, eq))
        used_old.add(i); used_new.add(j)
    leftover_old = [i for i in range(len(old_rows)) if i not in used_old]
    leftover_new = [j for j in range(len(new_rows)) if j not in used_new]
    return pairs, leftover_old, leftover_new

def build_diff_record(old_row, new_row):
    changes = []
    for col in COLS:
        ov = old_row["orig"].get(col)
        nv = new_row["orig"].get(col)
        if old_row["norm"].get(col) != new_row["norm"].get(col):
            changes.append(f"{col}열 '{ov}'→'{nv}'")
    msg = "; ".join(changes) if changes else "값 동일 (정규화 기준)"
    return {
        "기준행": old_row["_row"],
        "비교행": new_row["_row"],
        "변경요약": msg
    }

st.set_page_config(page_title="엑셀 행 재정렬 안전 비교 (A~K)", layout="wide")
st.title("📘 엑셀 행 재정렬 안전 비교 (A~K)")
st.caption("기준 엑셀을 먼저 업로드해 **행 내용을 저장**하고, 비교 엑셀을 올려 **정렬/순서 변경과 무관하게** 변경사항을 검출합니다.")

with st.expander("비교 옵션", expanded=True):
    trim_spaces = st.checkbox("앞뒤 공백 무시", value=True)
    case_sensitive = st.checkbox("대소문자 구분", value=True)
    exclude_rows_if_fill_changed = st.checkbox("색상(채우기) 변경된 행 제외", value=False)

st.subheader("1) 기준(이전) 파일 저장")
c1, c2 = st.columns(2)
with c1:
    file_old = st.file_uploader("기준 엑셀 파일", type=["xlsx"], key="old")
with c2:
    sheet_old = None
    if file_old:
        try:
            wb = load_workbook(file_old, read_only=True, data_only=True)
            sheet_old = st.selectbox("시트 선택(기준)", options=wb.sheetnames, index=0)
        except Exception as e:
            st.error(f"기준 파일 시트 읽기 실패: {e}")

if st.button("기준 데이터 저장", type="primary", disabled=not (file_old and sheet_old)):
    try:
        old_rows = read_sheet_values(file_old, sheet_old, trim_spaces, case_sensitive)
        st.session_state["old_rows"] = old_rows
        st.session_state["old_rows_norm_multiset"] = Counter([row_tuple(r["norm"]) for r in old_rows])
        from collections import defaultdict
        mapping = defaultdict(list)
        for idx, r in enumerate(old_rows):
            mapping[row_tuple(r["norm"])].append(idx)
        st.session_state["old_rows_by_tuple_indices"] = mapping
        if exclude_rows_if_fill_changed:
            st.session_state["old_fills"] = read_sheet_fills(file_old, sheet_old)
        st.success(f"기준 데이터 저장 완료: {len(old_rows)} 행")
    except Exception as e:
        st.exception(e)

st.subheader("2) 비교(이후) 파일 분석")
c3, c4 = st.columns(2)
with c3:
    file_new = st.file_uploader("비교 엑셀 파일", type=["xlsx"], key="new")
with c4:
    sheet_new = None
    if file_new:
        try:
            wb2 = load_workbook(file_new, read_only=True, data_only=True)
            sheet_new = st.selectbox("시트 선택(비교)", options=wb2.sheetnames, index=0)
        except Exception as e:
            st.error(f"비교 파일 시트 읽기 실패: {e}")

if st.button("변경 사항 분석 실행", type="primary", disabled=not (file_new and sheet_new and ("old_rows" in st.session_state))):
    try:
        old_rows = st.session_state["old_rows"]
        old_multiset = st.session_state["old_rows_norm_multiset"]
        old_tuple_to_indices = st.session_state["old_rows_by_tuple_indices"]
        new_rows = read_sheet_values(file_new, sheet_new, trim_spaces, case_sensitive)

        # Optional fill-change exclusion
        if exclude_rows_if_fill_changed:
            old_fills = st.session_state.get("old_fills", {})
            if not old_fills:
                # fallback to compute if user toggled after saving
                old_fills = read_sheet_fills(file_old, sheet_old) if (file_old and sheet_old) else {}
            new_fills = read_sheet_fills(file_new, sheet_new)
            max_row_new = max([r["_row"] for r in new_rows] or [0])
            fill_changed_rows = set()
            for r in range(1, max_row_new+1):
                changed = False
                for c in range(COL_START, COL_END+1):
                    if old_fills.get((r, c)) != new_fills.get((r, c)):
                        changed = True
                        break
                if changed:
                    fill_changed_rows.add(r)
        else:
            fill_changed_rows = set()

        remaining_old_indices = set(range(len(old_rows)))
        remaining_new_indices = set(range(len(new_rows)))
        exact_pairs = []
        temp_multiset = old_multiset.copy()
        temp_tuple_to_indices = {k: v.copy() for k, v in old_tuple_to_indices.items()}

        for j, nr in enumerate(new_rows):
            if nr["_row"] in fill_changed_rows:
                continue
            t = row_tuple(nr["norm"])
            if temp_multiset.get(t, 0) > 0:
                i = temp_tuple_to_indices[t].pop(0)
                temp_multiset[t] -= 1
                exact_pairs.append((i, j))
                remaining_old_indices.discard(i)
                remaining_new_indices.discard(j)

        old_left = [old_rows[i] for i in sorted(remaining_old_indices)]
        new_left = [new_rows[j] for j in sorted(remaining_new_indices) if new_rows[j]["_row"] not in fill_changed_rows]
        pairs, leftover_old_idx, leftover_new_idx = best_pairing(new_left, old_left)

        best_pairs = []
        sorted_old_left = sorted(remaining_old_indices)
        sorted_new_left = sorted(remaining_new_indices)
        for eq, i, j in sorted([(p[2], p[0], p[1]) for p in pairs], reverse=True):
            old_idx_global = sorted_old_left[i]
            new_idx_global = sorted_new_left[j]
            best_pairs.append((old_idx_global, new_idx_global, eq))

        unchanged_records = [{
            "기준행": old_rows[i]["_row"],
            "비교행": new_rows[j]["_row"],
            "상태": "동일(재정렬만)"
        } for i, j in exact_pairs]

        changes_records = []
        for i, j, eq in best_pairs:
            rec = build_diff_record(old_rows[i], new_rows[j])
            rec["일치열수"] = eq
            rec["상태"] = "변경"
            changes_records.append(rec)

        used_old = set([i for i, _, _ in best_pairs] + [i for i, _ in exact_pairs])
        used_new = set([j for _, j, _ in best_pairs] + [j for _, j in exact_pairs])

        removed_records = [{"기준행": old_rows[i]["_row"], "상태": "제거됨"} for i in range(len(old_rows)) if i not in used_old]
        added_records = [{"비교행": new_rows[j]["_row"], "상태": "추가됨"} for j in range(len(new_rows)) if (j not in used_new and new_rows[j]["_row"] not in fill_changed_rows)]

        df_unchanged = pd.DataFrame(unchanged_records)
        df_changes = pd.DataFrame(changes_records, columns=["기준행","비교행","일치열수","변경요약","상태"])
        df_removed = pd.DataFrame(removed_records)
        df_added = pd.DataFrame(added_records)

        st.success(f"결과: 동일(재정렬만) {len(df_unchanged)}건, 변경 {len(df_changes)}건, 제거 {len(df_removed)}건, 추가 {len(df_added)}건")
        st.write("### 동일(재정렬만)")
        st.dataframe(df_unchanged, use_container_width=True, hide_index=True)
        st.write("### 변경")
        st.dataframe(df_changes, use_container_width=True, hide_index=True)
        st.write("### 제거됨(기준에는 있었으나 비교에는 없음)")
        st.dataframe(df_removed, use_container_width=True, hide_index=True)
        st.write("### 추가됨(비교에는 있으나 기준에는 없음)")
        st.dataframe(df_added, use_container_width=True, hide_index=True)

        from io import BytesIO
        def to_xlsx(dfs, names, filename):
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as wr:
                for df, name in zip(dfs, names):
                    if not df.empty:
                        df.to_excel(wr, index=False, sheet_name=name)
                    else:
                        pd.DataFrame().to_excel(wr, index=False, sheet_name=name)
            return bio.getvalue()

        st.download_button("결과 통합 엑셀 다운로드",
                           data=to_xlsx([df_unchanged, df_changes, df_removed, df_added],
                                        ["unchanged", "changes", "removed", "added"],
                                        "diff_result.xlsx"),
                           file_name="diff_result.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.exception(e)

st.info("주의: '기준 데이터 저장'을 먼저 수행하세요. 그 다음 비교 파일을 업로드하고 '변경 사항 분석 실행'을 누르면, 행 순서가 달라도 정확히 식별합니다.")
