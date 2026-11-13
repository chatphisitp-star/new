import io
import re
import gc
import zipfile
import pandas as pd
import streamlit as st
import altair as alt
from openpyxl.styles import PatternFill

st.set_page_config(
    page_title="CRS Gain Validator",
    page_icon="⚡",
    layout="wide"
)

# ==================================================
# 0. SECURITY & MEMORY WIPE FUNCTION
# ==================================================
if "uploader_reset_count" not in st.session_state:
    st.session_state.uploader_reset_count = 0

def wipe_all_data():
    # 1. ล้างแคชข้อมูลที่ประมวลผลใน RAM
    st.cache_data.clear()
    # 2. รีเซ็ตตัวนับ Key ของ File Uploader เพื่อล้างไฟล์ทั้งหมดออกจากหน้าเว็บและ RAM
    reset_id = st.session_state.uploader_reset_count + 1
    # 3. ล้างค่าใน Session State ทั้งหมด
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.session_state.uploader_reset_count = reset_id
    st.session_state.num_weeks = 2
    # 4. บังคับคืนหน่วยความจำให้ระบบปฏิบัติการ
    gc.collect()

# ==================================================
# 1. SIDEBAR CONFIGURATION
# ==================================================
st.sidebar.header("⚙️ การตั้งค่าระบบ")

# ปุ่มล้างประวัติความปลอดภัยใน Sidebar
st.sidebar.error("🔒 ความปลอดภัยของข้อมูล")
if st.sidebar.button("🗑️ ล้างข้อมูลและประวัติทั้งหมด", use_container_width=True, type="primary"):
    wipe_all_data()
    st.rerun()
st.sidebar.caption("กดปุ่มนี้เมื่อใช้งานเสร็จ เพื่อลบไฟล์ที่อัปโหลด กราฟ และข้อมูลใน RAM ทั้งหมดทันที")
st.sidebar.divider()

with st.sidebar.expander("📝 ปรับแต่งข้อความหัวข้อ (Headers)", expanded=False):
    app_title = st.text_input("หัวข้อหลัก (Title):", value=" CRS Gain Validator ")
    app_subtitle = st.text_input(
        "หัวข้อย่อย (Subheader):", 
        value="ระบบตรวจเช็คและแก้ไข CRS Gain, pdschTypeBGain อัตโนมัติ"
    )
    section_manage_title = st.text_input("หัวข้อส่วนจัดการชุดไฟล์:", value="🛠️ จัดการชุดไฟล์ (Week)")
    section_chart_title = st.text_input("หัวข้อส่วนแสดงกราฟ:", value="CRS Gain, pdschTypeBGain")
    section_download_title = st.text_input("หัวข้อส่วนดาวน์โหลด:", value=" ดาวน์โหลดไฟล์ Excel")

with st.sidebar.expander("📄 ตั้งชื่อไฟล์ Export", expanded=False):
    prefix_checked = st.text_input("คำนำหน้าไฟล์ Checked:", value="checked_")
    prefix_fixed = st.text_input("คำนำหน้าไฟล์ Fixed:", value="fixed_")

with st.sidebar.expander("🎨 ปรับแต่งสีระบบ", expanded=True):
    st.subheader("1. สีในไฟล์ Excel")
    highlight_excel = st.checkbox("เปิดใช้งานไฮไลต์สีใน Excel", value=True)
    color_fail_excel = st.color_picker("สีไฮไลต์ Error 'f' (Checked File)", "#FFC7CE")
    color_fix_excel = st.color_picker("สีไฮไลต์จุดที่แก้ไข (Fixed File)", "#C6EFCE")

    st.subheader("2. สีกราฟแท่ง")
    color_incorrect_chart = st.color_picker("สีกราฟ Incorrect (f)", "#E05252")
    color_correct_chart = st.color_picker("สีกราฟ Correct (t)", "#2878D0")

st.title(app_title)
st.subheader(app_subtitle)

# ==================================================
# 2. STATE MANAGEMENT (บันทึกสถานะล่าสุดตลอด)
# ==================================================
if "num_weeks" not in st.session_state:
    st.session_state.num_weeks = 2

def update_weeks(delta: int):
    new_val = st.session_state.num_weeks + delta
    st.session_state.num_weeks = max(1, min(50, new_val))

def set_weeks_direct():
    st.session_state.num_weeks = st.session_state.direct_num_input

# ==================================================
# 3. FAST VALIDATION ENGINE
# ==================================================
VALID_PAIRS = {
    300: 1,
    477: 2,
    600: 3,
    -300: 1,
    177: 1,
    0: 0
}

VALID_PAIRS_SERIES = pd.Series(VALID_PAIRS)

def clean_col_name(name: str) -> str:
    if pd.isna(name):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

@st.cache_data(show_spinner=False)
def process_excel_fast(file_bytes: bytes, filename: str):
    df_full = pd.read_excel(io.BytesIO(file_bytes), header=None)
    
    if len(df_full) < 1:
        raise ValueError("ไฟล์ไม่มีข้อมูล")

    header_row_idx = None
    crs_col_idx = None
    pdsch_col_idx = None

    # ค้นหาคอลัมน์ในแถวที่ 1 ถึง 10 (Row Index 0 ถึง 9)
    max_search_rows = min(10, len(df_full))
    for r_idx in range(max_search_rows):
        row_values = df_full.iloc[r_idx]
        found_crs = None
        found_pdsch = None

        for c_idx, col in enumerate(row_values):
            clean_name = clean_col_name(col)
            if clean_name == "crsgain":
                found_crs = c_idx
            elif clean_name == "pdschtypebgain":
                found_pdsch = c_idx

        # เมื่อเจอทั้ง 2 คอลัมน์ในแถวเดียวกัน ให้บันทึกตำแหน่งแล้วออกจากลูป
        if found_crs is not None and found_pdsch is not None:
            header_row_idx = r_idx
            crs_col_idx = found_crs
            pdsch_col_idx = found_pdsch
            break

    if header_row_idx is None:
        raise ValueError("ไม่พบ 'crsGain' และ 'pdschTypeBGain' ในแถวที่ 1 ถึง 10")

    data_start_idx = header_row_idx + 1
    if data_start_idx >= len(df_full):
        raise ValueError(f"ไม่พบข้อมูลหลังแถว Header (Row {header_row_idx + 1})")

    # อ่านข้อมูลจริงเริ่มจากแถวถัดจาก Header
    crs_series = pd.to_numeric(df_full.iloc[data_start_idx:, crs_col_idx], errors='coerce')
    pdsch_series = pd.to_numeric(df_full.iloc[data_start_idx:, pdsch_col_idx], errors='coerce')

    expected_pdsch = crs_series.map(VALID_PAIRS_SERIES)
    valid_mask = crs_series.isin(VALID_PAIRS_SERIES.index) & (expected_pdsch == pdsch_series)
    non_null_mask = crs_series.notna() & pdsch_series.notna()

    results_series = pd.Series("", index=crs_series.index)
    results_series[non_null_mask] = valid_mask[non_null_mask].map({True: 't', False: 'f'})

    t_count = int((results_series == "t").sum())
    f_count = int((results_series == "f").sum())

    # สร้าง checked_df และเพิ่มคอลัมน์ result
    df_checked = df_full.copy()
    res_col = pd.Series("", index=df_checked.index)
    res_col.iloc[header_row_idx] = "result"
    res_col.iloc[data_start_idx:] = results_series.values
    df_checked[len(df_checked.columns)] = res_col

    # สร้าง fixed_df และอัปเดตค่า pdschTypeBGain
    df_fixed = df_full.copy()
    fixed_pdsch = pdsch_series.copy()
    fix_mask = crs_series.isin(VALID_PAIRS_SERIES.index) & crs_series.notna()
    fixed_pdsch[fix_mask] = expected_pdsch[fix_mask]
    
    df_fixed.iloc[data_start_idx:, pdsch_col_idx] = fixed_pdsch.values
    df_fixed[len(df_fixed.columns)] = res_col

    return {
        "filename": filename,
        "checked_df": df_checked,
        "fixed_df": df_fixed,
        "pdsch_col_idx": pdsch_col_idx,
        "header_row_idx": header_row_idx,
        "t": t_count,
        "f": f_count
    }

def df_to_excel_bytes_styled(df: pd.DataFrame, highlight_type=None, target_col_idx=None, header_row_idx=3, fail_color="#FFC7CE", fix_color="#C6EFCE") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False)
        
        if highlight_excel and highlight_type:
            worksheet = writer.sheets['Sheet1']
            hex_fail = fail_color.lstrip('#')
            hex_fix = fix_color.lstrip('#')
            
            fill_fail = PatternFill(start_color=hex_fail, end_color=hex_fail, fill_type="solid")
            fill_fix = PatternFill(start_color=hex_fix, end_color=hex_fix, fill_type="solid")

            last_col_idx = df.shape[1]
            data_start_idx = header_row_idx + 1
            res_series = df.iloc[data_start_idx:, last_col_idx - 1]
            fail_rows = res_series[res_series == "f"].index

            if highlight_type == "checked":
                for r_idx in fail_rows:
                    row_num = r_idx + 1
                    for col in range(1, last_col_idx + 1):
                        worksheet.cell(row=row_num, column=col).fill = fill_fail

            elif highlight_type == "fixed" and target_col_idx is not None:
                for r_idx in fail_rows:
                    row_num = r_idx + 1
                    worksheet.cell(row=row_num, column=target_col_idx + 1).fill = fill_fix

    output.seek(0)
    return output.getvalue()

# ==================================================
# 4. CONTROLS FOR ADDING/REMOVING WEEKS & GLOBAL UPLOAD
# ==================================================
st.markdown(f"### {section_manage_title}")

c1, c2, c3, c4, c5, c6 = st.columns(6)

with c1:
    st.button("➕ เพิ่ม 1", use_container_width=True, on_click=update_weeks, args=(1,))
with c2:
    st.button("➕ เพิ่ม 5", use_container_width=True, on_click=update_weeks, args=(5,))
with c3:
    st.button("➕ เพิ่ม 10", use_container_width=True, on_click=update_weeks, args=(10,))
with c4:
    st.button("➖ ลด 1", use_container_width=True, on_click=update_weeks, args=(-1,))
with c5:
    st.button("➖ ลด 5", use_container_width=True, on_click=update_weeks, args=(-5,))
with c6:
    st.button("➖ ลด 10", use_container_width=True, on_click=update_weeks, args=(-10,))

st.number_input(
    "หรือระบุจำนวนชุดอัปโหลดโดยตรง (1 - 50):", 
    min_value=1, 
    max_value=50, 
    value=st.session_state.num_weeks,
    key="direct_num_input",
    on_change=set_weeks_direct
)

st.caption(f"จำนวนชุดอัปโหลดปัจจุบัน: **{st.session_state.num_weeks}** ชุด")

# --- ช่องอัปโหลดส่วนกลาง (ใช้ร่วมกันทุก Week) ---
reset_suffix = st.session_state.uploader_reset_count
st.info("⚡ **อัปโหลดด่วนสำหรับทุก Week:** ลากไฟล์วางช่องด้านล่างนี้เพื่อกระจายไฟล์ให้ทุก Week ใช้งานพร้อมกันทันที")
global_up_files = st.file_uploader(
    "📂 ลากไฟล์ Excel ที่ต้องการใช้ร่วมกันในทุก Week (อัปโหลดครั้งเดียวได้ทุกชุด)",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"global_uploader_{reset_suffix}"
)

st.divider()

# ==================================================
# 5. DYNAMIC UPLOAD SECTIONS
# ==================================================
weeks_data_list = []
cols_per_row = 3

for i in range(0, st.session_state.num_weeks, cols_per_row):
    cols = st.columns(cols_per_row)
    for j in range(cols_per_row):
        week_idx = i + j
        if week_idx < st.session_state.num_weeks:
            with cols[j]:
                st.subheader(f"📅 ชุดที่ {week_idx + 1}")
                
                name_key = f"w_name_{week_idx}"
                if name_key not in st.session_state:
                    st.session_state[name_key] = f"2026{16 + week_idx}"

                w_name = st.text_input(
                    f"ชื่อ Week {week_idx + 1}:",
                    key=name_key
                )
                up_files = st.file_uploader(
                    f"📂 ลากไฟล์ Excel เพิ่มเติมเฉพาะชุดที่ {week_idx + 1}",
                    type=["xlsx"],
                    accept_multiple_files=True,
                    key=f"uploader_{week_idx}_{reset_suffix}"
                )
                
                # รวมไฟล์ส่วนกลางเข้ากับไฟล์เฉพาะของ Week นั้นๆ
                combined_files = []
                if global_up_files:
                    combined_files.extend(global_up_files)
                if up_files:
                    combined_files.extend(up_files)

                weeks_data_list.append({
                    "name": w_name,
                    "files": combined_files
                })

# ==================================================
# 6. PROCESS ALL WEEKS
# ==================================================
chart_rows = []
all_weeks_results = {}

has_files = any(len(w["files"]) > 0 for w in weeks_data_list)

if has_files:
    with st.spinner("🚀 กำลังประมวลผลข้อมูลในความเร็วสูง..."):
        for w_data in weeks_data_list:
            w_name = w_data["name"]
            files = w_data["files"]

            if files:
                w_results = []
                for uploaded_file in files[:50]:
                    try:
                        file_bytes = uploaded_file.getvalue()
                        res = process_excel_fast(file_bytes, uploaded_file.name)
                        w_results.append(res)
                    except Exception as e:
                        st.error(f"❌ [{w_name}] {uploaded_file.name}: {str(e)}")

                if w_results:
                    all_weeks_results[w_name] = w_results
                    tot_t = sum(r["t"] for r in w_results)
                    tot_f = sum(r["f"] for r in w_results)
                    chart_rows.append({
                        "Week": w_name,
                        "f - Incorrect": tot_f,
                        "t - Correct": tot_t
                    })

# ==================================================
# 7. COMBINED CHART
# ==================================================
if chart_rows:
    st.divider()

    combined_df = pd.DataFrame(chart_rows)
    
    df_melted = combined_df.melt(
        id_vars=["Week"], 
        value_vars=["f - Incorrect", "t - Correct"],
        var_name="Status", 
        value_name="Count"
    )

    chart = (
        alt.Chart(df_melted)
        .mark_bar()
        .encode(
            x=alt.X("Week:N", title=None, axis=alt.Axis(labelAngle=-90)),
            y=alt.Y("Count:Q", title="Count"),
            color=alt.Color(
                "Status:N",
                scale=alt.Scale(
                    domain=["f - Incorrect", "t - Correct"],
                    range=[color_incorrect_chart, color_correct_chart]
                ),
                legend=alt.Legend(orient="bottom", title=None)
            ),
            tooltip=["Week", "Status", "Count"]
        )
        .properties(
            title=alt.TitleParams(
                text=section_chart_title,
                fontSize=20,
                anchor="start",
                offset=15
            ),
            height=500
        )
    )

    st.altair_chart(chart, use_container_width=True)
    st.dataframe(combined_df, use_container_width=True, hide_index=True)

    # ==================================================
    # 8. DOWNLOAD SECTION
    # ==================================================
    st.divider()
    st.header(section_download_title)

    tab_titles = [f"📦 {w_name}" for w_name in all_weeks_results.keys()]
    if tab_titles:
        tabs = st.tabs(tab_titles)
        for tab_idx, (tab, (w_name, w_results)) in enumerate(zip(tabs, all_weeks_results.items())):
            with tab:
                zip_b = io.BytesIO()
                with zipfile.ZipFile(zip_b, "w", zipfile.ZIP_DEFLATED) as z:
                    for r in w_results:
                        chk_b = df_to_excel_bytes_styled(
                            r["checked_df"], highlight_type="checked", 
                            header_row_idx=r["header_row_idx"],
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )
                        fix_b = df_to_excel_bytes_styled(
                            r["fixed_df"], highlight_type="fixed", target_col_idx=r["pdsch_col_idx"],
                            header_row_idx=r["header_row_idx"],
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )
                        z.writestr(f"{prefix_checked}{r['filename']}", chk_b)
                        z.writestr(f"{prefix_fixed}{r['filename']}", fix_b)

                st.download_button(
                    f"📦 ดาวน์โหลดไฟล์ทั้งหมดของ {w_name} (ZIP รวม)",
                    data=zip_b.getvalue(),
                    file_name=f"crs_{w_name}_results.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                    key=f"dl_zip_{w_name}_{tab_idx}"
                )

                st.markdown("---")
                st.subheader(f"📄 รายการไฟล์และปุ่มดาวน์โหลดแยก ({len(w_results)} ไฟล์)")

                for file_idx, r in enumerate(w_results):
                    with st.expander(f"📄 {r['filename']} | ✅ Correct (t): {r['t']} | ❌ Incorrect (f): {r['f']}", expanded=False):
                        btn_col1, btn_col2 = st.columns(2)
                        
                        chk_bytes = df_to_excel_bytes_styled(
                            r["checked_df"], highlight_type="checked", 
                            header_row_idx=r["header_row_idx"],
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )
                        fix_bytes = df_to_excel_bytes_styled(
                            r["fixed_df"], highlight_type="fixed", target_col_idx=r["pdsch_col_idx"],
                            header_row_idx=r["header_row_idx"],
                            fail_color=color_fail_excel, fix_color=color_fix_excel
                        )

                        with btn_col1:
                            st.download_button(
                                label=f"⬇️ โหลด Checked ({r['filename']})",
                                data=chk_bytes,
                                file_name=f"{prefix_checked}{r['filename']}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"btn_chk_{tab_idx}_{file_idx}_{r['filename']}"
                            )

                        with btn_col2:
                            st.download_button(
                                label=f"⬇️ โหลด Fixed ({r['filename']})",
                                data=fix_bytes,
                                file_name=f"{prefix_fixed}{r['filename']}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key=f"btn_fix_{tab_idx}_{file_idx}_{r['filename']}"
                            )

    st.success("🎉 ประมวลผลเรียบร้อยแล้ว!")

    # ปุ่มล้างข้อมูลด้านล่างสุดหลังดาวน์โหลดเสร็จ
    st.divider()
    if st.button("🗑️ ดาวน์โหลดเสร็จแล้ว — ล้างข้อมูลและประวัติทั้งหมดออกจากระบบทันที", type="primary", use_container_width=True):
        wipe_all_data()
        st.rerun()
