import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

st.set_page_config(page_title="Mail Merge 式證書生成器", layout="wide")

st.title("✉️ Mail Merge 式多欄位證書生成器")
st.markdown("**像 Word Mail Merge 一樣：選取人員 + 多欄位疊加 + 即時預覽（縮放控制正常運作！）**")

# 左右分欄
left_col, right_col = st.columns([2, 3])

with left_col:
    st.header("🛠️ 設定區")

    # 上傳背景圖片
    background_file = st.file_uploader("上傳背景圖片模板（JPG/PNG，必填）", type=["jpg", "png", "jpeg"])
    if not background_file:
        st.stop()
    background = Image.open(background_file)

    # 上傳資料檔
    data_file = st.file_uploader("上傳資料檔（CSV 或 Excel，必填）", type=["csv", "xlsx", "xls"])
    if not data_file:
        st.info("請上傳資料檔後繼續")
        st.stop()

    # 讀取資料
    try:
        if data_file.name.lower().endswith(".csv"):
            df = pd.read_csv(data_file)
        else:
            df = pd.read_excel(data_file)
        st.success(f"資料上傳成功！共 {len(df)} 筆記錄")
    except Exception as e:
        st.error(f"讀取失敗：{str(e)}")
        st.stop()

    # 上傳字體（可選）
    font_file = st.file_uploader("（可選）上傳中文字體檔（.ttf，避免亂碼）", type=["ttf"])

    # Mail Merge 人員選擇
    st.subheader("✉️ Mail Merge 人員選擇")
    filter_column = st.selectbox("用哪一欄作為收件人識別？（例如「姓名」）", df.columns)
    all_options = df[filter_column].astype(str).unique().tolist()
    selected_names = st.multiselect(
        "選擇需要生成的收件人（支援搜尋，不選則全部）",
        options=all_options,
        default=[],
        placeholder="輸入搜尋或選擇..."
    )
    target_df = df[df[filter_column].astype(str).isin(selected_names)] if selected_names else df
    st.write(f"將生成 **{len(target_df)}** 張證書")

    # 多欄位選擇
    st.subheader("📌 要疊加的欄位")
    selected_columns = st.multiselect("選擇要顯示的欄位（可多選）", df.columns)
    if not selected_columns:
        st.warning("請至少選擇一個欄位！")
        st.stop()

    # 設定儲存
    if "settings" not in st.session_state:
        st.session_state.settings = {}

    for col in selected_columns:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {
                "x": background.width // 2,
                "y": background.height // 2 + selected_columns.index(col) * 120,
                "size": 80,
                "color": "#000000",
                "align": "中"
            }

        st.markdown(f"**{col}** 設定")
        c1, c2 = st.columns(2)
        with c1:
            st.session_state.settings[col]["x"] = st.number_input(f"X 位置", 0, background.width, st.session_state.settings[col]["x"], key=f"x_{col}")
            st.session_state.settings[col]["y"] = st.number_input(f"Y 位置", 0, background.height, st.session_state.settings[col]["y"], key=f"y_{col}")
        with c2:
            st.session_state.settings[col]["size"] = st.slider(f"字體大小", 20, 200, st.session_state.settings[col]["size"], key=f"size_{col}")
            st.session_state.settings[col]["color"] = st.color_picker(f"顏色", st.session_state.settings[col]["color"], key=f"color_{col}")
        st.session_state.settings[col]["align"] = st.selectbox(f"對齊方式", ["左", "中", "右"], index=["左","中","右"].index(st.session_state.settings[col]["align"]), key=f"align_{col}")

    # 預覽縮放控制（已修復）
    st.subheader("🔍 預覽控制")
    preview_scale = st.slider("預覽圖縮放比例（僅影響顯示，生成仍為100%原圖）", 20, 200, 100)

    # 生成按鈕
    generate_btn = st.button("🔥 開始批量生成所有證書", type="primary", use_container_width=True)

# 右欄：即時預覽區
with right_col:
    st.header("🔍 即時預覽區（調整即更新）")

    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        preview_img = background.copy()
        draw = ImageDraw.Draw(preview_img)

        # 字體載入（優化）
        if font_file:
            try:
                base_font = ImageFont.truetype(font_file, 80)
            except:
                base_font = ImageFont.load_default()
                st.warning("自訂字體載入失敗，使用預設")
        else:
            try:
                base_font = ImageFont.truetype("arial.ttf", 80)
            except:
                base_font = ImageFont.load_default()
                st.info("建議上傳 .ttf 字體以支援中文")

        # 繪製文字
        for col in selected_columns:
            settings = st.session_state.settings[col]
            text = str(preview_row[col])
            try:
                font = base_font.font_variant(size=settings["size"]) if hasattr(base_font, "font_variant") else ImageFont.truetype(font_file if font_file else "arial.ttf", settings["size"])
            except:
                font = ImageFont.load_default()

            x = settings["x"]
            if settings["align"] == "中":
                bbox = draw.textbbox((0, 0), text, font=font)
                x -= (bbox[2] - bbox[0]) // 2
            elif settings["align"] == "右":
                bbox = draw.textbbox((0, 0), text, font=font)
                x -= (bbox[2] - bbox[0])

            draw.text((x, settings["y"]), text, font=font, fill=settings["color"])

        # 縮放預覽圖（修復重點）
        display_img = preview_img.copy()
        if preview_scale != 100:
            new_width = int(background.width * preview_scale / 100)
            new_height = int(background.height * preview_scale / 100)
            display_img = display_img.resize((new_width, new_height), Image.LANCZOS)

        st.image(display_img, caption=f"即時預覽（顯示 {preview_scale}%）・生成時為100%原圖", use_container_width=True)
    else:
        st.info("無資料可預覽")

# 生成邏輯
if generate_btn:
    with st.spinner(f"正在生成 {len(target_df)} 張..."):
        output_images = []
        # 生成用字體
        gen_font_base = base_font if 'base_font' in locals() else ImageFont.load_default()

        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)

            for col in selected_columns:
                settings = st.session_state.settings[col]
                text = str(row[col])
                try:
                    font = gen_font_base.font_variant(size=settings["size"]) if hasattr(gen_font_base, "font_variant") else ImageFont.truetype(font_file if font_file else "arial.ttf", settings["size"])
                except:
                    font = ImageFont.load_default()

                final_x = settings["x"]
                if settings["align"] == "中":
                    bbox = draw.textbbox((0, 0), text, font=font)
                    final_x -= (bbox[2] - bbox[0]) // 2
                elif settings["align"] == "右":
                    bbox = draw.textbbox((0, 0), text, font=font)
                    final_x -= (bbox[2] - bbox[0])

                draw.text((final_x, settings["y"]), text, font=font, fill=settings["color"])

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            safe_name = str(row.get(filter_column, idx+1)).replace("/", "_").replace("\\", "_")
            filename = f"證書_{safe_name}.png"
            output_images.append((filename, buf))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, buf in output_images:
                buf.seek(0)
                zf.writestr(name, buf.read())
        zip_buffer.seek(0)

        st.download_button("📥 下載所有證書（ZIP）", zip_buffer, "certificates.zip", "application/zip")
        st.success("生成完成！")
        st.balloons()

st.caption("安全高效：資料僅臨時處理，不儲存。")
