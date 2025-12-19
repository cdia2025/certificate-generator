import streamlit as st
import pandas as pd  # 確保這行在最上面
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

st.set_page_config(page_title="Mail Merge 式證書生成器", layout="wide")

st.title("✉️ Mail Merge 式多欄位證書生成器")
st.markdown("**像 Word Mail Merge 一樣：選取特定人員 + 多欄位疊加 + 即時大預覽（大小可調）**")

# 左右分欄
left_col, right_col = st.columns([2, 3])

with left_col:
    st.header("🛠️ 設定區")

    # 1. 上傳背景圖片
    background_file = st.file_uploader("上傳背景圖片模板（JPG/PNG，必填）", type=["jpg", "png", "jpeg"])
    if not background_file:
        st.stop()
    background = Image.open(background_file)

    # 2. 上傳資料檔
    data_file = st.file_uploader("上傳資料檔（CSV 或 Excel，必填）", type=["csv", "xlsx", "xls"])
    if not data_file:
        st.info("請上傳資料檔後繼續")
        st.stop()

    # 正確讀取資料（關鍵修復處）
    try:
        if data_file.name.lower().endswith(".csv"):
            df = pd.read_csv(data_file)
        else:  # .xlsx 或 .xls
            df = pd.read_excel(data_file)
        st.success(f"資料上傳成功！共 {len(df)} 筆記錄")
    except Exception as e:
        st.error(f"讀取資料檔失敗：{str(e)}")
        st.stop()

    # 3. 上傳字體（可選）
    font_file = st.file_uploader("（可選）上傳中文字體檔（.ttf，避免亂碼）", type=["ttf"])

    # 4. Mail Merge 人員選擇
    st.subheader("✉️ Mail Merge 人員選擇")
    filter_column = st.selectbox("用哪一欄作為收件人識別？（例如「姓名」）", df.columns)
    all_options = df[filter_column].astype(str).unique().tolist()
    selected_names = st.multiselect(
        "選擇需要生成的收件人（支援搜尋，不選則全部生成）",
        options=all_options,
        default=[],
        placeholder="開始輸入搜尋或選擇..."
    )
    target_df = df[df[filter_column].astype(str).isin(selected_names)] if selected_names else df
    st.write(f"將生成 **{len(target_df)}** 張個人化證書")

    # 5. 多欄位選擇
    st.subheader("📌 要疊加的欄位")
    selected_columns = st.multiselect("選擇要顯示的欄位（可多選）", df.columns)

    if not selected_columns:
        st.warning("請至少選擇一個欄位！")
        st.stop()

    # 儲存設定
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

    # 6. 預覽大小控制
    st.subheader("🔍 預覽控制")
    preview_scale = st.slider("預覽圖縮放比例（僅影響顯示，不影響生成品質）", 20, 200, 100)

    # 生成按鈕
    generate_btn = st.button("🔥 開始批量生成所有證書", type="primary", use_container_width=True)

# 右欄：即時預覽區
with right_col:
    st.header("🔍 即時預覽區（調整即更新）")

    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        preview_img = background.copy()
        draw = ImageDraw.Draw(preview_img)

        # 載入字體
        try:
            if font_file:
                base_font = ImageFont.truetype(font_file, 80)
            else:
                base_font = ImageFont.truetype("arial.ttf", 80)
        except:
            base_font = ImageFont.load_default()
            if font_file is None:
                st.warning("未上傳字體，使用預設（可能中文亂碼），建議上傳 .ttf")

        # 繪製所有欄位
        for col in selected_columns:
            settings = st.session_state.settings[col]
            text = str(preview_row[col])
            try:
                font = base_font.font_variant(size=settings["size"]) if hasattr(base_font, 'font_variant') else ImageFont.truetype(font_file or "arial.ttf", settings["size"]) if font_file else base_font
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

        # 縮放預覽圖
        if preview_scale != 100:
            new_width = int(background.width * preview_scale / 100)
            new_height = int(background.height * preview_scale / 100)
            preview_img = preview_img.resize((new_width, new_height), Image.LANCZOS)

        st.image(preview_img, caption=f"即時預覽（顯示 {preview_scale}%）", use_container_width=True)
    else:
        st.info("無資料可預覽")

# 生成邏輯
if generate_btn:
    with st.spinner(f"正在生成 {len(target_df)} 張證書..."):
        output_images = []
        # 重置字體（生成時用原大小）
        try:
            if font_file:
                gen_base_font = ImageFont.truetype(font_file, 80)
            else:
                gen_base_font = ImageFont.truetype("arial.ttf", 80)
        except:
            gen_base_font = ImageFont.load_default()

        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)

            for col in selected_columns:
                settings = st.session_state.settings[col]
                text = str(row[col])
                try:
                    font = gen_base_font.font_variant(size=settings["size"]) if hasattr(gen_base_font, 'font_variant') else ImageFont.truetype(font_file or "arial.ttf", settings["size"]) if font_file else gen_base_font
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
        st.success("所有證書生成完成！")
        st.balloons()

st.caption("安全．高效：類似 Mail Merge，快速產生個人化證書，資料不儲存。")
