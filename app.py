import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

st.set_page_config(page_title="多欄位證書生成器", layout="wide")

st.title("📝 多欄位批量證書生成器（左右分欄 + 大預覽區）")
st.markdown("**左邊調整設定，右邊即時大圖預覽，所見即所得！圖片自動縮放適應畫面**")

# 左右分欄
left_col, right_col = st.columns([1, 1.5])  # 左1 : 右1.5，可調整比例

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
        st.stop()
    if data_file.name.endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)

    st.success(f"資料上傳成功！共 {len(df)} 筆")

    # 上傳字體（可選）
    font_file = st.file_uploader("（可選）上傳中文字體檔（.ttf，避免亂碼）", type=["ttf"])

    # 過濾人員
    st.subheader("人員過濾")
    filter_column = st.selectbox("用哪一欄過濾？（例如「姓名」）", df.columns)
    selected_names = st.multiselect("選擇特定人員（不選則全部）", df[filter_column].unique().tolist())
    target_df = df[df[filter_column].isin(selected_names)] if selected_names else df
    st.write(f"將生成 {len(target_df)} 張")

    # 多欄位選擇與設定
    st.subheader("📌 文字欄位設定")
    selected_columns = st.multiselect("選擇要疊加的欄位（可多選）", df.columns)

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
                "y": background.height // 2 + selected_columns.index(col) * 100,
                "size": 80,
                "color": "#000000",
                "align": "中"
            }

        st.markdown(f"**{col}**")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.session_state.settings[col]["x"] = st.number_input(f"X", 0, background.width, st.session_state.settings[col]["x"], key=f"x_{col}")
        with c2:
            st.session_state.settings[col]["y"] = st.number_input(f"Y", 0, background.height, st.session_state.settings[col]["y"], key=f"y_{col}")
        with c3:
            st.session_state.settings[col]["size"] = st.slider(f"大小", 20, 200, st.session_state.settings[col]["size"], key=f"size_{col}")
        with c4:
            st.session_state.settings[col]["color"] = st.color_picker(f"顏色", st.session_state.settings[col]["color"], key=f"color_{col}")
        with c5:
            st.session_state.settings[col]["align"] = st.selectbox(f"對齊", ["左", "中", "右"], index=["左","中","右"].index(st.session_state.settings[col]["align"]), key=f"align_{col}")

    # 生成按鈕（放在左欄底部）
    generate_btn = st.button("🔥 開始批量生成", type="primary", use_container_width=True)

# 右欄：專屬預覽區
with right_col:
    st.header("🔍 即時預覽工作區")
    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        preview_img = background.copy()
        draw = ImageDraw.Draw(preview_img)

        # 載入字體
        if font_file:
            base_font = ImageFont.truetype(font_file, 80)
        else:
            try:
                base_font = ImageFont.truetype("arial.ttf", 80)
            except:
                base_font = ImageFont.load_default()
                st.warning("建議上傳 .ttf 字體避免亂碼")

        # 繪製所有欄位文字
        for col in selected_columns:
            settings = st.session_state.settings[col]
            text = str(preview_row[col])

            try:
                font = base_font.font_variant(size=settings["size"])
            except:
                font = ImageFont.truetype(font_file or "arial.ttf", settings["size"]) if font_file else ImageFont.load_default()

            x = settings["x"]
            if settings["align"] == "中":
                bbox = draw.textbbox((0, 0), text, font=font)
                x -= (bbox[2] - bbox[0]) // 2
            elif settings["align"] == "右":
                bbox = draw.textbbox((0, 0), text, font=font)
                x -= (bbox[2] - bbox[0])

            draw.text((x, settings["y"]), text, font=font, fill=settings["color"])

        # 預覽圖自動縮放填滿右欄
        st.image(preview_img, caption="即時預覽（所有圖片將以此效果生成）", use_container_width=True)
    else:
        st.info("無資料可預覽")

# 生成邏輯（放在外面，點按鈕後執行）
if generate_btn:
    with st.spinner("正在批量生成..."):
        output_images = []
        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)

            for col in selected_columns:
                settings = st.session_state.settings[col]
                text = str(row[col])
                try:
                    font = base_font.font_variant(size=settings["size"])
                except:
                    font = ImageFont.truetype(font_file or "arial.ttf", settings["size"]) if font_file else ImageFont.load_default()

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
            filename = f"證書_{idx+1}.png"
            output_images.append((filename, buf))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, buf in output_images:
                buf.seek(0)
                zf.writestr(name, buf.read())
        zip_buffer.seek(0)

        st.download_button("📥 下載所有圖片（ZIP）", zip_buffer, "certificates.zip", "application/zip")
        st.success("生成完成！")
        st.balloons()

st.caption("安全私密：資料只在臨時環境處理，不儲存。")
