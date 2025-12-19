import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

st.set_page_config(page_title="多欄位證書生成器", layout="wide")

st.title("📝 多欄位批量證書生成器（即時多文字預覽）")
st.markdown("**支援多個資料欄位（如姓名、電話、職位），每個獨立調整位置/大小/顏色，即時預覽效果！**")

# 上傳背景圖片
background_file = st.file_uploader("1. 上傳背景圖片模板（JPG/PNG，必填）", type=["jpg", "png", "jpeg"])
if not background_file:
    st.stop()

background = Image.open(background_file)
st.image(background, caption="背景模板", use_column_width=True)

# 上傳資料檔
data_file = st.file_uploader("2. 上傳資料檔（CSV 或 Excel，必填）", type=["csv", "xlsx", "xls"])
if not data_file:
    st.stop()

if data_file.name.endswith(".csv"):
    df = pd.read_csv(data_file)
else:
    df = pd.read_excel(data_file)

st.success(f"資料上傳成功！共 {len(df)} 筆")
st.dataframe(df.head(10))

# 上傳字體（可選，全域適用）
font_file = st.file_uploader("3. （可選）上傳中文字體檔（.ttf，避免所有文字亂碼）", type=["ttf"])

# 選擇要生成的資料行（過濾人員）
st.write("### 選擇要生成的人員（不選則全部）")
name_column_for_filter = st.selectbox("用哪一欄過濾人員？（例如「姓名」）", df.columns)
selected_names = st.multiselect(f"選擇特定人員", df[name_column_for_filter].unique().tolist())
target_df = df[df[name_column_for_filter].isin(selected_names)] if selected_names else df
st.write(f"將生成 {len(target_df)} 張圖片")

# 多欄位設定
st.subheader("📌 多欄位文字設定（每欄獨立調整）")
selected_columns = st.multiselect("選擇要疊加的資料欄位（可多選）", df.columns)

if not selected_columns:
    st.warning("請至少選擇一個欄位！")
    st.stop()

# 儲存每個欄位的設定（用 session_state 持久化）
if "field_settings" not in st.session_state:
    st.session_state.field_settings = {}

for col in selected_columns:
    if col not in st.session_state.field_settings:
        st.session_state.field_settings[col] = {
            "x": background.width // 2,
            "y": background.height // 2 + (selected_columns.index(col) * 100),  # 預設垂直間隔
            "size": 80,
            "color": "#000000",
            "align": "中"
        }

    st.markdown(f"#### ⚙️ 【{col}】欄位設定")
    cols = st.columns(5)
    with cols[0]:
        st.session_state.field_settings[col]["x"] = st.number_input(f"{col} - X 位置", min_value=0, max_value=background.width, value=st.session_state.field_settings[col]["x"], key=f"x_{col}")
    with cols[1]:
        st.session_state.field_settings[col]["y"] = st.number_input(f"{col} - Y 位置", min_value=0, max_value=background.height, value=st.session_state.field_settings[col]["y"], key=f"y_{col}")
    with cols[2]:
        st.session_state.field_settings[col]["size"] = st.slider(f"{col} - 字體大小", 20, 200, st.session_state.field_settings[col]["size"], key=f"size_{col}")
    with cols[3]:
        st.session_state.field_settings[col]["color"] = st.color_picker(f"{col} - 顏色", st.session_state.field_settings[col]["color"], key=f"color_{col}")
    with cols[4]:
        st.session_state.field_settings[col]["align"] = st.selectbox(f"{col} - 對齊", ["左", "中", "右"], index=["左", "中", "右"].index(st.session_state.field_settings[col]["align"]), key=f"align_{col}")

# 載入字體（全域）
if font_file:
    base_font = ImageFont.truetype(font_file, 80)  # 基礎大小，之後會依每個欄位縮放
    st.success("已載入自訂字體")
else:
    try:
        base_font = ImageFont.truetype("arial.ttf", 80)
    except:
        base_font = ImageFont.load_default()
        st.warning("使用系統預設字體（可能中文亂碼），強烈建議上傳 .ttf 字體檔")

# 即時預覽（用第一筆資料顯示所有欄位文字）
st.subheader("🔍 即時預覽效果（調整後立即更新，所有圖片將以此為準）")
if len(target_df) > 0:
    preview_row = target_df.iloc[0]
    preview_img = background.copy()
    draw = ImageDraw.Draw(preview_img)

    for col in selected_columns:
        settings = st.session_state.field_settings[col]
        text = str(preview_row[col])
        font = ImageFont.truetype(font_file) if font_file else base_font
        font = font.font_variant(size=settings["size"]) if font_file else ImageFont.load_default().font_variant(size=settings["size"]) if hasattr(ImageFont.load_default(), 'font_variant') else base_font  # 簡化處理

        try:
            # Pillow 9.0+ 支援 font_variant
            font = base_font.font_variant(size=settings["size"])
        except:
            font = ImageFont.truetype(font_file or "arial.ttf", settings["size"])

        x = settings["x"]
        if settings["align"] == "中":
            bbox = draw.textbbox((0, 0), text, font=font)
            x -= (bbox[2] - bbox[0]) // 2
        elif settings["align"] == "右":
            bbox = draw.textbbox((0, 0), text, font=font)
            x -= (bbox[2] - bbox[0])

        draw.text((x, settings["y"]), text, font=font, fill=settings["color"])

    st.image(preview_img, use_column_width=True)
else:
    st.info("無資料可預覽")

# 生成按鈕
if st.button("🔥 開始批量生成所有圖片", type="primary"):
    with st.spinner("正在生成，請稍候..."):
        output_images = []
        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)

            for col in selected_columns:
                settings = st.session_state.field_settings[col]
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

        st.download_button("📥 下載所有圖片（ZIP）", zip_buffer, "multi_field_certificates.zip", "application/zip")
        st.success("生成完成！")
        st.balloons()

st.caption("安全．私密：所有處理都在臨時環境，資料不儲存。")
