import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

st.set_page_config(page_title="批量證書生成器", layout="wide")

st.title("🔥 批量證書/圖片文字疊加生成器（即時預覽版）")
st.markdown("**所見即所得：調整位置/大小後立即預覽效果！完全基於上傳檔案處理**")

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

# 上傳字體（可選）
font_file = st.file_uploader("3. （可選）上傳中文字體檔（.ttf，避免亂碼）", type=["ttf"])

# 選擇欄位與過濾
columns = df.columns.tolist()
selected_column = st.selectbox("選擇要疊加的文字欄位（例如「姓名」）", columns)

st.write("### 選擇要生成的人員（不選則全部）")
selected_names = st.multiselect(f"從「{selected_column}」選擇", df[selected_column].unique().tolist())
target_df = df[df[selected_column].isin(selected_names)] if selected_names else df
st.write(f"將生成 {len(target_df)} 張圖片")

# 文字設定
st.subheader("文字調整（調整後下方即時預覽）")
col1, col2, col3, col4 = st.columns(4)
with col1:
    pos_x = st.number_input("X 位置（從左）", min_value=0, max_value=background.width, value=background.width // 2)
with col2:
    pos_y = st.number_input("Y 位置（從上）", min_value=0, max_value=background.height, value=background.height // 2)
with col3:
    font_size = st.slider("字體大小", 20, 200, 80)
with col4:
    text_color = st.color_picker("文字顏色", "#000000")

align = st.selectbox("文字對齊方式", ["左", "中", "右"])

# 載入字體
if font_file:
    font = ImageFont.truetype(font_file, font_size)
    st.success("已載入自訂字體")
else:
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()
        st.warning("使用系統預設字體（可能中文亂碼），建議上傳 .ttf 字體檔")

# 即時預覽（用第一筆資料）
if len(target_df) > 0:
    sample_text = str(target_df.iloc[0][selected_column])
else:
    sample_text = "預覽文字"

preview_img = background.copy()
draw = ImageDraw.Draw(preview_img)

# 計算對齊後 X
x = pos_x
if align == "中":
    bbox = draw.textbbox((0, 0), sample_text, font=font)
    x -= (bbox[2] - bbox[0]) // 2
elif align == "右":
    bbox = draw.textbbox((0, 0), sample_text, font=font)
    x -= (bbox[2] - bbox[0])

draw.text((x, pos_y), sample_text, font=font, fill=text_color)
st.image(preview_img, caption="🔍 即時預覽效果（所有圖片將以此設定為準）", use_column_width=True)

# 生成按鈕
if st.button("🔥 開始批量生成所有圖片", type="primary"):
    with st.spinner(f"正在生成 {len(target_df)} 張..."):
        output_images = []
        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)
            text = str(row[selected_column])

            final_x = pos_x
            if align == "中":
                bbox = draw.textbbox((0, 0), text, font=font)
                final_x -= (bbox[2] - bbox[0]) // 2
            elif align == "右":
                bbox = draw.textbbox((0, 0), text, font=font)
                final_x -= (bbox[2] - bbox[0])

            draw.text((final_x, pos_y), text, font=font, fill=text_color)

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            filename = f"{text.replace('/', '_')}_{idx+1}.png"
            output_images.append((filename, buf))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, buf in output_images:
                buf.seek(0)
                zf.writestr(name, buf.read())
        zip_buffer.seek(0)

        st.download_button(
            "📥 下載所有圖片（ZIP）",
            zip_buffer,
            "certificates.zip",
            "application/zip"
        )
        st.success("生成完成！下載後解壓即可列印～")
        st.balloons()

st.caption("安全提示：所有處理都在臨時環境，資料不會永久儲存。")
