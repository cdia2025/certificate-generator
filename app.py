import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
from streamlit_draggable import draggable

st.set_page_config(page_title="拖拉式證書生成器", layout="wide")

st.title("🖱️ 拖拉式批量證書生成器（所見即所得）")
st.markdown("**直接在圖片上拖拉文字位置，即時預覽，超直覺！完全基於上傳檔案處理**")

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
font_file = st.file_uploader("3. （可選）上傳中文字體檔（.ttf 避免亂碼）", type=["ttf"])

# 選擇欄位
columns = df.columns.tolist()
selected_column = st.selectbox("選擇要顯示的文字欄位（例如「姓名」）", columns)

# 過濾人員
st.write("### 選擇要生成的人員（不選則全部）")
selected_names = st.multiselect(f"從「{selected_column}」選擇", df[selected_column].unique().tolist())
target_df = df[df[selected_column].isin(selected_names)] if selected_names else df
st.write(f"將生成 {len(target_df)} 張圖片")

# 文字設定
st.subheader("文字樣式調整（拖拉 + 滑桿）")
col1, col2, col3 = st.columns(3)
with col1:
    font_size = st.slider("字體大小", 20, 200, 80)
with col2:
    text_color = st.color_picker("文字顏色", "#000000")
with col3:
    align = st.selectbox("文字對齊", ["左", "中", "右"])

# 載入字體
if font_file:
    font = ImageFont.truetype(font_file, font_size)
else:
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        font = ImageFont.load_default()
        st.warning("使用預設字體，可能中文亂碼，建議上傳 .ttf")

# 取第一筆資料作為拖拉預覽樣本
sample_text = str(target_df.iloc[0][selected_column]) if len(target_df) > 0 else "預覽文字"

# 拖拉元件（在圖片上拖文字）
st.write("### 🖱️ 拖拉文字到想要的位置（即時預覽）")
draggable_text = draggable(
    sample_text,
    x=int(background.width / 2),
    y=int(background.height / 2),
    font_size=font_size,
    font_color=text_color,
    background_image=background,
    align=align.lower(),
    key="draggable_text"
)

# 顯示即時預覽圖（帶文字）
preview_img = background.copy()
draw = ImageDraw.Draw(preview_img)

x = draggable_text["x"]
y = draggable_text["y"]

if align == "中":
    bbox = draw.textbbox((0, 0), sample_text, font=font)
    text_width = bbox[2] - bbox[0]
    x -= text_width // 2
elif align == "右":
    bbox = draw.textbbox((0, 0), sample_text, font=font)
    text_width = bbox[2] - bbox[0]
    x -= text_width

draw.text((x, y), sample_text, font=font, fill=text_color)
st.image(preview_img, caption="即時預覽（所有圖片將以此位置為準）", use_column_width=True)

# 生成按鈕
if st.button("🔥 開始批量生成所有圖片", type="primary"):
    pos_x = draggable_text["x"]
    pos_y = draggable_text["y"]

    with st.spinner(f"正在生成 {len(target_df)} 張圖片..."):
        output_images = []
        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)
            text = str(row[selected_column])

            # 計算對齊
            final_x = pos_x
            if align == "中":
                bbox = draw.textbbox((0, 0), text, font=font)
                final_x = pos_x - (bbox[2] - bbox[0]) // 2
            elif align == "右":
                bbox = draw.textbbox((0, 0), text, font=font)
                final_x = pos_x - (bbox[2] - bbox[0])

            draw.text((final_x, pos_y), text, font=font, fill=text_color)

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            filename = f"{text.replace('/', '_')}_{idx+1}.png"
            output_images.append((filename, buf))

        # ZIP 下載
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
        st.success("生成完成！")
        st.balloons()
