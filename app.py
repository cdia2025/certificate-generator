import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile

st.set_page_config(page_title="批量證書生成器", layout="centered")

st.title("🔒 安全批量證書/圖片文字疊加生成器")
st.markdown("**完全基於您上傳的檔案處理，無任何預設資料或模板，保護隱私！**")

st.info("""
### 使用說明：
1. 先上傳證書背景圖片（JPG/PNG）。
2. 上傳包含資料的 CSV 或 Excel 檔（至少有一欄如「姓名」）。
3. （可選）上傳中文字體檔（.ttf）以避免亂碼，推薦如「Noto Sans TC」或「Microsoft JhengHei」。
4. 選擇欄位、過濾人員、調整文字位置/大小/顏色。
5. 點擊生成 → 下載 ZIP 包（內含所有個人化圖片）。
""")

# 上傳背景圖片（必須）
background_file = st.file_uploader("1. 上傳背景圖片模板（JPG/PNG，必填）", type=["jpg", "png", "jpeg"])
if background_file:
    background = Image.open(background_file)
    st.image(background, caption="您的背景模板預覽", use_column_width=True)
else:
    st.stop()

# 上傳資料檔（必須）
data_file = st.file_uploader("2. 上傳資料檔（CSV 或 Excel，必填）", type=["csv", "xlsx", "xls"])
if data_file:
    if data_file.name.endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)
    
    st.success("資料上傳成功！")
    st.dataframe(df.head(10))
    st.write(f"總共 {len(df)} 筆資料")
else:
    st.stop()

# （可選）上傳自訂字體
font_file = st.file_uploader("3. （可選）上傳中文字體檔（.ttf，推薦用來避免亂碼）", type=["ttf"])

# 選擇要疊加的欄位
columns = df.columns.tolist()
selected_column = st.selectbox("選擇要疊加的文字欄位（例如「姓名」）", columns)

# 過濾特定人員
st.write("### 選擇要生成的資料（不選則全部生成）")
selected_names = st.multiselect(f"從「{selected_column}」欄位選擇特定人員", df[selected_column].unique().tolist())

if selected_names:
    target_df = df[df[selected_column].isin(selected_names)]
else:
    target_df = df

st.write(f"即將生成 **{len(target_df)}** 張圖片")

# 文字設定
st.subheader("文字樣式設定")
col1, col2, col3, col4 = st.columns(4)
with col1:
    pos_x = st.number_input("X 位置（從左邊算起）", min_value=0, value=int(background.width / 2))
with col2:
    pos_y = st.number_input("Y 位置（從上邊算起）", min_value=0, value=int(background.height / 2))
with col3:
    font_size = st.number_input("字體大小", min_value=10, value=80)
with col4:
    text_color = st.color_picker("文字顏色", "#000000")

# 文字對齊
align = st.selectbox("文字對齊方式", ["左", "中", "右"])

# 載入字體
if font_file:
    font = ImageFont.truetype(font_file, font_size)
    st.success("已使用您上傳的自訂字體")
else:
    try:
        font = ImageFont.truetype("arial.ttf", font_size)
    except:
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
            st.warning("未上傳字體，使用系統預設（可能中文亂碼），建議上傳 .ttf 字體檔")

# 生成按鈕
if st.button("🔥 開始批量生成所有圖片", type="primary"):
    with st.spinner("正在生成，請稍候..."):
        output_images = []
        for idx, row in target_df.iterrows():
            img = background.copy()
            draw = ImageDraw.Draw(img)
            text = str(row[selected_column])
            
            # 計算對齊位置
            if align == "中":
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                x = pos_x - text_width // 2
            elif align == "右":
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                x = pos_x - text_width
            else:
                x = pos_x
            
            draw.text((x, pos_y), text, font=font, fill=text_color)
            
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            buf.seek(0)
            filename = f"{text.replace('/', '_')}_{idx+1}.png"
            output_images.append((filename, buf))

        if output_images:
            st.image(output_images[0][1], caption="第一張預覽（其他相似）")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for name, buf in output_images:
                buf.seek(0)
                zip_file.writestr(name, buf.read())
        zip_buffer.seek(0)

        st.download_button(
            label="📥 下載所有圖片（ZIP 壓縮包）",
            data=zip_buffer,
            file_name="generated_certificates.zip",
            mime="application/zip"
        )
        
        st.success("生成完成！下載後解壓即可列印。")
        st.balloons()

st.caption("此應用完全在您的瀏覽器與臨時伺服器處理，資料不會永久儲存。")