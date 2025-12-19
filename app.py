import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile

st.set_page_config(page_title="Mail Merge 式證書生成器", layout="wide")

st.title("✉️ Mail Merge 式多欄位證書生成器")
st.markdown("**多欄位疊加 + 中文字體支援 + 即時預覽 + 拖拽定位**")

# 主容器
main_container = st.container()
with main_container:
    # 上傳區域
    upload_col1, upload_col2, upload_col3 = st.columns([2, 2, 1])
    
    with upload_col1:
        background_file = st.file_uploader("📁 上傳背景圖片", type=["jpg", "png", "jpeg"])
    
    with upload_col2:
        data_file = st.file_uploader("📊 上傳資料檔", type=["csv", "xlsx", "xls"])
    
    with upload_col3:
        font_file = st.file_uploader("🔤 字體檔 (.ttf)", type=["ttf"])

# 檢查必要檔案
if not background_file or not data_file:
    st.info("請先上傳背景圖片和資料檔案")
    st.stop()

background = Image.open(background_file)

# 讀取資料
try:
    if data_file.name.lower().endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)
    st.success(f"資料載入成功！共 {len(df)} 筆記錄")
except Exception as e:
    st.error(f"資料讀取失敗：{str(e)}")
    st.stop()

# 分割線
st.divider()

# 主要功能區域
main_left, main_right = st.columns([1, 1], gap="large")

with main_left:
    st.header("🔧 設定區域")
    
    # Mail Merge 選擇
    filter_column = st.selectbox("👤 選擇識別欄位", df.columns)
    all_options = df[filter_column].astype(str).unique().tolist()
    selected_names = st.multiselect(
        "👥 選擇目標對象",
        options=all_options,
        placeholder="選擇要生成的對象...",
        default=all_options[:5] if len(all_options) > 5 else all_options
    )
    target_df = df[df[filter_column].astype(str).isin(selected_names)]
    st.info(f"將生成 **{len(target_df)}** 張")

    # 欄位選擇
    st.subheader("📋 選擇要顯示的欄位")
    selected_columns = st.multiselect("選擇欄位", df.columns, default=df.columns[:3] if len(df.columns) >= 3 else df.columns)
    
    if not selected_columns:
        st.warning("請至少選擇一個欄位！")
        st.stop()

    # 初始化設定
    if "settings" not in st.session_state:
        st.session_state.settings = {}
    
    for col in selected_columns:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {
                "x": background.width // 3,
                "y": background.height // 3 + selected_columns.index(col) * 100,
                "size": 60,
                "color": "#000000",
                "align": "左",
                "bold": False,
                "italic": False
            }

    # 欄位設定
    st.subheader("⚙️ 欄位詳細設定")
    
    # 創建可滾動的設定區域
    settings_container = st.container()
    
    with settings_container:
        for i, col in enumerate(selected_columns):
            with st.expander(f"📝 {col}", expanded=True):
                # 位置設定
                pos1, pos2 = st.columns(2)
                with pos1:
                    st.session_state.settings[col]["x"] = st.slider(
                        "X 座標", 0, background.width, 
                        st.session_state.settings[col]["x"], 
                        key=f"x_{i}_{col}"
                    )
                with pos2:
                    st.session_state.settings[col]["y"] = st.slider(
                        "Y 座標", 0, background.height, 
                        st.session_state.settings[col]["y"], 
                        key=f"y_{i}_{col}"
                    )
                
                # 字體設定
                font1, font2 = st.columns(2)
                with font1:
                    st.session_state.settings[col]["size"] = st.slider(
                        "字體大小", 20, 150, 
                        st.session_state.settings[col]["size"], 
                        key=f"size_{i}_{col}"
                    )
                with font2:
                    st.session_state.settings[col]["color"] = st.color_picker(
                        "文字顏色", 
                        st.session_state.settings[col]["color"], 
                        key=f"color_{i}_{col}"
                    )
                
                # 對齊與樣式
                align_style = st.columns(3)
                with align_style[0]:
                    st.session_state.settings[col]["align"] = st.radio(
                        "對齊", ["左", "中", "右"], 
                        index=["左", "中", "右"].index(st.session_state.settings[col]["align"]),
                        key=f"align_{i}_{col}",
                        horizontal=True
                    )
                with align_style[1]:
                    st.session_state.settings[col]["bold"] = st.checkbox(
                        "粗體", 
                        value=st.session_state.settings[col]["bold"], 
                        key=f"bold_{i}_{col}"
                    )
                with align_style[2]:
                    st.session_state.settings[col]["italic"] = st.checkbox(
                        "斜體", 
                        value=st.session_state.settings[col]["italic"], 
                        key=f"italic_{i}_{col}"
                    )

    # 配置管理
    st.subheader("💾 配置管理")
    config_col1, config_col2 = st.columns(2)
    with config_col1:
        config_data = {"settings": st.session_state.settings, "selected_columns": selected_columns}
        st.download_button(
            "💾 保存配置",
            data=json.dumps(config_data, ensure_ascii=False, indent=2),
            file_name="certificate_config.json",
            mime="application/json"
        )
    with config_col2:
        uploaded_config = st.file_uploader("📁 載入配置", type=["json"], key="config_upload")
        if uploaded_config:
            try:
                loaded_config = json.load(uploaded_config)
                st.session_state.settings.update(loaded_config["settings"])
                st.success("配置載入成功！")
            except Exception as e:
                st.error(f"配置載入失敗：{str(e)}")

    # 預覽控制
    st.subheader("🔍 預覽控制")
    preview_scale = st.slider("預覽縮放", 30, 150, 80, key="preview_scale")

    # 生成按鈕
    if st.button("🚀 開始生成", type="primary", use_container_width=True):
        st.session_state.generate_clicked = True

# 右側預覽區域
with main_right:
    st.header("👁️ 即時預覽")
    
    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        
        # 嘗試載入字體
        def load_font(size):
            try:
                if font_file:
                    # 保存上傳的字體到臨時檔案
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.ttf') as tmp_font:
                        tmp_font.write(font_file.getvalue())
                        return ImageFont.truetype(tmp_font.name, size)
                else:
                    # 嘗試系統字體
                    for font_path in [
                        "/System/Library/Fonts/Arial Unicode.ttf",  # macOS
                        "/System/Library/Fonts/Helvetica.ttc",     # macOS
                        "C:/Windows/Fonts/msyh.ttc",               # Windows 中易黑體
                        "C:/Windows/Fonts/simhei.ttf",             # Windows 黑體
                        "C:/Windows/Fonts/msyhbd.ttc",             # Windows 粗體黑體
                        "C:/Windows/Fonts/arial.ttf",              # Windows Arial
                        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
                        "/usr/share/fonts/TTF/DejaVuSans.ttf"      # Linux alternative
                    ]:
                        if os.path.exists(font_path):
                            try:
                                return ImageFont.truetype(font_path, size)
                            except:
                                continue
            except:
                pass
            # 如果都失敗，返回默認字體
            return ImageFont.load_default()

        # 創建預覽圖片
        preview_img = background.copy()
        draw = ImageDraw.Draw(preview_img)

        # 繪製每個選定的欄位
        for col in selected_columns:
            if col in st.session_state.settings:
                settings = st.session_state.settings[col]
                text = str(preview_row[col])
                
                # 載入字體
                font = load_font(settings["size"])
                
                # 計算文字寬度（用於對齊）
                try:
                    bbox = draw.textbbox((0, 0), text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                except:
                    # 如果計算失敗，使用替代方法
                    text_width = len(text) * settings["size"] * 0.6
                    text_height = settings["size"]
                
                # 計算最終 X 位置（根據對齊方式）
                final_x = settings["x"]
                if settings["align"] == "中":
                    final_x = settings["x"] - text_width // 2
                elif settings["align"] == "右":
                    final_x = settings["x"] - text_width
                
                # 繪製文字（如果有粗體需求）
                if settings["bold"]:
                    # 繪製多層來模擬粗體效果
                    for dx in [-1, 0, 1]:
                        for dy in [-1, 0, 1]:
                            if dx != 0 or dy != 0:
                                draw.text((final_x + dx, settings["y"] + dy), 
                                        text, font=font, fill=settings["color"])
                
                # 繪製主文字
                draw.text((final_x, settings["y"]), text, font=font, fill=settings["color"])

        # 縮放並顯示預覽
        new_w = int(background.width * preview_scale / 100)
        new_h = int(background.height * preview_scale / 100)
        display_img = preview_img.resize((new_w, new_h), Image.LANCZOS)
        
        st.image(display_img, caption=f"預覽 ({preview_scale}%)", use_column_width=True)
        
        # 顯示當前資料
        st.subheader("📋 預覽資料")
        for col in selected_columns:
            if col in preview_row:
                st.write(f"**{col}**: `{preview_row[col]}`")
    else:
        st.info("無資料可預覽")

# 生成功能
if hasattr(st.session_state, 'generate_clicked') and st.session_state.generate_clicked:
    with st.spinner("正在生成證書..."):
        output_images = []
        total_count = len(target_df)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, row in target_df.iterrows():
            status_text.text(f"生成中... ({idx+1}/{total_count})")
            progress_bar.progress((idx + 1) / total_count)
            
            # 創建單張證書
            img = background.copy()
            draw = ImageDraw.Draw(img)
            
            for col in selected_columns:
                if col in st.session_state.settings:
                    settings = st.session_state.settings[col]
                    text = str(row[col])
                    
                    # 載入字體
                    font = load_font(settings["size"])
                    
                    # 計算文字尺寸
                    try:
                        bbox = draw.textbbox((0, 0), text, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                    except:
                        text_width = len(text) * settings["size"] * 0.6
                        text_height = settings["size"]
                    
                    # 計算位置
                    final_x = settings["x"]
                    if settings["align"] == "中":
                        final_x = settings["x"] - text_width // 2
                    elif settings["align"] == "右":
                        final_x = settings["x"] - text_width
                    
                    # 繪製粗體效果
                    if settings["bold"]:
                        for dx in [-1, 0, 1]:
                            for dy in [-1, 0, 1]:
                                if dx != 0 or dy != 0:
                                    draw.text((final_x + dx, settings["y"] + dy), 
                                            text, font=font, fill=settings["color"])
                    
                    # 繪製主文字
                    draw.text((final_x, settings["y"]), text, font=font, fill=settings["color"])
            
            # 保存圖片
            buf = io.BytesIO()
            img.save(buf, format="PNG", dpi=(300, 300))
            buf.seek(0)
            
            # 安全的檔案名稱
            safe_name = str(row.get(filter_column, f"cert_{idx+1}"))
            safe_name = "".join(c for c in safe_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"證書_{safe_name}.png"
            output_images.append((filename, buf))
        
        # 創建 ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, buf in output_images:
                buf.seek(0)
                zf.writestr(name, buf.read())
        zip_buffer.seek(0)
        
        # 下載
        st.download_button(
            label="📥 下載所有證書",
            data=zip_buffer,
            file_name="certificates.zip",
            mime="application/zip"
        )
        st.success(f"✅ 生成完成！共 {len(output_images)} 張證書")
        st.balloons()
        
        # 重置生成狀態
        delattr(st.session_state, 'generate_clicked')

st.caption("🔒 資料僅在本機處理，不會上傳至任何地方")
