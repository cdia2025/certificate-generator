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
st.markdown("**精準座標定位 + 中心對齊優化 + 自由調整預覽**")

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
bg_width, bg_height = background.size

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
                "x": bg_width // 2,  # 預設置中
                "y": bg_height // 3 + selected_columns.index(col) * 100,
                "size": 60,
                "color": "#000000",
                "align": "中",  # 預設置中對齊
                "bold": False,
                "italic": False,
                "anchor": "center"  # 使用英文，避免中文key問題
            }

    # 欄位設定
    st.subheader("⚙️ 欄位詳細設定")
    
    settings_container = st.container()
    
    with settings_container:
        for i, col in enumerate(selected_columns):
            with st.expander(f"📝 {col}", expanded=True):
                # 座標與對齊說明
                st.caption(f"背景尺寸: {bg_width}×{bg_height}px | 置中座標: ({bg_width//2}, {bg_height//2})")
                
                # 位置設定
                pos1, pos2 = st.columns(2)
                with pos1:
                    st.session_state.settings[col]["x"] = st.slider(
                        "X 座標", 0, bg_width, 
                        st.session_state.settings[col]["x"], 
                        key=f"x_{i}_{col}",
                        help=f"範圍: 0~{bg_width}, 置中點: {bg_width//2}"
                    )
                with pos2:
                    st.session_state.settings[col]["y"] = st.slider(
                        "Y 座標", 0, bg_height, 
                        st.session_state.settings[col]["y"], 
                        key=f"y_{i}_{col}",
                        help=f"範圍: 0~{bg_height}, 置中點: {bg_height//2}"
                    )
                
                # 對齊方式與錨點
                align_col1, align_col2 = st.columns(2)
                with align_col1:
                    align_options = ["左", "中", "右"]
                    current_align_index = 0
                    if st.session_state.settings[col]["align"] in align_options:
                        current_align_index = align_options.index(st.session_state.settings[col]["align"])
                    st.session_state.settings[col]["align"] = st.radio(
                        "文字對齊", align_options, 
                        index=current_align_index,
                        key=f"align_{i}_{col}",
                        horizontal=True
                    )
                with align_col2:
                    anchor_options = ["left_top", "center", "right_bottom"]
                    anchor_labels = ["左上", "中心", "右下"]
                    current_anchor = st.session_state.settings[col]["anchor"]
                    current_anchor_index = 0
                    if current_anchor in anchor_options:
                        current_anchor_index = anchor_options.index(current_anchor)
                    
                    selected_anchor_index = st.radio(
                        "錨點", anchor_labels, 
                        index=current_anchor_index,
                        key=f"anchor_{i}_{col}",
                        horizontal=True
                    )
                    # 轉換回英文key
                    st.session_state.settings[col]["anchor"] = anchor_options[current_anchor_index]

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
                
                # 樣式
                style_col1, style_col2 = st.columns(2)
                with style_col1:
                    st.session_state.settings[col]["bold"] = st.checkbox(
                        "粗體", 
                        value=st.session_state.settings[col]["bold"], 
                        key=f"bold_{i}_{col}"
                    )
                with style_col2:
                    st.session_state.settings[col]["italic"] = st.checkbox(
                        "斜體", 
                        value=st.session_state.settings[col]["italic"], 
                        key=f"italic_{i}_{col}"
                    )
                
                # 座標計算說明
                align_desc = {
                    "左": "左對齊：文字從指定 X 座標開始",
                    "中": "置中對齊：文字以指定 X 座標為中心",
                    "右": "右對齊：文字在指定 X 座標結束"
                }
                st.caption(f"說明：{align_desc.get(st.session_state.settings[col]['align'], '未設定')}")

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
                # 更新設定，保持向後兼容
                for col_key, settings_val in loaded_config["settings"].items():
                    if col_key not in st.session_state.settings:
                        st.session_state.settings[col_key] = settings_val
                    else:
                        st.session_state.settings[col_key].update(settings_val)
                st.success("配置載入成功！")
            except Exception as e:
                st.error(f"配置載入失敗：{str(e)}")

    # 預覽控制
    st.subheader("🔍 預覽控制")
    preview_scale = st.slider("預覽縮放", 20, 200, 80, key="preview_scale")
    
    # 預覽尺寸調整
    preview_size = st.select_slider(
        "預覽尺寸",
        options=["小", "中", "大", "超大"],
        value="中"
    )
    
    size_map = {"小": 300, "中": 500, "大": 700, "超大": 900}
    max_display_width = size_map[preview_size]

    # 生成按鈕
    if st.button("🚀 開始生成", type="primary", use_container_width=True):
        st.session_state.generate_clicked = True

# 右側預覽區域
with main_right:
    st.header("👁️ 即時預覽")
    
    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        
        # 字體載入函數
        def load_font(size):
            try:
                if font_file:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.ttf') as tmp_font:
                        tmp_font.write(font_file.getvalue())
                        return ImageFont.truetype(tmp_font.name, size)
                else:
                    # 系統字體嘗試順序
                    font_paths = [
                        "/System/Library/Fonts/Arial Unicode.ttf",  # macOS
                        "/System/Library/Fonts/Helvetica.ttc",     # macOS
                        "C:/Windows/Fonts/msyh.ttc",               # Windows 中易黑體
                        "C:/Windows/Fonts/simhei.ttf",             # Windows 黑體
                        "C:/Windows/Fonts/msyhbd.ttc",             # Windows 粗體黑體
                        "C:/Windows/Fonts/arial.ttf",              # Windows Arial
                        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",  # Linux
                        "/usr/share/fonts/TTF/DejaVuSans.ttf"      # Linux alternative
                    ]
                    for font_path in font_paths:
                        if os.path.exists(font_path):
                            try:
                                return ImageFont.truetype(font_path, size)
                            except:
                                continue
            except:
                pass
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
                
                # 計算文字尺寸
                try:
                    bbox = draw.textbbox((0, 0), text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                except:
                    # 備用計算方式
                    text_width = len(text) * settings["size"] * 0.6
                    text_height = settings["size"]
                
                # 根據對齊方式計算最終位置
                final_x = settings["x"]
                if settings["align"] == "中":
                    final_x = settings["x"] - text_width // 2
                elif settings["align"] == "右":
                    final_x = settings["x"] - text_width
                
                # 根據錨點調整 Y 座標
                final_y = settings["y"]
                if settings["anchor"] == "center":
                    final_y = settings["y"] - text_height // 2
                elif settings["anchor"] == "right_bottom":
                    final_y = settings["y"] - text_height
                
                # 繪製粗體效果
                if settings["bold"]:
                    for dx in [-1, 0, 1]:
                        for dy in [-1, 0, 1]:
                            if dx != 0 or dy != 0:
                                draw.text((final_x + dx, final_y + dy), 
                                        text, font=font, fill=settings["color"])
                
                # 繪製主文字
                draw.text((final_x, final_y), text, font=font, fill=settings["color"])
                
                # 繪製定位參考線（虛線）
                # 水平線
                draw.line([(0, final_y), (bg_width, final_y)], fill="#FF0000", width=1)
                # 垂直線
                draw.line([(final_x, 0), (final_x, bg_height)], fill="#0000FF", width=1)

        # 計算顯示尺寸
        aspect_ratio = bg_height / bg_width
        display_width = min(max_display_width, bg_width)
        display_height = int(display_width * aspect_ratio)
        
        # 縮放並顯示預覽
        scaled_preview = preview_img.resize((display_width, display_height), Image.LANCZOS)
        
        st.image(scaled_preview, 
                caption=f"預覽 ({preview_scale}% | {display_width}×{display_height}px)", 
                use_column_width=True)
        
        # 顯示當前資料和座標信息
        st.subheader("📋 預覽資料 & 座標信息")
        col_info1, col_info2 = st.columns(2)
        
        with col_info1:
            st.write("**原始座標**")
            for col in selected_columns:
                settings = st.session_state.settings[col]
                st.write(f"{col}: ({settings['x']}, {settings['y']})")
        
        with col_info2:
            st.write("**實際繪製座標**")
            for col in selected_columns:
                settings = st.session_state.settings[col]
                text = str(preview_row[col])
                
                # 重新計算實際座標
                font = load_font(settings["size"])
                try:
                    bbox = draw.textbbox((0, 0), text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                except:
                    text_width = len(text) * settings["size"] * 0.6
                    text_height = settings["size"]
                
                actual_x = settings["x"]
                if settings["align"] == "中":
                    actual_x = settings["x"] - text_width // 2
                elif settings["align"] == "右":
                    actual_x = settings["x"] - text_width
                
                actual_y = settings["y"]
                if settings["anchor"] == "center":
                    actual_y = settings["y"] - text_height // 2
                elif settings["anchor"] == "right_bottom":
                    actual_y = settings["y"] - text_height
                
                st.write(f"{col}: ({actual_x:.0f}, {actual_y:.0f})")

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
                    
                    # 計算實際繪製位置
                    final_x = settings["x"]
                    if settings["align"] == "中":
                        final_x = settings["x"] - text_width // 2
                    elif settings["align"] == "右":
                        final_x = settings["x"] - text_width
                    
                    final_y = settings["y"]
                    if settings["anchor"] == "center":
                        final_y = settings["y"] - text_height // 2
                    elif settings["anchor"] == "right_bottom":
                        final_y = settings["y"] - text_height
                    
                    # 繪製粗體效果
                    if settings["bold"]:
                        for dx in [-1, 0, 1]:
                            for dy in [-1, 0, 1]:
                                if dx != 0 or dy != 0:
                                    draw.text((final_x + dx, final_y + dy), 
                                            text, font=font, fill=settings["color"])
                    
                    # 繪製主文字
                    draw.text((final_x, final_y), text, font=font, fill=settings["color"])
            
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
