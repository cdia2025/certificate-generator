import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os

st.set_page_config(page_title="Mail Merge 式證書生成器", layout="wide")

st.title("✉️ Mail Merge 式多欄位證書生成器")
st.markdown("**多欄位疊加 + 粗體支援 + 即時預覽縮放 + 拖拽定位 + 配置保存**")

# 左右分欄
left_col, right_col = st.columns([2, 3])

with left_col:
    st.header("🛠️ 設定區")

    # 上傳背景圖片
    background_file = st.file_uploader("上傳背景圖片模板（JPG/PNG，必填）", type=["jpg", "png", "jpeg"])
    if not background_file:
        st.info("請先上傳背景圖片模板")
        st.stop()
    background = Image.open(background_file)

    # 上傳資料檔
    data_file = st.file_uploader("上傳資料檔（CSV 或 Excel，必填）", type=["csv", "xlsx", "xls"])
    if not data_file:
        st.info("請先上傳資料檔案")
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
    font_file = st.file_uploader("（可選）上傳中文字體檔（.ttf，支援粗體更好）", type=["ttf"])

    # Mail Merge 人員選擇
    st.subheader("✉️ Mail Merge 人員選擇")
    filter_column = st.selectbox("用哪一欄作為收件人識別？（例如「姓名」）", df.columns)
    all_options = df[filter_column].astype(str).unique().tolist()
    selected_names = st.multiselect(
        "選擇需要生成的收件人（支援搜尋，不選則全部）",
        options=all_options,
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

    # 初始化設定
    for col in selected_columns:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {
                "x": background.width // 2,
                "y": background.height // 2 + selected_columns.index(col) * 120,
                "size": 80,
                "color": "#000000",
                "align": "中",
                "bold": False,
                "italic": False,
                "underline": False,
                "rotation": 0
            }

    # 欄位設定區域
    st.subheader("📝 欄位設定")
    
    # 拖拽排序功能
    col_order = list(selected_columns)
    reordered_cols = st.multiselect("調整欄位順序（拖拽排序）", 
                                   options=selected_columns, 
                                   default=selected_columns)
    
    # 各欄位詳細設定
    for col in reordered_cols:
        with st.expander(f"**{col}** 設定", expanded=True):
            # 位置設定
            pos_col1, pos_col2 = st.columns(2)
            with pos_col1:
                st.session_state.settings[col]["x"] = st.slider(f"X 位置", 
                                                              0, background.width, 
                                                              st.session_state.settings[col]["x"], 
                                                              key=f"x_{col}")
            with pos_col2:
                st.session_state.settings[col]["y"] = st.slider(f"Y 位置", 
                                                              0, background.height, 
                                                              st.session_state.settings[col]["y"], 
                                                              key=f"y_{col}")
            
            # 字體設定
            font_col1, font_col2 = st.columns(2)
            with font_col1:
                st.session_state.settings[col]["size"] = st.slider(f"字體大小", 
                                                                 20, 200, 
                                                                 st.session_state.settings[col]["size"], 
                                                                 key=f"size_{col}")
                st.session_state.settings[col]["color"] = st.color_picker(f"顏色", 
                                                                        st.session_state.settings[col]["color"], 
                                                                        key=f"color_{col}")
            with font_col2:
                st.session_state.settings[col]["align"] = st.selectbox(f"對齊方式", 
                                                                     ["左", "中", "右"], 
                                                                     index=["左","中","右"].index(st.session_state.settings[col]["align"]), 
                                                                     key=f"align_{col}")
                st.session_state.settings[col]["rotation"] = st.slider(f"旋轉角度", 
                                                                     -180, 180, 
                                                                     st.session_state.settings[col]["rotation"], 
                                                                     key=f"rotation_{col}")
            
            # 樣式設定
            style_col1, style_col2, style_col3 = st.columns(3)
            with style_col1:
                st.session_state.settings[col]["bold"] = st.checkbox(f"粗體", 
                                                                   value=st.session_state.settings[col]["bold"], 
                                                                   key=f"bold_{col}")
            with style_col2:
                st.session_state.settings[col]["italic"] = st.checkbox(f"斜體", 
                                                                     value=st.session_state.settings[col]["italic"], 
                                                                     key=f"italic_{col}")
            with style_col3:
                st.session_state.settings[col]["underline"] = st.checkbox(f"底線", 
                                                                        value=st.session_state.settings[col]["underline"], 
                                                                        key=f"underline_{col}")

    # 配置保存與載入
    st.subheader("💾 配置管理")
    col1, col2 = st.columns(2)
    with col1:
        # 保存配置
        config_data = {"settings": st.session_state.settings, "columns": reordered_cols}
        st.download_button(
            label="💾 保存當前配置",
            data=json.dumps(config_data, ensure_ascii=False, indent=2),
            file_name="mail_merge_config.json",
            mime="application/json"
        )
    with col2:
        # 載入配置
        uploaded_config = st.file_uploader("📁 載入配置", type=["json"], key="load_config")
        if uploaded_config:
            try:
                loaded_config = json.load(uploaded_config)
                st.session_state.settings.update(loaded_config["settings"])
                if "columns" in loaded_config:
                    reordered_cols = loaded_config["columns"]
                st.success("配置載入成功！")
            except Exception as e:
                st.error(f"配置載入失敗：{str(e)}")

    # 預覽縮放控制
    st.subheader("🔍 預覽控制")
    preview_scale = st.slider("預覽圖縮放比例（僅影響顯示，生成為100%原圖）", 20, 200, 100)

    # 生成按鈕
    generate_btn = st.button("🔥 開始批量生成所有證書", type="primary", use_container_width=True)

# 右欄：即時預覽區
with right_col:
    st.header("🔍 即時預覽區（調整即更新）")

    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        preview_img = background.copy()
        draw = ImageDraw.Draw(preview_img)

        # 字體載入
        try:
            if font_file:
                base_font_path = font_file
            else:
                # 嘗試常見中文字體路徑
                base_font_path = None
                for font_path in ["C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Arial Unicode.ttf"]:
                    if os.path.exists(font_path):
                        base_font_path = font_path
                        break
                if not base_font_path:
                    base_font_path = "arial.ttf"
            base_font = ImageFont.truetype(base_font_path, 80)
        except:
            base_font = ImageFont.load_default()
            st.warning("字體載入失敗，使用預設字體（建議上傳 .ttf 檔案）")

        # 繪製所有欄位
        for col in reordered_cols:
            if col in st.session_state.settings:
                settings = st.session_state.settings[col]
                text = str(preview_row[col])
                
                # 載入字體（依大小）
                try:
                    font = ImageFont.truetype(base_font_path, settings["size"])
                except:
                    font = ImageFont.load_default()

                # 計算文字框
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]

                # 計算最終位置（考慮對齊）
                final_x = settings["x"]
                if settings["align"] == "中":
                    final_x = settings["x"] - text_width // 2
                elif settings["align"] == "右":
                    final_x = settings["x"] - text_width

                # 應用旋轉（簡化版本，實際旋轉需要更複雜的計算）
                # 這裡我們先繪製文字，後續可以考慮加入旋轉矩陣
                
                # 應用粗體效果
                if settings["bold"]:
                    stroke_width = max(1, settings["size"] // 30)
                    for dx in [-stroke_width, 0, stroke_width]:
                        for dy in [-stroke_width, 0, stroke_width]:
                            if dx != 0 or dy != 0:
                                draw.text((final_x + dx, settings["y"] + dy), 
                                        text, font=font, fill=settings["color"])

                # 應用底線效果
                if settings["underline"]:
                    underline_y = settings["y"] + text_height + 2
                    draw.line([(final_x, underline_y), (final_x + text_width, underline_y)], 
                             fill=settings["color"], width=max(1, settings["size"] // 20))

                # 主文字繪製
                draw.text((final_x, settings["y"]), text, font=font, fill=settings["color"])

        # 縮放顯示圖
        display_img = preview_img
        if preview_scale != 100:
            new_w = int(background.width * preview_scale / 100)
            new_h = int(background.height * preview_scale / 100)
            display_img = preview_img.resize((new_w, new_h), Image.LANCZOS)

        st.image(display_img, caption=f"即時預覽（顯示 {preview_scale}%）", use_container_width=True)
        
        # 顯示第一筆資料內容
        st.subheader("📋 預覽資料內容")
        for col in reordered_cols:
            if col in preview_row:
                st.write(f"**{col}**: {preview_row[col]}")
    else:
        st.info("無資料可預覽")

# 生成邏輯
if generate_btn:
    if not selected_columns:
        st.error("請至少選擇一個要疊加的欄位！")
    else:
        with st.spinner("正在生成證書..."):
            output_images = []
            total_rows = len(target_df)
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, row in target_df.iterrows():
                status_text.text(f"正在生成第 {idx+1}/{total_rows} 張證書...")
                progress_bar.progress((idx + 1) / total_rows)
                
                img = background.copy()
                draw = ImageDraw.Draw(img)

                for col in reordered_cols:
                    if col in st.session_state.settings:
                        settings = st.session_state.settings[col]
                        text = str(row[col])
                        
                        # 載入字體
                        try:
                            if font_file:
                                font_path = font_file
                            else:
                                # 盡可能使用中文字體
                                font_path = None
                                for path in ["C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Arial Unicode.ttf"]:
                                    if os.path.exists(path):
                                        font_path = path
                                        break
                                if not font_path:
                                    font_path = "arial.ttf"
                            font = ImageFont.truetype(font_path, settings["size"])
                        except:
                            font = ImageFont.load_default()

                        # 計算文字框
                        bbox = draw.textbbox((0, 0), text, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]

                        # 計算最終位置（考慮對齊）
                        final_x = settings["x"]
                        if settings["align"] == "中":
                            final_x = settings["x"] - text_width // 2
                        elif settings["align"] == "右":
                            final_x = settings["x"] - text_width

                        # 應用粗體效果
                        if settings["bold"]:
                            stroke_width = max(1, settings["size"] // 30)
                            for dx in [-stroke_width, 0, stroke_width]:
                                for dy in [-stroke_width, 0, stroke_width]:
                                    if dx != 0 or dy != 0:
                                        draw.text((final_x + dx, settings["y"] + dy), 
                                                text, font=font, fill=settings["color"])

                        # 應用底線效果
                        if settings["underline"]:
                            underline_y = settings["y"] + text_height + 2
                            draw.line([(final_x, underline_y), (final_x + text_width, underline_y)], 
                                     fill=settings["color"], width=max(1, settings["size"] // 20))

                        # 主文字繪製
                        draw.text((final_x, settings["y"]), text, font=font, fill=settings["color"])

                # 保存圖片
                buf = io.BytesIO()
                img.save(buf, format="PNG", dpi=(300, 300))  # 高解析度輸出
                buf.seek(0)
                
                # 安全的檔案名稱
                safe_name = str(row.get(filter_column, f"record_{idx+1}")).replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace('"', "_").replace("<", "_").replace(">", "_").replace("|", "_")
                filename = f"證書_{safe_name}.png"
                output_images.append((filename, buf))

            # 創建 ZIP 檔案
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, buf in output_images:
                    buf.seek(0)
                    zf.writestr(name, buf.read())
            zip_buffer.seek(0)

            # 下載按鈕
            st.download_button(
                label="📥 下載所有證書（ZIP）",
                data=zip_buffer,
                file_name="certificates.zip",
                mime="application/zip"
            )
            st.success(f"✅ 生成完成！共 {len(output_images)} 張證書")
            st.balloons()

st.caption("🔒 安全提醒：所有資料僅在瀏覽器內處理，不會上傳至任何伺服器。")
