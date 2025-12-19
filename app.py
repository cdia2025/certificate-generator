import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile
import requests
from pathlib import Path

st.set_page_config(page_title="Mail Merge 證書生成器 V2", layout="wide")

# --- 函數定義 ---

def get_font_path(font_type, font_file=None):
    """
    獲取可用的字體路徑，若在 Linux 環境且無字體則下載思源黑體
    """
    if font_type == "custom" and font_file is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.ttf') as tmp:
            tmp.write(font_file.getvalue())
            return tmp.name

    # 系統路徑清單
    paths = {
        "msjh": ["C:/Windows/Fonts/msjh.ttc", "/System/Library/Fonts/STHeiti Light.ttc", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"],
        "dfkai": ["C:/Windows/Fonts/DFKai-SB.ttf", "/System/Library/Fonts/Kaiti.ttc", "/usr/share/fonts/truetype/arphic/uming.ttc"],
        "pmingliu": ["C:/Windows/Fonts/pmingliu.ttc", "/System/Library/Fonts/Songti.ttc"],
        "arial": ["C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Arial.ttf"]
    }

    selected_paths = paths.get(font_type, [])
    for p in selected_paths:
        if os.path.exists(p):
            return p
    
    # 備用方案：如果都沒有，下載思源黑體 (Noto Sans TC)
    backup_font = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(backup_font):
        try:
            url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
            response = requests.get(url)
            with open(backup_font, "wb") as f:
                f.write(response.content)
            return backup_font
        except:
            return None
    return backup_font

def load_font_safely(size, font_path):
    try:
        if font_path and os.path.exists(font_path):
            return ImageFont.truetype(font_path, size)
    except:
        pass
    return ImageFont.load_default()

# --- 介面開始 ---

st.title("✉️ Mail Merge 式多欄位證書生成器")

# 初始化 session_state 用於記憶設定
if "settings" not in st.session_state:
    st.session_state.settings = {}

# 1. 上傳區域
upload_col1, upload_col2, upload_col3 = st.columns([2, 2, 1])

with upload_col1:
    background_file = st.file_uploader("📁 1. 上傳背景圖片", type=["jpg", "png", "jpeg"])

with upload_col2:
    data_file = st.file_uploader("📊 2. 上傳資料檔", type=["csv", "xlsx", "xls"])

with upload_col3:
    font_file = st.file_uploader("🔤 3. 字體檔 (選填 .ttf)", type=["ttf"])

if not background_file or not data_file:
    st.info("請先上傳背景圖與資料檔以開始操作")
    st.stop()

# 讀取背景
background = Image.open(background_file)
bg_width, bg_height = background.size

# 讀取資料
try:
    if data_file.name.lower().endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)
except Exception as e:
    st.error(f"資料讀取失敗：{e}")
    st.stop()

st.divider()

# 2. 主要佈局
main_left, main_right = st.columns([2, 3], gap="large")

with main_left:
    st.header("🔧 參數設定")
    
    # --- 3. 批量選擇名單 ---
    st.subheader("👥 對象選擇")
    filter_column = st.selectbox("識別欄位 (用於檔名)", df.columns)
    all_options = df[filter_column].astype(str).unique().tolist()
    
    select_all = st.checkbox("全選所有名單", value=False)
    if select_all:
        selected_names = st.multiselect("目標對象", options=all_options, default=all_options)
    else:
        selected_names = st.multiselect("目標對象", options=all_options, default=all_options[:3] if len(all_options)>3 else all_options)
    
    target_df = df[df[filter_column].astype(str).isin(selected_names)]
    st.caption(f"已選擇 {len(target_df)} 筆資料")

    # --- 欄位選擇與設定 (記憶功能) ---
    st.subheader("📋 顯示欄位設定")
    selected_columns = st.multiselect("要顯示在證書上的欄位", df.columns, default=[df.columns[0]])
    
    # 字體選擇
    font_options = {"微軟正黑體": "msjh", "標楷體": "dfkai", "新細明體": "pmingliu", "Arial": "arial", "自訂字體": "custom"}
    selected_font_key = st.selectbox("選擇字體類型", list(font_options.keys()))
    current_font_path = get_font_path(font_options[selected_font_key], font_file)

    for col in selected_columns:
        # 如果該欄位之前沒有設定過，才給予預設值 (保留上次設定)
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {
                "x": bg_width // 2, "y": bg_height // 2,
                "size": 60, "color": "#000000", "align": "中", "bold": False
            }
        
        with st.expander(f"設定: {col}", expanded=True):
            c1, c2 = st.columns(2)
            with c1:
                st.session_state.settings[col]["x"] = st.number_input(f"{col} X 座標", 0, bg_width, st.session_state.settings[col]["x"], key=f"x_{col}")
                st.session_state.settings[col]["size"] = st.number_input(f"{col} 大小", 10, 500, st.session_state.settings[col]["size"], key=f"s_{col}")
            with c2:
                st.session_state.settings[col]["y"] = st.number_input(f"{col} Y 座標", 0, bg_height, st.session_state.settings[col]["y"], key=f"y_{col}")
                st.session_state.settings[col]["color"] = st.color_picker(f"{col} 顏色", st.session_state.settings[col]["color"], key=f"c_{col}")
            
            st.session_state.settings[col]["align"] = st.radio(f"{col} 對齊", ["左", "中", "右"], index=["左", "中", "right"].index("中") if st.session_state.settings[col]["align"]=="中" else 0, horizontal=True, key=f"a_{col}")
            st.session_state.settings[col]["bold"] = st.checkbox("粗體效果 (模擬)", value=st.session_state.settings[col]["bold"], key=f"b_{col}")

with main_right:
    st.header("👁️ 即時預覽")
    
    # 1. 預覽圖大小控制
    preview_scale = st.slider("調整工作圖示預覽大小 (%)", 10, 100, 60)
    
    if len(target_df) > 0:
        preview_row = target_df.iloc[0]
        preview_img = background.copy()
        draw = ImageDraw.Draw(preview_img)
        
        for col in selected_columns:
            s = st.session_state.settings[col]
            text = str(preview_row[col])
            font = load_font_safely(s["size"], current_font_path)
            
            # 計算座標
            try:
                bbox = draw.textbbox((0, 0), text, font=font)
                w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
            except:
                w, h = len(text) * s["size"] * 0.7, s["size"]
            
            final_x = s["x"]
            if s["align"] == "中": final_x -= w // 2
            elif s["align"] == "右": final_x -= w
            
            # 繪製文字
            if s["bold"]:
                for dx, dy in [(-1,-1), (1,1), (-1,1), (1,-1)]:
                    draw.text((final_x+dx, s["y"]+dy), text, font=font, fill=s["color"])
            draw.text((final_x, s["y"]), text, font=font, fill=s["color"])
            
            # 輔助線
            draw.line([(0, s["y"]), (bg_width, s["y"])], fill="#FF000033", width=2)
            draw.line([(s["x"], 0), (s["x"], bg_height)], fill="#0000FF33", width=2)

        # 顯示預覽圖 (套用縮放寬度)
        st.image(preview_img, use_column_width=False, width=int(bg_width * (preview_scale/100)))
        st.caption(f"預覽顯示為原始尺寸的 {preview_scale}%")

# 3. 生成與下載
st.divider()
if st.button("🚀 開始大量生成所有選定證書", type="primary", use_container_width=True):
    if not target_df.empty:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            progress_bar = st.progress(0)
            for idx, (i, row) in enumerate(target_df.iterrows()):
                img = background.copy()
                draw = ImageDraw.Draw(img)
                
                for col in selected_columns:
                    s = st.session_state.settings[col]
                    text = str(row[col])
                    font = load_font_safely(s["size"], current_font_path)
                    
                    try:
                        bbox = draw.textbbox((0, 0), text, font=font)
                        w = bbox[2] - bbox[0]
                    except: w = len(text) * s["size"] * 0.7
                    
                    final_x = s["x"]
                    if s["align"] == "中": final_x -= w // 2
                    elif s["align"] == "右": final_x -= w
                    
                    if s["bold"]:
                        for dx, dy in [(-1,-1), (1,1), (-1,1), (1,-1)]:
                            draw.text((final_x+dx, s["y"]+dy), text, font=font, fill=s["color"])
                    draw.text((final_x, s["y"]), text, font=font, fill=s["color"])
                
                img_buf = io.BytesIO()
                img.save(img_buf, format="PNG")
                filename = f"{row[filter_column]}.png".replace("/", "_")
                zf.writestr(filename, img_buf.getvalue())
                progress_bar.progress((idx + 1) / len(target_df))
        
        st.download_button(
            "📥 下載打包好的證書 (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="certificates_export.zip",
            mime="application/zip",
            use_container_width=True
        )
        st.balloons()
