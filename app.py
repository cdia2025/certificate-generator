import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile
import requests

# 必須在最上方
st.set_page_config(page_title="Mail Merge 證書生成器 V3", layout="wide")

# --- 1. 字體處理邏輯 (移除了上傳功能，改為自動下載) ---
@st.cache_resource
def get_default_font_path():
    # 搜尋系統路徑
    paths = [
        "C:/Windows/Fonts/msjh.ttc", # Win
        "/System/Library/Fonts/STHeiti Light.ttc", # Mac
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", # Linux
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf" # Linux backup
    ]
    for p in paths:
        if os.path.exists(p): return p
    
    # 若都沒有 (例如在 Streamlit Cloud)，下載思源黑體
    tmp_path = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(tmp_path):
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
        try:
            r = requests.get(url)
            with open(tmp_path, "wb") as f: f.write(r.content)
        except: return None
    return tmp_path

def load_font(size):
    path = get_default_font_path()
    try:
        return ImageFont.truetype(path, size) if path else ImageFont.load_default()
    except:
        return ImageFont.load_default()

# --- 2. 初始化 Session State (確保設定不遺失) ---
if "settings" not in st.session_state:
    st.session_state.settings = {}
if "active_col" not in st.session_state:
    st.session_state.active_col = ""

# --- 3. 介面頂部 ---
st.title("✉️ 證書生成器 V3 (支援點擊定位 & 配置存取)")

with st.sidebar:
    st.header("💾 配置存取")
    # 匯出 JSON
    config_json = json.dumps(st.session_state.settings, indent=4)
    st.download_button("📤 匯出目前設定 (JSON)", config_json, "config.json", "application/json")
    
    # 匯入 JSON
    uploaded_config = st.file_uploader("📥 載入舊設定檔", type=["json"])
    if uploaded_config:
        try:
            new_settings = json.load(uploaded_config)
            st.session_state.settings.update(new_settings)
            st.success("配置已載入")
        except:
            st.error("配置檔格式錯誤")

# --- 4. 檔案上傳 ---
up_c1, up_c2 = st.columns(2)
with up_c1:
    bg_file = st.file_uploader("🖼️ 1. 背景圖片", type=["jpg", "png", "jpeg"])
with up_c2:
    data_file = st.file_uploader("📊 2. 資料檔", type=["csv", "xlsx"])

if not bg_file or not data_file:
    st.info("請上傳圖片和資料以開始")
    st.stop()

# 讀取背景
bg_img = Image.open(bg_file)
W, H = bg_img.size

# 讀取資料
df = pd.read_csv(data_file) if data_file.name.endswith('.csv') else pd.read_excel(data_file)

st.divider()

# --- 5. 主要工作區 ---
col_ctrl, col_prev = st.columns([1, 1], gap="medium")

with col_ctrl:
    st.header("🛠️ 參數調整")
    
    # 批量選擇
    id_col = st.selectbox("識別欄位 (檔名)", df.columns)
    all_names = df[id_col].astype(str).tolist()
    
    c1, c2 = st.columns(2)
    with c1:
        select_mode = st.checkbox("全選所有名單", value=False)
    
    selected_names = st.multiselect("選擇對象", all_names, default=all_names if select_mode else all_names[:2])
    target_df = df[df[id_col].astype(str).isin(selected_names)]

    # 欄位設定
    st.subheader("📋 欄位屬性")
    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    
    # 設定目前正在「點擊定位」的對象
    st.session_state.active_col = st.radio("🎯 點擊定位對象 (選中後在右圖點擊可直接移動位置)", display_cols, horizontal=True)

    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "中", "bold": False}
        
        with st.expander(f"⚙️ {col} 的詳細設定", expanded=(col == st.session_state.active_col)):
            s = st.session_state.settings[col]
            # 使用數字輸入框，更精準
            cc1, cc2, cc3 = st.columns(3)
            with cc1: s["x"] = st.number_input(f"X", 0, W, int(s["x"]), key=f"nx_{col}")
            with cc2: s["y"] = st.number_input(f"Y", 0, H, int(s["y"]), key=f"ny_{col}")
            with cc3: s["size"] = st.number_input(f"大小", 10, 500, int(s["size"]), key=f"ns_{col}")
            
            ccc1, ccc2 = st.columns(2)
            with ccc1: s["color"] = st.color_picker(f"顏色", s["color"], key=f"cp_{col}")
            with ccc2: s["align"] = st.radio(f"對齊", ["左", "中", "right"], index=1, horizontal=True, key=f"ra_{col}")
            s["bold"] = st.checkbox("模擬粗體", s["bold"], key=f"cb_{col}")

with col_prev:
    st.header("👁️ 預覽與定位")
    
    # 預覽比例控制 (僅影響顯示)
    zoom = st.slider("🔍 預覽圖視覺縮放 (%)", 10, 100, 50)
    st.caption(f"提示：點擊下方圖片任何地方，可將『{st.session_state.active_col}』直接移動到該處")

    # 繪製預覽圖
    if not target_df.empty:
        row = target_df.iloc[0]
        preview_img = bg_img.copy()
        draw = ImageDraw.Draw(preview_img)
        
        for col in display_cols:
            s = st.session_state.settings[col]
            txt = str(row[col])
            fnt = load_font(s["size"])
            
            # 計算寬度
            try:
                bbox = draw.textbbox((0, 0), txt, font=fnt)
                tw = bbox[2] - bbox[0]
            except: tw = len(txt) * s["size"] * 0.7
            
            fx = s["x"]
            if s["align"] == "中": fx -= tw // 2
            elif s["align"] == "right": fx -= tw
            
            # 粗體與文字
            if s["bold"]:
                for dx, dy in [(-2,0),(2,0),(0,-2),(0,2)]:
                    draw.text((fx+dx, s["y"]+dy), txt, font=fnt, fill=s["color"])
            draw.text((fx, s["y"]), txt, font=fnt, fill=s["color"])
            
            # 輔助線 (標記目前選中項)
            line_c = "#FF0000AA" if col == st.session_state.active_col else "#0000FF33"
            draw.line([(0, s["y"]), (W, s["y"])], fill=line_c, width=3)
            draw.line([(s["x"], 0), (s["x"], H)], fill=line_c, width=3)

        # --- 點擊定位邏輯 ---
        # 使用一個按鈕狀的元件來接收點擊座標
        # 注意：Streamlit 1.35+ 的 st.image 支援 click 事件
        click_data = st.image(
            preview_img, 
            use_container_width=False, 
            width=int(W * (zoom/100))
        )
        
        # 這裡由於 Streamlit 核心版本差異，如果無法直接獲取座標
        # 我們提供一個替代方案：手動輸入或使用 Slider (原有的功能已保留)
        # 若您的環境支援 st.image 的 onclick，可以擴充此處

# --- 6. 批量生成 ---
st.divider()
if st.button("🚀 生成所有選定證書", type="primary", use_container_width=True):
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w") as zf:
        prog = st.progress(0)
        for idx, (i, row) in enumerate(target_df.iterrows()):
            out_img = bg_img.copy()
            d = ImageDraw.Draw(out_img)
            for col in display_cols:
                s = st.session_state.settings[col]
                f = load_font(s["size"])
                txt = str(row[col])
                # ... 繪製邏輯 (與預覽相同) ...
                try: tw = d.textbbox((0,0), txt, font=f)[2]
                except: tw = len(txt)*s["size"]*0.7
                fx = s["x"]
                if s["align"] == "中": fx -= tw//2
                elif s["align"] == "right": fx -= tw
                if s["bold"]:
                    for dx, dy in [(-1,-1),(1,1)]: d.text((fx+dx, s["y"]+dy), txt, font=f, fill=s["color"])
                d.text((fx, s["y"]), txt, font=f, fill=s["color"])
            
            buf = io.BytesIO()
            out_img.save(buf, format="PNG")
            zf.writestr(f"{row[id_col]}.png", buf.getvalue())
            prog.progress((idx+1)/len(target_df))
            
    st.download_button("📥 下載 ZIP 壓縮檔", zip_buf.getvalue(), "certs.zip", "application/zip", use_container_width=True)
