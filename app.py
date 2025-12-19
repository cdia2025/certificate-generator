import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile
import requests

# 頁面基本設定
st.set_page_config(page_title="證書生成器 V4", layout="wide")

# --- 1. 強化版中文字體載入器 ---
@st.cache_resource
def get_font_resource():
    """確保環境中一定有中文字體可用"""
    # 1. 定義系統可能存在的路徑
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc",            # Windows 微軟正黑
        "C:/Windows/Fonts/dfkai-sb.ttf",        # Windows 標楷體
        "/System/Library/Fonts/STHeiti Light.ttc", # macOS 華文黑體
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", # Linux
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"  # Linux
    ]
    
    for p in font_paths:
        if os.path.exists(p):
            return p

    # 2. 如果系統路徑都沒有，從網路下載思源黑體 (Noto Sans TC)
    target_path = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(target_path):
        # 這是 Google Fonts 的原始下載鏈接 (繁體中文)
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
        try:
            with st.spinner("正在初始化中文字體庫 (僅需執行一次)..."):
                response = requests.get(url, timeout=15)
                with open(target_path, "wb") as f:
                    f.write(response.content)
            return target_path
        except Exception as e:
            st.error(f"字體下載失敗，請檢查網路連線: {e}")
            return None
    return target_path

def load_font(size):
    font_path = get_font_resource()
    try:
        if font_path:
            return ImageFont.truetype(font_path, size)
    except Exception:
        pass
    return ImageFont.load_default()

# --- 2. 初始化 Session State ---
if "settings" not in st.session_state:
    st.session_state.settings = {}

# --- 3. 介面頂部與檔案上傳 ---
st.title("✉️ 專業證書生成器 V4")

# 側邊欄：配置管理
with st.sidebar:
    st.header("💾 配置管理")
    if st.session_state.settings:
        config_json = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
        st.download_button("📤 匯出目前設定 (JSON)", config_json, "cert_config.json", "application/json")
    
    uploaded_config = st.file_uploader("📥 載入舊設定檔", type=["json"])
    if uploaded_config:
        try:
            st.session_state.settings.update(json.load(uploaded_config))
            st.success("配置已載入！")
        except:
            st.error("配置檔解析失敗")

# 上傳區
up1, up2 = st.columns(2)
with up1:
    bg_file = st.file_uploader("🖼️ 1. 上傳證書背景圖", type=["jpg", "png", "jpeg"])
with up2:
    data_file = st.file_uploader("📊 2. 上傳資料檔 (Excel/CSV)", type=["xlsx", "csv"])

if not bg_file or not data_file:
    st.info("👋 請先上傳背景圖片和 Excel/CSV 資料檔開始工作。")
    st.stop()

# 讀取檔案
bg_img = Image.open(bg_file)
W, H = bg_img.size
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

st.divider()

# --- 4. 工作區佈局 ---
col_ctrl, col_prev = st.columns([1, 1], gap="large")

with col_ctrl:
    st.header("🛠️ 參數調整")
    
    # 名單選擇
    id_col = st.selectbox("選擇識別欄位 (用於檔案命名)", df.columns)
    all_items = df[id_col].astype(str).tolist()
    
    c1, c2 = st.columns([1, 2])
    with c1:
        is_all = st.checkbox("全選所有名單")
    with c2:
        selected_items = st.multiselect("選取生成名單", all_items, default=all_items if is_all else all_items[:2])
    
    target_df = df[df[id_col].astype(str).isin(selected_items)]

    # 欄位內容設定
    st.subheader("📋 顯示欄位設定")
    display_cols = st.multiselect("要在證書上顯示的欄位", df.columns, default=[df.columns[0]])
    
    for col in display_cols:
        # 記憶上次設定值
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {
                "x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False
            }
        
        with st.expander(f"📝 欄位：{col}", expanded=True):
            s = st.session_state.settings[col]
            
            # 座標設定
            cc1, cc2 = st.columns(2)
            with cc1:
                s["x"] = st.slider(f"{col} X 位置", 0, W, int(s["x"]), key=f"x_{col}")
            with cc2:
                s["y"] = st.slider(f"{col} Y 位置", 0, H, int(s["y"]), key=f"y_{col}")
            
            # 樣式設定
            cc3, cc4, cc5 = st.columns([1, 1, 1])
            with cc3:
                s["size"] = st.number_input(f"字體大小", 10, 1000, int(s["size"]), key=f"sz_{col}")
            with cc4:
                s["color"] = st.color_picker(f"顏色", s["color"], key=f"cl_{col}")
            with cc5:
                s["align"] = st.selectbox(f"對齊", ["左對齊", "居中", "右對齊"], index=1, key=f"al_{col}")
            
            s["bold"] = st.checkbox("模擬粗體 (文字加粗)", value=s["bold"], key=f"bd_{col}")

with col_prev:
    st.header("👁️ 即時預覽")
    
    # 預覽縮放滑桿 - 修正縮放比例問題
    zoom_percent = st.slider("🔍 調整右側預覽圖顯示大小 (不影響輸出)", 10, 100, 50)
    
    if not target_df.empty:
        # 取第一筆資料做預覽
        row = target_df.iloc[0]
        preview_canvas = bg_img.copy()
        draw = ImageDraw.Draw(preview_canvas)
        
        for col in display_cols:
            s = st.session_state.settings[col]
            txt = str(row[col])
            font = load_font(s["size"])
            
            # 計算寬度以處理對齊
            try:
                # 取得文字框範圍
                left, top, right, bottom = draw.textbbox((0, 0), txt, font=font)
                tw = right - left
            except:
                tw = len(txt) * s["size"] * 0.7 # 估計值備援
            
            final_x = s["x"]
            if s["align"] == "居中":
                final_x -= tw // 2
            elif s["align"] == "右對齊":
                final_x -= tw
            
            # 繪製模擬粗體
            if s["bold"]:
                for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                    draw.text((final_x + dx, s["y"] + dy), txt, font=font, fill=s["color"])
            
            # 繪製主文字
            draw.text((final_x, s["y"]), txt, font=font, fill=s["color"])
            
            # 繪製輔助紅線 (讓用戶知道精確點在哪)
            draw.line([(0, s["y"]), (W, s["y"])], fill="#FF000055", width=2)
            draw.line([(s["x"], 0), (s["x"], H)], fill="#0000FF55", width=2)

        # 顯示預覽圖
        display_w = int(W * (zoom_percent / 100))
        st.image(preview_canvas, width=display_w, caption=f"預覽模式 (第一位對象：{row[id_col]})")

# --- 5. 批量生成功能 ---
st.divider()
if st.button("🚀 開始批量生成並打包下載", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先選擇要生成的名單")
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, (i, row) in enumerate(target_df.iterrows()):
                status_text.text(f"正在製作: {row[id_col]} ({idx+1}/{len(target_df)})")
                
                # 繪製單張證書
                cert_img = bg_img.copy()
                d = ImageDraw.Draw(cert_img)
                
                for col in display_cols:
                    s = st.session_state.settings[col]
                    f = load_font(s["size"])
                    t = str(row[col])
                    
                    try:
                        l, tp, r, b = d.textbbox((0, 0), t, font=f)
                        tw = r - l
                    except: tw = len(t) * s["size"] * 0.7
                    
                    fx = s["x"]
                    if s["align"] == "居中": fx -= tw // 2
                    elif s["align"] == "右對齊": fx -= tw
                    
                    if s["bold"]:
                        for dx, dy in [(-1,-1), (1,1)]:
                            d.text((fx+dx, s["y"]+dy), t, font=f, fill=s["color"])
                    d.text((fx, s["y"]), t, font=f, fill=s["color"])
                
                # 存入 ZIP
                img_io = io.BytesIO()
                cert_img.save(img_io, format="PNG", optimize=True)
                zf.writestr(f"{str(row[id_col]).replace('/', '_')}.png", img_io.getvalue())
                
                progress_bar.progress((idx + 1) / len(target_df))
            
            status_text.text("✅ 全部製作完成！")
        
        st.download_button(
            "📥 點此下載 ZIP 壓縮檔",
            zip_buffer.getvalue(),
            file_name="certificates_pack.zip",
            mime="application/zip",
            use_container_width=True
        )
        st.balloons()
