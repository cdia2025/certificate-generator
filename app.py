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
st.set_page_config(page_title="專業證書生成器 V5.1", layout="wide")

# --- 1. 強化版中文字體載入與斜體支援 ---
@st.cache_resource
def get_font_resource():
    """確保環境中一定有中文字體可用"""
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc",            # Windows 微軟正黑
        "C:/Windows/Fonts/dfkai-sb.ttf",        # Windows 標楷體
        "/System/Library/Fonts/STHeiti Light.ttc", # macOS 華文黑體
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
    ]
    for p in font_paths:
        if os.path.exists(p): return p

    target_path = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(target_path):
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
        try:
            response = requests.get(url, timeout=15)
            with open(target_path, "wb") as f: f.write(response.content)
            return target_path
        except: return None
    return target_path

def draw_styled_text(draw, text, pos, font, color, align="居中", bold=False, italic=False):
    """處理模擬粗體與模擬斜體的繪製函數"""
    try:
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        tw = right - left
        th = bottom - top
    except:
        tw, th = len(text) * font.size * 0.7, font.size

    x, y = pos
    if align == "居中": x -= tw // 2
    elif align == "右對齊": x -= tw

    if italic:
        # 創建暫存文字層以進行矩陣變換
        txt_img = Image.new("RGBA", (int(tw * 1.5) + 20, int(th * 2) + 20), (255, 255, 255, 0))
        d_txt = ImageDraw.Draw(txt_img)
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                d_txt.text((10+dx, 10+dy), text, font=font, fill=color)
        d_txt.text((10, 10), text, font=font, fill=color)
        
        m = 0.3 # 斜率
        txt_img = txt_img.transform(txt_img.size, Image.AFFINE, (1, m, -10*m, 0, 1, 0))
        return (txt_img, (int(x - 10), int(y - 10)))
    else:
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                draw.text((x + dx, y + dy), text, font=font, fill=color)
        draw.text((x, y), text, font=font, fill=color)
        return None

# --- 2. 初始化 Session State ---
if "settings" not in st.session_state:
    st.session_state.settings = {}
if "linked_layers" not in st.session_state:
    st.session_state.linked_layers = []

# --- 3. 側邊欄與檔案上傳 ---
with st.sidebar:
    st.header("💾 設定存檔")
    if st.session_state.settings:
        config_json = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
        st.download_button("📤 匯出目前設定 (JSON)", config_json, "cert_config.json", "application/json")
    
    uploaded_config = st.file_uploader("📥 載入舊設定檔", type=["json"])
    if uploaded_config:
        try:
            loaded_data = json.load(uploaded_config)
            st.session_state.settings.update(loaded_data)
            st.success("配置已載入")
        except Exception as e:
            st.error(f"載入失敗: {e}")

st.title("✉️ 專業證書生成器 V5.1")
up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 上傳證書背景圖", type=["jpg", "png", "jpeg"])
with up2: data_file = st.file_uploader("📊 2. 上傳資料檔 (Excel/CSV)", type=["xlsx", "csv"])

if not bg_file or not data_file:
    st.info("👋 請上傳背景圖片和資料檔以開始。")
    st.stop()

# 讀取檔案
bg_img = Image.open(bg_file).convert("RGBA")
W, H = bg_img.size
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# --- 4. 參數自動補全與初始化 ---
display_cols = st.multiselect("選擇要在證書上顯示的欄位", df.columns, default=[df.columns[0]])

for col in display_cols:
    if col not in st.session_state.settings:
        st.session_state.settings[col] = {
            "x": W//2, "y": H//2, "size": 60, "color": "#000000", 
            "align": "居中", "bold": False, "italic": False
        }
    else:
        # 防錯機制：如果載入的是舊 JSON，手動補上缺失的 key
        defaults = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        for key, val in defaults.items():
            if key not in st.session_state.settings[col]:
                st.session_state.settings[col][key] = val

st.divider()

# --- 5. Photoshop 批量控制區 ---
st.header("🔗 Photoshop 圖層工具 (批量修改)")
col_link1, col_link2 = st.columns([1, 2])
with col_link1:
    st.session_state.linked_layers = st.multiselect("選取要『連結』的欄位", display_cols)
with col_link2:
    lc1, lc2, lc3 = st.columns(3)
    with lc1: move_x = st.number_input("左右位移 (px)", value=0)
    with lc2: move_y = st.number_input("上下位移 (px)", value=0)
    with lc3: change_size = st.number_input("縮放大小", value=0)
    if st.button("✅ 執行批量套用", use_container_width=True):
        for col in st.session_state.linked_layers:
            st.session_state.settings[col]["x"] += move_x
            st.session_state.settings[col]["y"] += move_y
            st.session_state.settings[col]["size"] += change_size
        st.rerun()

st.divider()

# --- 6. 工作區佈局 ---
col_ctrl, col_prev = st.columns([1, 1], gap="large")

with col_ctrl:
    st.header("🛠️ 單獨調整")
    id_col = st.selectbox("識別欄位 (用於檔案命名)", df.columns)
    all_names = df[id_col].astype(str).tolist()
    selected_items = st.multiselect("選取預覽/生成對象", all_names, default=all_names[:1])
    target_df = df[df[id_col].astype(str).isin(selected_items)]

    for col in display_cols:
        with st.expander(f"📝 欄位：{col} {' (🔗 已連結)' if col in st.session_state.linked_layers else ''}"):
            s = st.session_state.settings[col]
            cc1, cc2 = st.columns(2)
            with cc1: s["x"] = st.slider(f"X 位置", 0, W, int(s.get("x", W//2)), key=f"x_{col}")
            with cc2: s["y"] = st.slider(f"Y 位置", 0, H, int(s.get("y", H//2)), key=f"y_{col}")
            
            cc3, cc4 = st.columns(2)
            with cc3: s["size"] = st.number_input(f"字體大小", 10, 1000, int(s.get("size", 60)), key=f"sz_{col}")
            with cc4: s["color"] = st.color_picker(f"顏色", s.get("color", "#000000"), key=f"cl_{col}")
            
            cc5, cc6 = st.columns(2)
            with cc5: s["bold"] = st.checkbox("粗體 (Bold)", s.get("bold", False), key=f"bd_{col}")
            # 使用 .get() 確保安全讀取
            with cc6: s["italic"] = st.checkbox("斜體 (Italic)", s.get("italic", False), key=f"it_{col}")
            
            align_options = ["左對齊", "居中", "右對齊"]
            current_align = s.get("align", "居中")
            s["align"] = st.selectbox(f"對齊", align_options, index=align_options.index(current_align) if current_align in align_options else 1, key=f"al_{col}")

with col_prev:
    st.header("👁️ 即時預覽")
    zoom = st.slider("🔍 預覽圖縮放", 10, 100, 50)
    
    if not target_df.empty:
        row = target_df.iloc[0]
        preview_canvas = bg_img.copy()
        draw = ImageDraw.Draw(preview_canvas)
        
        for col in display_cols:
            s = st.session_state.settings[col]
            font = load_font(s["size"])
            text_val = str(row[col])
            
            res = draw_styled_text(
                draw, text_val, (s["x"], s["y"]), font, s["color"], 
                s["align"], s["bold"], s.get("italic", False)
            )
            if res:
                img_l, pos_l = res
                preview_canvas.alpha_composite(img_l, dest=pos_l)
            
            color_line = "#FF000088" if col in st.session_state.linked_layers else "#0000FF33"
            draw.line([(0, s["y"]), (W, s["y"])], fill=color_line, width=2)
            draw.line([(s["x"], 0), (s["x"], H)], fill=color_line, width=2)

        st.image(preview_canvas, width=int(W * (zoom/100)))

# --- 7. 批量生成 ---
st.divider()
if st.button("🚀 生成所有選定證書", type="primary", use_container_width=True):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        prog = st.progress(0)
        for idx, (i, row) in enumerate(target_df.iterrows()):
            cert_img = bg_img.copy()
            d = ImageDraw.Draw(cert_img)
            for col in display_cols:
                s = st.session_state.settings[col]
                f = load_font(s["size"])
                res = draw_styled_text(d, str(row[col]), (s["x"], s["y"]), f, s["color"], s["align"], s.get("bold", False), s.get("italic", False))
                if res:
                    img_l, pos_l = res
                    cert_img.alpha_composite(img_l, dest=pos_l)
            
            img_io = io.BytesIO()
            cert_img.convert("RGB").save(img_io, format="JPEG", quality=95)
            zf.writestr(f"{str(row[id_col])}.jpg", img_io.getvalue())
            prog.progress((idx+1)/len(target_df))
    st.download_button("📥 下載 ZIP 壓縮檔", zip_buffer.getvalue(), "certs.zip", "application/zip", use_container_width=True)
