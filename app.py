import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile
import requests

# 必須放在最首行
st.set_page_config(page_title="專業證書生成器 V5.3", layout="wide")

# --- 1. 字體處理與樣式繪製 ---

@st.cache_resource
def get_font_resource():
    """確保系統一定有中文字體，若無則下載思源黑體"""
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc",            # Windows
        "C:/Windows/Fonts/dfkai-sb.ttf",        # Windows 標楷
        "/System/Library/Fonts/STHeiti Light.ttc", # macOS
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", # Linux
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
    ]
    for p in font_paths:
        if os.path.exists(p): return p

    target_path = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(target_path):
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
        try:
            with st.spinner("正在初始化中文字體..."):
                response = requests.get(url, timeout=20)
                with open(target_path, "wb") as f:
                    f.write(response.content)
            return target_path
        except:
            return None
    return target_path

@st.cache_data
def get_font_object(size):
    """緩存字體對象以提升渲染效能"""
    path = get_font_resource()
    try:
        if path: return ImageFont.truetype(path, size)
    except: pass
    return ImageFont.load_default()

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
        # 斜體需透過矩陣變幻，空間需加大防止邊緣切斷
        padding = 40
        txt_img = Image.new("RGBA", (int(tw * 1.5) + padding, int(th * 2) + padding), (255, 255, 255, 0))
        d_txt = ImageDraw.Draw(txt_img)
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                d_txt.text((padding//2+dx, padding//2+dy), text, font=font, fill=color)
        d_txt.text((padding//2, padding//2), text, font=font, fill=color)
        
        m = 0.3 # 傾斜率
        txt_img = txt_img.transform(txt_img.size, Image.AFFINE, (1, m, -padding//2*m, 0, 1, 0))
        return (txt_img, (int(x - padding//2), int(y - padding//2)))
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

# --- 3. 側邊欄與檔案處理 ---
with st.sidebar:
    st.header("💾 設定存檔")
    if st.session_state.settings:
        config_json = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
        st.download_button("📤 匯出設定 (JSON)", config_json, "cert_config.json", "application/json")
    
    uploaded_config = st.file_uploader("📥 載入舊設定檔", type=["json"], key="config_loader")
    if uploaded_config:
        try:
            st.session_state.settings.update(json.load(uploaded_config))
            st.success("配置已載入")
        except:
            st.error("載入失敗")

st.title("✉️ 專業證書生成器 V5.3")
up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖片", type=["jpg", "png", "jpeg"], key="bg_uploader")
with up2: data_file = st.file_uploader("📊 2. 資料檔 (Excel/CSV)", type=["xlsx", "csv"], key="data_uploader")

if not bg_file or not data_file:
    st.info("👋 請上傳背景圖片和資料檔以開始。")
    st.stop()

# 讀取檔案
bg_img = Image.open(bg_file).convert("RGBA")
W, H = bg_img.size
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# --- 4. 欄位補全 ---
display_cols = st.multiselect("要在證書上顯示的欄位", df.columns, default=[df.columns[0]], key="display_cols_select")

for col in display_cols:
    if col not in st.session_state.settings:
        st.session_state.settings[col] = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
    else:
        # 自動相容新版本參數
        defaults = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        for k, v in defaults.items():
            if k not in st.session_state.settings[col]:
                st.session_state.settings[col][k] = v

st.divider()

# --- 5. Photoshop 批量工具 ---
st.header("🔗 圖層連結工具 (批量移動)")
col_link1, col_link2 = st.columns([1, 2])
with col_link1:
    st.session_state.linked_layers = st.multiselect("選取要『連結』的欄位", display_cols, key="linked_selector")
with col_link2:
    lc1, lc2, lc3 = st.columns(3)
    with lc1: m_x = st.number_input("左右位移 (px)", value=0, key="batch_x")
    with lc2: m_y = st.number_input("上下位移 (px)", value=0, key="batch_y")
    with lc3: m_s = st.number_input("縮放字體", value=0, key="batch_s")
    if st.button("✅ 執行批量套用", use_container_width=True, key="batch_apply_btn"):
        for c in st.session_state.linked_layers:
            st.session_state.settings[c]["x"] += m_x
            st.session_state.settings[c]["y"] += m_y
            st.session_state.settings[c]["size"] += m_s
        st.rerun()

st.divider()

# --- 6. 工作區佈局 ---
col_ctrl, col_prev = st.columns([1, 1], gap="large")

with col_ctrl:
    st.header("🛠️ 參數調整")
    id_col = st.selectbox("識別欄位 (檔名)", df.columns, key="id_col_select")
    all_names = df[id_col].astype(str).tolist()
    selected_items = st.multiselect("選取預覽對象", all_names, default=all_names[:1], key="preview_items_select")
    target_df = df[df[id_col].astype(str).isin(selected_items)]

    for col in display_cols:
        link_tag = " (🔗 已連結)" if col in st.session_state.linked_layers else ""
        with st.expander(f"📝 欄位：{col}{link_tag}"):
            s = st.session_state.settings[col]
            cc1, cc2 = st.columns(2)
            with cc1: s["x"] = st.slider(f"X 位置", 0, W, int(s["x"]), key=f"x_s_{col}")
            with cc2: s["y"] = st.slider(f"Y 位置", 0, H, int(s["y"]), key=f"y_s_{col}")
            
            cc3, cc4 = st.columns(2)
            with cc3: s["size"] = st.number_input(f"字體大小", 10, 1000, int(s["size"]), key=f"sz_n_{col}")
            with cc4: s["color"] = st.color_picker(f"顏色", s["color"], key=f"cl_p_{col}")
            
            cc5, cc6 = st.columns(2)
            with cc5: s["bold"] = st.checkbox("粗體 (Bold)", s.get("bold", False), key=f"bd_c_{col}")
            with cc6: s["italic"] = st.checkbox("斜體 (Italic)", s.get("italic", False), key=f"it_c_{col}")
            
            opts = ["左對齊", "居中", "右對齊"]
            s["align"] = st.selectbox(f"對齊", opts, index=opts.index(s.get("align", "居中")), key=f"al_s_{col}")

with col_prev:
    st.header("👁️ 即時預覽")
    # 修正：縮放範圍改為 50 - 250，並加上 key 以穩定狀態
    zoom = st.slider("🔍 預覽圖視覺縮放 (%)", 50, 250, 100, step=10, key="zoom_slider")
    
    if not target_df.empty:
        row = target_df.iloc[0]
        preview_canvas = bg_img.copy()
        draw = ImageDraw.Draw(preview_canvas)
        
        for col in display_cols:
            s = st.session_state.settings[col]
            f_obj = get_font_object(s["size"])
            t_val = str(row[col])
            
            res = draw_styled_text(draw, t_val, (s["x"], s["y"]), f_obj, s["color"], s["align"], s["bold"], s["italic"])
            if res:
                l_img, l_pos = res
                preview_canvas.alpha_composite(l_img, dest=l_pos)
            
            l_color = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
            draw.line([(0, s["y"]), (W, s["y"])], fill=l_color, width=2)
            draw.line([(s["x"], 0), (s["x"], H)], fill=l_color, width=2)

        # 穩定顯示預覽圖，使用固定寬度計算
        display_width = int(W * (zoom / 100))
        st.image(preview_canvas, width=display_width)

# --- 7. 生成功能 ---
st.divider()
if st.button("🚀 開始大量製作所有選定證書", type="primary", use_container_width=True, key="generate_btn"):
    if target_df.empty:
        st.warning("請先選擇名單")
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            prog = st.progress(0)
            for idx, (i, row) in enumerate(target_df.iterrows()):
                out_img = bg_img.copy()
                d = ImageDraw.Draw(out_img)
                for col in display_cols:
                    s = st.session_state.settings[col]
                    f = get_font_object(s["size"])
                    res = draw_styled_text(d, str(row[col]), (s["x"], s["y"]), f, s["color"], s["align"], s["bold"], s["italic"])
                    if res:
                        l_img, l_pos = res
                        out_img.alpha_composite(l_img, dest=l_pos)
                
                final_buf = io.BytesIO()
                out_img.convert("RGB").save(final_buf, format="JPEG", quality=95)
                zf.writestr(f"{str(row[id_col])}.jpg", final_buf.getvalue())
                prog.progress((idx+1)/len(target_df))
        
        st.download_button("📥 下載證書打包檔 (ZIP)", zip_buf.getvalue(), "certificates.zip", "application/zip", use_container_width=True, key="download_btn")
