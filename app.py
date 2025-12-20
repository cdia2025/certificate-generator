import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile
import requests
import math

# ==========================================
# 1. 系統初始化與頁面設定
# ==========================================
st.set_page_config(page_title="專業證書生成器 V6.2 最終優化版", layout="wide")

DPI = 300
PX_PER_CM = DPI / 2.54 
A4_W_PX = int(21.0 * PX_PER_CM)
A4_H_PX = int(29.7 * PX_PER_CM)

# 初始化 Session State
if "settings" not in st.session_state: st.session_state.settings = {}
if "linked_layers" not in st.session_state: st.session_state.linked_layers = []
if "master_selection" not in st.session_state: st.session_state.master_selection = []

def reset_project():
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

def sync_coord(col, axis, trigger):
    nk, sk = f"num_{axis}_{col}", f"sl_{axis}_{col}"
    if trigger == 'num': st.session_state[sk] = st.session_state[nk]
    else: st.session_state[nk] = st.session_state[sk]
    st.session_state.settings[col][axis] = st.session_state[nk]

# ==========================================
# 2. 字體處理與繪製 (支援斜體模擬)
# ==========================================
@st.cache_resource
def get_font_resource():
    font_paths = ["C:/Windows/Fonts/msjh.ttc", "C:/Windows/Fonts/dfkai-sb.ttf", "/System/Library/Fonts/STHeiti Light.ttc", "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"]
    for p in font_paths:
        if os.path.exists(p): return p
    tp = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(tp):
        try:
            r = requests.get("https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf", timeout=20)
            with open(tp, "wb") as f: f.write(r.content)
            return tp
        except: return None
    return tp

@st.cache_data
def get_font_obj(size):
    p = get_font_resource()
    return ImageFont.truetype(p, size) if p else ImageFont.load_default()

def draw_styled_text(draw, text, pos, font, color, align="居中", bold=False, italic=False):
    try:
        l, t, r, b = draw.textbbox((0, 0), text, font=font)
        tw, th = r - l, b - t
    except: tw, th = len(text) * font.size * 0.7, font.size
    x, y = pos
    if align == "居中": x -= tw // 2
    elif align == "右對齊": x -= tw
    if italic:
        p = 60
        txt_img = Image.new("RGBA", (int(tw * 1.5) + p, int(th * 2) + p), (255, 255, 255, 0))
        d_txt = ImageDraw.Draw(txt_img)
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]: d_txt.text((p//2+dx, p//2+dy), text, font=font, fill=color)
        d_txt.text((p//2, p//2), text, font=font, fill=color)
        txt_img = txt_img.transform(txt_img.size, Image.AFFINE, (1, 0.3, -p//2*0.3, 0, 1, 0))
        return (txt_img, (int(x - p//2), int(y - p//2)))
    else:
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]: draw.text((x + dx, y + dy), text, font=font, fill=color)
        draw.text((x, y), text, font=font, fill=color)
        return None

# ==========================================
# 3. 檔案上傳
# ==========================================
st.title("✉️ 專業證書生成器 V6.2")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 歡迎！上傳檔案後，請使用左側側邊欄進行詳細調整。側邊欄寬度可手動拖動。")
    st.stop()

bg_img = Image.open(bg_file).convert("RGBA")
W, H = float(bg_img.size[0]), float(bg_img.size[1])
mid_x, mid_y = W / 2, H / 2
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄控制面板
# ==========================================
with st.sidebar:
    if st.button("🆕 新專案 / 重新重置", use_container_width=True): reset_project()
    st.header("⚙️ 屬性面板")
    
    with st.expander("💾 配置管理"):
        if st.session_state.settings:
            js = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
            st.download_button("📤 匯出設定 (JSON)", js, "config.json", "application/json")
        uploaded_config = st.file_uploader("📥 載入舊設定", type=["json"])
        if uploaded_config:
            st.session_state.settings.update(json.load(uploaded_config))
            for k, v in st.session_state.settings.items():
                st.session_state[f"num_x_{k}"] = st.session_state[f"sl_x_{k}"] = float(v["x"])
                st.session_state[f"num_y_{k}"] = st.session_state[f"sl_y_{k}"] = float(v["y"])
            st.success("配置已載入")

    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": mid_x, "y": mid_y, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        for ax in ['x', 'y']:
            k = f"num_{ax}_{col}"
            if k not in st.session_state: st.session_state[k] = st.session_state[f"sl_{ax}_{col}"] = float(st.session_state.settings[col][ax])

    st.divider()
    st.subheader("📝 單獨圖層設定")
    for col in display_cols:
        with st.expander(f"圖層：{col}"):
            s = st.session_state.settings[col]
            # 座標控制
            st.write(f"**X 座標** (中心參考: {mid_x:.0f})")
            c1, c2 = st.columns([1, 2])
            with c1: st.number_input("X", 0.0, W, key=f"num_x_{col}", on_change=sync_coord, args=(col, 'x', 'num'), label_visibility="collapsed")
            with c2: st.slider("X Slider", 0.0, W, key=f"sl_x_{col}", on_change=sync_coord, args=(col, 'x', 'sl'), label_visibility="collapsed")
            st.write(f"**Y 座標** (中心參考: {mid_y:.0f})")
            c1, c2 = st.columns([1, 2])
            with c1: st.number_input("Y", 0.0, H, key=f"num_y_{col}", on_change=sync_coord, args=(col, 'y', 'num'), label_visibility="collapsed")
            with c2: st.slider("Y Slider", 0.0, H, key=f"sl_y_{col}", on_change=sync_coord, args=(col, 'y', 'sl'), label_visibility="collapsed")
            
            f1, f2 = st.columns(2)
            with f1: s["size"] = st.number_input("大小", 10, 2000, int(s["size"]), key=f"sz_{col}")
            with f2: s["color"] = st.color_picker("顏色", s["color"], key=f"cp_{col}")
            sc1, sc2 = st.columns(2)
            with sc1: s["bold"] = st.checkbox("粗體", s["bold"], key=f"bd_{col}")
            with sc2: s["italic"] = st.checkbox("斜體", s["italic"], key=f"it_{col}")
            s["align"] = st.selectbox("對齊", ["左對齊", "居中", "右對齊"], index=["左對齊", "居中", "右對齊"].index(s["align"]), key=f"al_{col}")

    st.divider()
    with st.expander("🔗 批量位移工具", expanded=False):
        st.session_state.linked_layers = st.multiselect("批量連結對象", display_cols)
        lc1, lc2 = st.columns(2)
        with lc1: b_x = st.number_input("批量 X 位移", value=0.0)
        with lc2: b_y = st.number_input("批量 Y 位移", value=0.0)
        b_s = st.number_input("批量縮放", value=0)
        if st.button("✅ 執行批量套用"):
            for c in st.session_state.linked_layers:
                nx, ny = max(0.0, min(W, st.session_state.settings[c]["x"] + b_x)), max(0.0, min(H, st.session_state.settings[c]["y"] + b_y))
                ns = max(10, st.session_state.settings[c]["size"] + b_s)
                st.session_state.settings[c].update({"x": nx, "y": ny, "size": ns})
                st.session_state[f"num_x_{c}"] = st.session_state[f"sl_x_{c}"] = nx
                st.session_state[f"num_y_{c}"] = st.session_state[f"sl_y_{c}"] = ny
            st.rerun()

# ==========================================
# 5. 主頁面：預覽與搜尋 (V6.2 邏輯)
# ==========================================
st.divider()
st.header("🔍 名單搜尋與選取")
id_col = st.selectbox("選擇主識別欄位 (用於檔案命名基準)", df.columns, key="id_sel")

# 搜尋邏輯
search_term = st.text_input("輸入關鍵字 (會比對 Excel 的所有欄位)", placeholder="例如：姓名、學號、部門...").strip().lower()

# 過濾出符合搜尋條件的 DataFrame
if search_term:
    mask = df.astype(str).apply(lambda x: x.str.lower().str.contains(search_term)).any(axis=1)
    filtered_df = df[mask]
    filtered_ids = filtered_df[id_col].astype(str).tolist()
else:
    filtered_ids = []

# 功能按鈕
col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
with col_btn1:
    if st.button("➕ 將搜尋結果全部加入"):
        # 合併清單並去重
        st.session_state.master_selection = list(set(st.session_state.master_selection + filtered_ids))
with col_btn2:
    if st.button("➖ 從選取中移除搜尋結果"):
        st.session_state.master_selection = [x for x in st.session_state.master_selection if x not in filtered_ids]
with col_btn3:
    if st.button("🗑️ 清空所有已選"):
        st.session_state.master_selection = []

# 最終名單管理 (Multiselect)
st.session_state.master_selection = st.multiselect(
    f"✅ 已選取的製作清單 (目前共有 {len(st.session_state.master_selection)} 筆)",
    options=df[id_col].astype(str).tolist(),
    default=st.session_state.master_sele
