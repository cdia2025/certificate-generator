import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import zipfile
import json
import os
import tempfile
import requests

# ==========================================
# 1. 頁面設定與系統初始化
# ==========================================
st.set_page_config(page_title="專業證書生成器 V5.9.3", layout="wide")

# --- 重置專案功能 ---
def reset_project():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

# 初始化 Session State
if "settings" not in st.session_state:
    st.session_state.settings = {}
if "linked_layers" not in st.session_state:
    st.session_state.linked_layers = []

# --- 同步函數：座標連動邏輯 ---
def sync_coord(col, axis, trigger):
    num_key = f"num_{axis}_{col}"
    sl_key = f"sl_{axis}_{col}"
    if trigger == 'num':
        st.session_state[sl_key] = st.session_state[num_key]
    else:
        st.session_state[num_key] = st.session_state[sl_key]
    st.session_state.settings[col][axis] = st.session_state[num_key]

# ==========================================
# 2. 字體處理與進階繪製
# ==========================================

@st.cache_resource
def get_font_resource():
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc", "C:/Windows/Fonts/dfkai-sb.ttf",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
    ]
    for p in font_paths:
        if os.path.exists(p): return p
    target_path = os.path.join(tempfile.gettempdir(), "NotoSansTC-Regular.otf")
    if not os.path.exists(target_path):
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf"
        try:
            with st.spinner("正在初始化中文字體庫..."):
                r = requests.get(url, timeout=20)
                with open(target_path, "wb") as f: f.write(r.content)
            return target_path
        except: return None
    return target_path

@st.cache_data
def get_font_object(size):
    path = get_font_resource()
    try:
        if path: return ImageFont.truetype(path, size)
    except: pass
    return ImageFont.load_default()

def draw_styled_text(draw, text, pos, font, color, align="居中", bold=False, italic=False):
    try:
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        tw, th = right - left, bottom - top
    except:
        tw, th = len(text) * font.size * 0.7, font.size

    x, y = pos
    if align == "居中": x -= tw // 2
    elif align == "右對齊": x -= tw

    if italic:
        padding = 60
        txt_img = Image.new("RGBA", (int(tw * 1.5) + padding, int(th * 2) + padding), (255, 255, 255, 0))
        d_txt = ImageDraw.Draw(txt_img)
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                d_txt.text((padding//2+dx, padding//2+dy), text, font=font, fill=color)
        d_txt.text((padding//2, padding//2), text, font=font, fill=color)
        m = 0.3 
        txt_img = txt_img.transform(txt_img.size, Image.AFFINE, (1, m, -padding//2*m, 0, 1, 0))
        return (txt_img, (int(x - padding//2), int(y - padding//2)))
    else:
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                draw.text((x + dx, y + dy), text, font=font, fill=color)
        draw.text((x, y), text, font=font, fill=color)
        return None

# ==========================================
# 3. 檔案上傳區
# ==========================================
st.title("✉️ 專業證書生成器 V5.9.3")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 上傳背景圖片", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 上傳資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 歡迎！請先上傳背景圖與資料檔，設定功能將顯示於左側側邊欄。")
    st.stop()

# 載入核心數據
bg_img = Image.open(bg_file).convert("RGBA")
W, H = float(bg_img.size[0]), float(bg_img.size[1])
mid_x, mid_y = W / 2, H / 2
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄控制面板
# ==========================================
with st.sidebar:
    if st.button("🆕 新專案 / 重新重置", use_container_width=True):
        reset_project()
    
    st.header("⚙️ 屬性面板")
    
    with st.expander("💾 配置管理"):
        if st.session_state.settings:
            js = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
            st.download_button("📤 匯出設定 (JSON)", js, "config.json", "application/json")
        uploaded_config = st.file_uploader("📥 載入舊設定", type=["json"])
        if uploaded_config:
            st.session_state.settings.update(json.load(uploaded_config))
            for k, v in st.session_state.settings.items():
                st.session_state[f"num_x_{k}"] = float(v["x"])
                st.session_state[f"sl_x_{k}"] = float(v["x"])
                st.session_state[f"num_y_{k}"] = float(v["y"])
                st.session_state[f"sl_y_{k}"] = float(v["y"])
            st.success("配置已載入")

    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    
    # 補全參數
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": mid_x, "y": mid_y, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        if f"num_x_{col}" not in st.session_state: st.session_state[f"num_x_{col}"] = float(st.session_state.settings[col]["x"])
        if f"sl_x_{col}" not in st.session_state: st.session_state[f"sl_x_{col}"] = float(st.session_state.settings[col]["x"])
        if f"num_y_{col}" not in st.session_state: st.session_state[f"num_y_{col}"] = float(st.session_state.settings[col]["y"])
        if f"sl_y_{col}" not in st.session_state: st.session_state[f"sl_y_{col}"] = float(st.session_state.settings[col]["y"])

    st.divider()

    # --- 單獨圖層調整 (現在位於上方) ---
    st.subheader("📝 圖層屬性設定")
    for col in display_cols:
        tag = " (🔗)" if col in st.session_state.linked_layers else ""
        with st.expander(f"圖層：{col}{tag}"):
            s = st.session_state.settings[col]
            st.caption(f"📍 中心位置參考：X={mid_x:.1f}, Y={mid_y:.1f}")
            
            # X 控制
            st.write("**X 座標控制**")
            cx1, cx2 = st.columns([1, 2])
            with cx1: st.number_input("數值", 0.0, W, key=f"num_x_{col}", on_change=sync_coord, args=(col, 'x', 'num'), label_visibility="collapsed")
            with cx2: st.slider("滑桿", 0.0, W, key=f"sl_x_{col}", on_change=sync_coord, args=(col, 'x', 'sl'), label_visibility="collapsed")
            
            # Y 控制
            st.write("**Y 座標控制**")
            cy1, cy2 = st.columns([1, 2])
            with cy1: st.number_input("數值", 0.0, H, key=f"num_y_{col}", on_change=sync_coord, args=(col, 'y', 'num'), label_visibility="collapsed")
            with cy2: st.slider("滑桿", 0.0, H, key=f"sl_y_{col}", on_change=sync_coord, args=(col, 'y', 'sl'), label_visibility="collapsed")
            
            st.divider()
            f1, f2 = st.columns([1, 1])
            with f1: s["size"] = st.number_input("大小", 10, 1000, int(s["size"]), key=f"sz_{col}")
            with f2: s["color"] = st.color_picker("顏色", s["color"], key=f"cp_{col}")
            sc1, sc2 = st.columns(2)
            with sc1: s["bold"] = st.checkbox("粗體", s["bold"], key=f"bd_{col}")
            with sc2: s["italic"] = st.checkbox("斜體", s["italic"], key=f"it_{col}")
            opts = ["左對齊", "居中", "右對齊"]
            s["align"] = st.selectbox(f"對齊", opts, index=opts.index(s["align"]), key=f"al_{col}")

    st.divider()

    # --- Photoshop 批量工具 (現在移至下方，且預設閉合) ---
    with st.expander("🔗 批量連結與位移工具", expanded=False):
        st.info(f"📍 中心點參考：X={mid_x:.1f}, Y={mid_y:.1f}")
        st.session_state.linked_layers = st.multiselect("選取要同時移動的對象", display_cols)
        lc1, lc2 = st.columns(2)
        with lc1: b_x = st.number_input("批量 X 位移", value=0.0, key="batch_x")
        with lc2: b_y = st.number_input("批量 Y 位移", value=0.0, key="batch_y")
        b_s = st.number_input("批量縮放大小", value=0, key="batch_s")
        
        if st.button("✅ 執行批量套用", use_container_width=True):
            for c in st.session_state.linked_layers:
                nx = max(0.0, min(W, float(st.session_state.settings[c]["x"] + b_x)))
                ny = max(0.0, min(H, float(st.session_state.settings[c]["y"] + b_y)))
                ns = max(10, min(1000, int(st.session_state.settings[c]["size"] + b_s)))
                st.session_state.settings[c].update({"x": nx, "y": ny, "size": ns})
                st.session_state[f"num_x_{c}"] = nx
                st.session_state[f"sl_x_{c}"] = nx
                st.session_state[f"num_y_{c}"] = ny
                st.session_state[f"sl_y_{c}"] = ny
            st.rerun()

# ==========================================
# 5. 主頁面：預覽與畫布
# ==========================================
st.divider()
p1, p2 = st.columns([1, 1])
with p1: id_col = st.selectbox("命名依據欄位", df.columns, key="id_sel")
with p2:
    all_n = df[id_col].astype(str).tolist()
    sel_n = st.multiselect("預覽名單", all_n, default=all_n[:1], key="pre_sel")
    target_df = df[df[id_col].astype(str).isin(sel_n)]

st.subheader("👁️ 即時畫布預覽")
zoom = st.slider("🔍 視覺縮放 (%)", 50, 250, 100, step=10, key="zoom_sl")

if not target_df.empty:
    row = target_df.iloc[0]
    canvas = bg_img.copy()
    draw = ImageDraw.Draw(canvas)
    for col in display_cols:
        # 抓取連動的最新座標
        cur_x = st.session_state[f"num_x_{col}"]
        cur_y = st.session_state[f"num_y_{col}"]
        sv = st.session_state.settings[col]
        f_obj = get_font_object(sv["size"])
        res = draw_styled_text(draw, str(row[col]), (cur_x, cur_y), f_obj, sv["color"], sv["align"], sv["bold"], sv["italic"])
        if res: canvas.alpha_composite(res[0], dest=res[1])
        g_c = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw.line([(0, cur_y), (W, cur_y)], fill=g_c, width=2)
        draw.line([(cur_x, 0), (cur_x, H)], fill=g_c, width=2)
    st.image(canvas, width=int(W * (zoom / 100)))

# ==========================================
# 6. 生成與打包
# ==========================================
st.divider()
if st.button("🚀 開始批量製作所有選定證書", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先選取預覽對象")
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            prog = st.progress(0)
            for idx, (i, row) in enumerate(target_df.iterrows()):
                out = bg_img.copy()
                d_f = ImageDraw.Draw(out)
                for col in display_cols:
                    sv = st.session_state.settings[col]
                    cx, cy = st.session_state[f"num_x_{col}"], st.session_state[f"num_y_{col}"]
                    res = draw_styled_text(d_f, str(row[col]), (cx, cy), get_font_object(sv["size"]), sv["color"], sv["align"], sv["bold"], sv["italic"])
                    if res: out.alpha_composite(res[0], dest=res[1])
                fb = io.BytesIO()
                out.convert("RGB").save(fb, format="JPEG", quality=95)
                zf.writestr(f"{str(row[id_col])}.jpg", fb.getvalue())
                prog.progress((idx + 1) / len(target_df))
        st.download_button("📥 下載打包檔 (ZIP)", zip_buf.getvalue(), "certs.zip", "application/zip", use_container_width=True)
