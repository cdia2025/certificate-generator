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
# 1. 系統初始化與頁面設定
# ==========================================
st.set_page_config(page_title="專業證書生成器 V5.6 安全修復版", layout="wide")

if "settings" not in st.session_state:
    st.session_state.settings = {}
if "linked_layers" not in st.session_state:
    st.session_state.linked_layers = []

# ==========================================
# 2. 字體處理與繪製邏輯
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
st.title("✉️ 專業證書生成器 V5.6")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 上傳證書背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 上傳資料檔 (Excel/CSV)", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("💡 提示：上傳檔案後，使用左側「側邊欄」調整。側邊欄邊界可滑鼠拖拽寬度。")
    st.stop()

bg_img = Image.open(bg_file).convert("RGBA")
W, H = float(bg_img.size[0]), float(bg_img.size[1]) # 轉為 float 確保與 slider 類型一致
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄控制面板
# ==========================================
with st.sidebar:
    st.header("⚙️ 參數調整面板")
    
    with st.expander("💾 配置存取"):
        if st.session_state.settings:
            js = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
            st.download_button("📤 匯出設定 (JSON)", js, "config.json", "application/json")
        uploaded_config = st.file_uploader("📥 載入舊設定", type=["json"])
        if uploaded_config:
            st.session_state.settings.update(json.load(uploaded_config))
            st.success("配置已載入")

    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    
    # 確保每個欄位都有設定且數值在邊界內
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": W/2, "y": H/2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        else:
            # 安全防護：補全缺失參數並強制鉗制座標在當前圖片範圍內
            s_dict = st.session_state.settings[col]
            s_dict["x"] = max(0.0, min(W, float(s_dict.get("x", W/2))))
            s_dict["y"] = max(0.0, min(H, float(s_dict.get("y", H/2))))
            if "size" not in s_dict: s_dict["size"] = 60
            if "color" not in s_dict: s_dict["color"] = "#000000"
            if "align" not in s_dict: s_dict["align"] = "居中"
            if "bold" not in s_dict: s_dict["bold"] = False
            if "italic" not in s_dict: s_dict["italic"] = False

    st.divider()

    # --- Photoshop 批量工具 (加入邊界防護) ---
    with st.expander("🔗 Photoshop 批量連結工具", expanded=True):
        st.session_state.linked_layers = st.multiselect("選取要同時移動的欄位", display_cols)
        lc1, lc2 = st.columns(2)
        with lc1: b_x = st.number_input("左右位移 (px)", value=0.0)
        with lc2: b_y = st.number_input("上下位移 (px)", value=0.0)
        b_s = st.number_input("字體縮放", value=0)
        
        if st.button("✅ 執行批量套用", use_container_width=True):
            for c in st.session_state.linked_layers:
                # 運算並進行範圍鉗制
                nx = max(0.0, min(W, float(st.session_state.settings[c]["x"] + b_x)))
                ny = max(0.0, min(H, float(st.session_state.settings[c]["y"] + b_y)))
                ns = max(10, min(1000, int(st.session_state.settings[c]["size"] + b_s)))
                
                st.session_state.settings[c]["x"] = nx
                st.session_state.settings[c]["y"] = ny
                st.session_state.settings[c]["size"] = ns
                
                # 同步更新內部 Key，確保 Slider 讀取正確
                st.session_state[f"x_{c}"] = nx
                st.session_state[f"y_{c}"] = ny
                st.session_state[f"s_{c}"] = ns
            st.rerun()

    st.divider()

    # --- 單獨圖層設定 (加入邊界防護) ---
    st.subheader("📝 單獨圖層設定")
    for col in display_cols:
        link_tag = " (🔗)" if col in st.session_state.linked_layers else ""
        with st.expander(f"圖層：{col}{link_tag}"):
            s = st.session_state.settings[col]
            
            # 渲染 Slider 前再次確認數值合法，避免崩潰
            cur_x = max(0.0, min(W, float(s["x"])))
            cur_y = max(0.0, min(H, float(s["y"])))
            
            s["x"] = st.slider(f"X 座標", 0.0, W, cur_x, key=f"x_{col}")
            s["y"] = st.slider(f"Y 座標", 0.0, H, cur_y, key=f"y_{col}")
            s["size"] = st.number_input(f"字體大小", 10, 1000, int(s["size"]), key=f"s_{col}")
            s["color"] = st.color_picker(f"文字顏色", s["color"], key=f"c_{col}")
            
            sc1, sc2 = st.columns(2)
            with sc1: s["bold"] = st.checkbox("粗體", s["bold"], key=f"b_{col}")
            with sc2: s["italic"] = st.checkbox("斜體", s["italic"], key=f"i_{col}")
            
            opts = ["左對齊", "居中", "右對齊"]
            s["align"] = st.selectbox(f"對齊方式", opts, index=opts.index(s["align"]), key=f"a_{col}")

# ==========================================
# 5. 主頁面：預覽與畫布
# ==========================================
st.divider()
p1, p2 = st.columns([1, 1])
with p1: id_col = st.selectbox("命名依據欄位", df.columns)
with p2:
    all_n = df[id_col].astype(str).tolist()
    sel_n = st.multiselect("預覽名單", all_n, default=all_n[:1])
    target_df = df[df[id_col].astype(str).isin(sel_n)]

st.subheader("👁️ 畫布即時預覽")
zoom = st.slider("🔍 視覺縮放 (%)", 50, 250, 100, step=10, key="zoom_sl")

if not target_df.empty:
    row_data = target_df.iloc[0]
    canvas = bg_img.copy()
    draw = ImageDraw.Draw(canvas)
    
    for col in display_cols:
        set_val = st.session_state.settings[col]
        font_obj = get_font_object(set_val["size"])
        res = draw_styled_text(draw, str(row_data[col]), (set_val["x"], set_val["y"]), font_obj, set_val["color"], set_val["align"], set_val["bold"], set_val["italic"])
        if res:
            canvas.alpha_composite(res[0], dest=res[1])
        
        guide_c = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw.line([(0, set_val["y"]), (W, set_val["y"])], fill=guide_c, width=2)
        draw.line([(set_val["x"], 0), (set_val["x"], H)], fill=guide_c, width=2)

    st.image(canvas, width=int(W * (zoom / 100)))

# ==========================================
# 6. 生成
# ==========================================
st.divider()
if st.button("🚀 開始批量製作選定證書", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先選取名單")
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            prog = st.progress(0)
            for idx, (i, row) in enumerate(target_df.iterrows()):
                f_img = bg_img.copy()
                d_f = ImageDraw.Draw(f_img)
                for col in display_cols:
                    sv = st.session_state.settings[col]
                    res = draw_styled_text(d_f, str(row[col]), (sv["x"], sv["y"]), get_font_object(sv["size"]), sv["color"], sv["align"], sv["bold"], sv["italic"])
                    if res: f_img.alpha_composite(res[0], dest=res[1])
                img_io = io.BytesIO()
                f_img.convert("RGB").save(img_io, format="JPEG", quality=95)
                zf.writestr(f"{str(row[id_col])}.jpg", img_io.getvalue())
                prog.progress((idx + 1) / len(target_df))
        st.download_button("📥 下載 ZIP 檔", zip_buf.getvalue(), "certs.zip", "application/zip", use_container_width=True)
