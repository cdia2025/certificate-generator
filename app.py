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
st.set_page_config(page_title="專業證書生成器 V5.4", layout="wide")

# --- 1. 字體與繪製邏輯 (支援斜體與緩存) ---

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
            with st.spinner("正在初始化中文字體..."):
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
        m = 0.3 # 斜率
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

# --- 3. 檔案上傳 (置於主頁上方) ---
st.title("✉️ 專業證書生成器 V5.4")
up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖片", type=["jpg", "png", "jpeg"], key="bg_up")
with up2: data_file = st.file_uploader("📊 2. 資料檔", type=["xlsx", "csv"], key="data_up")

if not bg_file or not data_file:
    st.info("💡 請上傳圖片和資料後，使用左側「側邊欄」進行詳細調整。側邊欄邊框可滑鼠拖動調整寬度。")
    st.stop()

bg_img = Image.open(bg_file).convert("RGBA")
W, H = bg_img.size
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# --- 4. 側邊欄控制台 (重點：支援滑鼠調整闊度) ---
with st.sidebar:
    st.header("⚙️ 參數調整面板")
    st.caption("👈 滑動此欄邊框可調整與預覽圖比例")
    
    # 設定管理
    with st.expander("💾 配置存取", expanded=False):
        if st.session_state.settings:
            js = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
            st.download_button("📤 匯出設定 (JSON)", js, "config.json", "application/json")
        uploaded_config = st.file_uploader("📥 載入設定", type=["json"])
        if uploaded_config:
            st.session_state.settings.update(json.load(uploaded_config))
            st.success("配置已更新")

    st.divider()

    # 欄位選取
    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    
    # 補全參數
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        else:
            defaults = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
            for k, v in defaults.items():
                if k not in st.session_state.settings[col]: st.session_state.settings[col][k] = v

    # 批量工具
    with st.expander("🔗 Photoshop 批量連結工具", expanded=False):
        st.session_state.linked_layers = st.multiselect("連結欄位", display_cols)
        lc1, lc2 = st.columns(2)
        with lc1: bx = st.number_input("X 位移", value=0)
        with lc2: by = st.number_input("Y 位移", value=0)
        bs = st.number_input("縮放字體", value=0)
        if st.button("✅ 執行批量套用", use_container_width=True):
            for c in st.session_state.linked_layers:
                st.session_state.settings[c]["x"] += bx
                st.session_state.settings[c]["y"] += by
                st.session_state.settings[c]["size"] += bs
            st.rerun()

    st.divider()

    # 個別欄位調整
    st.subheader("📝 單獨圖層屬性")
    for col in display_cols:
        link_tag = " (🔗)" if col in st.session_state.linked_layers else ""
        with st.expander(f"圖層：{col}{link_tag}"):
            s = st.session_state.settings[col]
            s["x"] = st.slider(f"X 位置", 0, W, int(s["x"]), key=f"x_{col}")
            s["y"] = st.slider(f"Y 位置", 0, H, int(s["y"]), key=f"y_{col}")
            s["size"] = st.number_input(f"大小", 10, 1000, int(s["size"]), key=f"s_{col}")
            s["color"] = st.color_picker(f"顏色", s["color"], key=f"c_{col}")
            
            c1, c2 = st.columns(2)
            with c1: s["bold"] = st.checkbox("粗體", s["bold"], key=f"b_{col}")
            with c2: s["italic"] = st.checkbox("斜體", s["italic"], key=f"i_{col}")
            
            opts = ["左對齊", "居中", "右對齊"]
            s["align"] = st.selectbox(f"對齊", opts, index=opts.index(s["align"]), key=f"a_{col}")

# --- 5. 主頁面預覽區 ---
st.divider()
prev_col1, prev_col2 = st.columns([1, 1]) # 僅用於顯示名單選取
with prev_col1:
    id_col = st.selectbox("識別欄位 (用於命名)", df.columns)
with prev_col2:
    all_n = df[id_col].astype(str).tolist()
    sel_n = st.multiselect("選取對象", all_n, default=all_n[:1])
    target_df = df[df[id_col].astype(str).isin(sel_n)]

st.subheader("👁️ 即時畫布預覽")
zoom = st.slider("🔍 畫布視覺縮放 (%)", 50, 250, 100, step=10)

if not target_df.empty:
    row = target_df.iloc[0]
    canvas = bg_img.copy()
    draw = ImageDraw.Draw(canvas)
    
    for col in display_cols:
        s = st.session_state.settings[col]
        f = get_font_object(s["size"])
        res = draw_styled_text(draw, str(row[col]), (s["x"], s["y"]), f, s["color"], s["align"], s["bold"], s["italic"])
        if res:
            l_img, l_pos = res
            canvas.alpha_composite(l_img, dest=l_pos)
        
        # 輔助線
        l_c = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw.line([(0, s["y"]), (W, s["y"])], fill=l_c, width=2)
        draw.line([(s["x"], 0), (s["x"], H)], fill=l_c, width=2)

    st.image(canvas, width=int(W * (zoom / 100)))

# --- 6. 生成 ---
st.divider()
if st.button("🚀 開始批量製作所有選定證書", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先選擇對象")
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            prog = st.progress(0)
            for idx, (i, row) in enumerate(target_df.iterrows()):
                out = bg_img.copy()
                d = ImageDraw.Draw(out)
                for col in display_cols:
                    s = st.session_state.settings[col]
                    res = draw_styled_text(d, str(row[col]), (s["x"], s["y"]), get_font_object(s["size"]), s["color"], s["align"], s["bold"], s["italic"])
                    if res: out.alpha_composite(res[0], dest=res[1])
                
                fb = io.BytesIO()
                out.convert("RGB").save(fb, format="JPEG", quality=95)
                zf.writestr(f"{str(row[id_col])}.jpg", fb.getvalue())
                prog.progress((idx+1)/len(target_df))
        st.download_button("📥 下載打包檔 (ZIP)", zip_buf.getvalue(), "certs.zip", "application/zip", use_container_width=True)
