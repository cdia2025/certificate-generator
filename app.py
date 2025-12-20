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
st.set_page_config(page_title="專業證書生成器 V6.1 拼板修正版", layout="wide")

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
# 2. 字體與繪製邏輯
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
st.title("✉️ 專業證書生成器 V6.1")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 請上傳檔案。V6.1 修復了透明拼板背景變黑的問題。")
    st.stop()

bg_img = Image.open(bg_file).convert("RGBA")
W, H = float(bg_img.size[0]), float(bg_img.size[1])
mid_x, mid_y = W / 2, H / 2
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄控制
# ==========================================
with st.sidebar:
    if st.button("🆕 新專案 / 重新重置", use_container_width=True): reset_project()
    
    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": mid_x, "y": mid_y, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        for ax in ['x', 'y']:
            k = f"num_{ax}_{col}"
            if k not in st.session_state: st.session_state[k] = st.session_state[f"sl_{ax}_{col}"] = float(st.session_state.settings[col][ax])

    st.subheader("📝 圖層屬性")
    for col in display_cols:
        with st.expander(f"圖層：{col}"):
            s = st.session_state.settings[col]
            c1, c2 = st.columns([1, 2])
            with c1: st.number_input("X", 0.0, W, key=f"num_x_{col}", on_change=sync_coord, args=(col, 'x', 'num'), label_visibility="collapsed")
            with c2: st.slider("X Slider", 0.0, W, key=f"sl_x_{col}", on_change=sync_coord, args=(col, 'x', 'sl'), label_visibility="collapsed")
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

# ==========================================
# 5. 主頁面：預覽與搜尋 (優化版)
# ==========================================
st.divider()
st.header("🔍 名單搜尋與選取")
id_col = st.selectbox("選擇主識別欄位 (檔名基準)", df.columns, key="id_sel")

search_term = st.text_input("輸入關鍵字搜尋 (會搜尋 Excel 內的所有欄位內容)", "").strip().lower()

# 多欄位搜尋邏輯
if search_term:
    mask = df.astype(str).apply(lambda x: x.str.lower().str.contains(search_term)).any(axis=1)
    filtered_df = df[mask]
    filtered_list = filtered_df[id_col].astype(str).tolist()
else:
    filtered_list = []

c_search1, c_search2, c_search3 = st.columns([2, 1, 1])
with c_search1:
    st.caption(f"搜尋結果: 找到 {len(filtered_list)} 筆相關資料")
with c_search2:
    if st.button("➕ 將搜尋結果全部加入選取"):
        # 合併目前的 master_selection 和過濾出來的清單
        combined = list(set(st.session_state.master_selection + filtered_list))
        st.session_state.master_selection = combined
with c_search3:
    if st.button("🗑️ 清空所有已選"):
        st.session_state.master_selection = []

# 已選取清單
st.session_state.master_selection = st.multiselect(
    "✅ 目前已選取的製作名單 (可手動刪除)",
    options=df[id_col].astype(str).tolist(),
    default=st.session_state.master_selection
)

target_df = df[df[id_col].astype(str).isin(st.session_state.master_selection)]

if not target_df.empty:
    st.subheader("👁️ 畫布即時預覽")
    zoom = st.slider("🔍 預覽縮放 (%)", 50, 250, 100, step=10, key="zoom_sl")
    row = target_df.iloc[0]
    canvas = bg_img.copy()
    draw = ImageDraw.Draw(canvas)
    for col in display_cols:
        cx, cy, sv = st.session_state[f"num_x_{col}"], st.session_state[f"num_y_{col}"], st.session_state.settings[col]
        res = draw_styled_text(draw, str(row[col]), (cx, cy), get_font_obj(sv["size"]), sv["color"], sv["align"], sv["bold"], sv["italic"])
        if res: canvas.alpha_composite(res[0], dest=res[1])
        gc = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw.line([(0, cy), (W, cy)], fill=gc, width=2); draw.line([(cx, 0), (cx, H)], fill=gc, width=2)
    st.image(canvas, width=int(W * (zoom / 100)))

# ==========================================
# 6. 進階批量輸出功能 (修復全黑問題)
# ==========================================
st.divider()
st.header("🚀 批量輸出設定")
out_c1, out_c2, out_c3 = st.columns(3)

with out_c1:
    out_mode = st.radio("輸出模式", ["完整 (背景+文字)", "透明 (僅限文字)"])
    out_layout = st.radio("排版方式", ["單張圖片 (ZIP)", "A4 自動拼板 (Print Ready)"])

with out_c2:
    target_width_cm = st.number_input("輸出寬度 (CM)", 1.0, 50.0, 10.0)
    a4_margin_cm = st.number_input("A4 邊界 (CM)", 0.0, 5.0, 1.0)

with out_c3:
    item_w_px = int(target_width_cm * PX_PER_CM)
    item_h_px = int(item_w_px * (H / W))
    st.info(f"解析度: 300 DPI\n單張大小: {item_w_px}x{item_h_px} 像素")

if st.button("🔥 開始批量生成", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先選取製作名單")
    else:
        results = []
        prog = st.progress(0); status = st.empty()
        
        for idx, (i, row) in enumerate(target_df.iterrows()):
            status.text(f"製作中: {idx+1}/{len(target_df)}")
            canvas = bg_img.copy() if out_mode == "完整 (背景+文字)" else Image.new("RGBA", (int(W), int(H)), (0,0,0,0))
            draw = ImageDraw.Draw(canvas)
            for col in display_cols:
                sv = st.session_state.settings[col]
                res = draw_styled_text(draw, str(row[col]), (sv["x"], sv["y"]), get_font_obj(sv["size"]), sv["color"], sv["align"], sv["bold"], sv["italic"])
                if res: canvas.alpha_composite(res[0], dest=res[1])
            
            resized = canvas.resize((item_w_px, item_h_px), Image.LANCZOS)
            results.append((str(row[id_col]), resized))
            prog.progress((idx + 1) / len(target_df))

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            if out_layout == "單張圖片 (ZIP)":
                for name, img in results:
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    zf.writestr(f"{name}.png", buf.getvalue())
            else:
                # A4 拼板核心修正：處理透明疊加
                margin_px = int(a4_margin_cm * PX_PER_CM)
                gap_px = 10 
                # 先用 RGBA 建立 A4 底色為白色，方便疊加透明圖片
                curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                cx, cy, max_rh, page_idx = margin_px, margin_px, 0, 1
                
                for idx, (name, img) in enumerate(results):
                    # 換行
                    if cx + item_w_px > A4_W_PX - margin_px:
                        cx = margin_px
                        cy += max_rh + gap_px
                        max_rh = 0
                    
                    # 換頁
                    if cy + item_h_px > A4_H_PX - margin_px:
                        # 儲存當前頁 (轉回 RGB)
                        final_page = curr_page.convert("RGB")
                        buf = io.BytesIO(); final_page.save(buf, format="JPEG", quality=95)
                        zf.writestr(f"A4_Print_Page_{page_idx}.jpg", buf.getvalue())
                        # 重置新頁面
                        curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                        cx, cy, max_rh, page_idx = margin_px, margin_px, 0, page_idx + 1
                    
                    # 使用 alpha_composite 或 paste(mask) 以支援透明度
                    # 我們使用透明度遮罩貼上，這樣不論透明或完整模式都不會變黑
                    curr_page.paste(img, (cx, cy), img)
                    
                    max_rh = max(max_rh, item_h_px)
                    cx += item_w_px + gap_px
                
                # 最後一頁
                final_page = curr_page.convert("RGB")
                buf = io.BytesIO(); final_page.save(buf, format="JPEG", quality=95)
                zf.writestr(f"A4_Print_Page_{page_idx}.jpg", buf.getvalue())

        status.text("✅ 生成完成！")
        st.download_button("📥 下載打包檔案 (ZIP)", zip_buf.getvalue(), "output_v6_1.zip", "application/zip", use_container_width=True)
