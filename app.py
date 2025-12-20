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
st.set_page_config(page_title="專業證書生成器 V6.0 印刷增強版", layout="wide")

# 常數定義 (300 DPI 標準)
DPI = 300
PX_PER_CM = DPI / 2.54  # 約 118.11 像素/公分
A4_W_PX = int(21.0 * PX_PER_CM)
A4_H_PX = int(29.7 * PX_PER_CM)

if "settings" not in st.session_state: st.session_state.settings = {}
if "linked_layers" not in st.session_state: st.session_state.linked_layers = []

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
# 3. 介面與檔案處理
# ==========================================
st.title("✉️ 專業證書生成器 V6.0 (印刷拼板版)")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 請上傳檔案。本版本支援 A4 自動拼板與透明背景輸出。")
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
    st.header("⚙️ 屬性面板")
    
    with st.expander("💾 配置管理"):
        if st.session_state.settings:
            st.download_button("📤 匯出設定 (JSON)", json.dumps(st.session_state.settings, indent=4, ensure_ascii=False), "config.json", "application/json")
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
    st.subheader("📝 圖層屬性")
    for col in display_cols:
        with st.expander(f"圖層：{col}"):
            s = st.session_state.settings[col]
            st.caption(f"📍 中心點參考：X={mid_x:.1f}, Y={mid_y:.1f}")
            c1, c2 = st.columns([1, 2])
            with c1: st.number_input("X 數值", 0.0, W, key=f"num_x_{col}", on_change=sync_coord, args=(col, 'x', 'num'), label_visibility="collapsed")
            with c2: st.slider("X 滑桿", 0.0, W, key=f"sl_x_{col}", on_change=sync_coord, args=(col, 'x', 'sl'), label_visibility="collapsed")
            c1, c2 = st.columns([1, 2])
            with c1: st.number_input("Y 數值", 0.0, H, key=f"num_y_{col}", on_change=sync_coord, args=(col, 'y', 'num'), label_visibility="collapsed")
            with c2: st.slider("Y 滑桿", 0.0, H, key=f"sl_y_{col}", on_change=sync_coord, args=(col, 'y', 'sl'), label_visibility="collapsed")
            f1, f2 = st.columns(2)
            with f1: s["size"] = st.number_input("大小", 10, 2000, int(s["size"]), key=f"sz_{col}")
            with f2: s["color"] = st.color_picker("顏色", s["color"], key=f"cp_{col}")
            sc1, sc2 = st.columns(2)
            with sc1: s["bold"] = st.checkbox("粗體", s["bold"], key=f"bd_{col}")
            with sc2: s["italic"] = st.checkbox("斜體", s["italic"], key=f"it_{col}")
            s["align"] = st.selectbox("對齊", ["左對齊", "居中", "右對齊"], index=["左對齊", "居中", "右對齊"].index(s["align"]), key=f"al_{col}")

    st.divider()
    with st.expander("🔗 批量位移工具"):
        st.session_state.linked_layers = st.multiselect("連結對象", display_cols)
        lc1, lc2 = st.columns(2)
        with lc1: b_x = st.number_input("批量 X", value=0.0)
        with lc2: b_y = st.number_input("批量 Y", value=0.0)
        b_s = st.number_input("批量縮放", value=0)
        if st.button("執行批量套用"):
            for c in st.session_state.linked_layers:
                nx, ny = max(0.0, min(W, st.session_state.settings[c]["x"] + b_x)), max(0.0, min(H, st.session_state.settings[c]["y"] + b_y))
                ns = max(10, st.session_state.settings[c]["size"] + b_s)
                st.session_state.settings[c].update({"x": nx, "y": ny, "size": ns})
                st.session_state[f"num_x_{c}"] = st.session_state[f"sl_x_{c}"] = nx
                st.session_state[f"num_y_{c}"] = st.session_state[f"sl_y_{c}"] = ny
            st.rerun()

# ==========================================
# 5. 主頁面：預覽與搜尋
# ==========================================
st.divider()
id_col = st.selectbox("識別欄位 (用於命名)", df.columns, key="id_sel")
all_n = df[id_col].astype(str).tolist()

# 搜尋功能
search_term = st.text_input("🔍 搜尋預覽名單...", "").strip().lower()
filtered_n = [n for n in all_n if search_term in n.lower()] if search_term else all_n

p1, p2 = st.columns([1, 1])
with p1: st.checkbox("全選過濾後的名單", key="pre_all_chk", on_change=lambda: st.session_state.update({"pre_sel": filtered_n if st.session_state.pre_all_chk else []}))
with p2: sel_n = st.multiselect(f"已選取預覽名單 ({len(filtered_n)} 筆相符)", filtered_n, key="pre_sel")

target_df = df[df[id_col].astype(str).isin(sel_n)]
zoom = st.slider("🔍 視覺縮放 (%)", 50, 250, 100, step=10, key="zoom_sl")

if not target_df.empty:
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
# 6. 進階批量輸出功能
# ==========================================
st.divider()
st.header("🚀 批量輸出設定")
out_c1, out_c2, out_c3 = st.columns(3)

with out_c1:
    out_mode = st.radio("輸出內容", ["完整 (背景+文字)", "透明 (僅限文字)"])
    out_type = st.radio("輸出佈局", ["單張圖片 (ZIP)", "A4 自動拼板 (Print Ready)"])

with out_c2:
    target_width_cm = st.number_input("指定輸出寬度 (CM)", 1.0, 50.0, 10.0, help="僅影響最終輸出檔案的尺寸")
    a4_margin_cm = st.number_input("A4 邊界留白 (CM)", 0.0, 5.0, 1.0)

with out_c3:
    st.write("**A4 拼板預估：**")
    item_w_px = int(target_width_cm * PX_PER_CM)
    item_h_px = int(item_w_px * (H / W))
    cols = max(1, int((21.0 - 2 * a4_margin_cm) // target_width_cm))
    rows = max(1, int((29.7 - 2 * a4_margin_cm) // (target_width_cm * (H / W))))
    st.info(f"每張 A4 可容納: {cols}x{rows} = {cols*rows} 張")

if st.button("🔥 開始批量生成", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先選取對象名單")
    else:
        results = []
        prog = st.progress(0); status = st.empty()
        
        for idx, (i, row) in enumerate(target_df.iterrows()):
            status.text(f"處理中: {idx+1}/{len(target_df)}")
            # 建立畫布
            canvas = bg_img.copy() if out_mode == "完整 (背景+文字)" else Image.new("RGBA", (int(W), int(H)), (0,0,0,0))
            draw = ImageDraw.Draw(canvas)
            for col in display_cols:
                sv = st.session_state.settings[col]
                res = draw_styled_text(draw, str(row[col]), (sv["x"], sv["y"]), get_font_obj(sv["size"]), sv["color"], sv["align"], sv["bold"], sv["italic"])
                if res: canvas.alpha_composite(res[0], dest=res[1])
            
            # 縮放至指定 CM
            resized = canvas.resize((item_w_px, item_h_px), Image.LANCZOS)
            results.append((str(row[id_col]), resized))
            prog.progress((idx + 1) / len(target_df))

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            if out_type == "單張圖片 (ZIP)":
                for name, img in results:
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    zf.writestr(f"{name}.png", buf.getvalue())
            else:
                # A4 拼板邏輯
                margin_px = int(a4_margin_cm * PX_PER_CM)
                gap_px = 10 # 圖片間距
                curr_page = Image.new("RGB", (A4_W_PX, A4_H_PX), "white")
                cx, cy, max_rh, page_idx = margin_px, margin_px, 0, 1
                
                for idx, (name, img) in enumerate(results):
                    # 換行檢查
                    if cx + item_w_px > A4_W_PX - margin_px:
                        cx = margin_px
                        cy += max_rh + gap_px
                        max_rh = 0
                    
                    # 換頁檢查
                    if cy + item_h_px > A4_H_PX - margin_px:
                        buf = io.BytesIO(); curr_page.save(buf, format="JPEG", quality=95)
                        zf.writestr(f"A4_Layout_Page_{page_idx}.jpg", buf.getvalue())
                        curr_page = Image.new("RGB", (A4_W_PX, A4_H_PX), "white")
                        cx, cy, max_rh, page_idx = margin_px, margin_px, 0, page_idx + 1
                    
                    curr_page.paste(img.convert("RGB"), (cx, cy))
                    max_rh = max(max_rh, item_h_px)
                    cx += item_w_px + gap_px
                
                # 存最後一頁
                buf = io.BytesIO(); curr_page.save(buf, format="JPEG", quality=95)
                zf.writestr(f"A4_Layout_Page_{page_idx}.jpg", buf.getvalue())

        status.text("✅ 生成完成！")
        st.download_button("📥 下載打包檔案 (ZIP)", zip_buf.getvalue(), "batch_output.zip", "application/zip", use_container_width=True)
