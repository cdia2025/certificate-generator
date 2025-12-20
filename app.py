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
st.set_page_config(page_title="專業證書生成器 V6.4 高效選取版", layout="wide")

DPI = 300
PX_PER_CM = DPI / 2.54 
A4_W_PX = int(21.0 * PX_PER_CM)
A4_H_PX = int(29.7 * PX_PER_CM)

# 初始化 Session State
if "settings" not in st.session_state: st.session_state.settings = {}
if "linked_layers" not in st.session_state: st.session_state.linked_layers = []
if "master_selection" not in st.session_state: st.session_state.master_selection = set()

def reset_project():
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

def sync_coord(col, axis, trigger):
    nk, sk = f"num_{axis}_{col}", f"sl_{axis}_{col}"
    if trigger == 'num': st.session_state[sk] = st.session_state[nk]
    else: st.session_state[nk] = st.session_state[sk]
    st.session_state.settings[col][axis] = st.session_state[nk]

# ==========================================
# 2. 字體處理與繪製邏輯
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
st.title("✉️ 專業證書生成器 V6.4")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 上傳背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 上傳資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 歡迎！請上傳檔案後開始設計。V6.4 支援表格化勾選名單與寬度打字輸入。")
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
# 5. 主頁面：名單管理與預覽 (表格化優化)
# ==========================================
st.divider()
st.header("👥 製作名單選取")
id_col = st.selectbox("選擇主識別欄位 (檔案命名基準)", df.columns, key="id_sel")

# 建立選擇用資料表
if "selection_df" not in st.session_state:
    st.session_state.selection_df = pd.DataFrame({"選取": False, id_col: df[id_col].astype(str)})

# 全選與清空邏輯
c_btn1, c_btn2, _ = st.columns([1, 1, 4])
with c_btn1:
    if st.button("🔳 全選所有", use_container_width=True):
        st.session_state.selection_df["選取"] = True
with c_btn2:
    if st.button("🗑️ 清空選取", use_container_width=True):
        st.session_state.selection_df["選取"] = False

# 搜尋過濾功能
search_q = st.text_input("🔍 搜尋名單 (輸入關鍵字過濾下方表格)", "")
filtered_selection_df = st.session_state.selection_df[
    st.session_state.selection_df[id_col].str.contains(search_q, case=False)
]

# 表格化選取區 (Data Editor)
st.write("請在下方表格勾選要製作的人員：")
edited_df = st.data_editor(
    filtered_selection_df,
    column_config={"選取": st.column_config.CheckboxColumn(required=True)},
    disabled=[id_col],
    hide_index=True,
    use_container_width=True,
    key="list_editor"
)

# 同步回 session_state 的主表
st.session_state.selection_df.update(edited_df)

# 取得最終選取的名單
final_selected_ids = st.session_state.selection_df[st.session_state.selection_df["選取"] == True][id_col].tolist()
target_df = df[df[id_col].astype(str).isin(final_selected_ids)]

# 即時預覽
if not target_df.empty:
    st.subheader(f"👁️ 畫布即時預覽 (已選取 {len(final_selected_ids)} 筆)")
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
# 6. 生成與進階排版 (支援打字輸入寬度)
# ==========================================
st.divider()
st.header("🚀 批量輸出設定")
out_c1, out_c2, out_c3 = st.columns(3)

with out_c1:
    out_mode = st.radio("輸出內容", ["完整 (背景+文字)", "透明 (僅限文字)"])
    out_layout = st.radio("佈局方式", ["單張圖片 (ZIP)", "A4 自動拼板 (Print Ready)"])

with out_c2:
    st.write("**物件輸出寬度 (CM)**")
    w_col1, w_col2 = st.columns([1, 2])
    with w_col1:
        # 數值輸入框，支援打字
        target_width_cm = st.number_input("CM 數值", 1.0, 50.0, 10.0, step=0.1, label_visibility="collapsed")
    with w_col2:
        # 同步的滑桿，支援拖動
        target_width_cm = st.slider("CM 滑桿", 1.0, 50.0, target_width_cm, step=0.1, label_visibility="collapsed")
    
    a4_margin_cm = st.number_input("A4 頁邊界 (CM)", 0.0, 5.0, 1.0, step=0.1)

with out_c3:
    item_w_px = int(target_width_cm * PX_PER_CM)
    item_h_px = int(item_w_px * (H / W))
    st.info(f"解析度: 300 DPI\n單一物件像素: {item_w_px}x{item_h_px}")

if st.button("🔥 開始執行生成任務", type="primary", use_container_width=True):
    if not final_selected_ids:
        st.warning("請先在表格中勾選名單！")
    else:
        results = []
        prog = st.progress(0); status = st.empty()
        
        for idx, (i, row) in enumerate(target_df.iterrows()):
            status.text(f"處理中: {idx+1}/{len(target_df)}")
            canvas = bg_img.copy() if out_mode == "完整 (背景+文字)" else Image.new("RGBA", (int(W), int(H)), (0,0,0,0))
            draw = ImageDraw.Draw(canvas)
            for col in display_cols:
                sv = st.session_state.settings[col]
                res = draw_styled_text(draw, str(row[col]), (sv["x"], sv["y"]), get_font_obj(sv["size"]), sv["color"], sv["align"], sv["bold"], sv["italic"])
                if res: canvas.alpha_composite(res[0], dest=res[1])
            
            # 高品質縮放
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
                margin_px = int(a4_margin_cm * PX_PER_CM)
                gap_px = 10 
                curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                cx, cy, max_rh, page_idx = margin_px, margin_px, 0, 1
                
                for idx, (name, img) in enumerate(results):
                    if cx + item_w_px > A4_W_PX - margin_px:
                        cx = margin_px
                        cy += max_rh + gap_px
                        max_rh = 0
                    if cy + item_h_px > A4_H_PX - margin_px:
                        final_page = curr_page.convert("RGB")
                        buf = io.BytesIO(); final_page.save(buf, format="JPEG", quality=95)
                        zf.writestr(f"A4_Layout_Page_{page_idx}.jpg", buf.getvalue())
                        curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                        cx, cy, max_rh, page_idx = margin_px, margin_px, 0, page_idx + 1
                    
                    curr_page.paste(img, (cx, cy), img)
                    max_rh = max(max_rh, item_h_px)
                    cx += item_w_px + gap_px
                
                final_page = curr_page.convert("RGB")
                buf = io.BytesIO(); final_page.save(buf, format="JPEG", quality=95)
                zf.writestr(f"A4_Layout_Page_{page_idx}.jpg", buf.getvalue())

        status.text("✅ 生成任務已完成！")
        st.download_button("📥 下載產出的壓縮包 (ZIP)", zip_buf.getvalue(), "output_v6_4.zip", "application/zip", use_container_width=True)
