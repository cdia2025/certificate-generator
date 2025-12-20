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
st.set_page_config(page_title="專業證書生成器 V7.6 即時連動版", layout="wide")

DPI = 300
PX_PER_CM = DPI / 2.54 
A4_W_PX = int(21.0 * PX_PER_CM)
A4_H_PX = int(29.7 * PX_PER_CM)

# 初始化 Session State 核心數據
if "settings" not in st.session_state: st.session_state.settings = {}
if "linked_layers" not in st.session_state: st.session_state.linked_layers = []
if "out_w_cm" not in st.session_state: st.session_state.out_w_cm = 10.0

# 初始化批量位移基準點 (用於計算 Delta 位移量)
if "last_batch_x" not in st.session_state: st.session_state.last_batch_x = 0.0
if "last_batch_y" not in st.session_state: st.session_state.last_batch_y = 0.0
if "last_batch_s" not in st.session_state: st.session_state.last_batch_s = 0

# --- 【核心邏輯 1】個別座標同步 ---
def sync_widget(col, axis, source):
    num_key = f"nx_{col}" if axis == 'x' else f"ny_{col}"
    sl_key = f"sx_{col}" if axis == 'x' else f"sy_{col}"
    if source == 'num':
        st.session_state[sl_key] = st.session_state[num_key]
    else:
        st.session_state[num_key] = st.session_state[sl_key]
    st.session_state.settings[col][axis] = st.session_state[num_key]

# --- 【核心邏輯 2】即時批量同步 (即拖即動，不需按鍵) ---
def batch_sync_live(axis):
    """根據滑桿變動的差值 (Delta)，即時更新所有連結圖層"""
    if not st.session_state.linked_layers:
        return

    if axis == 'x':
        current = st.session_state.batch_sl_x
        delta = current - st.session_state.last_batch_x
        st.session_state.last_batch_x = current
        param, nk, sk = 'x', 'nx_', 'sx_'
    elif axis == 'y':
        current = st.session_state.batch_sl_y
        delta = current - st.session_state.last_batch_y
        st.session_state.last_batch_y = current
        param, nk, sk = 'y', 'ny_', 'sy_'
    else:
        current = st.session_state.batch_sl_s
        delta = current - st.session_state.last_batch_s
        st.session_state.last_batch_s = current
        param, nk, sk = 'size', 'size_', None

    # 更新所有勾選的圖層
    for c in st.session_state.linked_layers:
        new_val = st.session_state.settings[c][param] + delta
        # 限制範圍
        if axis in ['x', 'y']:
            limit = float(st.session_state.bg_width if axis == 'x' else st.session_state.bg_height)
            new_val = max(0.0, min(limit, new_val))
            st.session_state[f"{nk}{c}"] = new_val
            st.session_state[f"{sk}{c}"] = new_val
        else:
            new_val = int(max(10, min(5000, new_val)))
            st.session_state[f"{nk}{c}"] = new_val
        
        st.session_state.settings[c][param] = new_val

def reset_project():
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

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
st.title("✉️ 專業證書生成器 V7.6")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 上傳背景圖與資料後，即可在側邊欄進行詳細調整。批量工具已更新為即時連動模式。")
    st.stop()

bg_img = Image.open(bg_file).convert("RGBA")
W, H = float(bg_img.size[0]), float(bg_img.size[1])
st.session_state.bg_width, st.session_state.bg_height = W, H
mid_x, mid_y = W / 2, H / 2
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄控制
# ==========================================
with st.sidebar:
    if st.button("🆕 新專案 / 重新重置", use_container_width=True): reset_project()
    st.header("⚙️ 屬性面板")
    
    with st.expander("💾 配置管理"):
        js = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
        st.download_button("📤 匯出設定 (JSON)", js, "config.json", "application/json")
        uploaded_config = st.file_uploader("📥 載入舊設定", type=["json"])
        if uploaded_config:
            st.session_state.settings.update(json.load(uploaded_config))
            # 清理 key 強制重繪
            for k in list(st.session_state.keys()):
                if any(x in k for x in ["nx_", "sx_", "ny_", "sy_", "size_"]): del st.session_state[k]
            st.rerun()

    display_cols = st.multiselect("顯示欄位", df.columns, default=[df.columns[0]])
    
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": mid_x, "y": mid_y, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        
        s_val = st.session_state.settings[col]
        if f"nx_{col}" not in st.session_state: st.session_state[f"nx_{col}"] = float(s_val["x"])
        if f"sx_{col}" not in st.session_state: st.session_state[f"sx_{col}"] = float(s_val["x"])
        if f"ny_{col}" not in st.session_state: st.session_state[f"ny_{col}"] = float(s_val["y"])
        if f"sy_{col}" not in st.session_state: st.session_state[f"sy_{col}"] = float(s_val["y"])
        if f"size_{col}" not in st.session_state: st.session_state[f"size_{col}"] = int(s_val["size"])

    st.divider()
    st.subheader("📝 個別圖層設定")
    for col in display_cols:
        link_tag = " (🔗)" if col in st.session_state.linked_layers else ""
        with st.expander(f"圖層：{col}{link_tag}"):
            s = st.session_state.settings[col]
            st.caption(f"📍 中心點參考：X={mid_x:.0f}, Y={mid_y:.0f}")
            s["x"] = st.number_input("X座標數值", 0.0, W, key=f"nx_{col}", on_change=sync_widget, args=(col, 'x', 'num'))
            s["x"] = st.slider("X座標滑桿", 0.0, W, key=f"sx_{col}", on_change=sync_widget, args=(col, 'x', 'sl'), label_visibility="collapsed")
            s["y"] = st.number_input("Y座標數值", 0.0, H, key=f"ny_{col}", on_change=sync_widget, args=(col, 'y', 'num'))
            s["y"] = st.slider("Y座標滑桿", 0.0, H, key=f"sy_{col}", on_change=sync_widget, args=(col, 'y', 'sl'), label_visibility="collapsed")
            f1, f2 = st.columns(2)
            with f1: s["size"] = st.number_input("字體大小", 10, 5000, key=f"size_{col}")
            with f2: s["color"] = st.color_picker("文字顏色", s["color"], key=f"color_{col}")
            sc1, sc2 = st.columns(2)
            with sc1: s["bold"] = st.checkbox("加粗", s["bold"], key=f"bold_{col}")
            with sc2: s["italic"] = st.checkbox("傾斜", s["italic"], key=f"italic_{col}")
            s["align"] = st.selectbox("對齊", ["左對齊", "居中", "右對齊"], index=["左對齊", "居中", "右對齊"].index(s.get("align", "居中")), key=f"align_{col}")

    st.divider()
    # 🔗 即時批量位移工具 (無需按鈕，獨立觸發)
    with st.expander("🔗 批量即時連動工具 (免按鍵)", expanded=False):
        st.info(f"📍 中心參考：X={mid_x:.0f}, Y={mid_y:.0f}")
        st.session_state.linked_layers = st.multiselect("選取同步對象", display_cols)
        
        st.write("**⬅️ 左右批量位移 ➡️**")
        st.slider("調整中...", -W, W, 0.0, key="batch_sl_x", on_change=batch_sync_live, args=('x',))
        
        st.write("**⬆️ 上下批量位移 ⬇️**")
        st.slider("調整中...", -H, H, 0.0, key="batch_sl_y", on_change=batch_sync_live, args=('y',))
        
        st.write("**🔍 批量縮放大小 🔍**")
        st.slider("調整中...", -1000, 1000, 0, key="batch_sl_s", on_change=batch_sync_live, args=('s',))
        
        st.caption("提示：此工具會根據滑桿移動的相對量，即時改變上方已選取圖層的位置。")

# ==========================================
# 5. 主頁面：名單選取
# ==========================================
st.divider()
st.header("👥 製作名單選取")
id_col = st.selectbox("選擇主識別欄位 (檔名基準)", df.columns, key="id_sel")

if "selection_df" not in st.session_state or st.session_state.get('last_id_col') != id_col:
    st.session_state.selection_df = pd.DataFrame({"選取": False}, index=df[id_col].astype(str).unique())
    st.session_state.last_id_col = id_col

c_btn1, c_btn2, _ = st.columns([1, 1, 4])
with c_btn1:
    if st.button("🔳 全選名單", use_container_width=True): st.session_state.selection_df["選取"] = True
with c_btn2:
    if st.button("🗑️ 清空選取", use_container_width=True): st.session_state.selection_df["選取"] = False

search_q = st.text_input("🔍 搜尋名單...", "")
view_df = st.session_state.selection_df.copy()
if search_q: view_df = view_df[view_df.index.str.contains(search_q, case=False)]

edited_view = st.data_editor(view_df, column_config={"選取": st.column_config.CheckboxColumn("選取", default=False, required=True)}, use_container_width=True, key="list_editor_v76")
if not edited_view.equals(view_df):
    st.session_state.selection_df.update(edited_view)
    st.rerun()

final_selected_ids = st.session_state.selection_df[st.session_state.selection_df["選取"] == True].index.tolist()
target_df = df[df[id_col].astype(str).isin(final_selected_ids)]

# 即時預覽
if not target_df.empty:
    st.subheader(f"👁️ 即時畫布預覽 (第一位：{target_df.iloc[0][id_col]})")
    zoom = st.slider("🔍 畫布縮放 (%)", 50, 250, 100, step=10, key="zoom_sl")
    row = target_df.iloc[0]
    canvas = bg_img.copy()
    draw = ImageDraw.Draw(canvas)
    for col in display_cols:
        # 抓取最新狀態
        cur_x = st.session_state.get(f"nx_{col}", mid_x)
        cur_y = st.session_state.get(f"ny_{col}", mid_y)
        s_dict = st.session_state.settings[col]
        f_obj = get_font_obj(int(st.session_state.get(f"size_{col}", s_dict['size'])))
        res = draw_styled_text(draw, str(row[col]), (cur_x, cur_y), f_obj, s_dict["color"], s_dict["align"], s_dict["bold"], s_dict["italic"])
        if res: canvas.alpha_composite(res[0], dest=res[1])
        gc = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw.line([(0, cur_y), (W, cur_y)], fill=gc, width=2)
        draw.line([(cur_x, 0), (cur_x, H)], fill=gc, width=2)
    st.image(canvas, width=int(W * (zoom / 100)))

# ==========================================
# 6. 生成與導出
# ==========================================
st.divider()
st.header("🚀 批量輸出設定")
out_c1, out_c2, out_c3 = st.columns(3)

with out_c1:
    out_mode = st.radio("輸出內容", ["完整 (背景+文字)", "透明 (僅限文字)"])
    out_layout = st.radio("排版佈局", ["單張圖片 (ZIP)", "A4 自動拼板 (Print Ready)"])

with out_c2:
    w_num = st.number_input("物件輸出寬度 (CM)", 1.0, 100.0, float(st.session_state.out_w_cm), step=0.1, key="w_num_in")
    w_sl = st.slider("快速拖動寬度", 1.0, 100.0, float(w_num), step=0.1, key="w_sl_in", label_visibility="collapsed")
    st.session_state.out_w_cm = w_sl
    a4_margin_cm = st.number_input("A4 頁邊留白 (CM)", 0.0, 5.0, 1.0, step=0.1)
    item_gap_mm = st.number_input("圖塊間距 (MM)", 0.0, 10.0, 0.5, step=0.1)

with out_c3:
    item_w_px = int(st.session_state.out_w_cm * PX_PER_CM)
    item_h_px = int(item_w_px * (H / W))
    st.info(f"解析度: 300 DPI\n單張尺寸: {item_w_px}x{item_h_px} 像素")

if st.button("🔥 開始批量生成所有選定檔案", type="primary", use_container_width=True):
    if not final_selected_ids:
        st.warning("請先勾選製作名單！")
    else:
        results = []
        prog = st.progress(0); status = st.empty()
        for idx, (i, row) in enumerate(target_df.iterrows()):
            status.text(f"正在製作: {idx+1}/{len(target_df)} ({row[id_col]})")
            canvas = bg_img.copy() if out_mode == "完整 (背景+文字)" else Image.new("RGBA", (int(W), int(H)), (0,0,0,0))
            draw = ImageDraw.Draw(canvas)
            for col in display_cols:
                # 使用最新 UI 數據
                cx = st.session_state.get(f"nx_{col}", mid_x)
                cy = st.session_state.get(f"ny_{col}", mid_y)
                sz = int(st.session_state.get(f"size_{col}", 60))
                s_final = st.session_state.settings[col]
                res = draw_styled_text(draw, str(row[col]), (cx, cy), get_font_obj(sz), s_final["color"], s_final["align"], s_final["bold"], s_final["italic"])
                if res: canvas.alpha_composite(res[0], dest=res[1])
            resized = canvas.resize((item_w_px, item_h_px), Image.LANCZOS)
            results.append((str(row[id_col]), resized))
            prog.progress((idx + 1) / len(target_df))

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as zf:
            if out_layout == "單張圖片 (ZIP)":
                for name, img in results:
                    buf = io.BytesIO(); img.save(buf, format="PNG"); zf.writestr(f"{name}.png", buf.getvalue())
            else:
                margin_px = int(a4_margin_cm * PX_PER_CM)
                gap_px = int((item_gap_mm / 10) * PX_PER_CM) 
                curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                cx, cy, max_rh, page_idx = margin_px, margin_px, 0, 1
                for idx, (name, img) in enumerate(results):
                    if cx + item_w_px > A4_W_PX - margin_px: cx, cy, max_rh = margin_px, cy + max_rh + gap_px, 0
                    if cy + item_h_px > A4_H_PX - margin_px:
                        buf = io.BytesIO(); curr_page.convert("RGB").save(buf, format="JPEG", quality=95); zf.writestr(f"A4_Layout_{page_idx}.jpg", buf.getvalue())
                        curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                        cx, cy, max_rh, page_idx = margin_px, margin_px, 0, page_idx + 1
                    curr_page.paste(img, (cx, cy), img)
                    max_rh = max(max_rh, item_h_px); cx += item_w_px + gap_px
                buf = io.BytesIO(); curr_page.convert("RGB").save(buf, format="JPEG", quality=95); zf.writestr(f"A4_Layout_{page_idx}.jpg", buf.getvalue())

        status.text("✅ 生成完成！")
        st.download_button("📥 下載生成的壓縮包 (ZIP)", zip_buf.getvalue(), "output_v7_6.zip", "application/zip", use_container_width=True)
