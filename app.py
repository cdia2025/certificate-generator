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
st.set_page_config(page_title="專業證書生成器 V7.0 穩定版", layout="wide")

DPI = 300
PX_PER_CM = DPI / 2.54 
A4_W_PX = int(21.0 * PX_PER_CM)
A4_H_PX = int(29.7 * PX_PER_CM)

# 初始化 Session State
if "settings" not in st.session_state: st.session_state.settings = {}
if "linked_layers" not in st.session_state: st.session_state.linked_layers = []
if "out_w_cm" not in st.session_state: st.session_state.out_w_cm = 10.0

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
# 3. 檔案上傳區
# ==========================================
st.title("✉️ 專業證書生成器 V7.0")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 背景圖片", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 資料檔", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("👋 上傳背景圖與資料檔後，即可在側邊欄進行詳細排版。")
    st.stop()

bg_img = Image.open(bg_file).convert("RGBA")
W, H = float(bg_img.size[0]), float(bg_img.size[1])
mid_x, mid_y = W / 2, H / 2
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄控制面板 (穩定版狀態管理)
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
            st.rerun()

    display_cols = st.multiselect("選擇要在證書上顯示的欄位", df.columns, default=[df.columns[0]])
    
    # 預先初始化字典數據，避免 KeyError
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": mid_x, "y": mid_y, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        else:
            # 確保所有參數完整
            defaults = {"x": mid_x, "y": mid_y, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
            for k, v in defaults.items():
                if k not in st.session_state.settings[col]: st.session_state.settings[col][k] = v

    st.divider()
    # 📝 個別圖層屬性
    st.subheader("📝 個別圖層設定")
    for col in display_cols:
        link_tag = " (🔗)" if col in st.session_state.linked_layers else ""
        with st.expander(f"圖層：{col}{link_tag}"):
            s = st.session_state.settings[col]
            st.caption(f"📍 中心參考：X={mid_x:.0f}, Y={mid_y:.0f}")
            
            # 使用 value 參數綁定字典，不使用 key 直接綁定 session_state 以防止崩潰
            s["x"] = st.number_input(f"X座標數值", 0.0, W, float(s["x"]), key=f"nx_{col}")
            s["x"] = st.slider(f"X座標滑桿", 0.0, W, float(s["x"]), key=f"sx_{col}", label_visibility="collapsed")
            
            s["y"] = st.number_input(f"Y座標數值", 0.0, H, float(s["y"]), key=f"ny_{col}")
            s["y"] = st.slider(f"Y座標滑桿", 0.0, H, float(s["y"]), key=f"sy_{col}", label_visibility="collapsed")
            
            f1, f2 = st.columns(2)
            with f1: s["size"] = st.number_input("字體大小", 10, 3000, int(s["size"]), key=f"size_{col}")
            with f2: s["color"] = st.color_picker("文字顏色", s["color"], key=f"color_{col}")
            
            sc1, sc2 = st.columns(2)
            with sc1: s["bold"] = st.checkbox("加粗", s["bold"], key=f"bold_{col}")
            with sc2: s["italic"] = st.checkbox("傾斜", s["italic"], key=f"italic_{col}")
            
            s["align"] = st.selectbox("對齊", ["左對齊", "居中", "右對齊"], index=["左對齊", "居中", "右對齊"].index(s["align"]), key=f"align_{col}")

    st.divider()
    # 🔗 批量位移工具 (穩定版：不直接修改元件 Key)
    with st.expander("🔗 批量連結與位移工具", expanded=False):
        st.info(f"📍 畫布中心參考：X={mid_x:.0f}, Y={mid_y:.0f}")
        st.session_state.linked_layers = st.multiselect("選取要同時移動的對象", display_cols)
        
        st.write("**批量位移 (相對偏移)**")
        bx = st.slider("左右偏移", -W, W, 0.0, key="batch_sl_x")
        by = st.slider("上下偏移", -H, H, 0.0, key="batch_sl_y")
        bs = st.slider("大小增減", -500, 500, 0, key="batch_sl_s")
        
        if st.button("🚀 執行批量套用變更", use_container_width=True):
            if not st.session_state.linked_layers:
                st.warning("請先選取要連結的圖層")
            else:
                for c in st.session_state.linked_layers:
                    st.session_state.settings[c]["x"] = max(0.0, min(W, st.session_state.settings[c]["x"] + bx))
                    st.session_state.settings[c]["y"] = max(0.0, min(H, st.session_state.settings[c]["y"] + by))
                    st.session_state.settings[c]["size"] = max(10, st.session_state.settings[c]["size"] + bs)
                st.success("批量修改已更新至數據字典")
                st.rerun() # 透過重新執行腳本來更新所有元件顯示

# ==========================================
# 5. 主頁面：製作名單 (表格選取)
# ==========================================
st.divider()
st.header("👥 製作名單選取")
id_col = st.selectbox("選擇主識別欄位 (檔名基準)", df.columns, key="id_sel")

if "selection_df" not in st.session_state:
    st.session_state.selection_df = pd.DataFrame({"選取": False, id_col: df[id_col].astype(str)})

c_btn1, c_btn2, _ = st.columns([1, 1, 4])
with c_btn1:
    if st.button("🔳 全選名單", use_container_width=True): st.session_state.selection_df["選取"] = True
with c_btn2:
    if st.button("🗑️ 清空選取", use_container_width=True): st.session_state.selection_df["選取"] = False

search_q = st.text_input("🔍 搜尋名單...", "")
filtered_selection_df = st.session_state.selection_df[st.session_state.selection_df[id_col].str.contains(search_q, case=False)]

edited_df = st.data_editor(
    filtered_selection_df,
    column_config={"選取": st.column_config.CheckboxColumn(required=True)},
    disabled=[id_col], hide_index=True, use_container_width=True, key="list_editor"
)
st.session_state.selection_df.update(edited_df)

final_selected_ids = st.session_state.selection_df[st.session_state.selection_df["選取"] == True][id_col].tolist()
target_df = df[df[id_col].astype(str).isin(final_selected_ids)]

# 即時預覽區
if not target_df.empty:
    st.subheader(f"👁️ 即時預覽 (已選取 {len(final_selected_ids)} 筆)")
    zoom = st.slider("🔍 畫布視覺縮放 (%)", 50, 250, 100, step=10, key="zoom_sl")
    row = target_df.iloc[0]
    canvas = bg_img.copy()
    draw = ImageDraw.Draw(canvas)
    for col in display_cols:
        s = st.session_state.settings[col]
        f_obj = get_font_obj(s["size"])
        res = draw_styled_text(draw, str(row[col]), (s["x"], s["y"]), f_obj, s["color"], s["align"], s["bold"], s["italic"])
        if res: canvas.alpha_composite(res[0], dest=res[1])
        gc = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw.line([(0, s["y"]), (W, s["y"])], fill=gc, width=2)
        draw.line([(s["x"], 0), (s["x"], H)], fill=gc, width=2)
    st.image(canvas, width=int(W * (zoom / 100)))

# ==========================================
# 6. 生成與排版 (印刷優化)
# ==========================================
st.divider()
st.header("🚀 批量輸出設定")
out_c1, out_c2, out_c3 = st.columns(3)

with out_c1:
    out_mode = st.radio("輸出模式", ["完整 (背景+文字)", "透明 (僅限文字)"])
    out_layout = st.radio("排版佈局", ["單張圖片 (ZIP)", "A4 自動拼板 (Print Ready)"])

with out_c2:
    st.write("**物件輸出寬度 (CM)**")
    # 輸出寬度同步邏輯 (穩定版)
    cur_w = st.session_state.out_w_cm
    w_num = st.number_input("CM 精確輸入", 1.0, 100.0, float(cur_w), step=0.1, key="w_num_input")
    w_sl = st.slider("CM 快速拖動", 1.0, 100.0, float(w_num), step=0.1, key="w_sl_input", label_visibility="collapsed")
    st.session_state.out_w_cm = w_sl
    
    a4_margin_cm = st.number_input("A4 頁邊距留白 (CM)", 0.0, 5.0, 1.0, step=0.1)
    item_gap_mm = st.number_input("物件間距 (MM)", 0.0, 10.0, 0.5, step=0.1)

with out_c3:
    item_w_px = int(st.session_state.out_w_cm * PX_PER_CM)
    item_h_px = int(item_w_px * (H / W))
    st.info(f"解析度: 300 DPI\n單一圖塊像素: {item_w_px}x{item_h_px}")

if st.button("🔥 開始批量生成任務", type="primary", use_container_width=True):
    if not final_selected_ids:
        st.warning("請先勾選名單！")
    else:
        results = []
        prog = st.progress(0); status = st.empty()
        for idx, (i, row) in enumerate(target_df.iterrows()):
            status.text(f"正在製作: {idx+1}/{len(target_df)} ({row[id_col]})")
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
                    buf = io.BytesIO(); img.save(buf, format="PNG"); zf.writestr(f"{name}.png", buf.getvalue())
            else:
                # A4 拼板邏輯
                margin_px = int(a4_margin_cm * PX_PER_CM)
                gap_px = int((item_gap_mm / 10) * PX_PER_CM) 
                curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                cx, cy, max_rh, page_idx = margin_px, margin_px, 0, 1
                for idx, (name, img) in enumerate(results):
                    if cx + item_w_px > A4_W_PX - margin_px: cx, cy, max_rh = margin_px, cy + max_rh + gap_px, 0
                    if cy + item_h_px > A4_H_PX - margin_px:
                        buf = io.BytesIO(); curr_page.convert("RGB").save(buf, format="JPEG", quality=95); zf.writestr(f"A4_Page_{page_idx}.jpg", buf.getvalue())
                        curr_page = Image.new("RGBA", (A4_W_PX, A4_H_PX), (255, 255, 255, 255))
                        cx, cy, max_rh, page_idx = margin_px, margin_px, 0, page_idx + 1
                    curr_page.paste(img, (cx, cy), img)
                    max_rh = max(max_rh, item_h_px); cx += item_w_px + gap_px
                buf = io.BytesIO(); curr_page.convert("RGB").save(buf, format="JPEG", quality=95); zf.writestr(f"A4_Page_{page_idx}.jpg", buf.getvalue())

        status.text("✅ 全部任務已完成！")
        st.download_button("📥 下載生成的壓縮包 (ZIP)", zip_buf.getvalue(), "batch_output.zip", "application/zip", use_container_width=True)
