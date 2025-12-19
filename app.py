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
st.set_page_config(page_title="專業證書生成器 V5.5 最終版", layout="wide")

# 初始化 Session State 儲存空間
if "settings" not in st.session_state:
    st.session_state.settings = {}
if "linked_layers" not in st.session_state:
    st.session_state.linked_layers = []

# ==========================================
# 2. 字體處理與進階繪製邏輯
# ==========================================

@st.cache_resource
def get_font_resource():
    """自動偵測系統字體或下載思源黑體，確保支援中文"""
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc",              # Windows 微軟正黑
        "C:/Windows/Fonts/dfkai-sb.ttf",          # Windows 標楷體
        "/System/Library/Fonts/STHeiti Light.ttc",   # macOS 華文黑體
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", # Linux
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"  # Linux
    ]
    for p in font_paths:
        if os.path.exists(p): return p

    # 雲端環境備援：下載思源黑體 (Noto Sans TC)
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
    """快取字體對象以優化渲染速度"""
    path = get_font_resource()
    try:
        if path: return ImageFont.truetype(path, size)
    except: pass
    return ImageFont.load_default()

def draw_styled_text(draw, text, pos, font, color, align="居中", bold=False, italic=False):
    """繪製支援對齊、模擬粗體與物理斜體變換的文字"""
    try:
        # 計算文字尺寸
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        tw, th = right - left, bottom - top
    except:
        tw, th = len(text) * font.size * 0.7, font.size

    x, y = pos
    if align == "居中": x -= tw // 2
    elif align == "右對齊": x -= tw

    if italic:
        # 斜體需透過矩陣變幻 (Affine Transform)
        padding = 60
        txt_img = Image.new("RGBA", (int(tw * 1.5) + padding, int(th * 2) + padding), (255, 255, 255, 0))
        d_txt = ImageDraw.Draw(txt_img)
        
        # 繪製文字 (包含模擬粗體)
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                d_txt.text((padding//2+dx, padding//2+dy), text, font=font, fill=color)
        d_txt.text((padding//2, padding//2), text, font=font, fill=color)
        
        # 物理斜體變換：矩陣 (1, 0.3, -offset, 0, 1, 0)
        m = 0.3 
        txt_img = txt_img.transform(txt_img.size, Image.AFFINE, (1, m, -padding//2*m, 0, 1, 0))
        return (txt_img, (int(x - padding//2), int(y - padding//2)))
    else:
        # 非斜體：直接在畫布繪製
        if bold:
            for dx, dy in [(-1,-1), (1,1), (1,-1), (-1,1)]:
                draw.text((x + dx, y + dy), text, font=font, fill=color)
        draw.text((x, y), text, font=font, fill=color)
        return None

# ==========================================
# 3. 主頁面：檔案上傳區
# ==========================================
st.title("✉️ 專業證書生成器 V5.5")

up1, up2 = st.columns(2)
with up1: bg_file = st.file_uploader("🖼️ 1. 上傳證書背景圖", type=["jpg", "png", "jpeg"], key="main_bg")
with up2: data_file = st.file_uploader("📊 2. 上傳資料檔 (Excel/CSV)", type=["xlsx", "csv"], key="main_data")

if not bg_file or not data_file:
    st.info("💡 操作提示：上傳檔案後，使用左側「側邊欄」調整參數。側邊欄邊界可用滑鼠拖動調整寬度。")
    st.stop()

# 載入背景圖與資料
bg_img = Image.open(bg_file).convert("RGBA")
W, H = bg_img.size
df = pd.read_excel(data_file) if data_file.name.endswith('xlsx') else pd.read_csv(data_file)

# ==========================================
# 4. 側邊欄：控制面板 (可滑鼠調整闊度)
# ==========================================
with st.sidebar:
    st.header("⚙️ 參數調整面板")
    
    # --- 配置管理 ---
    with st.expander("💾 設定存檔與載入"):
        if st.session_state.settings:
            js = json.dumps(st.session_state.settings, indent=4, ensure_ascii=False)
            st.download_button("📤 匯出設定 (JSON)", js, "cert_config.json", "application/json")
        uploaded_config = st.file_uploader("📥 載入舊設定", type=["json"])
        if uploaded_config:
            st.session_state.settings.update(json.load(uploaded_config))
            st.success("配置已載入")

    # --- 欄位選取與補全 ---
    display_cols = st.multiselect("選擇顯示欄位", df.columns, default=[df.columns[0]])
    
    for col in display_cols:
        if col not in st.session_state.settings:
            st.session_state.settings[col] = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
        else:
            # 相容性補全 (確保舊 JSON 有新功能參數)
            defaults = {"x": W//2, "y": H//2, "size": 60, "color": "#000000", "align": "居中", "bold": False, "italic": False}
            for k, v in defaults.items():
                if k not in st.session_state.settings[col]: st.session_state.settings[col][k] = v

    st.divider()

    # --- Photoshop 批量工具 (修正自動返回問題) ---
    with st.expander("🔗 Photoshop 批量連結工具", expanded=True):
        st.session_state.linked_layers = st.multiselect("選取要同時移動的欄位", display_cols)
        lc1, lc2 = st.columns(2)
        with lc1: batch_x = st.number_input("左右位移 (px)", value=0, key="batch_move_x")
        with lc2: batch_y = st.number_input("上下位移 (px)", value=0, key="batch_move_y")
        batch_s = st.number_input("字體縮放", value=0, key="batch_zoom_s")
        
        if st.button("✅ 執行批量套用", use_container_width=True):
            for c in st.session_state.linked_layers:
                # 1. 更新資料字典
                st.session_state.settings[c]["x"] += batch_x
                st.session_state.settings[c]["y"] += batch_y
                st.session_state.settings[c]["size"] += batch_s
                
                # 2. 同步更新組件 Key 狀態，防止 Slider 在頁面重整後跳回舊值
                st.session_state[f"x_{c}"] = float(st.session_state.settings[c]["x"])
                st.session_state[f"y_{c}"] = float(st.session_state.settings[c]["y"])
                st.session_state[f"s_{c}"] = int(st.session_state.settings[c]["size"])
            
            st.success("批量修改已生效")
            st.rerun()

    st.divider()

    # --- 個別圖層屬性 ---
    st.subheader("📝 單獨圖層設定")
    for col in display_cols:
        link_tag = " (🔗)" if col in st.session_state.linked_layers else ""
        with st.expander(f"圖層：{col}{link_tag}"):
            s = st.session_state.settings[col]
            # 綁定 Key 並與 settings 字典同步
            s["x"] = st.slider(f"X 座標", 0, W, float(s["x"]), key=f"x_{col}")
            s["y"] = st.slider(f"Y 座標", 0, H, float(s["y"]), key=f"y_{col}")
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
with p1: 
    id_col = st.selectbox("命名依據欄位 (識別證書檔案名稱)", df.columns)
with p2:
    all_names = df[id_col].astype(str).tolist()
    sel_names = st.multiselect("預覽對象選取", all_names, default=all_names[:1])
    target_df = df[df[id_col].astype(str).isin(sel_names)]

st.subheader("👁️ 畫布即時預覽")
# 視覺縮放滑桿 (範圍 50% - 250%)
zoom_lvl = st.slider("🔍 畫布視覺縮放 (%)", 50, 250, 100, step=10, key="main_zoom_slider")

if not target_df.empty:
    row_data = target_df.iloc[0]
    preview_canvas = bg_img.copy()
    draw_obj = ImageDraw.Draw(preview_canvas)
    
    for col in display_cols:
        set_val = st.session_state.settings[col]
        font_obj = get_font_object(set_val["size"])
        text_content = str(row_data[col])
        
        # 繪製文字
        render_res = draw_styled_text(
            draw_obj, text_content, (set_val["x"], set_val["y"]), 
            font_obj, set_val["color"], set_val["align"], 
            set_val["bold"], set_val["italic"]
        )
        
        # 若為斜體層則疊加
        if render_res:
            preview_canvas.alpha_composite(render_res[0], dest=render_res[1])
        
        # 繪製十字輔助線 (紅色代表已連結)
        guide_color = "#FF0000BB" if col in st.session_state.linked_layers else "#0000FF44"
        draw_obj.line([(0, set_val["y"]), (W, set_val["y"])], fill=guide_color, width=2)
        draw_obj.line([(set_val["x"], 0), (set_val["x"], H)], fill=guide_color, width=2)

    # 渲染預覽圖
    st.image(preview_canvas, width=int(W * (zoom_lvl / 100)))

# ==========================================
# 6. 生成與導出功能
# ==========================================
st.divider()
if st.button("🚀 開始批量製作所有選定證書", type="primary", use_container_width=True):
    if target_df.empty:
        st.warning("請先在上方『選取對象』中選擇名單")
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            prog_bar = st.progress(0)
            status = st.empty()
            
            for idx, (i, row) in enumerate(target_df.iterrows()):
                status.text(f"製作進度: {idx+1}/{len(target_df)} - {row[id_col]}")
                
                final_img = bg_img.copy()
                d_final = ImageDraw.Draw(final_img)
                
                for col in display_cols:
                    s_final = st.session_state.settings[col]
                    f_final = get_font_object(s_final["size"])
                    res_final = draw_styled_text(
                        d_final, str(row[col]), (s_final["x"], s_final["y"]), 
                        f_final, s_final["color"], s_final["align"], 
                        s_final["bold"], s_final["italic"]
                    )
                    if res_final:
                        final_img.alpha_composite(res_final[0], dest=res_final[1])
                
                # 轉為 JPEG 並存入 ZIP
                img_io = io.BytesIO()
                final_img.convert("RGB").save(img_io, format="JPEG", quality=95)
                zf.writestr(f"{str(row[id_col])}.jpg", img_io.getvalue())
                prog_bar.progress((idx + 1) / len(target_df))
            
            status.text("✅ 生成完成！")
        
        st.download_button(
            "📥 下載證書打包檔 (ZIP)", 
            zip_buffer.getvalue(), 
            "certificates_pack.zip", 
            "application/zip", 
            use_container_width=True
        )
        st.balloons()
