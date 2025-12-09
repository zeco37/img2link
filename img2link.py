# app.py
import io, os, re, zipfile, posixpath, warnings, csv, base64, datetime
from pathlib import Path
from typing import Optional, Dict, Tuple, List
from xml.etree import ElementTree as ET

import streamlit as st
import pandas as pd
import cloudinary
import cloudinary.uploader
import requests
from PIL import Image

# ───────────────────────────────────────────────────────────────────────────────
# 1) Brand & look (CHANGE THESE)
# ───────────────────────────────────────────────────────────────────────────────
AUTHOR_NAME  = "Zakaria Belalioui"           # <- change
COMPANY_NAME = "Ora Technologies"        # <- change
COMPANY_URL  = "https://www.kooul.ma/"                    # optional, e.g. "https://yourcompany.com"

# Logo for favicon + header (local path OR URL). Keep PNG (256x256) if possible.
LOGO_SOURCE  = "assets/attachment-clip-page-paper-icon-vector-design-png_117669.jpg"     # e.g. "https://res.cloudinary.com/xxx/image/upload/logo.png"

# Background image (URL is best; you can also embed a local image)
BACKGROUND_URL = (
    "https://res.cloudinary.com/dqye9uju0/image/upload/v1758554635/NsvG1713971804597-Artboard20220copy100_gocy5z.jpg"
)

# ───────────────────────────────────────────────────────────────────────────────
# 2) Page, favicon, background & CSS
# ───────────────────────────────────────────────────────────────────────────────
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet.header_footer")

def _pil_icon_from_source(src: str) -> Optional[Image.Image]:
    try:
        if src.startswith("http"):
            data = requests.get(src, timeout=15).content
            return Image.open(io.BytesIO(data))
        else:
            return Image.open(src)
    except Exception:
        return None

def _data_uri_from_source(src: str) -> Optional[str]:
    try:
        if src.startswith("http"):
            data = requests.get(src, timeout=15).content
        else:
            with open(src, "rb") as f:
                data = f.read()
        b64 = base64.b64encode(data).decode()
        mime = "image/png" if src.lower().endswith(".png") else "image/jpeg"
        return f"data:{mime};base64,{b64}"
    except Exception:
        return None

ICON = _pil_icon_from_source(LOGO_SOURCE)
st.set_page_config(
    page_title="Image → Link Converter",
    page_icon=ICON if ICON else "🔗",
    layout="centered",
)

def set_background_from_url(url: str, dark_overlay: float = 0.35):
    st.markdown(f"""
        <style>
        .stApp {{
            background:
              linear-gradient(rgba(0,0,0,{dark_overlay}), rgba(0,0,0,{dark_overlay})),
              url('{url}');
            background-size: cover;
            background-attachment: fixed;
            background-position: center center;
        }}
        </style>
    """, unsafe_allow_html=True)

set_background_from_url(BACKGROUND_URL, dark_overlay=0.35)

# Global polish
st.markdown("""
<style>
.main .block-container { max-width: 1120px; }
section.main > div.block-container {
  background: rgba(255,255,255,0.90);
  border-radius: 20px;
  padding: 26px 28px;
  box-shadow: 0 12px 40px rgba(0,0,0,.18);
  backdrop-filter: blur(6px);
}
div[data-baseweb="tab-list"] button[role="tab"] {
  border-radius: 999px !important;
  padding: 6px 14px !important;
  margin-right: 6px;
  background: rgba(255,255,255,.65);
  border: 1px solid rgba(0,0,0,.06);
}
div[data-baseweb="tab-list"] button[aria-selected="true"] {
  background: #6D28D9; color: #fff; border-color: #6D28D9;
}
.stButton > button, .stDownloadButton > button {
  border-radius: 12px; padding: 10px 16px; font-weight: 600;
  box-shadow: 0 6px 20px rgba(109,40,217,.18);
}
.stNumberInput input, .stTextInput input, .stSelectbox [data-baseweb="select"] {
  border-radius: 10px;
}
.stFileUploader > section { border-radius: 16px; background: rgba(255,255,255,.75); }

/* Brand header */
.brand-row {display:flex; align-items:center; gap:12px; margin:-4px 0 10px;}
.brand-row .brand-logo {width:40px; height:40px; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,.12);}
.brand-title {font-size:26px; font-weight:800; letter-spacing:.2px;}
.brand-by {color:#6366F1; font-weight:600; font-size:12px;}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Brand header (logo + title + your name/company)
logo_data_uri = _data_uri_from_source(LOGO_SOURCE)
logo_html = f'<img class="brand-logo" src="{logo_data_uri}"/>' if logo_data_uri else ""

company_html = (
    f' · <a href="{COMPANY_URL}" target="_blank" style="color:#4338CA;text-decoration:none;">{COMPANY_NAME}</a>'
    if COMPANY_URL else f" · {COMPANY_NAME}" if COMPANY_NAME else ""
)
st.markdown(f"""
<div class="brand-row">
  {logo_html}
  <div>
    <div class="brand-title">Image → Link Converter</div>
    <div class="brand-by">by <strong>{AUTHOR_NAME}</strong>{company_html}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ───────────────────────────────────────────────────────────────────────────────
# 3) Cloudinary config & common helpers
# ───────────────────────────────────────────────────────────────────────────────
URL_RE = re.compile(r'https?://[^\s")]+', re.IGNORECASE)

def init_cloudinary_from_secrets() -> bool:
    cloudinary.config(
        cloud_name=st.secrets.get("CLOUDINARY_CLOUD_NAME"),
        api_key=st.secrets.get("CLOUDINARY_API_KEY"),
        api_secret=st.secrets.get("CLOUDINARY_API_SECRET"),
        secure=True,
    )
    os.environ["FOLDER"] = st.secrets.get("FOLDER", "Products")
    cfg = cloudinary.config()
    if not (cfg.cloud_name and cfg.api_key and cfg.api_secret):
        st.error(
            "Cloudinary secrets missing. Add to Streamlit **Secrets** (TOML):\n\n"
            'CLOUDINARY_CLOUD_NAME = "your_cloud"\n'
            'CLOUDINARY_API_KEY    = "your_key"\n'
            'CLOUDINARY_API_SECRET = "your_secret"\n'
            'FOLDER                = "Products"\n'
        )
        return False
    return True

def upload_bytes_to_cloudinary(data: bytes, filename: str) -> str:
    res = cloudinary.uploader.upload(
        io.BytesIO(data),
        resource_type="image",
        public_id=None,
        folder=(os.getenv("FOLDER") or None),
        filename=filename,
        use_filename=True,
        unique_filename=True,
        overwrite=False,
    )
    return res.get("secure_url")

def clean_filename(s: str) -> str:
    s = (s or "image").strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:120] or "image"

def extract_url_from_cell(val: str) -> Optional[str]:
    if not val: return None
    m = re.search(r'=IMAGE\(\s*"([^"]+)"', str(val), re.IGNORECASE)  # Google Sheets IMAGE()
    if m: return m.group(1)
    m = URL_RE.search(str(val))
    if m: return m.group(0)
    return None

def norm(s: str) -> str: return re.sub(r'[^a-z0-9]+', '', (s or '').casefold())

def find_col_idx(headers: List[str], primary: List[str], synonyms: List[str]) -> Optional[int]:
    H = [norm(h) for h in headers]
    for i, h in enumerate(H):
        for p in primary:
            if H[i].startswith(norm(p)): return i
    for i, h in enumerate(H):
        for p in primary + synonyms:
            if norm(p) in h: return i
    return None

def auto_detect_columns(df: pd.DataFrame,
                        name_hint: Optional[str],
                        img_hint: Optional[str]) -> Tuple[int, int]:
    headers = [str(c) for c in df.columns]
    name_primary = [name_hint] if name_hint else ["produit","products","product","pruducts","name","le nom"]
    img_primary  = [img_hint]  if img_hint  else ["photo","image","images","img","picture"]
    name_syn = ["libelle","designation","article","titre","item"]
    img_syn  = ["link","imagelink","imageurl","photourl","lienimage"]
    n_idx = find_col_idx(headers, name_primary, name_syn)
    i_idx = find_col_idx(headers, img_primary, img_syn)
    if n_idx is None or i_idx is None:
        raise ValueError(f"Auto-detect failed. Headers: {headers}")
    return n_idx, i_idx

# ───────────────────────────────────────────────────────────────────────────────
# 4) XLSX drawing XML mapping (strict / any-col / smart)
# ───────────────────────────────────────────────────────────────────────────────
NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

def _et_from_zip(z: zipfile.ZipFile, path: str):
    try: return ET.fromstring(z.read(path))
    except KeyError: return None

def _rels_map(z: zipfile.ZipFile, rels_path: str) -> Dict[str,str]:
    root = _et_from_zip(z, rels_path)
    if root is None: return {}
    m = {}
    for rel in root:
        if rel.tag.endswith("Relationship"):
            m[rel.attrib["Id"]] = rel.attrib["Target"]
    return m

def embedded_images_by_row_via_xml(xlsx_bytes: bytes, sheet_name: str, image_col_idx_1b: int) -> Dict[int, Tuple[str, bytes]]:
    mapping: Dict[int, Tuple[str, bytes]] = {}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as z:
        wb = _et_from_zip(z, "xl/workbook.xml")
        if wb is None: return mapping
        wb_rels = _rels_map(z, "xl/_rels/workbook.xml.rels")

        sheet_path = None
        for s in wb.findall("{*}sheets/{*}sheet"):
            if s.attrib.get("name") == sheet_name:
                rid = s.attrib.get(f"{{{NS_R}}}id")
                tgt = wb_rels.get(rid)
                if tgt: sheet_path = posixpath.normpath(posixpath.join("xl", tgt))
                break
        if not sheet_path: return mapping

        srels = _rels_map(z, posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path)+".rels"))
        drawing_target = None
        for _, t in srels.items():
            if "drawing" in t:
                drawing_target = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), t))
                break
        if not drawing_target: return mapping
        if not drawing_target.startswith("xl/"):
            drawing_target = posixpath.normpath(posixpath.join("xl", drawing_target))

        dr = _et_from_zip(z, drawing_target)
        if dr is None: return mapping
        drrels = _rels_map(z, posixpath.join(posixpath.dirname(drawing_target), "_rels", posixpath.basename(drawing_target)+".rels"))

        anchors = list(dr.findall(f"{{{NS_XDR}}}twoCellAnchor")) + list(dr.findall(f"{{{NS_XDR}}}oneCellAnchor"))
        for a in anchors:
            frm = a.find(f"{{{NS_XDR}}}from")
            if frm is None: continue
            c = frm.find(f"{{{NS_XDR}}}col"); r = frm.find(f"{{{NS_XDR}}}row")
            if c is None or r is None: continue
            col_1b = int(c.text) + 1
            row_1b = int(r.text) + 1
            if col_1b != image_col_idx_1b: continue

            pic = a.find(f"{{{NS_XDR}}}pic")
            if pic is None: continue
            blip = pic.find(f"{{{NS_XDR}}}blipFill/{{{NS_A}}}blip")
            if blip is None: continue
            rid = blip.attrib.get(f"{{{NS_R}}}embed")
            if not rid: continue
            tgt = drrels.get(rid);  # path in drawing rels
            if not tgt: continue

            media_path = posixpath.normpath(posixpath.join(posixpath.dirname(drawing_target), tgt))
            if not media_path.startswith("xl/"):
                media_path = posixpath.normpath(posixpath.join("xl/drawings", tgt)).replace("drawings/../","")
            try:
                data = z.read(media_path)
            except KeyError:
                continue
            mapping[row_1b] = (Path(media_path).name, data)
    return mapping

def embedded_images_by_row_anycol(xlsx_bytes: bytes, sheet_name: str) -> Dict[int, Tuple[str, bytes]]:
    mapping: Dict[int, Tuple[str, bytes]] = {}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as z:
        wb = _et_from_zip(z, "xl/workbook.xml")
        if wb is None: return mapping
        wb_rels = _rels_map(z, "xl/_rels/workbook.xml.rels")
        sheet_path = None
        for s in wb.findall("{*}sheets/{*}sheet"):
            if s.attrib.get("name") == sheet_name:
                rid = s.attrib.get(f"{{{NS_R}}}id")
                tgt = wb_rels.get(rid)
                if tgt: sheet_path = posixpath.normpath(posixpath.join("xl", tgt))
                break
        if not sheet_path: return mapping
        srels = _rels_map(z, posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path)+".rels"))
        drawing_target = None
        for _, t in srels.items():
            if "drawing" in t:
                drawing_target = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), t))
                break
        if not drawing_target: return mapping
        if not drawing_target.startswith("xl/"): drawing_target = posixpath.normpath(posixpath.join("xl", drawing_target))
        dr = _et_from_zip(z, drawing_target)
        if dr is None: return mapping
        drrels = _rels_map(z, posixpath.join(posixpath.dirname(drawing_target), "_rels", posixpath.basename(drawing_target)+".rels"))
        anchors = list(dr.findall(f"{{{NS_XDR}}}twoCellAnchor")) + list(dr.findall(f"{{{NS_XDR}}}oneCellAnchor"))
        for a in anchors:
            frm = a.find(f"{{{NS_XDR}}}from");  r = frm.find(f"{{{NS_XDR}}}row") if frm is not None else None
            if r is None: continue
            row_1b = int(r.text) + 1
            pic = a.find(f"{{{NS_XDR}}}pic");  blip = pic.find(f"{{{NS_XDR}}}blipFill/{{{NS_A}}}blip") if pic is not None else None
            if blip is None: continue
            rid = blip.attrib.get(f"{{{NS_R}}}embed");  tgt = drrels.get(rid) if rid else None
            if not tgt: continue
            media_path = posixpath.normpath(posixpath.join(posixpath.dirname(drawing_target), tgt))
            if not media_path.startswith("xl/"): media_path = posixpath.normpath(posixpath.join("xl/drawings", tgt)).replace("drawings/../","")
            try: data = z.read(media_path)
            except KeyError: continue
            mapping[row_1b] = (Path(media_path).name, data)
    return mapping

def embedded_images_by_row_smart(xlsx_bytes: bytes, sheet_name: str,
                                 data_start_1b: int, data_end_1b: int,
                                 prefer_col_1b: Optional[int]) -> Dict[int, Tuple[str, bytes]]:
    out: Dict[int, Tuple[str, bytes]] = {}
    tmp: Dict[int, Tuple[str, bytes, int]] = {}
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as z:
        wb = _et_from_zip(z, "xl/workbook.xml")
        if wb is None: return out
        wb_rels = _rels_map(z, "xl/_rels/workbook.xml.rels")
        sheet_path = None
        for s in wb.findall("{*}sheets/{*}sheet"):
            if s.attrib.get("name") == sheet_name:
                rid = s.attrib.get(f"{{{NS_R}}}id")
                tgt = wb_rels.get(rid)
                if tgt: sheet_path = posixpath.normpath(posixpath.join("xl", tgt))
                break
        if not sheet_path: return out
        srels = _rels_map(z, posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path)+".rels"))
        drawing_target = None
        for _, t in srels.items():
            if "drawing" in t: drawing_target = posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), t)); break
        if not drawing_target: return out
        if not drawing_target.startswith("xl/"): drawing_target = posixpath.normpath(posixpath.join("xl", drawing_target))
        dr = _et_from_zip(z, drawing_target)
        if dr is None: return out
        drrels = _rels_map(z, posixpath.join(posixpath.dirname(drawing_target), "_rels", posixpath.basename(drawing_target)+".rels"))
        anchors = list(dr.findall(f"{{{NS_XDR}}}twoCellAnchor")) + list(dr.findall(f"{{{NS_XDR}}}oneCellAnchor"))
        for a in anchors:
            frm = a.find(f"{{{NS_XDR}}}from");  fr = frm.find(f"{{{NS_XDR}}}row") if frm is not None else None
            fc = frm.find(f"{{{NS_XDR}}}col") if frm is not None else None
            if fr is None or fc is None: continue
            fr0 = int(fr.text); fc0 = int(fc.text)
            to = a.find(f"{{{NS_XDR}}}to")
            if to is not None:
                tr = to.find(f"{{{NS_XDR}}}row"); tc = to.find(f"{{{NS_XDR}}}col")
                tr0 = int(tr.text) if tr is not None else fr0
                tc0 = int(tc.text) if tc is not None else fc0
            else:
                tr0, tc0 = fr0, fc0
            center_row_1b = int(round((fr0 + tr0) / 2.0)) + 1
            center_col_1b = int(round((fc0 + tc0) / 2.0)) + 1
            snap_row_1b   = min(max(center_row_1b, data_start_1b), data_end_1b)
            pic = a.find(f"{{{NS_XDR}}}pic"); blip = pic.find(f"{{{NS_XDR}}}blipFill/{{{NS_A}}}blip") if pic is not None else None
            if blip is None: continue
            rid = blip.attrib.get(f"{{{NS_R}}}embed"); tgt = drrels.get(rid) if rid else None
            if not tgt: continue
            media_path = posixpath.normpath(posixpath.join(posixpath.dirname(drawing_target), tgt))
            if not media_path.startswith("xl/"): media_path = posixpath.normpath(posixpath.join("xl/drawings", tgt)).replace("drawings/../","")
            try: data = z.read(media_path)
            except KeyError: continue
            if snap_row_1b not in tmp:
                tmp[snap_row_1b] = (Path(media_path).name, data, center_col_1b)
            else:
                if prefer_col_1b is not None:
                    _, _, old_cc = tmp[snap_row_1b]
                    if abs(center_col_1b - prefer_col_1b) < abs(old_cc - prefer_col_1b):
                        tmp[snap_row_1b] = (Path(media_path).name, data, center_col_1b)
    for r, (name, data, _) in tmp.items():
        out[r] = (name, data)
    return out

# ───────────────────────────────────────────────────────────────────────────────
# 5) CSV helper + image processing (background + upscaling)
# ───────────────────────────────────────────────────────────────────────────────
def read_csv_safely(uploaded_file, enc_choice: str, delim_choice: str):
    raw = uploaded_file.getvalue()
    enc_candidates = ["utf-8-sig", "utf-8", "cp1252", "latin1"] if enc_choice == "auto" else [enc_choice]
    if delim_choice == "auto":
        try:
            sample = raw[:4096].decode("utf-8", errors="ignore")
            sniffed = csv.Sniffer().sniff(sample).delimiter
        except Exception:
            sniffed = ","
        delim_candidates = [sniffed, ",", ";", "|", "\t"]
    else:
        delim_candidates = ["\t" if delim_choice == "\\t" else delim_choice]

    last_err = None
    for enc in enc_candidates:
        for delim in delim_candidates:
            try:
                buf = io.BytesIO(raw)
                df = pd.read_csv(buf, dtype=str, keep_default_na=False,
                                 encoding=enc, sep=delim, engine="python")
                return df, enc, delim
            except Exception as e:
                last_err = e
    raise last_err or RuntimeError("CSV parse failed")

def _get_rembg_remove():
    try:
        from rembg import remove as _remove
        return _remove
    except Exception:
        return None

def _ensure_mode(img: Image.Image, want_alpha: bool) -> Image.Image:
    return img.convert("RGBA" if want_alpha else "RGB")

def _upscale(img: Image.Image, up_mode: str, w: int, h: int, transparent: bool) -> Image.Image:
    if up_mode == "none": return img
    W, H = img.size
    if up_mode == "scale_min":
        s = max(w / W if W < w else 1.0, h / H if H < h else 1.0)
        return img.resize((int(round(W*s)), int(round(H*s))), Image.LANCZOS) if s > 1 else img
    if up_mode in ("pad_box", "fill_box"):
        s = min(w/W, h/H) if up_mode == "pad_box" else max(w/W, h/H)
        new = (int(round(W*s)), int(round(H*s)))
        scaled = img.resize(new, Image.LANCZOS)
        if up_mode == "pad_box":
            canvas = Image.new("RGBA" if transparent else "RGB", (w,h), (0,0,0,0) if transparent else (255,255,255))
            ox = (w - new[0]) // 2; oy = (h - new[1]) // 2
            if scaled.mode == "RGBA" and transparent: canvas.paste(scaled, (ox,oy), scaled)
            else: canvas.paste(scaled, (ox,oy))
            return canvas
        left = (new[0]-w)//2; top = (new[1]-h)//2
        return scaled.crop((left, top, left+w, top+h))
    return img

def process_image_bytes(data: bytes, bg_mode: str, up_mode: str, up_w: int, up_h: int) -> Tuple[bytes, str]:
    try: base = Image.open(io.BytesIO(data))
    except Exception: return data, ".jpg"
    rembg_remove = _get_rembg_remove()
    if bg_mode == "remove" and rembg_remove is not None:
        try:
            removed = rembg_remove(data)
            base = Image.open(io.BytesIO(removed)).convert("RGBA")
        except Exception:
            base = base.convert("RGBA")
    elif bg_mode == "white":
        base = base.convert("RGBA")
        bg = Image.new("RGBA", base.size, (255,255,255,255))
        bg.paste(base, mask=base.split()[-1] if base.mode == "RGBA" else None)
        base = bg.convert("RGB")
    else:
        base = base.convert("RGBA" if base.mode == "RGBA" else "RGB")
    transparent = (bg_mode == "remove")
    base = _ensure_mode(base, transparent)
    base = _upscale(base, up_mode, up_w, up_h, transparent)
    buf = io.BytesIO()
    if bg_mode == "remove":
        base.save(buf, format="PNG");  return buf.getvalue(), ".png"
    else:
        base.convert("RGB").save(buf, format="JPEG", quality=95, optimize=True); return buf.getvalue(), ".jpg"

def get_bytes_from_url(url: str) -> Optional[bytes]:
    try:
        r = requests.get(url, timeout=20)
        if r.ok: return r.content
    except Exception:
        pass
    return None

# ───────────────────────────────────────────────────────────────────────────────
# 6) UI + logic
# ───────────────────────────────────────────────────────────────────────────────
st.caption("Convert tables (XLSX/CSV) that contain images to public links, or upload a folder/ZIP of images and get a link mapping.")

if not init_cloudinary_from_secrets():
    st.stop()

bg_choice = st.selectbox(
    "Background",
    ["No change", "Remove background (transparent PNG)", "Set white background (JPG)"],
    index=0
)
bg_mode = {"No change":"none","Remove background (transparent PNG)":"remove","Set white background (JPG)":"white"}[bg_choice]

up_choice = st.selectbox(
    "Upscale (agrandissage)",
    ["No upscale", "Scale up if smaller (keep ratio)", "Pad to box (no distortion)", "Fill & crop to box"],
    index=0
)
up_mode = {"No upscale":"none","Scale up if smaller (keep ratio)":"scale_min","Pad to box (no distortion)":"pad_box","Fill & crop to box":"fill_box"}[up_choice]

c1, c2 = st.columns(2)
with c1: up_w = st.number_input("Target width (px)", 64, 8000, 1200, 50)
with c2: up_h = st.number_input("Target height (px)", 64, 8000, 1200, 50)

tab_table, tab_images, tab_tool2 = st.tabs([
    "📄 Table (.xlsx/.csv)", 
    "📁 Images / ZIP",
    "📥 Download Images from CSV"
])

# ───────────────────────────────────────────────────────────────────────────────
# 7) TOOL 3 : DOWNLOAD IMAGES FROM CSV → ZIP (UPDATED FINAL)
# ───────────────────────────────────────────────────────────────────────────────
with tab_tool2:
    st.header("📥 Download Images From CSV/XLSX")
    st.caption("Upload a CSV or Excel file. The tool will read product names and image URLs, download all images, and generate a ZIP file.")

    file2 = st.file_uploader("Upload CSV or XLSX", type=["csv", "xlsx"], key="tool3_upload")

    if file2 is not None:
        suffix = Path(file2.name).suffix.lower()

        # ---- READ FILE ----
        if suffix == ".csv":
            try:
                df2 = pd.read_csv(file2, dtype=str).fillna("")
            except Exception as e:
                st.error(f"❌ Could not read CSV: {e}")
                st.stop()

        elif suffix == ".xlsx":
            try:
                xls = pd.ExcelFile(file2)
                sheet = st.selectbox("Select sheet", xls.sheet_names)
                df2 = pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna("")
            except Exception as e:
                st.error(f"❌ Could not read XLSX: {e}")
                st.stop()

        st.write("📌 Columns detected:", list(df2.columns))

        # ---- COLUMN AUTODETECT ----
        auto_product = None
        auto_url = None

        for col in df2.columns:
            low = col.lower()
            if any(x in low for x in ["product", "name", "title"]):
                auto_product = col
            if any(x in low for x in ["url", "image", "img", "link"]):
                auto_url = col

        product_col = st.selectbox("Select product column", df2.columns, index=df2.columns.get_loc(auto_product) if auto_product else 0)
        url_col = st.selectbox("Select image URL column", df2.columns, index=df2.columns.get_loc(auto_url) if auto_url else 0)

        # ---- OUTPUT DIR ----
        output_dir = "downloaded_images"
        os.makedirs(output_dir, exist_ok=True)

        # ---- SANITIZE ----
        def sanitize(name):
            return re.sub(r'[<>:"/\\|?*]', "_", name).strip()

        # ---- START ----
        if st.button("⬇️ Download Images & Generate ZIP"):
            downloaded = []
            prog = st.progress(0, text="Downloading images...")

            for i, row in df2.iterrows():
                name = str(row[product_col]).strip()
                raw_cell = str(row[url_col]).strip()

                # Extract URL using robust extractor
                url = extract_url_from_cell(raw_cell)

                if not url:
                    print("SKIPPED:", raw_cell)
                    continue

                try:
                    res = requests.get(url, timeout=20)
                    if res.ok:
                        img = Image.open(BytesIO(res.content))
                        if img.mode == "RGBA":
                            img = img.convert("RGB")

                        fname = sanitize(name) + ".jpg"
                        fpath = os.path.join(output_dir, fname)
                        img.save(fpath, "JPEG", quality=95)
                        downloaded.append(fpath)
                except Exception as e:
                    print("ERROR:", url, e)

                prog.progress(int((i + 1) / len(df2) * 100))

            # ---- ZIP ----
            zipname = "images_from_file.zip"
            with zipfile.ZipFile(zipname, "w") as z:
                for f in downloaded:
                    z.write(f, os.path.basename(f))

            st.success(f"🎉 Done! {len(downloaded)} images downloaded.")
            with open(zipname, "rb") as f:
                st.download_button("⬇️ Download ZIP", f, file_name=zipname)


# ── Table tab
with tab_table:
    table_file = st.file_uploader("Upload a .xlsx or .csv", type=["xlsx","csv"], key="table_uploader")
    if table_file is not None:
        suffix = Path(table_file.name).suffix.lower()
        # XLSX
        if suffix == ".xlsx":
            data = table_file.getvalue()
            xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
            sheet = st.selectbox("Sheet", xls.sheet_names, index=0)
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str, engine="openpyxl").fillna("")
            headers = [str(c) for c in df.columns]
            st.write("Detected headers:", headers)

            auto = st.checkbox("Auto-detect columns", value=False)
            i_idx_in = st.number_input("Image column index (1-based; 0 = auto)", 0, len(headers), min(3, len(headers)))
            n_idx_in = st.number_input("Product/Name column index (1-based; 0 = auto)", 0, len(headers), min(4, len(headers)))
            header_row = st.number_input("Header row number (1-based)", 1, 9999, 1)

            name_hint = st.text_input("Name header hint (optional)", "")
            img_hint  = st.text_input("Image header hint (optional)", "")

            n_idx = (n_idx_in - 1) if n_idx_in > 0 else None
            i_idx = (i_idx_in - 1) if i_idx_in > 0 else None
            if auto or n_idx is None or i_idx is None:
                n_auto, i_auto = auto_detect_columns(df, name_hint or None, img_hint or None)
                if n_idx is None: n_idx = n_auto
                if i_idx is None: i_idx = i_auto

            accept_any_col = st.checkbox("Accept any picture anchored on the row (ignore image column)", value=False)
            smart_snap     = st.checkbox("Smart row snap (use image vertical center)", value=True)

            if st.button("Convert (XLSX)", key="convert_xlsx"):
                data_start = header_row + 1
                data_end   = header_row + len(df)
                if smart_snap:
                    embedded = embedded_images_by_row_smart(
                        data, sheet, data_start, data_end,
                        None if accept_any_col else (i_idx + 1)
                    )
                else:
                    embedded = embedded_images_by_row_anycol(data, sheet) if accept_any_col \
                               else embedded_images_by_row_via_xml(data, sheet, i_idx + 1)

                matched = sum(1 for r in range(len(df)) if (header_row + 1 + r) in embedded)
                st.caption(f"Embedded pictures matched to rows: {matched}/{len(df)}")

                links = []
                prog = st.progress(0, text="Uploading…")
                for r in range(len(df)):
                    product = str(df.iat[r, n_idx] or "").strip()
                    img_txt = str(df.iat[r, i_idx] or "").strip()
                    excel_row_1b = header_row + 1 + r  # strict row mapping

                    emb_name, emb_bytes = embedded.get(excel_row_1b, (None, None))
                    url_in_cell = extract_url_from_cell(img_txt)

                    if url_in_cell:
                        raw = get_bytes_from_url(url_in_cell) if (bg_mode != "none" or up_mode != "none") else None
                        if raw:
                            processed, ext = process_image_bytes(raw, bg_mode, up_mode, up_w, up_h)
                            fname = clean_filename(product or "image") + ext
                            link = upload_bytes_to_cloudinary(processed, fname)
                        else:
                            link = url_in_cell
                    elif emb_bytes:
                        processed, ext = process_image_bytes(emb_bytes, bg_mode, up_mode, up_w, up_h)
                        base = Path(emb_name or "image").stem
                        fname = clean_filename(product or base) + ext
                        link = upload_bytes_to_cloudinary(processed, fname)
                    else:
                        link = ""  # keep alignment
                    links.append(link)
                    prog.progress(int((r + 1) / len(df) * 100))

                df.insert(i_idx + 1, "Image Link", links)

                out_xlsx = io.BytesIO()
                with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
                    df.to_excel(w, sheet_name=sheet, index=False)
                st.success("Done.")
                st.download_button("⬇️ Download XLSX", data=out_xlsx.getvalue(),
                                   file_name=Path(table_file.name).stem + "_with_links.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                out_csv = df.to_csv(index=False).encode("utf-8")
                st.download_button("⬇️ Download CSV", data=out_csv,
                                   file_name=Path(table_file.name).stem + "_with_links.csv",
                                   mime="text/csv")
        # CSV
        else:
            enc   = st.selectbox("CSV encoding",  ["auto","utf-8","utf-8-sig","cp1252","latin1"], 0, key="csv_enc")
            delim = st.selectbox("CSV delimiter", ["auto",",",";","|","\\t"], 0, key="csv_delim")
            try:
                df, used_enc, used_delim = read_csv_safely(table_file, enc, delim)
                st.caption(f"Parsed with **{used_enc}** and delimiter **{repr(used_delim)}**")
            except Exception as e:
                st.error(f"Could not read CSV: {e}")
                df = None

            if df is not None:
                headers = [str(c) for c in df.columns]
                st.write("Detected headers:", headers)

                auto = st.checkbox("Auto-detect columns", value=True, key="auto_csv")
                name_hint = st.text_input("Name header hint (optional)", "", key="name_hint_csv")
                img_hint  = st.text_input("Image header hint (optional)", "", key="img_hint_csv")

                n_idx_in = st.number_input("Product/Name column (1-based; 0 = auto)", 0, len(headers), 0 if auto else min(2, len(headers)), key="n_idx_csv")
                i_idx_in = st.number_input("Image column (1-based; URLs; 0 = auto)",   0, len(headers), 0 if auto else min(1, len(headers)), key="i_idx_csv")

                n_idx = (n_idx_in - 1) if n_idx_in > 0 else None
                i_idx = (i_idx_in - 1) if i_idx_in > 0 else None
                if auto or n_idx is None or i_idx is None:
                    try:
                        n_auto, i_auto = auto_detect_columns(df, name_hint or None, img_hint or None)
                        if n_idx is None: n_idx = n_auto
                        if i_idx is None: i_idx = i_auto
                    except Exception as e:
                        st.error(str(e))
                        n_idx, i_idx = 0, 0

                st.info("For CSV: the image column should be **URLs**. With background/size options, each URL is downloaded, processed, and re-uploaded to Cloudinary.")

                if st.button("Convert (CSV)", key="convert_csv"):
                    links = []; prog = st.progress(0, text="Processing…")
                    for r, row in enumerate(df.itertuples(index=False), 1):
                        product = str(row[n_idx] or "").strip()
                        img_txt = str(row[i_idx] or "").strip()
                        url = extract_url_from_cell(img_txt)
                        if url and (bg_mode != "none" or up_mode != "none"):
                            raw = get_bytes_from_url(url)
                            link = url
                            if raw:
                                processed, ext = process_image_bytes(raw, bg_mode, up_mode, up_w, up_h)
                                fname = clean_filename(product or "image") + ext
                                link = upload_bytes_to_cloudinary(processed, fname)
                        else:
                            link = url or ""
                        links.append(link); prog.progress(int(r/len(df)*100))
                    df.insert(i_idx + 1, "Image Link", links)

                    out_xlsx = io.BytesIO()
                    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
                        df.to_excel(w, index=False)
                    st.success("Done.")
                    st.download_button("⬇️ Download XLSX", data=out_xlsx.getvalue(),
                                       file_name=Path(table_file.name).stem + "_with_links.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    out_csv = df.to_csv(index=False).encode("utf-8")
                    st.download_button("⬇️ Download CSV", data=out_csv,
                                       file_name=Path(table_file.name).stem + "_with_links.csv",
                                       mime="text/csv")

# ── Images/ZIP tab
with tab_images:
    st.caption("Drop a **.zip** of your folder or select many images. (Browsers can’t read raw folders.)")
    IMG_TYPES = ["zip","jpg","jpeg","png","webp","bmp","tif","tiff","gif"]
    img_files = st.file_uploader("Drop a .zip OR select many images",
                                 type=IMG_TYPES, accept_multiple_files=True, key="images_uploader")

    name_source = st.selectbox(
        "Derive product name from",
        ["File name (without extension)", "ParentFolder/File name", "Full path inside ZIP (folders included)"],
        index=0, key="name_source",
    )

    def nice_label_from_path(p: str) -> str:
        stem = Path(p).stem
        s = re.sub(r"[/_\\-]+", " ", stem).strip()
        return re.sub(r"\s+", " ", s)

    if st.button("Upload images", key="upload_images") and img_files:
        to_process: List[Tuple[str, bytes]] = []
        for uf in img_files:
            if uf.name.lower().endswith(".zip"):
                try: z = zipfile.ZipFile(io.BytesIO(uf.getvalue()), "r")
                except Exception as e: st.error(f"Could not read ZIP {uf.name}: {e}"); continue
                for zi in z.infolist():
                    if zi.is_dir(): continue
                    if not re.search(r"\.(jpg|jpeg|png|webp|bmp|tif|tiff|gif)$", zi.filename, re.I): continue
                    try: data = z.read(zi)
                    except Exception: continue
                    if name_source == "File name (without extension)":
                        label = nice_label_from_path(zi.filename)
                    elif name_source == "ParentFolder/File name":
                        p = Path(zi.filename)
                        label = f"{p.parent.name}/{nice_label_from_path(zi.filename)}" if p.parent.name else nice_label_from_path(zi.filename)
                    else:
                        label = zi.filename.replace("\\", "/").lstrip("./")
                    to_process.append((label, data))
            else:
                data = uf.getvalue()
                label = nice_label_from_path(uf.name)
                to_process.append((label, data))

        if not to_process:
            st.warning("No images found in the files you provided.")
        else:
            results = []; prog = st.progress(0, text="Uploading images…")
            for i, (label, raw) in enumerate(to_process, 1):
                processed, ext = process_image_bytes(raw, bg_mode, up_mode, up_w, up_h)
                fname = clean_filename(label) + ext
                try: url = upload_bytes_to_cloudinary(processed, fname)
                except Exception as e: url = ""; st.error(f"Upload failed for {label}: {e}")
                results.append({"File": label, "Product": clean_filename(label), "Image Link": url})
                prog.progress(int(i/len(to_process)*100))

            df_map = pd.DataFrame(results, columns=["File","Product","Image Link"])
            st.success(f"Uploaded {df_map['Image Link'].astype(bool).sum()} / {len(df_map)} images.")
            st.dataframe(df_map, use_container_width=True)

            out_xlsx = io.BytesIO()
            with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
                df_map.to_excel(w, sheet_name="Links", index=False)
            st.download_button("⬇️ Download mapping (XLSX)", data=out_xlsx.getvalue(),
                               file_name="folder_links.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            out_csv = df_map.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download mapping (CSV)", data=out_csv,
                               file_name="folder_links.csv", mime="text/csv")

# Footer with your name/company
year = datetime.datetime.now().year
footer_company = (f' · <a href="{COMPANY_URL}" target="_blank"> {COMPANY_NAME}</a>'
                  if COMPANY_URL else f" · {COMPANY_NAME}" if COMPANY_NAME else "")
st.markdown(f"""
<style>.footer {{text-align:center; margin-top:28px; color:#475569; font-size:13px;}}
.footer a {{color:#6D28D9; text-decoration:none;}}</style>
<div class="footer">© {year} {AUTHOR_NAME}{footer_company}</div>
""", unsafe_allow_html=True)
