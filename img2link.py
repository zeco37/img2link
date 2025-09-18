# app.py
import io, os, re, zipfile, posixpath, warnings, csv
from pathlib import Path
from typing import Optional, Dict, Tuple, List
from xml.etree import ElementTree as ET

import streamlit as st
import pandas as pd
import cloudinary
import cloudinary.uploader
import requests
from PIL import Image

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet.header_footer")
st.set_page_config(page_title="Image → Link (Cloudinary)", page_icon="🔗", layout="centered")

URL_RE = re.compile(r'https?://[^\s")]+', re.IGNORECASE)

# ---------- Cloudinary ----------
def init_cloudinary_from_secrets():
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
            "Cloudinary secrets missing. Add TOML secrets like:\n\n"
            'CLOUDINARY_CLOUD_NAME = "your_cloud"\n'
            'CLOUDINARY_API_KEY    = "your_key"\n'
            'CLOUDINARY_API_SECRET = "your_secret"\n'
            'FOLDER                = "Products"'
        )
        st.stop()

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

# ---------- helpers ----------
def clean_filename(s: str) -> str:
    s = (s or "image").strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:120] or "image"

def extract_url_from_cell(val: str) -> Optional[str]:
    if not val: return None
    m = re.search(r'=IMAGE\(\s*"([^"]+)"', str(val), re.IGNORECASE)
    if m: return m.group(1)
    m = URL_RE.search(str(val))
    if m: return m.group(0)
    return None

def norm(s: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', (s or '').casefold())

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

# ---------- XLSX image mapping via drawing XML ----------
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
            tgt = drrels.get(rid)
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
        if not drawing_target.startswith("xl/"):
            drawing_target = posixpath.normpath(posixpath.join("xl", drawing_target))
        dr = _et_from_zip(z, drawing_target)
        if dr is None: return mapping
        drrels = _rels_map(z, posixpath.join(posixpath.dirname(drawing_target), "_rels", posixpath.basename(drawing_target)+".rels"))
        anchors = list(dr.findall(f"{{{NS_XDR}}}twoCellAnchor")) + list(dr.findall(f"{{{NS_XDR}}}oneCellAnchor"))
        for a in anchors:
            frm = a.find(f"{{{NS_XDR}}}from")
            if frm is None: continue
            r = frm.find(f"{{{NS_XDR}}}row")
            if r is None: continue
            row_1b = int(r.text) + 1
            pic = a.find(f"{{{NS_XDR}}}pic")
            if pic is None: continue
            blip = pic.find(f"{{{NS_XDR}}}blipFill/{{{NS_A}}}blip")
            if blip is None: continue
            rid = blip.attrib.get(f"{{{NS_R}}}embed")
            if not rid: continue
            tgt = drrels.get(rid)
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

# ---------- CSV reader (robust) ----------
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
                continue
    raise last_err or RuntimeError("Failed to parse CSV with given options.")

# ---------- Background processing (lazy rembg) ----------
def _get_rembg_remove():
    try:
        from rembg import remove as _remove
        return _remove
    except Exception:
        return None

def process_image_bytes(data: bytes, mode: str) -> Tuple[bytes, str]:
    """
    mode: 'none' | 'remove' | 'white'
    returns: (image_bytes, suggested_extension)
    """
    if mode == "none":
        return data, ".jpg"
    rembg_remove = _get_rembg_remove()
    if rembg_remove is None:
        # rembg/onnxruntime not available -> skip processing
        return data, ".jpg"
    try:
        out = rembg_remove(data)  # PNG with alpha
        if mode == "remove":
            return out, ".png"
        img = Image.open(io.BytesIO(out)).convert("RGBA")
        bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        buf = io.BytesIO()
        bg.convert("RGB").save(buf, format="JPEG", quality=95)
        return buf.getvalue(), ".jpg"
    except Exception:
        return data, ".jpg"

def get_bytes_from_url(url: str) -> Optional[bytes]:
    try:
        r = requests.get(url, timeout=20)
        if r.ok:
            return r.content
    except Exception:
        pass
    return None

# ---------- UI ----------
st.title("🔗 Image → Link Converter (Cloudinary)")
st.markdown(
    "- **XLSX**: supports embedded pictures (inserted over cells)\n"
    "- **CSV**: image column must contain **URLs**. If you choose a background option, URLs are downloaded → processed → re-uploaded.\n"
    "- Links are public (unlisted) on Cloudinary."
)

init_cloudinary_from_secrets()

uploaded = st.file_uploader("Upload a .xlsx or .csv", type=["xlsx","csv"])
if not uploaded:
    st.stop()

bg_choice = st.selectbox(
    "Background",
    ["No change", "Remove background (transparent PNG)", "Set white background (JPG)"],
    index=0
)
bg_mode = {"No change": "none", "Remove background (transparent PNG)": "remove", "Set white background (JPG)": "white"}[bg_choice]

suffix = Path(uploaded.name).suffix.lower()

# ---------------- XLSX ----------------
if suffix == ".xlsx":
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sheet = st.selectbox("Sheet", xls.sheet_names, index=0)
    df = pd.read_excel(xls, sheet_name=sheet, dtype=str, engine="openpyxl").fillna("")
    headers = [str(c) for c in df.columns]
    st.write("Detected headers:", headers)

    auto = st.checkbox("Auto-detect columns", value=False)

    i_idx_in = st.number_input("Image column index (1-based; 0 = auto)", min_value=0, max_value=len(headers), value=min(3, len(headers)))
    n_idx_in = st.number_input("Product/Name column index (1-based; 0 = auto)", min_value=0, max_value=len(headers), value=min(4, len(headers)))
    header_row = st.number_input("Header row number (1-based)", min_value=1, value=1)

    name_hint = st.text_input("Name header hint (optional)", "")
    img_hint  = st.text_input("Image header hint (optional)", "")

    n_idx = (n_idx_in - 1) if n_idx_in > 0 else None
    i_idx = (i_idx_in - 1) if i_idx_in > 0 else None
    if auto or n_idx is None or i_idx is None:
        n_auto, i_auto = auto_detect_columns(df, name_hint or None, img_hint or None)
        if n_idx is None: n_idx = n_auto
        if i_idx is None: i_idx = i_auto

    accept_any_col = st.checkbox("Accept any picture anchored on the row (ignore image column)", value=False)

    if st.button("Convert"):
        embedded = embedded_images_by_row_anycol(data, sheet) if accept_any_col \
                   else embedded_images_by_row_via_xml(data, sheet, i_idx + 1)

        # First data row is header_row + 1
        matched = sum(1 for r in range(len(df)) if (header_row + 1 + r) in embedded)
        st.caption(f"Embedded pictures matched to rows: {matched}/{len(df)}")

        links = []
        prog = st.progress(0, text="Uploading…")

        for r in range(len(df)):
            product = str(df.iat[r, n_idx] or "").strip()
            img_txt = str(df.iat[r, i_idx] or "").strip()
            excel_row_1b = header_row + 1 + r  # strict row mapping (prevents shift)

            emb_name, emb_bytes = embedded.get(excel_row_1b, (None, None))
            url_in_cell = extract_url_from_cell(img_txt)

            if url_in_cell and bg_mode != "none":
                raw = get_bytes_from_url(url_in_cell)
                if raw:
                    processed, ext = process_image_bytes(raw, bg_mode)
                    fname = clean_filename(product or "image") + ext
                    link = upload_bytes_to_cloudinary(processed, fname)
                else:
                    link = url_in_cell
            elif url_in_cell:
                link = url_in_cell
            elif emb_bytes:
                processed, ext = process_image_bytes(emb_bytes, bg_mode)
                base = Path(emb_name or "image").stem
                fname = clean_filename(product or base) + ext
                link = upload_bytes_to_cloudinary(processed, fname)
            else:
                link = ""  # keep alignment when no image

            links.append(link)  # one output per row, always
            prog.progress(int((r + 1) / len(df) * 100))

        df.insert(i_idx + 1, "Image Link", links)

        out_xlsx = io.BytesIO()
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
            df.to_excel(w, sheet_name=sheet, index=False)
        st.success("Done.")
        st.download_button("⬇️ Download XLSX", data=out_xlsx.getvalue(),
                           file_name=Path(uploaded.name).stem + "_with_links.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        out_csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", data=out_csv,
                           file_name=Path(uploaded.name).stem + "_with_links.csv",
                           mime="text/csv")

# ---------------- CSV ----------------
else:
    enc = st.selectbox("CSV encoding", ["auto","utf-8","utf-8-sig","cp1252","latin1"], index=0)
    delim = st.selectbox("CSV delimiter", ["auto",",",";","|","\\t"], index=0)
    try:
        df, used_enc, used_delim = read_csv_safely(uploaded, enc, delim)
        st.caption(f"Parsed with encoding **{used_enc}** and delimiter **{repr(used_delim)}**")
    except Exception as e:
        st.error(f"Could not read CSV: {e}")
        st.stop()

    headers = [str(c) for c in df.columns]
    st.write("Detected headers:", headers)

    auto = st.checkbox("Auto-detect columns", value=True)
    name_hint = st.text_input("Name header hint (optional)", "")
    img_hint  = st.text_input("Image header hint (optional)", "")

    n_idx_in = st.number_input("Product/Name column (1-based; 0 = auto)", min_value=0, max_value=len(headers),
                               value=0 if auto else min(2, len(headers)))
    i_idx_in = st.number_input("Image column (1-based; URLs; 0 = auto)", min_value=0, max_value=len(headers),
                               value=0 if auto else min(1, len(headers)))

    n_idx = (n_idx_in - 1) if n_idx_in > 0 else None
    i_idx = (i_idx_in - 1) if i_idx_in > 0 else None
    if auto or n_idx is None or i_idx is None:
        n_auto, i_auto = auto_detect_columns(df, name_hint or None, img_hint or None)
        if n_idx is None: n_idx = n_auto
        if i_idx is None: i_idx = i_auto

    st.info("For CSV: the image column should be **URLs**. If you choose a background option, each URL is downloaded, processed, and re-uploaded to Cloudinary.")

    if st.button("Convert"):
        links = []
        prog = st.progress(0, text="Processing…")

        for r, row in enumerate(df.itertuples(index=False), 1):
            product = str(row[n_idx] or "").strip()
            img_txt = str(row[i_idx] or "").strip()
            url = extract_url_from_cell(img_txt)

            if url and bg_mode != "none":
                raw = get_bytes_from_url(url)
                if raw:
                    processed, ext = process_image_bytes(raw, bg_mode)
                    fname = clean_filename(product or "image") + ext
                    link = upload_bytes_to_cloudinary(processed, fname)
                else:
                    link = url
            else:
                link = url or ""

            links.append(link)
            prog.progress(int(r / len(df) * 100))

        df.insert(i_idx + 1, "Image Link", links)

        out_xlsx = io.BytesIO()
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
            df.to_excel(w, index=False)
        st.success("Done.")
        st.download_button("⬇️ Download XLSX", data=out_xlsx.getvalue(),
                           file_name=Path(uploaded.name).stem + "_with_links.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        out_csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", data=out_csv,
                           file_name=Path(uploaded.name).stem + "_with_links.csv",
                           mime="text/csv")
