import io, os, re, zipfile, posixpath, warnings
from pathlib import Path
from typing import Optional, Dict, Tuple, List
from xml.etree import ElementTree as ET

import streamlit as st
import pandas as pd
import cloudinary
import cloudinary.uploader

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet.header_footer")

st.set_page_config(page_title="Image → Link (Cloudinary)", page_icon="🔗", layout="centered")

URL_RE = re.compile(r'https?://[^\s")]+', re.IGNORECASE)
ALLOWED_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tiff"}

# ---------------------- Cloudinary ----------------------
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
        st.error("Cloudinary secrets missing. Add them in Settings → Secrets.")
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

# ---------------------- helpers ----------------------
def clean_filename(s: str) -> str:
    s = (s or "image").strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"\s+", " ", s)
    return s[:120] or "image"

def extract_url_from_cell(val: str) -> Optional[str]:
    if not val: return None
    m = re.search(r'=IMAGE\(\s*"([^"]+)"', str(val), re.IGNORECASE)  # Google Sheets style
    if m: return m.group(1)
    m = URL_RE.search(str(val))
    if m: return m.group(0)
    return None

def norm(s: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', (s or '').casefold())

def find_col_idx(headers: List[str], primary: List[str], synonyms: List[str]) -> Optional[int]:
    H = [norm(h) for h in headers]
    # prefix
    for i, h in enumerate(H):
        for p in primary:
            if H[i].startswith(norm(p)):
                return i
    # contains
    for i, h in enumerate(H):
        for p in primary + synonyms:
            if norm(p) in h:
                return i
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

# ------------- Robust XLSX image mapping via drawing XML -------------
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

# ---------------------- UI ----------------------
st.title("🔗 Image → Link Converter (Cloudinary)")

st.markdown(
    "- **XLSX**: supports embedded pictures (inserted over cells)\n"
    "- **CSV**: image column must contain **URLs** (web apps cannot read your local `C:\\...jpg` files)\n"
    "- Links are public (unlisted) on Cloudinary"
)

init_cloudinary_from_secrets()

uploaded = st.file_uploader("Upload a .xlsx or .csv", type=["xlsx","csv"])
if not uploaded:
    st.stop()

suffix = Path(uploaded.name).suffix.lower()

if suffix == ".xlsx":
    # Load workbook in memory
    data = uploaded.getvalue()
    xls = pd.ExcelFile(io.BytesIO(data), engine="openpyxl")
    sheet = st.selectbox("Sheet", xls.sheet_names, index=0)
    df = pd.read_excel(xls, sheet_name=sheet, dtype=str, engine="openpyxl").fillna("")

    headers = [str(c) for c in df.columns]
    st.write("Detected headers:", headers)

    auto = st.checkbox("Auto-detect columns", value=False)
    if auto:
        name_hint = st.text_input("Name header hint (optional)", "")
        img_hint  = st.text_input("Image header hint (optional)", "")
        try:
            n_idx, i_idx = auto_detect_columns(df, name_hint or None, img_hint or None)
        except Exception as e:
            st.error(str(e)); st.stop()
    else:
        i_idx = st.number_input("Image column index (1-based, the pictures column)", min_value=1, max_value=len(headers), value=3) - 1
        n_idx = st.number_input("Product/Name column index (1-based)", min_value=1, max_value=len(headers), value=4) - 1

    if st.button("Convert"):
        embedded = embedded_images_by_row_via_xml(data, sheet, i_idx + 1)

        links = []
        prog = st.progress(0, text="Uploading images…")
        for r in range(len(df)):
            product = str(df.iat[r, n_idx] or "").strip()
            img_txt = str(df.iat[r, i_idx] or "").strip()
            excel_row_1b = r + 2  # header on row 1
            emb_name, emb_bytes = (None, None)
            if excel_row_1b in embedded:
                emb_name, emb_bytes = embedded[excel_row_1b]

            url = extract_url_from_cell(img_txt)
            if url:
                link = url
            elif emb_bytes:
                ext = Path(emb_name or "image.png").suffix or ".png"
                fname = clean_filename(product or Path(emb_name or "image").stem) + ext
                link = upload_bytes_to_cloudinary(emb_bytes, fname)
            else:
                link = ""  # nothing to upload

            links.append(link)
            prog.progress(int((r+1)/len(df)*100))

        df.insert(i_idx + 1, "Image Link", links)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            df.to_excel(w, sheet_name=sheet, index=False)
        st.success("Done.")
        st.download_button("⬇️ Download XLSX", data=out.getvalue(),
                           file_name=Path(uploaded.name).stem + "_with_links.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    # CSV path
    enc = st.selectbox("CSV encoding", ["utf-8","utf-8-sig","cp1252","latin1"], index=1)
    delim = st.selectbox("CSV delimiter", [",",";","|","\\t"], index=0)
    delim = "\t" if delim == "\\t" else delim

    df = pd.read_csv(uploaded, dtype=str, keep_default_na=False, encoding=enc, sep=delim, engine="python")
    headers = [str(c) for c in df.columns]
    st.write("Detected headers:", headers)

    auto = st.checkbox("Auto-detect columns", value=True)
    if auto:
        name_hint = st.text_input("Name header hint (optional)", "")
        img_hint  = st.text_input("Image header hint (optional)", "")
        try:
            n_idx, i_idx = auto_detect_columns(df, name_hint or None, img_hint or None)
        except Exception as e:
            st.error(str(e)); st.stop()
    else:
        for i, h in enumerate(headers, start=1):
            st.write(f"{i}: {h}")
        n_idx = st.number_input("Product/Name column (1-based)", 1, len(headers), value=2) - 1
        i_idx = st.number_input("Image column (1-based, must be URLs)", 1, len(headers), value=1) - 1

    st.info("For CSV: the image column must be URLs; local file paths on your computer cannot be read by a web app.")

    if st.button("Convert"):
        links = []
        prog = st.progress(0, text="Processing…")
        for r, row in enumerate(df.itertuples(index=False), 1):
            product = str(row[n_idx] or "").strip()
            img_txt = str(row[i_idx] or "").strip()
            url = extract_url_from_cell(img_txt)
            # For CSV we keep the URL (do not attempt to read local files)
            links.append(url or "")
            prog.progress(int(r/len(df)*100))

        df.insert(i_idx + 1, "Image Link", links)

        if st.toggle("Download as XLSX", value=True):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                df.to_excel(w, index=False)
            st.success("Done.")
            st.download_button("⬇️ Download XLSX", data=out.getvalue(),
                               file_name=Path(uploaded.name).stem + "_with_links.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            csv_bytes = df.to_csv(index=False).encode("utf-8")
            st.success("Done.")
            st.download_button("⬇️ Download CSV", data=csv_bytes,
                               file_name=Path(uploaded.name).stem + "_with_links.csv",
                               mime="text/csv")
