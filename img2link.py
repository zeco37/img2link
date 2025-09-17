# img2link_wizard.py
from __future__ import annotations
import csv, io, os, re, zipfile, posixpath, warnings
from pathlib import Path
from typing import Optional, Dict, Tuple, List
from xml.etree import ElementTree as ET

import pandas as pd
from dotenv import load_dotenv
import cloudinary, cloudinary.uploader

from rich import print
from rich.prompt import Prompt, IntPrompt, Confirm
from rich.progress import Progress

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.worksheet.header_footer")

SCRIPT_DIR = Path(__file__).resolve().parent
load_dotenv(SCRIPT_DIR / ".env", override=True)

URL_RE = re.compile(r'https?://[^\s")]+', re.IGNORECASE)
ALLOWED_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tiff"}

# ---------- Cloudinary ----------
def init_cloudinary_interactive():
    cloud = os.getenv("CLOUDINARY_CLOUD_NAME") or Prompt.ask("[bold]Cloudinary cloud name[/bold] (e.g. dqye9uju0)")
    key   = os.getenv("CLOUDINARY_API_KEY")    or Prompt.ask("[bold]Cloudinary API key[/bold]")
    secret= os.getenv("CLOUDINARY_API_SECRET") or Prompt.ask("[bold]Cloudinary API secret[/bold]", password=True)
    folder= os.getenv("FOLDER") or Prompt.ask("Cloudinary folder (optional)", default="Products")

    cloudinary.config(cloud_name=cloud, api_key=key, api_secret=secret, secure=True)
    cfg = cloudinary.config()
    if not (cfg.cloud_name and cfg.api_key and cfg.api_secret):
        raise RuntimeError("Cloudinary credentials incomplete.")

    # Offer to save into .env for next time
    if Confirm.ask("Save these credentials to .env for next time?", default=True):
        env_path = SCRIPT_DIR / ".env"
        lines = [
            f"CLOUDINARY_CLOUD_NAME={cloud}\n",
            f"CLOUDINARY_API_KEY={key}\n",
            f"CLOUDINARY_API_SECRET={secret}\n",
            f"FOLDER={folder}\n",
        ]
        env_path.write_text("".join(lines), encoding="utf-8")
        print(f"[green]Saved[/green] to {env_path}")

    # Keep folder in env for this run
    os.environ["FOLDER"] = folder

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

def upload_file_to_cloudinary(path: str) -> str:
    res = cloudinary.uploader.upload(
        path,
        resource_type="image",
        folder=(os.getenv("FOLDER") or None),
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
    m = re.search(r'=IMAGE\(\s*"([^"]+)"', str(val), re.IGNORECASE)  # Sheets-style
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
    img_primary  = [img_hint]  if img_hint  else ["photo","image","images","img","picture","images","photos"]
    name_syn = ["libelle","designation","article","titre","item"]
    img_syn  = ["link","imagelink","imageurl","photourl","lienimage"]

    n_idx = find_col_idx(headers, name_primary, name_syn)
    i_idx = find_col_idx(headers, img_primary, img_syn)
    if n_idx is None or i_idx is None:
        raise ValueError(f"Auto-detect failed. Headers: {headers}")
    return n_idx, i_idx

# ---------- Robust XLSX image mapping via drawing XML ----------
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

def embedded_images_by_row_via_xml(xlsx_path: str, sheet_name: str, image_col_idx_1b: int) -> Dict[int, Tuple[str, bytes]]:
    mapping: Dict[int, Tuple[str, bytes]] = {}
    with zipfile.ZipFile(xlsx_path, "r") as z:
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

# ---------- core ----------
def handle_row(product: str, img_cell_text: str,
               embedded_bytes: Optional[bytes], embedded_name: Optional[str]) -> str:
    # 1) URL already
    url = extract_url_from_cell(img_cell_text)
    if url: return url
    # 2) local file path
    if img_cell_text and os.path.exists(img_cell_text):
        if Path(img_cell_text).suffix.lower() in ALLOWED_EXTS:
            return upload_file_to_cloudinary(img_cell_text)
    # 3) embedded bytes
    if embedded_bytes:
        ext = Path(embedded_name or "image.png").suffix or ".png"
        filename = clean_filename(product or Path(embedded_name or "image").stem) + ext
        return upload_bytes_to_cloudinary(embedded_bytes, filename)
    return ""

def choose_file_path() -> Path:
    while True:
        p = Prompt.ask("Enter [bold]input file path[/bold] (.csv or .xlsx)")
        path = Path(p).expanduser()
        if path.exists() and path.suffix.lower() in {".csv",".xlsx"}:
            return path
        print("[red]Path invalid or wrong extension. Try again.[/red]")

def sniff_csv_format(path: Path) -> Tuple[str, str]:
    # Try to guess encoding/delimiter
    encoding_choices = ["utf-8", "utf-8-sig", "cp1252", "latin1"]
    with path.open("rb") as f:
        sample = f.read(4096)
    try_delims = [",",";","|","\\t"]
    # delimiter
    try:
        dialect = csv.Sniffer().sniff(sample.decode("utf-8", errors="ignore"))
        delim = dialect.delimiter
    except Exception:
        delim = ","
    # encoding: try utf-8-sig, then cp1252
    for enc in ["utf-8-sig","utf-8","cp1252","latin1"]:
        try:
            sample.decode(enc)
            guessed = enc
            break
        except Exception:
            continue
    print(f"[cyan]Guessed CSV[/cyan]: encoding=[bold]{guessed}[/bold], delimiter=[bold]{delim}[/bold]")
    enc = Prompt.ask("Choose encoding", choices=encoding_choices, default=guessed)
    delim = Prompt.ask("Choose delimiter", choices=[",",";","|","\\t"], default=delim)
    if delim == "\\t": delim = "\t"
    return enc, delim

def wizard_csv(path: Path):
    enc, delim = sniff_csv_format(path)
    df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc, sep=delim, engine="python")
    headers = [str(c) for c in df.columns]
    print(f"[cyan]Headers[/cyan]: {headers}")

    mode = Prompt.ask("Select column mode", choices=["auto","manual"], default="auto")

    if mode == "auto":
        nh = Prompt.ask("Name header hint (optional)", default="")
        ih = Prompt.ask("Image header hint (optional)", default="")
        n_idx, i_idx = auto_detect_columns(df, nh or None, ih or None)
    else:
        for i, h in enumerate(headers, start=1):
            print(f"[dim]{i:>2}[/] : {h}")
        n_idx = IntPrompt.ask("Enter [bold]product/name[/bold] column number", default=2) - 1
        i_idx = IntPrompt.ask("Enter [bold]image[/bold] column number", default=1) - 1

    out_default = path.with_name(path.stem + "_with_links.csv")
    out = Path(Prompt.ask("Output file", default=str(out_default))).expanduser()

    # Cloudinary init (only if needed)
    init_cloudinary_interactive()

    links = []
    with Progress() as prog:
        task = prog.add_task("Converting…", total=len(df))
        for _, row in df.iterrows():
            product = str(row.iloc[n_idx] or "").strip()
            img_txt = str(row.iloc[i_idx] or "").strip()
            links.append(handle_row(product, img_txt, None, None))
            prog.advance(task)
    df.insert(i_idx + 1, "Image Link", links)
    if out.suffix.lower() == ".xlsx":
        df.to_excel(out, index=False)
    else:
        df.to_csv(out, index=False)
    print(f"[green]Done.[/green] Wrote: {out}")

def wizard_xlsx(path: Path):
    xls = pd.ExcelFile(path, engine="openpyxl")
    print(f"[cyan]Sheets[/cyan]: {xls.sheet_names}")
    sheet = Prompt.ask("Choose sheet", choices=xls.sheet_names, default=xls.sheet_names[0])
    df = pd.read_excel(xls, sheet_name=sheet, dtype=str, engine="openpyxl").fillna("")
    headers = [str(c) for c in df.columns]
    print(f"[cyan]Headers[/cyan]: {headers}")

    mode = Prompt.ask("Select column mode", choices=["auto","manual"], default="manual")

    if mode == "auto":
        nh = Prompt.ask("Name header hint (optional)", default="")
        ih = Prompt.ask("Image header hint (optional)", default="")
        n_idx, i_idx = auto_detect_columns(df, nh or None, ih or None)
    else:
        for i, h in enumerate(headers, start=1):
            print(f"[dim]{i:>2}[/] : {h}")
        # Often: images are in a column without real header -> user can still type the position
        i_idx = IntPrompt.ask("Enter [bold]image[/bold] column number (pictures column)", default=3) - 1
        n_idx = IntPrompt.ask("Enter [bold]product/name[/bold] column number", default=4) - 1

    out_default = path.with_name(path.stem + "_with_links.xlsx")
    out = Path(Prompt.ask("Output file", default=str(out_default))).expanduser()

    # Cloudinary init
    init_cloudinary_interactive()

    # Map embedded pictures for the chosen image column
    embedded = embedded_images_by_row_via_xml(str(path), sheet, i_idx + 1)

    links = []
    with Progress() as prog:
        task = prog.add_task("Converting…", total=len(df))
        for r in range(len(df)):
            product = str(df.iat[r, n_idx] or "").strip()
            img_txt = str(df.iat[r, i_idx] or "").strip()
            excel_row_1b = r + 2  # header row is 1
            emb_name, emb_bytes = (None, None)
            if excel_row_1b in embedded:
                emb_name, emb_bytes = embedded[excel_row_1b]
            links.append(handle_row(product, img_txt, emb_bytes, emb_name))
            prog.advance(task)

    df.insert(i_idx + 1, "Image Link", links)
    if out.suffix.lower() == ".csv":
        df.to_csv(out, index=False)
    else:
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            df.to_excel(w, sheet_name=sheet, index=False)
    print(f"[green]Done.[/green] Wrote: {out}")

def main():
    print("[bold cyan]Image → Link Wizard (Cloudinary)[/bold cyan]")
    path = choose_file_path()
    # Normalize Windows backslashes
    path = path.expanduser().resolve()
    if path.suffix.lower() == ".csv":
        wizard_csv(path)
    else:
        wizard_xlsx(path)

if __name__ == "__main__":
    main()
