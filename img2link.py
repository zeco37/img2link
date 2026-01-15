import io, os, re, zipfile, posixpath, warnings, csv, base64, datetime, uuid
from pathlib import Path
from typing import Optional, Dict, Tuple, List
from xml.etree import ElementTree as ET

import streamlit as st
import pandas as pd
import boto3
import requests
from PIL import Image
from botocore.exceptions import ClientError

# ─────────────────────────────────────────────────────────────
# BRAND / UI (UNCHANGED)
# ─────────────────────────────────────────────────────────────
AUTHOR_NAME  = "Zakaria Belalioui"
COMPANY_NAME = "Ora Technologies"
COMPANY_URL  = "https://www.kooul.ma/"
LOGO_SOURCE  = "assets/attachment-clip-page-paper-icon-vector-design-png_117669.jpg"
BACKGROUND_URL = "https://res.cloudinary.com/dqye9uju0/image/upload/v1758554635/NsvG1713971804597-Artboard20220copy100_gocy5z.jpg"

warnings.filterwarnings("ignore", category=UserWarning)

st.set_page_config(
    page_title="Image → Link Converter",
    page_icon="🔗",
    layout="centered",
)

# ─────────────────────────────────────────────────────────────
# BACKGROUND + STYLE (UNCHANGED)
# ─────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
.stApp {{
  background:
    linear-gradient(rgba(0,0,0,.35), rgba(0,0,0,.35)),
    url('{BACKGROUND_URL}');
  background-size: cover;
  background-attachment: fixed;
}}
.main .block-container {{
  max-width: 1120px;
  background: rgba(255,255,255,0.9);
  border-radius: 20px;
  padding: 26px;
  box-shadow: 0 12px 40px rgba(0,0,0,.18);
}}
footer {{visibility:hidden;}}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div style="display:flex;align-items:center;gap:12px;">
  <div style="font-size:26px;font-weight:800;">Image → Link Converter</div>
</div>
<div style="color:#6366F1;font-weight:600;">
  by {AUTHOR_NAME} · <a href="{COMPANY_URL}" target="_blank">{COMPANY_NAME}</a>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# S3 INITIALIZATION (REPLACES CLOUDINARY)
# ─────────────────────────────────────────────────────────────
@st.cache_resource
def init_s3():
    return boto3.client(
        "s3",
        region_name=st.secrets["AWS_REGION"],
        aws_access_key_id=st.secrets["AWS_ACCESS_KEY_ID"],
        aws_secret_access_key=st.secrets["AWS_SECRET_ACCESS_KEY"],
    )

s3 = init_s3()

S3_BUCKET = st.secrets["S3_BUCKET"]
S3_PREFIX = st.secrets["S3_PREFIX"].strip("/")
PUBLIC_BASE_URL = st.secrets["PUBLIC_BASE_URL"].rstrip("/")

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def sanitize_filename(name: str) -> str:
    return re.sub(r'[^A-Za-z0-9_-]+', "_", name).strip() or "image"

def now_utc():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

def upload_bytes_to_s3(data: bytes, filename: str, run_id: str) -> str:
    key = f"{S3_PREFIX}/{run_id}/{filename}"
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=key,
        Body=data,
        ContentType="image/jpeg" if filename.lower().endswith(".jpg") else "image/png",
    )
    return f"{PUBLIC_BASE_URL}/{run_id}/{filename}"

def ensure_rgb(img: Image.Image) -> Image.Image:
    if img.mode != "RGB":
        return img.convert("RGB")
    return img

# ─────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────
uploaded = st.file_uploader("Upload CSV or XLSX", type=["csv","xlsx"])

if uploaded:
    df = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)

    st.subheader("Detected columns")
    st.json(list(df.columns))

    product_col = st.selectbox("Product column", df.columns)
    url_col = st.selectbox("Image URL column", df.columns)

    if st.button("🚀 Process images"):
        run_id = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
        zip_buffer = io.BytesIO()
        server_urls = [None] * len(df)

        progress = st.progress(0.0)
        uploaded_count = 0
        skipped_count = 0

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                progress.progress((idx + 1) / len(df))

                product = str(row[product_col]).strip()
                url = str(row[url_col]).strip()

                if not url.startswith("http"):
                    skipped_count += 1
                    continue

                try:
                    r = requests.get(url, timeout=25)
                    r.raise_for_status()

                    img = Image.open(io.BytesIO(r.content))
                    img = ensure_rgb(img)

                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=95)
                    raw_bytes = buf.getvalue()

                    filename = sanitize_filename(product) + ".jpg"

                    public_url = upload_bytes_to_s3(raw_bytes, filename, run_id)
                    zipf.writestr(filename, raw_bytes)

                    server_urls[idx] = public_url
                    uploaded_count += 1

                except Exception as e:
                    skipped_count += 1

        df["Image Link"] = server_urls

        st.success(f"✅ Uploaded: {uploaded_count} | Skipped: {skipped_count}")

        st.download_button(
            "⬇️ Download ZIP",
            zip_buffer.getvalue(),
            file_name=f"images_{run_id}.zip",
            mime="application/zip",
        )

        st.download_button(
            "⬇️ Download CSV",
            df.to_csv(index=False).encode("utf-8"),
            file_name=f"updated_{run_id}.csv",
            mime="text/csv",
        )

# ─────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────
year = datetime.datetime.now().year
st.markdown(
    f"<div style='text-align:center;color:#475569;'>© {year} {AUTHOR_NAME} · {COMPANY_NAME}</div>",
    unsafe_allow_html=True,
)
