import base64
import hashlib
import html
import io
import os

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

from lampiran_core import (buat_docx, kunci_alami, parse_teks_per_nomor,
                           pilah_foto, urutkan_manual)

# ---------- Tampilan (edit di sini) ----------
NAMA_APP     = "LampirGen"                        # teks logo
TAGLINE      = "Lampiran foto otomatis ke Word"
FONT_LOGO    = "Orbitron"                         # nama font dari fonts.google.com
                                                  # (contoh lain: "Space Grotesk", "Audiowide", "Rubik Glitch", "Pacifico")
WARNA_LOGO_1 = "#4F8BF9"                          # warna gradasi kiri
WARNA_LOGO_2 = "#9B5DE5"                          # warna gradasi kanan
LOGO_FILE    = "logo.png"                         # logo buatan sendiri; taruh di repo (png/jpg/svg).
                                                  # kalau file tidak ada, dipakai logo tulisan di atas

st.set_page_config(page_title=NAMA_APP, page_icon="📎")

MODE_PILIH = "Sesuai urutan pilih / upload"
MODE_NAMA = "Urut nama file (A-Z)"
MODE_ANGKA = "Nama file berupa angka (1.jpg, 2.png, ...)"

PLACEHOLDER_BAWAH = """Ketik keterangan gambar di sini, contoh:
1: R1=220Ohm
2: V1=5V
3: Keterangan
dan seterusnya..."""

PLACEHOLDER_ATAS = """Opsional, contoh:
3: TABEL 1.2 BARIS 1"""


# ---------- Pratinjau ----------
@st.cache_data(show_spinner=False, max_entries=300)
def thumb_b64(data: bytes):
    """Thumbnail kecil (hanya untuk pratinjau; file Word tetap memakai foto asli)."""
    try:
        with Image.open(io.BytesIO(data)) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((240, 240))
            buf = io.BytesIO()
            im.convert("RGB").save(buf, "JPEG", quality=70)
        return base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return None


def tampilkan_pratinjau(item):
    """item = [(nomor, nama, bytes), ...]"""
    kotak = []
    for nomor, nama, data in item:
        b64 = thumb_b64(data)
        if b64:
            gambar = (f'<img src="data:image/jpeg;base64,{b64}" '
                      f'style="width:100%;border-radius:4px">')
        else:
            gambar = '<div style="padding:24px 0;border:1px dashed #888">?</div>'
        label = f"<b>{nomor}</b>" + (f" · {html.escape(nama)}" if nama else "")
        kotak.append(
            '<div style="width:31%;min-width:90px;text-align:center;font-size:12px;'
            f'word-break:break-all">{gambar}<div>{label}</div></div>'
        )
    st.markdown(
        '<div style="display:flex;flex-wrap:wrap;gap:8px">' + "".join(kotak) + "</div>",
        unsafe_allow_html=True,
    )


# ---------- Header kustom (logo) + sembunyikan menu bawaan ----------
def _logo_html():
    if LOGO_FILE and os.path.isfile(LOGO_FILE):
        ext = os.path.splitext(LOGO_FILE)[1].lower().lstrip(".")
        mime = {"svg": "image/svg+xml", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(ext, f"image/{ext}")
        with open(LOGO_FILE, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        return f'<img src="data:{mime};base64,{b64}" style="max-height:70px;max-width:85%">'
    return f'<div class="lg-text">{html.escape(NAMA_APP)}</div>'


def tampilkan_header():
    font_url = FONT_LOGO.replace(" ", "+")
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family={font_url}:wght@600;800&display=swap');
header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stMainMenu"],
[data-testid="stAppDeployButton"], [data-testid="stDecoration"], #MainMenu, footer {{
    display: none !important;
}}
.block-container {{ padding-top: 1.2rem; }}
.lg-wrap {{ text-align: center; margin: 0 0 1.2rem 0; }}
.lg-text {{
    font-family: '{FONT_LOGO}', sans-serif; font-weight: 800; font-size: 2.3rem;
    letter-spacing: .06em; line-height: 1.2;
    background: linear-gradient(90deg, {WARNA_LOGO_1}, {WARNA_LOGO_2});
    -webkit-background-clip: text; background-clip: text; color: transparent;
}}
.lg-tag {{ opacity: .7; font-size: .9rem; margin-top: .2rem; }}
</style>
<div class="lg-wrap">{_logo_html()}<div class="lg-tag">{html.escape(TAGLINE)}</div></div>
""",
        unsafe_allow_html=True,
    )


tampilkan_header()
st.caption("Pilih foto, atur teksnya, lalu download Word. Foto dimasukkan apa adanya (tanpa kompresi).")

# ---------- Upload ----------
berkas = st.file_uploader(
    "Pilih foto (bisa banyak sekaligus)",
    type=["jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff", "webp"],
    accept_multiple_files=True,
)
files = [(b.name, b.getvalue()) for b in (berkas or [])]

mode = st.radio(
    "Cara menentukan urutan foto",
    [MODE_PILIH, MODE_NAMA, MODE_ANGKA],
    help="Nomor foto (1, 2, 3, ...) menentukan posisi di tabel dan keterangan mana yang dipakai.",
)

foto = []          # [(nomor, bytes)]
tampil = []        # [(nomor, nama, bytes)]

if files:
    if mode == MODE_ANGKA:
        foto, diabaikan, duplikat = pilah_foto(files)
        tampil = [(n, "", d) for n, d in foto]
        if foto:
            nomor = [n for n, _ in foto]
            hilang = sorted(set(range(nomor[0], nomor[-1] + 1)) - set(nomor))
            if hilang:
                st.warning("Nomor yang loncat (tidak ada fotonya): " + ", ".join(map(str, hilang)))
        if duplikat:
            st.warning("Nomor ganda, yang terakhir dipakai: " + ", ".join(map(str, duplikat)))
        if diabaikan:
            st.warning("Dilewati karena namanya bukan angka: " + ", ".join(diabaikan))
    else:
        daftar = list(files)
        if mode == MODE_NAMA:
            daftar.sort(key=lambda f: kunci_alami(f[0]))
        else:
            st.caption("Urutan mengikuti urutan file yang dikirim browser/hape. "
                       "Cek pratinjau di bawah; kalau belum sesuai, ubah lewat 'Ubah urutan'.")

        nama_list = [n for n, _ in daftar]
        with st.expander("Ubah urutan (opsional)"):
            st.caption("Ubah angka di kolom Urutan. Foto dengan angka terkecil jadi nomor 1.")
            kunci = "urut_" + hashlib.md5((mode + "|".join(nama_list)).encode()).hexdigest()
            df = pd.DataFrame({"Urutan": list(range(1, len(daftar) + 1)), "Nama file": nama_list})
            edit = st.data_editor(
                df, hide_index=True, disabled=["Nama file"], key=kunci,
                column_config={"Urutan": st.column_config.NumberColumn(min_value=1, step=1)},
            )
            nilai = edit["Urutan"].tolist()

        urut = urutkan_manual(daftar, nilai)
        foto = [(n, item[1]) for n, item in urut]
        tampil = [(n, item[0], item[1]) for n, item in urut]

    if foto:
        st.success(f"{len(foto)} foto siap dipakai (nomor {foto[0][0]} sampai {foto[-1][0]})")
        with st.expander("Pratinjau urutan", expanded=True):
            tampilkan_pratinjau(tampil)

# ---------- Susunan & ukuran ----------
with st.expander("Susunan & ukuran", expanded=False):
    c1, c2 = st.columns(2)
    kolom = c1.number_input("Kolom (foto per baris)", 1, 6, 2)
    bph_on = c2.checkbox("Paksa jumlah baris per halaman")
    bph = c2.number_input("Baris per halaman", 1, 20, 3) if bph_on else None

    c3, c4 = st.columns(2)
    lebar = c3.number_input("Lebar foto (cm)", 1.0, 30.0, 8.53, step=0.1)
    tinggi = c4.number_input("Tinggi foto (cm)", 1.0, 30.0, 6.40, step=0.1)
    paksa = st.checkbox("Paksa ukuran persis (foto bisa gepeng)", value=False)

    gaya = st.selectbox(
        "Gaya garis",
        ["kotak_per_baris", "grid", "tanpa"],
        format_func={"kotak_per_baris": "Kotak per baris (seperti contoh)",
                     "grid": "Tiap foto punya kotak",
                     "tanpa": "Tanpa garis"}.get,
    )
    tebal = st.slider("Tebal garis (1/8 pt)", 2, 24, 12)

    c5, c6 = st.columns(2)
    kertas = c5.selectbox("Kertas", ["Letter", "A4"])
    landscape = c6.checkbox("Landscape")
    margin = st.number_input("Margin (cm)", 0.5, 5.0, 2.54, step=0.1)

# ---------- Teks ----------
with st.expander("Teks atas & bawah foto", expanded=True):
    tampil_atas = st.checkbox("Tampilkan teks atas", value=True)
    atas_default = st.text_input("Teks atas (default semua foto)", "TABEL 1.1 BARIS {n}",
                                 help="{n} otomatis diganti nomor foto")
    atas_khusus = st.text_area("Teks atas khusus (opsional)", "", height=80,
                               placeholder=PLACEHOLDER_ATAS,
                               help="Satu baris per foto, format  nomor: teks")

    tampil_bawah = st.checkbox("Tampilkan teks bawah", value=True)
    bawah_default = st.text_input("Teks bawah (default kalau nomor tidak terdaftar)", "")
    bawah_txt = st.text_area("Keterangan gambar (teks bawah)", "", height=220,
                             placeholder=PLACEHOLDER_BAWAH,
                             help="Satu baris per foto, format  nomor: teks. "
                                  "Nomor = nomor di pratinjau. Tulis \\n untuk baris baru.")

    t1, t2, t3 = st.columns(3)
    atas_bold = t1.checkbox("Atas tebal", value=True)
    bawah_bold = t2.checkbox("Bawah tebal", value=False)
    ukuran_teks = t3.number_input("Ukuran teks (pt)", 6, 30, 12)

# ---------- Gambar penutup ----------
with st.expander("Gambar penutup (opsional)"):
    penutup_file = st.file_uploader("Gambar di akhir dokumen", type=["png", "jpg", "jpeg", "webp"],
                                    key="penutup")
    p1, p2 = st.columns(2)
    p_lebar = p1.number_input("Lebar (cm)", 1.0, 30.0, 7.4, step=0.1)
    p_tinggi = p2.number_input("Tinggi (cm)", 1.0, 30.0, 22.9, step=0.1)
    p_baru = st.checkbox("Mulai di halaman baru", value=False)

nama_out = st.text_input("Nama file hasil", "Lampiran.docx")
if not nama_out.lower().endswith(".docx"):
    nama_out += ".docx"

# ---------- Buat ----------
if st.button("Buat Word", type="primary", disabled=not foto):
    cfg = {
        "kolom": int(kolom), "baris_per_halaman": int(bph) if bph else None,
        "lebar_foto_cm": float(lebar), "tinggi_foto_cm": float(tinggi), "paksa_ukuran": paksa,
        "gaya_garis": gaya, "tebal_garis": int(tebal),
        "kertas": kertas, "landscape": landscape, "margin_cm": float(margin),
        "tampil_atas": tampil_atas, "atas_default": atas_default,
        "atas": parse_teks_per_nomor(atas_khusus), "atas_bold": atas_bold,
        "atas_ukuran": int(ukuran_teks),
        "tampil_bawah": tampil_bawah, "bawah_default": bawah_default,
        "bawah": parse_teks_per_nomor(bawah_txt), "bawah_bold": bawah_bold,
        "bawah_ukuran": int(ukuran_teks),
        "penutup_lebar_cm": float(p_lebar), "penutup_tinggi_cm": float(p_tinggi),
        "penutup_halaman_baru": p_baru,
    }
    try:
        with st.spinner("Menyusun dokumen..."):
            hasil = buat_docx(foto, cfg, penutup_file.getvalue() if penutup_file else None)
        st.session_state["hasil"] = hasil
        st.session_state["nama"] = nama_out
    except Exception as e:
        st.session_state.pop("hasil", None)
        st.error(f"Gagal membuat dokumen: {e}")

if "hasil" in st.session_state:
    mb = len(st.session_state["hasil"]) / 1024 / 1024
    st.download_button(
        f"⬇️ Download {st.session_state['nama']} ({mb:.1f} MB)",
        st.session_state["hasil"],
        file_name=st.session_state["nama"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
