"""
Inti pembuat lampiran Word. Dipakai oleh app.py (web) dan bisa dites tanpa Streamlit.
Foto dimasukkan apa adanya (tanpa kompresi).
"""

import io
import os
import re

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image, ImageOps

EKSTENSI = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff")
RATA = {"kiri": WD_ALIGN_PARAGRAPH.LEFT,
        "tengah": WD_ALIGN_PARAGRAPH.CENTER,
        "kanan": WD_ALIGN_PARAGRAPH.RIGHT}

DEFAULT = {
    # halaman
    "kertas": "Letter", "landscape": False, "margin_cm": 2.54,
    "font": "Calibri", "ukuran_font": 12,
    # tabel
    "kolom": 2, "baris_per_halaman": None,
    "lebar_foto_cm": 8.53, "tinggi_foto_cm": 6.40, "paksa_ukuran": False,
    "gaya_garis": "kotak_per_baris", "tebal_garis": 12,
    # teks atas
    "tampil_atas": True, "atas_default": "TABEL 1.1 BARIS {n}", "atas": {},
    "atas_bold": True, "atas_ukuran": 12, "atas_rata": "tengah",
    # teks bawah
    "tampil_bawah": True, "bawah_default": "", "bawah": {},
    "bawah_bold": False, "bawah_ukuran": 12, "bawah_rata": "tengah",
    # gambar penutup
    "penutup_lebar_cm": 7.4, "penutup_tinggi_cm": 22.9, "penutup_halaman_baru": False,
}


def pilah_foto(files):
    """
    files = list of (nama_file, bytes).
    Return (foto_terurut, diabaikan, duplikat)
      foto_terurut : [(nomor, bytes), ...] urut angka
      diabaikan    : nama file yang bukan angka / bukan gambar
      duplikat     : nomor yang muncul lebih dari sekali
    """
    peta, diabaikan, duplikat = {}, [], []
    for nama, data in files:
        base, ext = os.path.splitext(nama)
        if ext.lower() in EKSTENSI and re.fullmatch(r"\d+", base):
            n = int(base)
            if n in peta:
                duplikat.append(n)
            peta[n] = data
        else:
            diabaikan.append(nama)
    return sorted(peta.items()), diabaikan, sorted(set(duplikat))


def parse_teks_per_nomor(teks):
    """
    Ubah teks multi-baris 'nomor: isi' menjadi dict {nomor: isi}.
    Tulis \\n di dalam isi untuk baris baru.
    """
    hasil = {}
    for baris in teks.splitlines():
        if ":" not in baris:
            continue
        kiri, kanan = baris.split(":", 1)
        kiri = kiri.strip()
        if kiri.isdigit():
            hasil[int(kiri)] = kanan.strip().replace("\\n", "\n")
    return hasil


_JFIF = b"\xff\xe0\x00\x10JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00"


def _siapkan_bytes(data):
    """
    Pastikan format gambar dikenali python-docx, TANPA mengubah/mengompres piksel.
    - JPEG tanpa penanda JFIF/Exif (umum pada foto WhatsApp): sisipkan penanda JFIF.
    - PNG/GIF/BMP/TIFF: dibiarkan.
    - Format lain (WebP dll): diubah ke PNG (lossless).
    """
    if data[:2] == b"\xff\xd8":
        if data[6:10] in (b"JFIF", b"Exif"):
            return data
        return data[:2] + _JFIF + data[2:]
    if (data[:8] == b"\x89PNG\r\n\x1a\n" or data[:6] in (b"GIF87a", b"GIF89a")
            or data[:2] == b"BM" or data[:4] in (b"II*\x00", b"MM\x00*")):
        return data
    with Image.open(io.BytesIO(data)) as im:
        im = ImageOps.exif_transpose(im)
        if im.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "PNG")
    return buf.getvalue()


def _siapkan_dokumen(cfg):
    doc = Document()
    sec = doc.sections[0]
    if cfg["kertas"].upper() == "A4":
        w, h = Cm(21.0), Cm(29.7)
    else:
        w, h = Cm(21.59), Cm(27.94)
    if cfg["landscape"]:
        sec.orientation = WD_ORIENT.LANDSCAPE
        w, h = h, w
    sec.page_width, sec.page_height = w, h
    for sisi in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, sisi, Cm(cfg["margin_cm"]))

    normal = doc.styles["Normal"]
    normal.font.name = cfg["font"]
    normal.font.size = Pt(cfg["ukuran_font"])
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), cfg["font"])
    return doc, sec.page_width - sec.left_margin - sec.right_margin


def _ukuran_gambar(data, cfg, maks_lebar_cm):
    """Ukuran gambar (EMU). Lebar tidak pernah melebihi isi kolom, supaya tidak kepotong."""
    lebar_cfg = cfg["lebar_foto_cm"]
    tinggi_cfg = cfg["tinggi_foto_cm"]
    if cfg["paksa_ukuran"]:
        k = min(1.0, maks_lebar_cm / lebar_cfg)
        return Cm(lebar_cfg * k), Cm(tinggi_cfg * k)
    with Image.open(io.BytesIO(data)) as im:
        w, h = im.size
        try:
            orientasi = im.getexif().get(0x0112, 1)
        except Exception:
            orientasi = 1
    if orientasi in (5, 6, 7, 8):      # foto HP yang disimpan miring (EXIF)
        w, h = h, w
    skala = min(min(lebar_cfg, maks_lebar_cm) / w, tinggi_cfg / h)
    return Cm(w * skala), Cm(h * skala)


def _garis_sel(cell, tebal, atas, kiri, bawah, kanan):
    tcPr = cell._tc.get_or_add_tcPr()
    lama = tcPr.find(qn("w:tcBorders"))
    if lama is not None:
        tcPr.remove(lama)
    borders = OxmlElement("w:tcBorders")
    for nama, ada in (("top", atas), ("left", kiri), ("bottom", bawah), ("right", kanan)):
        el = OxmlElement(f"w:{nama}")
        if ada:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(tebal))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tcPr.append(borders)


def _tulis_teks(sel, teks, bold, ukuran, rata, keep_next):
    p = sel.paragraphs[0]
    p.alignment = RATA[rata]
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.keep_with_next = keep_next
    run = p.add_run(teks)
    run.bold = bold
    run.font.size = Pt(ukuran)


def _buat_tabel(doc, lebar_konten, grup, cfg):
    kolom = cfg["kolom"]
    lebar_kolom = int(lebar_konten / kolom)

    jenis = []
    if cfg["tampil_atas"]:
        jenis.append("atas")
    jenis.append("gambar")
    if cfg["tampil_bawah"]:
        jenis.append("bawah")

    tabel = doc.add_table(rows=len(grup) * len(jenis), cols=kolom)
    tabel.style = "Table Grid"
    tabel.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabel.autofit = False
    for kol in tabel.columns:
        kol.width = lebar_kolom

    for gi, baris_foto in enumerate(grup):
        for ji, j in enumerate(jenis):
            row = tabel.rows[gi * len(jenis) + ji]
            row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
            terakhir = ji == len(jenis) - 1

            for ci in range(kolom):
                sel = row.cells[ci]
                sel.width = lebar_kolom
                sel.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                if cfg["gaya_garis"] == "tanpa":
                    _garis_sel(sel, cfg["tebal_garis"], False, False, False, False)
                elif cfg["gaya_garis"] == "grid":
                    _garis_sel(sel, cfg["tebal_garis"], True, True, True, True)
                else:
                    _garis_sel(sel, cfg["tebal_garis"],
                               atas=(ji == 0), kiri=(ci == 0),
                               bawah=terakhir, kanan=(ci == kolom - 1))

                p = sel.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                if ci >= len(baris_foto):
                    continue
                nomor, data = baris_foto[ci]
                keep = not terakhir

                if j == "atas":
                    teks = cfg["atas"].get(nomor, cfg["atas_default"]).replace("{n}", str(nomor))
                    _tulis_teks(sel, teks, cfg["atas_bold"], cfg["atas_ukuran"],
                                cfg["atas_rata"], keep)
                elif j == "gambar":
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.keep_with_next = keep
                    lebar, tinggi = _ukuran_gambar(data, cfg, lebar_kolom / 360000 - 0.4)
                    p.add_run().add_picture(io.BytesIO(data), width=lebar, height=tinggi)
                else:
                    teks = cfg["bawah"].get(nomor, cfg["bawah_default"]).replace("{n}", str(nomor))
                    _tulis_teks(sel, teks, cfg["bawah_bold"], cfg["bawah_ukuran"],
                                cfg["bawah_rata"], False)
    return tabel


def buat_docx(foto, cfg_user=None, penutup=None):
    """
    foto    : [(nomor, bytes), ...] sudah urut
    cfg_user: dict yang menimpa DEFAULT
    penutup : bytes gambar penutup (opsional)
    return  : bytes file .docx
    """
    cfg = dict(DEFAULT)
    cfg.update(cfg_user or {})
    if not foto:
        raise ValueError("Tidak ada foto.")

    foto_ok = []
    for nomor, data in foto:
        try:
            foto_ok.append((nomor, _siapkan_bytes(data)))
        except Exception as e:
            raise ValueError(f"Foto nomor {nomor} tidak bisa dibaca ({type(e).__name__}). "
                             f"Coba simpan ulang fotonya sebagai JPG/PNG biasa.")
    foto = foto_ok
    if penutup:
        try:
            penutup = _siapkan_bytes(penutup)
        except Exception as e:
            raise ValueError(f"Gambar penutup tidak bisa dibaca ({type(e).__name__}).")

    doc, lebar_konten = _siapkan_dokumen(cfg)

    kolom = cfg["kolom"]
    semua_baris = [foto[i:i + kolom] for i in range(0, len(foto), kolom)]
    bph = cfg["baris_per_halaman"]
    if bph:
        halaman = [semua_baris[i:i + bph] for i in range(0, len(semua_baris), bph)]
    else:
        halaman = [semua_baris]

    for idx, grup in enumerate(halaman):
        _buat_tabel(doc, lebar_konten, grup, cfg)
        if idx < len(halaman) - 1:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    if penutup:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.page_break_before = cfg["penutup_halaman_baru"]
        p.add_run().add_picture(io.BytesIO(penutup),
                                width=Cm(cfg["penutup_lebar_cm"]),
                                height=Cm(cfg["penutup_tinggi_cm"]))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
