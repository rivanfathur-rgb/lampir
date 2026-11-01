import streamlit as st

from lampiran_core import buat_docx, parse_teks_per_nomor, pilah_foto

st.set_page_config(page_title="Generator Lampiran", page_icon="📎")

# Isi awal kolom "teks bawah". Edit di sini kalau mau default permanen.
TEKS_BAWAH_AWAL = """1: A2=1 A1=1 A0=1
2: A2=1 A1=1 A0=1
3: A2=1 A1=1 A0=0
4: A2=1 A1=0 A0=1
5: A2=1 A1=0 A0=0
6: A2=0 A1=1 A0=1
7: A2=0 A1=1 A0=0
8: A2=0 A1=0 A0=1
9: A2=0 A1=0 A0=0"""


# ---------- Password (opsional) ----------
# Kalau di Secrets ada PASSWORD = "...", app minta password dulu.
def cek_password():
    try:
        pw = st.secrets["PASSWORD"]
    except Exception:
        return True                      # tidak ada password diset
    if st.session_state.get("lolos"):
        return True
    masuk = st.text_input("Password", type="password")
    if masuk:
        if masuk == pw:
            st.session_state["lolos"] = True
            st.rerun()
        else:
            st.error("Password salah.")
    return False


if not cek_password():
    st.stop()

st.title("📎 Generator Lampiran Foto")
st.caption("Nama file foto harus angka: 1.jpg, 2.png, 3.jpeg, dst. Foto dimasukkan apa adanya (tanpa kompresi).")

# ---------- Upload ----------
berkas = st.file_uploader(
    "Pilih foto (bisa banyak sekaligus)",
    type=["jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff"],
    accept_multiple_files=True,
)
files = [(b.name, b.getvalue()) for b in (berkas or [])]
foto, diabaikan, duplikat = pilah_foto(files)

if berkas:
    if foto:
        nomor = [n for n, _ in foto]
        st.success(f"Terbaca {len(foto)} foto: nomor {nomor[0]} sampai {nomor[-1]}")
        hilang = sorted(set(range(nomor[0], nomor[-1] + 1)) - set(nomor))
        if hilang:
            st.warning("Nomor yang loncat (tidak ada fotonya): " + ", ".join(map(str, hilang)))
    if duplikat:
        st.warning("Nomor ganda, yang terakhir dipakai: " + ", ".join(map(str, duplikat)))
    if diabaikan:
        st.warning("Dilewati karena namanya bukan angka: " + ", ".join(diabaikan))

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
                               help='Satu baris per foto, format  nomor: teks   contoh  10: TABEL 1.2 BARIS 1')

    tampil_bawah = st.checkbox("Tampilkan teks bawah", value=True)
    bawah_default = st.text_input("Teks bawah (default kalau nomor tidak terdaftar)", "")
    bawah_txt = st.text_area("Teks bawah per foto", TEKS_BAWAH_AWAL, height=220,
                             help='Satu baris per foto, format  nomor: teks.  Tulis \\n untuk baris baru.')

    t1, t2, t3 = st.columns(3)
    atas_bold = t1.checkbox("Atas tebal", value=True)
    bawah_bold = t2.checkbox("Bawah tebal", value=False)
    ukuran_teks = t3.number_input("Ukuran teks (pt)", 6, 30, 12)

# ---------- Gambar penutup ----------
with st.expander("Gambar penutup (opsional)"):
    penutup_file = st.file_uploader("Gambar di akhir dokumen", type=["png", "jpg", "jpeg"], key="penutup")
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
