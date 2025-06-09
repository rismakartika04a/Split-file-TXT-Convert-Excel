import streamlit as st
import re
import os
import tempfile
import zipfile
import pandas as pd

st.set_page_config(layout="wide")
st.title("Split file TXT & Convert ➔ Excel")

# 1) Upload file
uploaded_file = st.file_uploader("Upload file TXT", type=None)
if uploaded_file is None:
    st.info("Silakan upload file TXT...")
    st.stop()

# 2) Pilih header pemisah (optional)
split_header = st.selectbox(
    "Pilih header pemisah halaman (opsional):",
    ["", "GAJI KARYAWAN TETAP", "GAJI KARYAWAN KONTRAK"]
)

# 3) Baca isi file (UTF-8/Latin-1 fallback)
try:
    try:
        raw_text = uploaded_file.read().decode("utf-8")
    except UnicodeDecodeError:
        uploaded_file.seek(0)
        raw_text = uploaded_file.read().decode("latin-1")
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

# 4) Split berdasarkan header jika ada
if not split_header:
    sections = [raw_text]
    st.info("Tidak ada header dipilih → 1 section saja.")
else:
    pattern = rf'({re.escape(split_header)}.*?)(?={re.escape(split_header)}|\Z)'
    sections = re.findall(pattern, raw_text, flags=re.DOTALL)
    if not sections:
        sections = [raw_text]
        st.warning("Header tidak ditemukan, semua jadi 1 section.")
    else:
        st.success(f"Ditemukan {len(sections)} section berdasarkan header.")

with tempfile.TemporaryDirectory() as tmpdirname:
    # ZIP & Excel writer
    zip_path = os.path.join(tmpdirname, "split_files.zip")
    zipf = zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED)
    excel_path = os.path.join(tmpdirname, "hasil_split_gaji_karyawan.xlsx")
    writer = pd.ExcelWriter(excel_path, engine="xlsxwriter")
    preview_dfs = []
    filepaths = []

    for idx, section in enumerate(sections, start=1):
        fn_txt = f"gaji_karyawan_page_{idx}.txt"
        fp_txt = os.path.join(tmpdirname, fn_txt)
        with open(fp_txt, "w", encoding="utf-8") as f:
            f.write(section)
        zipf.write(fp_txt, arcname=fn_txt)
        filepaths.append((fn_txt, fp_txt, section))

        delimiter = "³"
        lines = [ln for ln in section.splitlines() if delimiter in ln]
        if not lines:
            st.warning(f"Section #{idx}: Tidak ada data (delimiter ³).")
            df = pd.DataFrame()
        else:
            # --- Cari header ASCII --- #
            header_idx = None
            for i, ln in enumerate(lines):
                if "NIK" in ln.upper() and "NAMA" in ln.upper():
                    header_idx = i
                    break
            if header_idx is None:
                header_idx = 0

            # --- Ambil 2 baris header (biarkan header ASCII tidak ikut)
            header1 = lines[header_idx].replace("Â", "").replace("Ã", "").replace("°", "")
            header2 = ""
            if header_idx+1 < len(lines):
                header2 = lines[header_idx+1].replace("Â", "").replace("Ã", "").replace("°", "")

            parts1 = [p.strip() for p in header1.split(delimiter)]
            parts2 = [p.strip() for p in header2.split(delimiter)] if header2 else [""]*len(parts1)
            maxlen = max(len(parts1), len(parts2))
            while len(parts1)<maxlen: parts1.append("")
            while len(parts2)<maxlen: parts2.append("")

            # Gabung header dua baris, lalu buat unik
            raw_cols = []
            for h1, h2 in zip(parts1, parts2):
                if h1 and h2: nm = f"{h1} {h2}".strip()
                else: nm = h1 or h2
                nm = re.sub(r"\s+", " ", nm).strip()
                raw_cols.append(nm if nm else "COL")
            seen = {}
            columns = []
            for col in raw_cols:
                if col not in seen:
                    seen[col]=1
                    columns.append(col)
                else:
                    seen[col] +=1
                    columns.append(f"{col}_{seen[col]}")
            n_cols = len(columns)

            # --- Parsing data
            parsed = []
            for ln in lines[header_idx+2:]:
                if "SUB TOTAL" in ln.upper():
                    continue
                row = ln.replace("Â", "").replace("Ã", "").replace("°", "")
                if row.startswith(delimiter):
                    row = row[len(delimiter):]
                parts = [p.strip().replace("³","") for p in row.split(delimiter)]
                if len(parts)<n_cols:
                    parts += [""]*(n_cols-len(parts))
                elif len(parts)>n_cols:
                    parts = parts[:n_cols]
                parsed.append(parts)

            # --- PATCH: Clean baris & kolom COL --- #
            if not parsed:
                df = pd.DataFrame(columns=columns)
            else:
                df = pd.DataFrame(parsed, columns=columns)

                # Hapus baris jika kolom NIK kosong/None/NIK
                if not df.empty:
                    nik_col = [c for c in df.columns if c.strip().upper() == 'NIK']
                    if nik_col:
                        nik_col = nik_col[0]
                        df = df[~df[nik_col].isin(["", None, "None", "nan", "NIK"])].copy()
                        df.reset_index(drop=True, inplace=True)
                    # Hapus baris yang seluruh kolomnya kosong
                    df = df[~(df.isnull() | (df == '')).all(axis=1)].copy()
                    df.reset_index(drop=True, inplace=True)

                # Split kolom NIK+Nama jadi dua
                if ("Nik" in df.columns or "NIK" in df.columns) and ("Nama" in df.columns or "NAMA" in df.columns):
                    cnik = [c for c in df.columns if c.strip().upper()=="NIK"][0]
                    cnama = [c for c in df.columns if c.strip().upper()=="NAMA"][0]
                    df[cnik] = df[cnik].astype(str).str.extract(r"^([A-Za-z0-9]+)")
                    df[cnama] = df[cnama].astype(str).str.strip()
                # Hapus kolom COL, COL_2 dst jika ada
                df = df[[c for c in df.columns if not c.startswith("COL")]]

        # Simpan ke Excel (sheet Page_{idx})
        sheet_name = f"Page_{idx}" if len(f"Page_{idx}")<=31 else f"Pg_{idx}"
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        preview_dfs.append((sheet_name, df))

        # Format kolom: semua ke text biar tidak ada “angka diubah Excel”
        wb = writer.book
        ws = writer.sheets[sheet_name]
        fmt_text = wb.add_format({"num_format":"@"})
        for i, c in enumerate(df.columns):
            if i==0: ws.set_column(i,i,12,fmt_text)
            elif i==1: ws.set_column(i,i,22,fmt_text)
            else: ws.set_column(i,i,16,fmt_text)

    writer.close()
    zipf.close()

    # ==== TOMBOL DOWNLOAD DITARUH DI ATAS ====
    col1, col2 = st.columns(2)
    with col1:
        with open(zip_path,"rb") as fzip:
            st.download_button("📦 Download Semua File TXT (.zip)", fzip, "hasil_split_gaji_karyawan_txt.zip","application/zip")
    with col2:
        with open(excel_path,"rb") as fexcel:
            st.download_button("📥 Download File Excel (.xlsx)", fexcel, "hasil_split_gaji_karyawan.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ==== PREVIEW SHEET ====
    for sheet_name, df in preview_dfs:
        with st.expander(f"🔍 Preview Sheet '{sheet_name}'"):
            if df.empty:
                st.write("_Sheet kosong_")
            else:
                st.dataframe(df, use_container_width=True)

    # Preview isi file txt mentah (optional, bisa di bawah)
    for fn_txt, fp_txt, content in filepaths:
        with st.expander(f"📄 Preview {fn_txt}"):
            preview = "\n".join(content.splitlines()[:20])
            st.text(preview)
            with open(fp_txt, "rb") as ftxt:
                st.download_button(
                    label=f"⬇️ Download {fn_txt}",
                    data=ftxt,
                    file_name=fn_txt,
                    mime="text/plain"
                )
