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
    zip_path = os.path.join(tmpdirname, "hasil_split_gaji_karyawan_txt.zip")
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

                # FILTER DATA: hanya baris yang punya NIK, Nama, dan data lain yang terisi
                nik_col = [c for c in df.columns if c.strip().upper() == 'NIK']
                nama_col = [c for c in df.columns if c.strip().upper() == 'NAMA']
                if nik_col and nama_col:
                    nik_col = nik_col[0]
                    nama_col = nama_col[0]
                    
                    # -------- PATCH BARU: Ambil NIK & Ekstra Kode (angka setelah spasi) --------
                    nik_asli = df[nik_col].astype(str)
                    # NIK (tanpa kode)
                    df[nik_col] = nik_asli.str.extract(r"^([A-Za-z0-9]+)")
                    # Kode = angka/kode setelah spasi
                    df['Kode'] = nik_asli.str.extract(r"^[A-Za-z0-9]+\s+(\d+)")
                    df['Kode'] = df['Kode'].fillna("")
                    df[nama_col] = df[nama_col].astype(str).str.strip()

                    # Hanya baris valid
                    df = df[
                        df[nik_col].notna() &
                        (df[nik_col].astype(str).str.strip() != "") &
                        (df[nik_col].astype(str).str.upper() != "NIK") &
                        (df[nama_col].astype(str).str.strip() != "") &
                        (df[nama_col].astype(str).str.upper() != "NAMA") &
                        (df.drop([nik_col, nama_col], axis=1).apply(lambda row: any([str(x).strip() != "" for x in row]), axis=1))
                    ].copy()
                    df.reset_index(drop=True, inplace=True)

                    # --- Susun ulang kolom agar Kode setelah NIK ---
                    cols = df.columns.tolist()
                    if 'Kode' in cols and nik_col in cols:
                        nik_idx = cols.index(nik_col)
                        cols.remove('Kode')
                        cols.insert(nik_idx+1, 'Kode')
                        df = df[cols]
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

    # ==== TOMBOL DOWNLOAD UTAMA (pakai nama default, tanpa rename) ====
    col1, col2 = st.columns(2)
    with col1:
        with open(excel_path,"rb") as fexcel:
            st.download_button(
                "📥 Download File Excel (.xlsx)",
                fexcel,
                file_name="hasil_split_gaji_karyawan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col2:
        with open(zip_path,"rb") as fzip:
            st.download_button(
                "📦 Download Semua File TXT (.zip)",
                fzip,
                file_name="hasil_split_gaji_karyawan_txt.zip",
                mime="application/zip"
            )

    # ==== PREVIEW SHEET + DOWNLOAD TIAP SHEET (ADA INPUT RENAME) ====
    for sheet_name, df in preview_dfs:
        with st.expander(f"🔍 Preview Sheet '{sheet_name}'"):
            if df.empty:
                st.write("_Sheet kosong_")
            else:
                st.dataframe(df, use_container_width=True)
                single_excel_name = st.text_input(
                    f"Nama file untuk sheet '{sheet_name}' (.xlsx)", value=f"{sheet_name}.xlsx", key=f"fname_{sheet_name}"
                )
                single_excel_path = os.path.join(tmpdirname, f"{sheet_name}.xlsx")
                with pd.ExcelWriter(single_excel_path, engine="xlsxwriter") as single_writer:
                    df.to_excel(single_writer, index=False, sheet_name=sheet_name)
                    wb = single_writer.book
                    ws = single_writer.sheets[sheet_name]
                    fmt_text = wb.add_format({"num_format": "@"})
                    for i, c in enumerate(df.columns):
                        if i==0: ws.set_column(i,i,12,fmt_text)
                        elif i==1: ws.set_column(i,i,22,fmt_text)
                        else: ws.set_column(i,i,16,fmt_text)
                with open(single_excel_path, "rb") as single_excel_file:
                    st.download_button(
                        label=f"⬇️ Download Sheet '{sheet_name}' (.xlsx)",
                        data=single_excel_file,
                        file_name=single_excel_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

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
