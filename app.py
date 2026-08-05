import streamlit as st
import pandas as pd
import io
import re

st.title("REPORT PSB")

uploaded_files = st.file_uploader("Upload satu atau beberapa file .txt", type="txt", accept_multiple_files=True)

def parse_txt_file_multiple_reports(file_content, source_file_name):
    text = file_content.decode("utf-8")
    sections = text.split("Detail Laporan")

    key_map = {
        "tanggal visit": "Tanggal Visit",
        "nama pelanggan": "Nama Pelanggan",
        "nama teknisi": "Nama Teknisi",
        "no.hasbel": "No.hasbel",
        "meteran awal": "Meteran Awal",
        "meteran akhir": "Meteran Akhir",
        "note": "Note"
    }

    records = []
    for section in sections:
        if ":" not in section:
            continue
        data = {v: None for v in key_map.values()}
        meteran_awal_count = 0
        lines = section.strip().splitlines()
        for line in lines:
            if ':' in line:
                raw_key, value = line.split(':', 1)
                key = raw_key.strip().lower()
                key = key.replace(":", "").strip()
                value = value.strip()
                if key == "meteran awal":
                    meteran_awal_count += 1
                    if meteran_awal_count == 1:
                        data["Meteran Awal"] = value
                    elif meteran_awal_count == 2:
                        data["Meteran Akhir"] = value
                elif key in key_map:
                    data[key_map[key]] = value
        data["Source File"] = source_file_name
        records.append(data)
    return records

# Fungsi cerdas untuk menyeragamkan format Husbel (HV, FA, ZL, LK, FT, dll)
def standarisasi_husbel(kode):
    if pd.isna(kode):
        return kode
    
    kode = str(kode).upper() # Jadikan kapital semua
    kode = re.sub(r'\s+', '', kode) # Hilangkan semua spasi
    
    # 1. Atasi tulisan dobel untuk SEMUA kode (misal: HV-HV-001, FA-FA-001)
    kode = re.sub(r'^([A-Z]+)-\1-', r'\1-', kode)
    
    # 2. Tambahkan tanda hubung otomatis jika lupa (misal: FA001 jadi FA-001)
    if '-' not in kode:
        kode = re.sub(r'^([A-Z]+)(\d+)', r'\1-\2', kode)
        
    return kode

if uploaded_files:
    records = []
    for uploaded_file in uploaded_files:
        content = uploaded_file.read()
        parsed_records = parse_txt_file_multiple_reports(content, uploaded_file.name)
        records.extend(parsed_records)

    df = pd.DataFrame(records)

    # Format Nama Pelanggan jadi kapital
    df["Nama Pelanggan"] = df["Nama Pelanggan"].str.title()

    # Ubah Tanggal ke format datetime
    df["Tanggal Visit"] = pd.to_datetime(df["Tanggal Visit"], dayfirst=True, errors='coerce')

    # [OTOMATISASI KODE HUSBEL]
    df["No.hasbel"] = df["No.hasbel"].apply(standarisasi_husbel)

    # Pastikan angka bisa dihitung (Wajib dilakukan SEBELUM sorting)
    df["Meteran Awal"] = pd.to_numeric(df["Meteran Awal"], errors='coerce')
    df["Meteran Akhir"] = pd.to_numeric(df["Meteran Akhir"], errors='coerce')

    # Hitung selisih (tarikan)
    df["Total Tarikan"] = df["Meteran Awal"] - df["Meteran Akhir"]

    # UBAH SORTIR DISINI: Urutkan by No.hasbel, lalu Meteran Awal Terbesar ke Terkecil
    df = df.sort_values(by=["No.hasbel", "Meteran Awal"], ascending=[True, False])

    # Reset index
    df = df.reset_index(drop=True)

    # Tambahkan kolom No (nomor urut dari 1)
    df.insert(0, "No", range(1, len(df) + 1))

    # Pindahkan kolom "Total Tarikan" setelah "Meteran Akhir"
    kolom_baru = df.columns.tolist()
    kolom_baru.insert(kolom_baru.index("Meteran Akhir") + 1, kolom_baru.pop(kolom_baru.index("Total Tarikan")))
    df = df[kolom_baru]

    # Format tanggal jadi '10 May 2025' (Dilakukan di akhir agar tidak mengganggu sortir)
    df["Tanggal Visit"] = df["Tanggal Visit"].dt.strftime("%d %B %Y")

    st.subheader("📊 Hasil Tabel")
    st.dataframe(df)

    # Ekspor ke Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        label="💾 Download Excel",
        data=output.getvalue(),
        file_name="laporan_kunjungan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
