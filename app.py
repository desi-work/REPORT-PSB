import streamlit as st
import pandas as pd
import io

st.title("📄 Import Laporan TXT Pelanggan (Multi Laporan per File)")

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

if uploaded_files:
    records = []
    for uploaded_file in uploaded_files:
        content = uploaded_file.read()
        parsed_records = parse_txt_file_multiple_reports(content, uploaded_file.name)
        records.extend(parsed_records)

    df = pd.DataFrame(records)

    # Format Nama Pelanggan jadi kapital
    df["Nama Pelanggan"] = df["Nama Pelanggan"].str.title()

    # Ubah Tanggal ke format datetime agar bisa disortir
    df["Tanggal Visit"] = pd.to_datetime(df["Tanggal Visit"], dayfirst=True, errors='coerce')

    # Urutkan berdasarkan Tanggal, No.hasbel, Nama Pelanggan
    df = df.sort_values(by=["Tanggal Visit", "No.hasbel", "Nama Pelanggan"], ascending=[True, True, True])

    # Reset index agar nomor urut rapi
    df = df.reset_index(drop=True)

    # Tambahkan kolom No (nomor urut dari 1)
    df.insert(0, "No", range(1, len(df) + 1))

    # Pastikan angka bisa dihitung
    df["Meteran Awal"] = pd.to_numeric(df["Meteran Awal"], errors='coerce')
    df["Meteran Akhir"] = pd.to_numeric(df["Meteran Akhir"], errors='coerce')

    # Hitung selisih (tarikan)
    df["Total Tarikan"] = df["Meteran Awal"] - df["Meteran Akhir"]

    # Pindahkan kolom "Total Tarikan" setelah "Meteran Akhir"
    kolom_baru = df.columns.tolist()
    kolom_baru.insert(kolom_baru.index("Meteran Akhir") + 1, kolom_baru.pop(kolom_baru.index("Total Tarikan")))
    df = df[kolom_baru]

    # Format tanggal jadi '10 May 2025'
    df["Tanggal Visit"] = df["Tanggal Visit"].dt.strftime("%d %B %Y")

    st.subheader("📊 Hasil Tabel")
    st.dataframe(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        label="💾 Download Excel",
        data=output.getvalue(),
        file_name="laporan_kunjungan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
