import pandas as pd

import streamlit as st
import xlwings as xw

import generate_qr_code
import scan_streamlit


# read and prepare the data
wb = xw.Book("data.xlsx")
worksheet = wb.sheets("Sheet1")

peserta = worksheet["A2"].expand("down").value

total_peserta = len(peserta)

scan, generate, stop = st.tabs(["Scan QR Code", "Generate QR Code", "Stop the Program"])

with stop:
    if st.button("Stop the Program"):
        st.success("Program stopped")
        st.stop()

with scan:
    "# Scan QR Code"
    if st.checkbox("Open Camera"):
        # Panggil fungsi scan_qr_code untuk memulai pemindaian
        scan_streamlit.scan(worksheet, peserta)

with generate:
    "# Generate New QR Code"
    # Data yang ingin diubah menjadi QR Code
    nama = st.text_input("Write you name", "")

    generate, delete = st.columns(2)

    with generate:
        if st.button("Sign Up and Generate QR Code"):
            # Tulis data ke excel
            if nama not in peserta:
                filename = nama + ".png"  # Nama file untuk menyimpan QR Code

                # Generate QR Code
                img = generate_qr_code.generate_qr_code(nama, filename)
                img.save("qrcode/" + filename)

                total_peserta += 1
                worksheet["A" + str(total_peserta + 1)].value = nama

                st.success(f"{nama} berhasil terdaftar")
                st.image("qrcode/" + filename, caption=filename)
            else:
                st.warning(f"{nama} sudah terdaftar")

    with delete:
        if st.button("Delete Name"):
            row = str(peserta.index(nama) + 2)
            for col in ('A', 'B', 'C'):
                worksheet.range(col + row).delete(shift='up')
            peserta.remove(nama)
            total_peserta -= 1

            st.success(nama + " berhasil dihapus.")
