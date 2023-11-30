import streamlit as st
from streamlit_qrcode_scanner import qrcode_scanner

import xlwings as xw
from datetime import datetime


def scan(worksheet, peserta):
    nama = qrcode_scanner(key="qrcode_scanner")
    if nama:
        try:
            id = peserta.index(nama)

            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

            if worksheet["B" + str(2 + id)].value == "HADIR":
                st.warning(
                    nama
                    + " -- SUDAH HADIR SEJAK -- "
                    + str(worksheet["C" + str(2 + id)].value)
                )
            else:
                worksheet["B" + str(2 + id)].value = "HADIR"
                worksheet["C" + str(2 + id)].value = dt_string
                st.success(" " + nama + " -- HADIR --" + dt_string,
                           icon='✅')

        except:
            st.error(" " + nama + " -- TIDAK TERDAFTAR", icon='🚨')