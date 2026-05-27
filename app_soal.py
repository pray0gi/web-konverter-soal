import requests # Tambahkan ini di paling atas bersama import lainnya
import streamlit as st
import docx
import pandas as pd
import re
import os
import shutil
from docx.oxml.ns import qn

# --- FUNGSI ENGINE PARSER ---
def ekstrak_gambar_dari_paragraf(paragraph, doc, folder_output, nama_file_basis):
    list_gambar = []
    counter = 1
    blips = paragraph._element.xpath('.//a:blip')
    for blip in blips:
        rId = blip.get(qn('r:embed'))
        if rId and rId in doc.part.related_parts:
            image_part = doc.part.related_parts[rId]
            image_data = image_part.blob
            ekstensi = image_part.content_type.split('/')[-1]
            nama_gambar = f"{nama_file_basis}_{counter}.{ekstensi}"
            path_gambar = os.path.join(folder_output, nama_gambar)
            
            with open(path_gambar, 'wb') as f:
                f.write(image_data)
            list_gambar.append(path_gambar)
            counter += 1
    return ";".join(list_gambar) if list_gambar else ""

def proses_konversi(file_upload, folder_gambar="gambar_soal"):
    if not os.path.exists(folder_gambar):
        os.makedirs(folder_gambar)
        
    doc = docx.Document(file_upload)
    data_soal = []
    soal_kosong = {
        "Pertanyaan": "", "Gambar_Pertanyaan": "",
        "A": "", "B": "", "C": "", "D": "", "E": "", "Kunci": ""
    }
    soal_sementara = soal_kosong.copy()
    id_soal = 0
    
    for para in doc.paragraphs:
        teks = para.text.strip()
        miliki_gambar = len(para._element.xpath('.//a:blip')) > 0
        if not teks and not miliki_gambar:
            continue
            
        if re.match(r'^\d+\.', teks):
            if soal_sementara["Pertanyaan"]:
                data_soal.append(soal_sementara.copy())
            id_soal += 1
            soal_sementara = soal_kosong.copy()
            soal_sementara["Pertanyaan"] = teks
            if miliki_gambar:
                soal_sementara["Gambar_Pertanyaan"] = ekstrak_gambar_dari_paragraf(
                    para, doc, folder_gambar, f"soal_{id_soal}_tanya"
                )
        elif re.match(r'^[A-E]\.', teks, re.IGNORECASE):
            huruf = teks[0].upper()
            soal_sementara[huruf] = teks[2:].strip()
        elif teks.lower().startswith("kunci:"):
            soal_sementara["Kunci"] = teks.split(":")[1].strip().upper()
        elif id_soal > 0 and not re.match(r'^[A-E]\.', teks, re.IGNORECASE) and not teks.lower().startswith("kunci:"):
            if teks:
                soal_sementara["Pertanyaan"] += "\n" + teks
            if miliki_gambar:
                path_gbr = ekstrak_gambar_dari_paragraf(para, doc, folder_gambar, f"soal_{id_soal}_tanya_tambahan")
                soal_sementara["Gambar_Pertanyaan"] += (";" + path_gbr) if soal_sementara["Gambar_Pertanyaan"] else path_gbr

    if soal_sementara["Pertanyaan"]:
        data_soal.append(soal_sementara)
        
    return pd.DataFrame(data_soal)

# --- ANTARMUKA STREAMLIT ---
st.set_page_config(page_title="Konverter Soal Ujian", layout="wide")

st.title("📝 Web Konverter Soal Word ke Excel")
st.markdown("Aplikasi untuk mengubah format soal ujian `.docx` menjadi `.xlsx` lengkap dengan ekstraksi gambarnya.")

file_word = st.file_uploader("Unggah File Word (.docx)", type=["docx"])

if file_word is not None:
    st.write("") 
    tombol_konversi = st.button("🚀 BUAT GOOGLE FORM SEKARANG", type="primary", use_container_width=True)
    
    if tombol_konversi:
        with st.spinner("Sedang membaca dokumen dan mengirim ke Google..."):
            # 1. Ekstrak Word menjadi Dataframe seperti biasa
            df_hasil = proses_konversi(file_word, "temp_gambar_soal")
            
            # 2. Ubah Dataframe menjadi format JSON (Dictionary Python)
            data_json = df_hasil.to_dict(orient='records')
            
            # 3. Kirim ke URL GAS Web App Anda
            URL_GAS = "https://script.google.com/macros/s/AKfycbxRUyFNd37oae91glDYewdP3-TFggDysiG1nSNOGnkXbGot0LHI9TmhSe8QKgLk8_Fgpw/exec"
            
            try:
                # Mengirim POST request
                respon = requests.post(URL_GAS, json=data_json)
                hasil_api = respon.json()
                
                if hasil_api.get("status") == "sukses":
                    st.success("🎉 Berhasil! Google Form Anda sudah siap.")
                    
                    # Menampilkan Link ke Pengguna
                    st.markdown(f"**🔗 [Klik di sini untuk Edit Form]({hasil_api['edit_url']})**")
                    st.markdown(f"**👀 [Klik di sini untuk Lihat Tampilan Form]({hasil_api['publish_url']})**")
                else:
                    st.error(f"Terjadi kesalahan di server Google: {hasil_api.get('pesan')}")
                    
            except Exception as e:
                st.error(f"Gagal menghubungi server: {e}")
