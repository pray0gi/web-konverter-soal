import requests # Tambahkan ini di paling atas bersama import lainnya
import urllib.parse # Tambahkan ini di deretan import atas
import streamlit as st
import docx
import pandas as pd
import re
import os
import shutil
from docx.oxml.ns import qn
import io # <-- Tambahkan ini untuk memanipulasi file di memori

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

# --- MENU SIDEBAR: DOWNLOAD TEMPLATE ---
with st.sidebar:
    st.header("📄 Format Standar Soal")
    st.markdown("Agar aplikasi dapat membaca soal dengan benar, pastikan Anda menggunakan template Word yang telah disediakan.")
    
    # Fungsi untuk membuat file Word di memori (BytesIO)
    def buat_template_memori():
        doc = docx.Document()
        doc.add_heading('Template Format Soal Ujian', 0)
        doc.add_paragraph('Petunjuk: Gunakan format di bawah ini dengan presisi. Jangan ubah pola penomoran (1., 2.) atau format huruf opsi (A., B.).')
        
        doc.add_paragraph('1. Tulis Soal disini')
        doc.add_paragraph('A. jawabanA')
        doc.add_paragraph('B. jawabanB')
        doc.add_paragraph('C. jawabanC')
        doc.add_paragraph('D. jawabanD')
        doc.add_paragraph('E. jawabanE')
        doc.add_paragraph('Kunci: A\n')
        
        doc.add_paragraph('2. Tulis Soal disini')
        doc.add_paragraph('A. jawabanA')
        doc.add_paragraph('B. jawabanB')
        doc.add_paragraph('C. jawabanC')
        doc.add_paragraph('D. jawabanD')
        doc.add_paragraph('E. jawabanE')
        doc.add_paragraph('Kunci: B\n')

        doc.add_paragraph('3. Dst')
        
        # Menyimpan ke memori (buffer) bukan ke harddisk
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    file_template = buat_template_memori()
    
    # Tombol Download
    st.download_button(
        label="⬇️ Unduh Template Word (.docx)",
        data=file_template,
        file_name="Template_Soal_Ujian.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )
    st.write("---")

st.title("📝 Web Konverter Soal Word ke Form")
st.markdown("Aplikasi untuk mengubah format soal ujian `.docx` menjadi `Google Form` lengkap dengan ekstraksi gambarnya.")

file_word = st.file_uploader("Unggah File Word (.docx)", type=["docx"])

if file_word is not None:
    st.write("---") 
    
    # Membuat 2 kolom sejajar untuk input
    col1, col2 = st.columns(2)
    with col1:
        nama_form = st.text_input("Nama File Form:", "Soal Ujian Teknik Ketenagalistrikan")
        poin_soal = st.number_input("Poin per soal benar:", min_value=1, max_value=100, value=10)
    with col2:
        nomor_wa = st.text_input("Nomor WA Tujuan (Opsional):", placeholder="Contoh: 081234567890")
    
    st.write("") 
    tombol_konversi = st.button("🚀 BUAT GOOGLE FORM SEKARANG", type="primary", use_container_width=True)
    
    if tombol_konversi:
        with st.spinner("Sedang membaca dokumen dan mengirim ke Google..."):
            df_hasil = proses_konversi(file_word, "temp_gambar_soal")
            data_json = df_hasil.to_dict(orient='records')
            
            # Menambahkan nama_form ke dalam payload
            payload_api = {
                "poin": poin_soal,
                "nama_form": nama_form,
                "soal": data_json
            }
            
            URL_GAS = "https://script.google.com/macros/s/AKfycbxRUyFNd37oae91glDYewdP3-TFggDysiG1nSNOGnkXbGot0LHI9TmhSe8QKgLk8_Fgpw/exec"
            
            try:
                respon = requests.post(URL_GAS, json=payload_api)
                hasil_api = respon.json()
                
                if hasil_api.get("status") == "sukses":
                    st.success(f"🎉 Berhasil! Form '{nama_form}' sudah siap.")
                    
                    link_edit = hasil_api['edit_url']
                    link_siswa = hasil_api['publish_url']
                    
                    st.markdown(f"**🔗 [Klik di sini untuk Edit Form]({link_edit})**")
                    st.markdown(f"**👀 [Klik di sini untuk Lihat Tampilan Form]({link_siswa})**")
                    
                    # LOGIKA PENGIRIMAN WHATSAPP
                    if nomor_wa:
                        pesan_wa = f"Halo! Google Form untuk *{nama_form}* sudah selesai dibuat.\n\n*Link Edit (Guru):*\n{link_edit}\n\n*Link Kuis (Siswa):*\n{link_siswa}"
                        
                        # Jika bot WhatsApp yang Anda jalankan dengan PM2 memiliki endpoint API masuk, 
                        # Anda bisa menghilangkan tanda pagar di bawah ini dan menyesuaikan URL-nya:
                        # requests.post("http://localhost:3000/api/sendText", json={"chatId": f"{nomor_wa}@c.us", "text": pesan_wa})
                        
                        # Alternatif instan: Generate link Click-to-Chat WhatsApp Web
                        if nomor_wa.startswith("08"):
                            nomor_wa = "628" + nomor_wa[2:] # Format ke 62
                        
                        pesan_terencode = urllib.parse.quote(pesan_wa)
                        link_wa_web = f"https://wa.me/{nomor_wa}?text={pesan_terencode}"
                        
                        st.info("Form berhasil dibuat. Klik tombol di bawah ini untuk meneruskan link via WhatsApp.")
                        st.markdown(f"**📱 [KIRIM LINK VIA WHATSAPP]({link_wa_web})**")
                        
                else:
                    st.error(f"Terjadi kesalahan di server Google: {hasil_api.get('pesan')}")
                    
            except Exception as e:
                st.error(f"Gagal menghubungi server: {e}")
