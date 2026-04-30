import os
import sys
import pandas as pd
from docxtpl import DocxTemplate
import ttkbootstrap as tb
from tkinter import messagebox, filedialog, ttk

# Fungsi mendapatkan direktori utama program
def get_base_path():
    if getattr(sys, 'frozen', False):  # Jika program dalam mode EXE
        return os.path.dirname(sys.executable)  # Direktori tempat EXE berada
    return os.path.dirname(os.path.abspath(__file__))

# Path file
BASE_PATH = get_base_path()
TEMPLATES_PATH = os.path.join(BASE_PATH, "templates")  # Path absolut templates
OUTPUT_PATH = os.path.join(os.path.expanduser("~"), "Documents", "Generated_Documents")  # Simpan di folder Documents

# Buat folder output jika belum ada
os.makedirs(OUTPUT_PATH, exist_ok=True)





# Fungsi untuk membuat dokumen
def create_document():
    data = {key: entries[key].get().strip() for key in fields}  # Ambil data dari field & hapus spasi awal/akhir

    # Validasi Input: Pastikan semua field terisi
    empty_fields = [key for key, value in data.items() if value == ""]
    if empty_fields:
        messagebox.showwarning("Peringatan", f"Harap isi semua field sebelum generate dokumen!\n\nField kosong: {', '.join(empty_fields)}")
        return  # Stop eksekusi jika ada field kosong

    # Minta pengguna memilih direktori output
    output_dir = filedialog.askdirectory(title="Pilih Folder untuk Menyimpan Dokumen", initialdir=OUTPUT_PATH)
    if not output_dir:  # Jika pengguna membatalkan dialog
        return

    # Pastikan nama tersedia untuk nama file
    nama_file = data.get('Nama', 'Dokumen')
    
    for template_name in ["FORM_BPKB.docx", "FORM_STNK.docx", "LEMBAR_FISIK.docx"]:
        template_path = os.path.join(TEMPLATES_PATH, template_name)

        if not os.path.exists(template_path):  # Jika template tidak ditemukan
            messagebox.showerror("Error", f"Template {template_name} tidak ditemukan di {TEMPLATES_PATH}")
            return

        doc = DocxTemplate(template_path)
        doc.render(data)

        output_file = os.path.join(output_dir, f"{nama_file}_{template_name}")
        doc.save(output_file)
        print(f"✅ Dokumen berhasil dibuat: {output_file}")

    messagebox.showinfo("Sukses", f"Dokumen berhasil dibuat di folder:\n{output_dir}")

# Membuat GUI dengan ttkbootstrap
root = tb.Window(themename="darkly")
root.title("Pembuatan Dokumen Otomatis")
# root.state("zoomed")  # Fullscreen otomatis
try:
    root.state("zoomed")  # Windows
except:
    root.attributes("-zoomed", True)  # Linux fallback


# Atur style tombol
style = tb.Style()
style.configure("TButton", font=("Bookman Old Style", 16, "bold"))  # Ukuran font tombol diperbesar

# Frame utama untuk form
main_frame = tb.Frame(root)
main_frame.pack(fill="both", expand=True, padx=50, pady=20)

left_frame = tb.Frame(main_frame)
left_frame.pack(side="left", expand=True, fill="both", padx=20)

middle_frame = tb.Frame(main_frame)
middle_frame.pack(side="left", expand=True, fill="both", padx=20)

right_frame = tb.Frame(main_frame)
right_frame.pack(side="left", expand=True, fill="both", padx=20)

# Daftar field
fields = [
    'nrkb_urutan', 'Nama', 'alamat', 'kota', 'no_telpon', 'nik', 'merek', 'tipe',
    'jenis', 'model', 'tahun_pembuatan', 'isi_silinder', 'warna_kendaraan', 'no_rangka',
    'no_mesin', 'bahan_bakar', 'jmlh_roda', 'no_sut', 'no_srut', 'warna_tnkb', 'tahun_registrasi'
]

entries = {}
split_1 = len(fields) // 3
split_2 = split_1 * 2

for i, field in enumerate(fields):
    if i < split_1:
        parent_frame = left_frame
    elif i < split_2:
        parent_frame = middle_frame
    else:
        parent_frame = right_frame

    lbl = tb.Label(parent_frame, text=field.capitalize(), font=("Bookman Old Style", 14, "bold"))
    lbl.pack(anchor="w", pady=5)

    ent = tb.Entry(parent_frame, width=40, font=("Bookman Old Style", 14))
    ent.pack(fill="x", padx=10, pady=5)
    entries[field] = ent

# Tombol Generate
btn_create = tb.Button(root, text="Generate Dokumen", command=create_document, bootstyle="success")
btn_create.pack(pady=30)

# Jalankan aplikasi
root.mainloop()