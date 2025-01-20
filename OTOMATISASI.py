import smtplib
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
from tkinter import filedialog, messagebox
import schedule
import threading
import time
from datetime import datetime

# Fungsi untuk mengirim email
def kirim_email():
    smtp_server = "smtp.mailtrap.io"  # Server SMTP tetap
    smtp_port = 2525  # Port SMTP tetap
    username = entry_username.get()
    password = entry_password.get()
    pengirim = entry_pengirim.get()
    penerima = entry_penerima.get()
    subjek = entry_subjek.get()
    isi = text_isi.get("1.0", "end-1c")
    nama_lampiran = entry_nama_lampiran.get()
    path_lampiran = entry_path_lampiran.get()

    try:
        msg = MIMEMultipart()
        msg['Subject'] = subjek
        msg['From'] = pengirim
        msg['To'] = penerima
        msg.attach(MIMEText(isi, 'plain'))

        if path_lampiran:
            try:
                with open(path_lampiran, 'rb') as f:
                    lampiran = MIMEApplication(f.read(), Name=nama_lampiran)
                lampiran['Content-Disposition'] = f'attachment; filename="{nama_lampiran}"'
                msg.attach(lampiran)
            except Exception as e:
                messagebox.showerror("Error", f"Terjadi kesalahan saat melampirkan file: {e}")
                return

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(username, password)
        server.sendmail(pengirim, penerima, msg.as_string())
        server.quit()

        messagebox.showinfo("Sukses", "Email berhasil dikirim!")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")

# Fungsi untuk menjalankan penjadwalan
def jadwalkan_pengiriman():
    waktu = entry_jadwal.get()
    try:
        jadwal_datetime = datetime.strptime(waktu, "%Y-%m-%d %H:%M:%S")
        sekarang = datetime.now()
        delay = (jadwal_datetime - sekarang).total_seconds()
        if delay <= 0:
            messagebox.showerror("Error", "Waktu penjadwalan harus di masa depan.")
            return

        threading.Timer(delay, kirim_email).start()
        messagebox.showinfo("Penjadwalan", f"Email dijadwalkan untuk dikirim pada {waktu}")
    except ValueError:
        messagebox.showerror("Error", "Format waktu tidak valid. Gunakan format YYYY-MM-DD HH:MM:SS")

# Fungsi untuk memilih file lampiran
def pilih_lampiran():
    path = filedialog.askopenfilename(title="Pilih Lampiran")
    entry_path_lampiran.delete(0, tk.END)
    entry_path_lampiran.insert(0, path)

# Membuat GUI dengan Tkinter
root = tk.Tk()
root.title("Aplikasi Pengiriman Email dengan Penjadwalan")

label_username = tk.Label(root, text="Username SMTP:")
label_username.grid(row=0, column=0)
entry_username = tk.Entry(root, width=40)
entry_username.grid(row=0, column=1)

label_password = tk.Label(root, text="Password SMTP:")
label_password.grid(row=1, column=0)
entry_password = tk.Entry(root, width=40, show="*")
entry_password.grid(row=1, column=1)

label_pengirim = tk.Label(root, text="Pengirim:")
label_pengirim.grid(row=2, column=0)
entry_pengirim = tk.Entry(root, width=40)
entry_pengirim.grid(row=2, column=1)

label_penerima = tk.Label(root, text="Penerima:")
label_penerima.grid(row=3, column=0)
entry_penerima = tk.Entry(root, width=40)
entry_penerima.grid(row=3, column=1)

label_subjek = tk.Label(root, text="Subjek:")
label_subjek.grid(row=4, column=0)
entry_subjek = tk.Entry(root, width=40)
entry_subjek.grid(row=4, column=1)

label_isi = tk.Label(root, text="Isi Pesan:")
label_isi.grid(row=5, column=0)
text_isi = tk.Text(root, height=5, width=40)
text_isi.grid(row=5, column=1)

label_nama_lampiran = tk.Label(root, text="Nama Lampiran:")
label_nama_lampiran.grid(row=6, column=0)
entry_nama_lampiran = tk.Entry(root, width=40)
entry_nama_lampiran.grid(row=6, column=1)

label_path_lampiran = tk.Label(root, text="Path Lampiran:")
label_path_lampiran.grid(row=7, column=0)
entry_path_lampiran = tk.Entry(root, width=40)
entry_path_lampiran.grid(row=7, column=1)

button_pilih_lampiran = tk.Button(root, text="Pilih Lampiran", command=pilih_lampiran)
button_pilih_lampiran.grid(row=7, column=2)

label_jadwal = tk.Label(root, text="Jadwal (YYYY-MM-DD HH:MM:SS):")
label_jadwal.grid(row=8, column=0)
entry_jadwal = tk.Entry(root, width=40)
entry_jadwal.grid(row=8, column=1)

button_jadwal = tk.Button(root, text="Jadwalkan Pengiriman", command=jadwalkan_pengiriman)
button_jadwal.grid(row=9, column=1)

button_kirim = tk.Button(root, text="Kirim Email Sekarang", command=kirim_email)
button_kirim.grid(row=10, column=1)

root.mainloop()
