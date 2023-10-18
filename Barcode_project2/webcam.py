# Kelompok 8 Tugas Besar Pengenalan Komputasi
# Anggota :
# Devon Josiah	                  16523081
# Muhammad Jibril Ibrahim	      19623277
# Luckman Fakhmanidris Arvasirri  19623053
# Hanif Muhammad Assegaf	      16523088
# Mohammad Najmutstsaqib	      16523228

# Deskripsi :
# Aplikasi Absensi kehadiran mahasiswa menggunakan qr code. 
# Aplikasi ini akan mengeluarkan file excel yang berisi nama nim mahasiswa beserta 
# kehadiran dan waktu absensi mereka.
# Aplikasi ini kami batasi untuk lingkup mahasiswa yang mengikuti kelas PRD 16 saja 
# namun dapat dikembangkan sesuai jumlah data mahasiswa yang diberikan

# KAMUS
# fakultas_path, kelasPRD_path, kelasPRDexcel_path, barcode_ico_path : string = berupa path untuk masing masing file
# df_kehadiran : pandas.dataframe = dataframe pandas yang berisi nama, nim, kehadiran, dan waktu absensi mahasiswa
# root : tkinter = sebuah class untuk membuat user interface saat menjalankan python
# app : WebcamApp = sebuah class untuk mengambil video dari kamera laptop

# import modul yang dibutuhkan
import cv2                          # modul ini untuk menangkap kamera video dari laptop
import tkinter as tk                # modul ini untuk membuat gui
import time                         # modul ini untuk menghentikan sejenak jalan program
import datetime as dt               # modul ini untuk mendapatkan waktu
import pandas as pd                 # modul ini untuk megatur dan mengolah data dari excel dan untuk membuat excel
import beepy                        # modul ini untuk membuat suara
from pyzbar.pyzbar import decode    # modul ini untuk mendecode qr code menjadi text
from PIL import Image, ImageTk      # modul ini untuk membuat gui menampilkan video kamera yang telah ditangkap

# variabel global
kelasPRDexcel_path = "Barcode_project2/kelasPRD.xlsx"    # relative path untuk kelasPRD.xlsx
barcode_ico_path = "Barcode_project2/Barcode.ico"        # relative path untuk Barcode.ico

df_kehadiran = pd.read_excel(kelasPRDexcel_path,index_col="NIM")   # dataframe dari file kelasPRD.xlsx 

class WebcamApp:                                     # class untuk menangkap video kamera
    def __init__(self, window, window_title):        # fungsi yang dijalankan saat inisiasi kelas
        # KAMUS LOKAL
        # self : WebcamApp = untuk merujuk ke class ini sendiri
        # window : tkinter = berupa gui untuk menampilkan video kamera
        # window_title : string = nama dari gui yang akan ditampilkan

        self.window = window                        # kita masukkan video kamera ke gui
        self.window.title(window_title)             # kita buat judul gui sesuai dengan window_title
        self.window.iconbitmap(barcode_ico_path)    # kita akan membuat ikon gui sesuai dengan ikon yang kita mau

        self.vid = cv2.VideoCapture(0)              # kita gunakan fungsi dari modul cv2 untuk mengambil video kamera laptop

        self.canvas = tk.Canvas(window, width=self.vid.get(3), height=self.vid.get(4))    # kita mengatur besar gui yang ditampilkan saat menjalankan aplikasi
        self.canvas.pack()                          # kita tampilkan gui ke layar

        self.update()                     # panggil fungsi update()
        self.window.mainloop()            # fungsi ini agar aplikasi berjalan terus selama tidak di close

    # fungsi untuk mengupdate gui agar gui yang ditampilkan sesuai dengan kamera laptop
    def update(self):
        ret, frame = self.vid.read()

        if ret:
            self.photo = ImageTk.PhotoImage(image=Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)))
            self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

        for i in decode(frame):
            value = i.data.decode('utf-8')      # kita decode qr code menjadi text utf-8
            time.sleep(0.5)                     # kita hentikan program sejenak agar tidak terus menerus mendecode qr code

            output = True
            try:                                # tes apakah data dari qr code berupa sebuah angka
                angka = int(value)
            except:
                output = False

            if len(value) == 16 and output:                   # cek apakah data yg didapatkan sesuai yang diharapkan
                if int(value[:8]) in df_kehadiran.index:      # cek apakah nim termasuk pada list di excel
                    df_kehadiran['Status'][int(value[:8])] = "Hadir"
                    df_kehadiran['Waktu Absensi'][int(value[:8])] = dt.datetime.now().strftime("%H:%M:%S")
                    beepy.beep(sound=5)
                    print("Berhasil hooray")

        self.window.after(1, self.update)

# membuat window gui lalu dimasukkan ke WebcamApp

root = tk.Tk()
app = WebcamApp(root, "QR Scanner")

#bikin pdf berisi absensi sesuai ktm
df_kehadiran.to_excel(f'Barcode_project2/excel absensi/Absensi Kehadiran PRD {dt.datetime.now().strftime("%d-%m-%Y pada %H.%M.%S")}.xlsx')
