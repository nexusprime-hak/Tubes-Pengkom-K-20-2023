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
import cv2
import tkinter as tk
import json
import time
import datetime as dt
import pandas as pd
import beepy
from pyzbar.pyzbar import decode
from PIL import Image, ImageTk

# variabel global
fakultas_path = "Barcode_project2/fakultas.json"         # relative path untuk fakultas.json
kelasPRD_path = "Barcode_project2/kelasPRD.json"         # relative path untuk kelasPRD.json
kelasPRDexcel_path = "Barcode_project2/kelasPRD.xlsx"    # relative path untuk kelasPRD.xlsx
barcode_ico_path = "Barcode_project2/Barcode.ico"        # relative path untuk Barcode.ico

df_kehadiran = pd.read_excel(kelasPRDexcel_path,index_col="NIM")   # dataframe dari file kelasPRD.xlsx 

class WebcamApp:                                     # class untuk menangkap video kamera
    def __init__(self, window, window_title):        # fungsi yang dijalankan saat inisiasi kelas

        with open(fakultas_path,'r') as data:        # membaca file fakultas.json dan memasukkannya ke variabel nama
            self.fakultas = json.load(data)

        with open(kelasPRD_path,'r') as data:
            self.nama = json.load(data)

        self.window = window
        self.window.title(window_title)
        self.window.iconbitmap(barcode_ico_path)

        self.vid = cv2.VideoCapture(0)

        self.canvas = tk.Canvas(window, width=self.vid.get(3), height=self.vid.get(4))
        self.canvas.pack()

        self.update()
        self.window.mainloop()

    def update(self):
        ret, frame = self.vid.read()

        if ret:
            self.photo = ImageTk.PhotoImage(image=Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)))
            self.canvas.create_image(0, 0, image=self.photo, anchor=tk.NW)

        for i in decode(frame):
            value = i.data.decode('utf-8')
            nim_code = value[:3]
            time.sleep(0.5)

            output = True
            try:
                angka = int(value)
            except:
                output = False

            if len(value) == 16 and output:
                if int(value[:8]) in df_kehadiran.index:
                    df_kehadiran['Status'][int(value[:8])] = "Hadir"
                    df_kehadiran['Waktu Absensi'][int(value[:8])] = dt.datetime.now().strftime("%H:%M:%S")
                    beepy.beep(sound=5)
                    print("Berhasil hooray")

        self.window.after(1, self.update)

# Create a window and pass it to the WebcamApp class

root = tk.Tk()
app = WebcamApp(root, "QR Scanner")

#bikin pdf berisi absensi sesuai ktm
df_kehadiran.to_excel(f'Barcode_project2/excel absensi/Absensi Kehadiran PRD {dt.datetime.now().strftime("%d-%m-%Y pada %H.%M.%S")}.xlsx')
