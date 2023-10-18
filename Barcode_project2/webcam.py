import cv2
import tkinter as tk
import json
import time
import datetime as dt
import pandas as pd
from pyzbar.pyzbar import decode
from PIL import Image, ImageTk

fakultas_path = "Barcode_project2/fakultas.json"
kelasPRD_path = "Barcode_project2/kelasPRD.json"
kelasPRDexcel_path = "Barcode_project2/kelasPRD.xlsx"
barcode_ico_path = "Barcode_project2/Barcode.ico"

df_kehadiran = pd.read_excel(kelasPRDexcel_path,index_col="NIM")

class WebcamApp:
    def __init__(self, window, window_title):

        with open(fakultas_path,'r') as data:
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
            time.sleep(1)
            print(value)

            output = True
            try:
                angka = int(value)
            except:
                output = False

            if len(value) == 16 and output:
                print(f"Nama : {self.nama[value[:8]]}")
                print(f"NIM : {value[:8]}")
                print(f"Fakultas : {self.fakultas[nim_code]}")
                print(f"Gerbang berhasil dibuka")
                print(f"waktu absen : {dt.datetime.now()}")
                df_kehadiran['Status'][int(value[:8])] = "Hadir"
                df_kehadiran['Waktu Absensi'][int(value[:8])] = dt.datetime.now().strftime("%H:%M:%S")

        self.window.after(1, self.update)

# Create a window and pass it to the WebcamApp class

root = tk.Tk()
app = WebcamApp(root, "QR Scanner")

#bikin pdf berisi absensi sesuai ktm
df_kehadiran.to_excel(f'excel absensi/Absensi Kehadiran PRD {dt.datetime.now().strftime("%d-%m-%Y pada %H.%M.%S")}.xlsx')
