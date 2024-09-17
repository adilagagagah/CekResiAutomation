import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time

# Inisialisasi WebDriver dalam mode headless (tanpa membuka jendela browser)
chrome_options = Options()
chrome_options.add_argument("--headless")  # Menjalankan browser tanpa GUI
driver = webdriver.Chrome(options=chrome_options)

# Fungsi untuk melakukan tracking resi
def track_resi(resi_number):
    # Buka halaman tracking di AUP
    driver.get('https://aup.co.id/index.php/5098-2/')

    # Masukkan nomor resi
    resi_input = driver.find_element(By.ID, 'inquiryNumber')
    resi_input.send_keys(resi_number)

    # Klik tombol submit
    submit_button = driver.find_element(By.NAME, 'form-submitted')
    submit_button.click()

    # Tunggu beberapa detik agar halaman dan data dimuat sepenuhnya
    time.sleep(5)

    try:
        # Ambil tabel hasil berdasarkan class 'ups-results'
        table = driver.find_element(By.CLASS_NAME, 'ups-results')

        # Ambil semua baris dari tabel
        rows = table.find_elements(By.TAG_NAME, 'tr')

        # Ambil <tr> terakhir dan <td> kedua
        if rows:
            last_row = rows[-1]
            tds = last_row.find_elements(By.TAG_NAME, 'td')
            
            if len(tds) >= 2:
                return tds[1].text
            else:
                return 'Baris terakhir tidak memiliki dua kolom.'
        else:
            return 'Tidak ada baris di tabel.'
    
    except Exception as e:
        return f'Error: {str(e)}'

# Membaca daftar nomor resi dari file CSV tanpa header
df = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiAUP', header=None)
nomor_resi_list = df[0].tolist()

# Tempatkan hasil pelacakan dalam list
hasil_tracking = []

for resi in nomor_resi_list:
    hasil = track_resi(resi)
    
    # Tentukan status berdasarkan apakah ada kata 'delivered' atau tidak
    if 'delivered' in hasil.lower():
        status = 'SELESAI'
    else:
        status = 'OTW'
    
    hasil_tracking.append({'No Resi': resi, 'Perjalanan Terakhir': hasil, 'Status': status})

# Tampilkan hasil tracking
for hasil in hasil_tracking:
    print(f"{hasil['No Resi']}  {hasil['Status']}")

# Jangan lupa tutup browser setelah selesai
driver.quit()
