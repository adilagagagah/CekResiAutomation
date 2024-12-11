from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
from datetime import datetime
import time

def setup_driver():
    chrome_options = Options()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--ignore-ssl-errors=yes")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--log-level=3")  # Hanya menampilkan error kritis
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def track_resi_lion(resi_number, driver, ts):
    url = f"https://lionparcel.com/track/stt?q={resi_number}"
    driver.get(url)
    time.sleep(ts)
    try:
        group_wrapper = driver.find_element(By.CLASS_NAME, 'group-wrapper')
        p_element = group_wrapper.find_element(By.TAG_NAME, 'p')
        return p_element.text
    except:
        return "'group-wrapper' atau 'p' tidak ditemukan."

def track_resi_AUP(resi_number, driver):
    driver.get('https://aup.co.id/index.php/5098-2/')

    # Masukkan nomor resi
    resi_input = driver.find_element(By.ID, 'inquiryNumber')
    resi_input.send_keys(resi_number)

    # Klik tombol submit
    submit_button = driver.find_element(By.NAME, 'form-submitted')
    submit_button.click()

    time.sleep(5)

    try:
        table = driver.find_element(By.CLASS_NAME, 'ups-results')
        rows = table.find_elements(By.TAG_NAME, 'tr')
        last_row = rows[-1]
        tds = last_row.find_elements(By.TAG_NAME, 'td')
        return tds[1].text
    except Exception as e:
        return f'Error: {str(e)}'

start_time = datetime.now()
# print("KURIR : LION PARCEL")
print("WAKTU :" ,start_time.strftime("%H:%M:%S %d/%m/%Y"))

driver = setup_driver()
df_lion = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiLION', header=None)
df_aup = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiAUP', header=None)
nomor_resi_list_lion = df_lion[0].tolist()
nomor_resi_list_aup = df_aup[0].tolist()

hasil_tracking_lion = []
hasil_tracking_aup = []

# lion
ts = 50
for resi in nomor_resi_list_lion:
    hasil = track_resi_lion(resi, driver, ts)
    
    if 'diterima oleh' in hasil.lower():
        status = 'SELESAI'
    elif 'tidak ditemukan' in hasil.lower():
        status = 'INVALID'
    else:
        status = 'OTW'
    
    hasil_tracking_lion.append({
        'No Resi': resi, 
        'Perjalanan Terakhir': hasil, 
        'Status': status
    })
    ts = 5

# aup
for resi in nomor_resi_list_aup:
    hasil = track_resi_AUP(resi, driver)
    
    if 'delivered' in hasil.lower():
        status = 'SELESAI'
    elif 'error' in hasil.lower():
        status = 'INVALID'
    else:
        status = 'OTW'
    
    hasil_tracking_aup.append({
        'No Resi': resi, 
        'Perjalanan Terakhir': hasil, 
        'Status': status
    })

driver.quit()

print("\n\nLION PARCEL")
for hasil in hasil_tracking_lion[1:]:
    print(f"{hasil['No Resi']} {hasil['Status']}")


print("\n\nAUP EXPRESS")
for hasil in hasil_tracking_aup[1:]:
    print(f"{hasil['No Resi']}  {hasil['Status']}")


end_time = datetime.now()
execution_time = end_time - start_time
execution_time = execution_time.total_seconds()
menit = int(execution_time // 60)
detik = int(execution_time % 60)

print("mulai   :", start_time.strftime("%H:%M:%S %d/%m/%Y"))
print("selesai :", end_time.strftime("%H:%M:%S %d/%m/%Y"))
print(f"Kode dieksekusi selama: {menit} menit {detik} detik")