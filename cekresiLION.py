from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--ignore-ssl-errors=yes")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--log-level=3")  # Hanya menampilkan error kritis
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def track_resi_lion(resi_number, driver):
    url = f"https://lionparcel.com/track/stt?q={resi_number}"
    driver.get(url)
    time.sleep(5)
    try:
        group_wrapper = driver.find_element(By.CLASS_NAME, 'group-wrapper')
        p_element = group_wrapper.find_element(By.TAG_NAME, 'p')
        return p_element.text
    except:
        return "'group-wrapper' atau 'p' tidak ditemukan."

start_time = time.time()

driver = setup_driver()
df = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiLION', header=None)
nomor_resi_list = df[0].tolist()

hasil_tracking = []

for resi in nomor_resi_list:
    hasil = track_resi_lion(resi, driver)
    
    if 'diterima oleh' in hasil.lower():
        status = 'SELESAI'
    else:
        status = 'OTW'
    
    hasil_tracking.append({
        'No Resi': resi, 
        'Perjalanan Terakhir': hasil, 
        'Status': status
    })

driver.quit()

for hasil in hasil_tracking:
    print(f"{hasil['No Resi']} {hasil['Status']}")
    
end_time = time.time()
execution_time = end_time - start_time
menit = int(execution_time // 60)
detik = int(execution_time % 60)

print(f"Kode dieksekusi selama: {menit} menit {detik} detik")
