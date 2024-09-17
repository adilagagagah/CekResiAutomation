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

driver = setup_driver()
df = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiLION', header=None)
nomor_resi_list = df[0].tolist()

hasil_tracking = []

for resi in nomor_resi_list:
    hasil = track_resi_lion(resi, driver)
    
    # Tentukan status berdasarkan apakah ada kata 'delivered' atau tidak
    # if 'delivered' in hasil.lower():
    #     status = 'SELESAI'
    # else:
    #     status = 'OTW'
    
    hasil_tracking.append({
        'No Resi': resi, 
        'Perjalanan Terakhir': hasil, 
        # 'Status': status
    })

driver.quit()

# Tampilkan hasil tracking
for hasil in hasil_tracking:
    print(f"{hasil['No Resi']}  {hasil['Perjalanan Terakhir']}")
