import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time

def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--ignore-ssl-errors=yes")
    chrome_options.add_argument("--ignore-certificate-errors")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def track_resi_AUP(resi_number, driver):
    driver.get('https://aup.co.id/index.php/5098-2/')

    # Masukkan nomor resi
    resi_input = driver.find_element(By.ID, 'inquiryNumber')
    resi_input.send_keys(resi_number)

    # Klik tombol submit
    submit_button = driver.find_element(By.NAME, 'form-submitted')
    submit_button.click()

    time.sleep(4)

    try:
        table = driver.find_element(By.CLASS_NAME, 'ups-results')
        rows = table.find_elements(By.TAG_NAME, 'tr')
        last_row = rows[-1]
        tds = last_row.find_elements(By.TAG_NAME, 'td')
        return tds[1].text
    except Exception as e:
        return f'Error: {str(e)}'

start_time = time.time()
print("KURIR : AUP EXPRESS")
print(f"WAKTU : {start_time}")

driver = setup_driver()
df = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiAUP', header=None)
nomor_resi_list = df[0].tolist()
hasil_tracking = []

for resi in nomor_resi_list:
    hasil = track_resi_AUP(resi, driver)
    
    if 'delivered' in hasil.lower():
        status = 'SELESAI'
    elif 'Error' in hasil.lower():
        status = 'INVALID'
    else:
        status = 'OTW'
    
    hasil_tracking.append({'No Resi': resi, 'Perjalanan Terakhir': hasil, 'Status': status})

driver.quit()

print("AUP EXPRESS")
for hasil in hasil_tracking:
    print(f"{hasil['No Resi']}  {hasil['Status']}")

end_time = time.time()
execution_time = end_time - start_time
menit = int(execution_time // 60)
detik = int(execution_time % 60)

print(f"mulai   : {start_time}")
print(f"selesai : {end_time}")
print(f"Kode dieksekusi selama: {menit} menit {detik} detik (AUP EXPRESS)")