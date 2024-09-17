import requests
import pandas as pd

def track_resi(resi_number):
    url = f"https://lionparcel.com/track/stt?q={resi_number}"
    response = requests.get(url)

    if response.status_code == 200:
        try:
            data = response.json()
            # Contoh, sesuaikan dengan struktur JSON yang sebenarnya
            last_status = data['data']['last_status']  
            return last_status
        except Exception as e:
            return f"Error decoding JSON: {str(e)}"
    else:
        return f"HTTP error: {response.status_code}"

df = pd.read_excel("../Cek Resi/cekresi.xlsx", sheet_name='cekresiLION', header=None)
nomor_resi_list = df[0].tolist()

hasil_tracking = []

for resi in nomor_resi_list:
    hasil = track_resi(resi)
    if 'delivered' in hasil.lower():
        status = 'SELESAI'
    else:
        status = 'OTW'
    
    hasil_tracking.append({'No Resi': resi, 'Perjalanan Terakhir': hasil, 'Status': status})

for hasil in hasil_tracking:
    print(f"{hasil['No Resi']}  {hasil['Perjalanan Terakhir']}")
