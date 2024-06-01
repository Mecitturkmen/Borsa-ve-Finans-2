import requests
from bs4 import BeautifulSoup
from isyatirimhisse import StockData, Financials
import pandas as pd
from tqdm import tqdm

# Hisse isimlerini çekme
hisseler = []

url = 'https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/sirket-karti.aspx?hisse=A1CAP'
r = requests.get(url)
s = BeautifulSoup(r.text, 'html.parser')
s1 = s.find('select', id='ddlAddCompare')
c1 = s1.findChild('optgroup').findAll('option')

for a in c1:
    hisseler.append(a.string)

# Financials sınıfını başlatın
financials = Financials()

# Tüm hisselerin finansal verilerini çekip bir DataFrame'e kaydedin
for hisse in tqdm(hisseler, desc="Veriler indiriliyor"):
    try:
        finansal_veri = financials.get_data(symbols=hisse, exchange='TRY')
        df = pd.DataFrame(finansal_veri[hisse])  # Finansal veriyi DataFrame'e dönüştür
        
        # 'itemCode' ve 'itemDescEng' sütunlarını silin
        df.drop(columns=['itemCode', 'itemDescEng'], inplace=True)

        # Hisse sembolünü sütun olarak ekleyin
        df['Hisse Kodu'] = hisse
        
        # Hisse kodu sütununu silin
        df.drop(columns=['Hisse Kodu'], inplace=True)
        
        # DataFrame'i CSV olarak kaydetme
        df.to_csv("C:/Users/q/Desktop/bılancolar/{}.csv".format(hisse), index=False)
    except KeyError:
        print(f"No data found for {hisse}. Skipping...")
