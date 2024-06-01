import pandas as pd
import numpy as np
import yfinance as yf
import os
import warnings
from datetime import datetime 

warnings.filterwarnings('ignore', message="The 'unit' keyword in TimedeltaIndex construction is deprecated", category=FutureWarning)

# ADX ve DI hesaplamak için fonksiyon
def calculate_adx(data, period):
    data['High_Low'] = data['High'] - data['Low']
    data['High_PreviousClose'] = abs(data['High'] - data['Close'].shift(1))
    data['Low_PreviousClose'] = abs(data['Low'] - data['Close'].shift(1))
    data['True_Range'] = data[['High_Low', 'High_PreviousClose', 'Low_PreviousClose']].max(axis=1)
    
    data['UpMove'] = data['High'] - data['High'].shift(1)
    data['DownMove'] = data['Low'].shift(1) - data['Low']
    data['PlusDM'] = np.where((data['UpMove'] > data['DownMove']) & (data['UpMove'] > 0), data['UpMove'], 0)
    data['MinusDM'] = np.where((data['DownMove'] > data['UpMove']) & (data['DownMove'] > 0), data['DownMove'], 0)
    
    data['PlusDI'] = 100 * (data['PlusDM'].rolling(window=period).mean() / data['True_Range'].rolling(window=period).mean())
    data['MinusDI'] = 100 * (data['MinusDM'].rolling(window=period).mean() / data['True_Range'].rolling(window=period).mean())
    
    data['DX'] = 100 * (abs(data['PlusDI'] - data['MinusDI']) / (data['PlusDI'] + data['MinusDI']))
    data['ADX'] = data['DX'].rolling(window=period).mean()
    
    return data['ADX'], data['PlusDI'], data['MinusDI']

# Dosya yollarını tanımla
dosya_klasoru = "C:/Users/q/Desktop/bılancolar"
trend_yolu = "C:/Users/q/Desktop/bt/Trendler.xlsx"
neva_veri_yolu = "C:/Users/q/Desktop/bt/NEVAData.xlsx"

# Periyotları tanımla
kisa_periyot = 14
orta_periyot = 30
uzun_periyot = 50

# Dosya isimlerini al
dosya_isimleri = os.listdir(dosya_klasoru)

#bugünün tarihini çek
bugun = datetime.now().strftime('%Y-%m-%d')
print(bugun) 
print("Tarihli veriler hazırlanıyor...")
# Her dosya için işlem yap
sonuclar = []
for dosya_adi in dosya_isimleri:
    sembol = dosya_adi.split('.')[0] + '.IS'
    try:
        veri = yf.download(sembol, start='2023-04-06', end=bugun)
        adx_kisa, plus_di_kisa, minus_di_kisa = calculate_adx(veri, kisa_periyot)
        adx_orta, plus_di_orta, minus_di_orta = calculate_adx(veri, orta_periyot)
        adx_uzun, plus_di_uzun, minus_di_uzun = calculate_adx(veri, uzun_periyot)
        
        trend_yonu_kisa = np.where(plus_di_kisa > minus_di_kisa, 'Yukarı Trend', 'Aşağı Trend')
        trend_yonu_orta = np.where(plus_di_orta > minus_di_orta, 'Yukarı Trend', 'Aşağı Trend')
        trend_yonu_uzun = np.where(plus_di_uzun > minus_di_uzun, 'Yukarı Trend', 'Aşağı Trend')
        
        sonuc = {
            'Hisse': sembol[:-3],
            'Kisa_Trend': trend_yonu_kisa[-1],
            'Orta_Trend': trend_yonu_orta[-1],
            'Uzun_Trend': trend_yonu_uzun[-1]
        }
        sonuclar.append(sonuc)
    except Exception as e:
        print(f"Hata: {dosya_adi} dosyası işlenirken bir hata oluştu: {e}")

# Sonuçları Excel'e kaydet
sonuclar_df = pd.DataFrame(sonuclar)
sonuclar_df.to_excel(trend_yolu, index=False)
print(f"Sonuçlar başarıyla '{trend_yolu}' adlı Excel dosyasına kaydedildi. (1/2)")


# NEVA verilerini trend sinyalleriyle birleştir
try:
    trend_sinyalleri = pd.read_excel(trend_yolu)
    neva_veri = pd.read_excel(neva_veri_yolu)
    birlesik_veri = pd.merge(neva_veri, trend_sinyalleri, on='Hisse', how='inner')
    birlesik_veri.to_excel(neva_veri_yolu, index=False)
    print(f"Sonuçlar başarıyla '{neva_veri_yolu}' adlı Excel dosyasına kaydedildi. (2/2)")
except FileNotFoundError:
    print("Dosya bulunamadı.")
