import pandas as pd

# Excel dosyasını oku
excel_dosya_adi = "C:/Users/q/Desktop/bt/NEVAData.xlsx"
veriler = pd.read_excel(excel_dosya_adi)

Temel_Analiz_Puan = []
Teknik_Analiz_Puan = []

def hesapla_temel_puan(row):
    puan = 0

    # Koşulları değerlendirerek puanı hesapla
    if row['Cari Oran'] >= 1.5:
        puan += 5  # Cari Oran koşulu için 5 puan ekledik

    if row['Cari Oran'] > row['Sektörel Cari Oran']:
        puan += 3  # Cari Oran sektör ortalamasının üzerinde olma koşuluna 3 puan ekledik

    if row['Asit-Test Oranı'] >= 1:
        puan += 2  # Asit-Test Oranı koşuluna 2 puan ekledik

    if row['Asit-Test Oranı'] > row['Sektörel Asit-Test Oran']:
        puan += 3

    if row['Nakit Oranı'] >= 0.2:
        puan += 4  # Nakit Oranı koşuluna 4 puan ekledik

    if row['Nakit Oranı'] > row['Sektörel Nakit Oran']:
        puan += 3   

    if row['Stok Bağımlılık Oranı'] > row['Sektörel Stok Bağımlılık Oran']:
        puan += 2   

    if row['Brüt Kar Marjı'] > row['Sektörel Brüt Kar Marjı']:
        puan += 3  # Brüt Kar Marjı sektör ortalamasının üzerinde olma koşuluna 3 puan ekledik

    if row['FAVÖK Marjı'] > row['Sektörel FAVÖK Marjı']:
        puan += 3

    if row['Net Kar Marjı'] > row['Sektörel Net Kar Marjı']:
        puan += 3  # Net Kar Marjı sektör ortalamasının üzerinde olma koşuluna 2 puan ekledik

    if row['Aktif Karlılık ROA (%)'] > row['Sektörel Aktif Karlılık']:
        puan += 4  # Aktif Karlılık ROA koşuluna 4 puan ekledik

    if row['Özsermaye Karlılığı (ROE Yıllık)'] > row['Sektörel Özsermaye Karlılık']:
        puan += 3  # Özsermaye Karlılığı ROE koşuluna 3 puan ekledik

    if row['Kaldıraç Oranı'] <= 0.5:
        puan += 4  # Kaldıraç Oranı koşuluna 2 puan ekledik

    if row['Finansman Oranı'] >= 1:
        puan += 2  # Finansman Oranı koşuluna 2 puan ekledik

    if row['Borç/Özsermaye Oranı'] < 1:
        puan += 3  # Borç/Özsermaye Oranı koşuluna 3 puan ekledik

    if row['M/Ö Oranı'] >= 0.70:
        puan += 2  # M/Ö Oranı koşuluna 2 puan ekledik

    if row['Finansal Borç Oranı'] >= 0.5:
        puan += 1  # Finansal Borç Oranı koşuluna 1 puan ekledik

    if row['Aktif Devir Hızı'] > row['Sektörel Aktif Devir Hızı']:
        puan += 3  # Aktif Devir Hızı koşuluna 3 puan ekledik

    if 0 < row['F/K Oranı'] < 40:
        puan += 4  # F/K Oranı koşuluna 4 puan ekledik

    if row['F/K Oranı'] < row['Sektörel F/K Oranı']:
        puan += 3

    if row['PD/DD Oranı'] > 0:
        puan += 2  # PD/DD Oranı koşuluna 2 puan ekledik

    if row['PD/DD Oranı'] < row['Sektörel PD/DD Oranı']:
        puan += 3
        
    if row['Temettü Verimi'] > 0:
        puan += 1  # Temettü Verimi koşuluna 1 puan ekledik

    if row['Halka Açıklık Oranı (%)'] < 50:
        puan += 2  # Halka Açıklık Oranı koşuluna 2 puan ekledik

    if row['Karı Sermayeden Yüksekler'] > 0:
        puan += 3  # Karı Sermayeden Yüksekler koşuluna 3 puan ekledik

    if row['Satışlar Değişimi (Yıllık)'] > 0:
        puan += 2  # Satışlar Değişimi (Yıllık) koşuluna 2 puan ekledik

    if row['Brüt Kar Değişimi (Yıllık)'] > 0:
        puan += 2  # Brüt Kar Değişimi (Yıllık) koşuluna 2 puan ekledik

    if row['Net Kar Değişimi (Yıllık)'] > 0:
        puan += 2  # Net Kar Değişimi (Yıllık) koşuluna 2 puan ekledik

    if row['Satışlar Değişimi (Çeyreklik)'] > 0:
        puan += 1  # Satışlar Değişimi (Çeyreklik) koşuluna 1 puan ekledik

    if row['Brüt Kar Değişimi (Çeyreklik)'] > 0:
        puan += 1  # Brüt Kar Değişimi (Çeyreklik) koşuluna 1 puan ekledik

    if row['Net Kar Değişimi (Çeyreklik)'] > 0:
        puan += 1  # Net Kar Değişimi (Çeyreklik) koşuluna 1 puan ekledik

    if row['Dönen Varlıklar Değişim (Çeyreklik)'] > 0:
        puan += 2  # Dönen Varlıklar Değişim (Çeyreklik) koşuluna 2 puan ekledik

    if row['Duran Varlıklar Değişimi (Çeyreklik)'] > 0:
        puan += 2  # Duran Varlıklar Değişim (Çeyreklik) koşuluna 2 puan ekledik

    if row['Özkaynaklar Değişim (Çeyreklik)'] > 0:
        puan += 3  # Özkaynaklar Değişim (Çeyreklik) koşuluna 3 puan ekledik

    if row['Toplam Varlıklar Değişim (Çeyreklik)'] > 0:
        puan += 3  # Toplam Varlıklar Değişim (Çeyreklik) koşuluna 3 puan ekledik

    return puan

def hesapla_teknik_puan(row):
    puan = 0

    # Koşulları değerlendirerek puanı hesapla
    if row['CCI_Signal'] == 'Buy':
        puan += 15 

    if row['RSI_Signal'] == 'Buy':
        puan += 10 

    if row['MOM_Signal'] == 'Buy':
        puan += 8 

    if row['STO_Signal'] == 'Buy':
        puan += 10 

    if row['MACD_Signal'] == 'Buy':
        puan += 12 

    if row['SMA_Signal'] == 'Buy':
        puan += 10 

    if row['BBand_Signal'] == 'Buy':
        puan += 10 

    if row['Kisa_Trend'] == 'Yukarı Trend':
        puan += 7 

    if row['Orta_Trend'] == 'Yukarı Trend':
        puan += 7 

    if row['Uzun_Trend'] == 'Yukarı Trend':
        puan += 11 


    return puan

# Her satır için temel analiz ve teknik analiz puanlarını hesapla
veriler['Temel Analiz Puan'] = veriler.apply(hesapla_temel_puan, axis=1)
veriler['Teknik Analiz Puan'] = veriler.apply(hesapla_teknik_puan, axis=1)

# Toplam Puanı hesapla (Temel Analiz Puanı %70 + Teknik Analiz Puanı %30)
veriler['Toplam Puan'] = (0.7 * veriler['Temel Analiz Puan']) + (0.3 * veriler['Teknik Analiz Puan'])

# Puanları içeren verileri başka bir Excel dosyasına yazdır
output_excel_dosya_adi = "C:/Users/q/Desktop/bt/NEVADataWithScores.xlsx"
veriler.to_excel(output_excel_dosya_adi, index=False)

print(f"Analiz puanları başarıyla hesaplandı ve '{output_excel_dosya_adi}' dosyasına yazıldı.")