import subprocess
import os
import time

# Dosyaların bulunduğu yol
kaynak_dizin = r"C:\Users\q\Desktop\NEVAD"

# Çalıştırılacak dosyaların listesi
dosya_listesi = [ "1)Hızlı Bilançolar.py", "2)BIST Sektörel Finansal Veriler.py", "3)Hedef Fiyat Tablosu.py", "4)Teknik Analiz Güncel.py", "5) Trendler.py", "6) Temel ve Teknik Analiz Puanlama.py"]

# İşlem başlangıç zamanını kaydet
baslangic_zamani = time.time()

# Her bir dosyayı sırayla çalıştırma
for dosya in dosya_listesi:
    dosya_yolu = os.path.join(kaynak_dizin, dosya)  # Dosya yolunu oluştur
    subprocess.run(["python", dosya_yolu])

# İşlem bitiş zamanını kaydet
bitis_zamani = time.time()

# İşlem süresini hesapla
islem_suresi_saniye = bitis_zamani - baslangic_zamani

# Saniye cinsinden işlem süresini dakika ve saniyeye dönüştür
islem_suresi_dakika = islem_suresi_saniye // 60
kalan_saniye = islem_suresi_saniye % 60

print("İşlem süresi:", int(islem_suresi_dakika), "dakika", int(kalan_saniye), "saniye")
