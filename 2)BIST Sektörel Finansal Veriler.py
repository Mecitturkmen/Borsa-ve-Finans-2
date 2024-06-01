import os
import pandas as pd
import requests
from io import StringIO

# İsyatirim web sitesinden hisse senetleri tablosunu çek
url = "https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/Temel-Degerler-Ve-Oranlar.aspx#page-1"
response = requests.get(url)
html_string = response.text
tablo = pd.read_html(StringIO(html_string))[2]
sektör = pd.DataFrame({"Hisse": tablo["Kod"], "Sektör": tablo["Sektör"]})
hissetamad = pd.DataFrame({"Hisse": tablo["Hisse Adı"], "Hisse Tam Adı": tablo["Hisse Adı"]})
halkaacıklık = pd.DataFrame({"Hisse": tablo["Hisse Adı"], "Hisse Tam Adı": tablo["Halka Açıklık Oranı (%)"]})

tablo2 = pd.read_html(StringIO(html_string))[7]
gunlukg = pd.DataFrame({"Hisse": tablo2["Kod"], "Sektör": tablo2["Günlük Getiri (%)"]})
haftalıkg = pd.DataFrame({"Hisse": tablo2["Kod"], "Sektör": tablo2["Haftalık Getiri (%)"]})
aylıkg = pd.DataFrame({"Hisse": tablo2["Kod"], "Sektör": tablo2["Aylık Getiri (%)"]})
yıllıkg = pd.DataFrame({"Hisse": tablo2["Kod"], "Sektör": tablo2["Yıl İçi Getiri (%)"]})

gunluk_getırı = dict (zip(tablo2["Kod"], tablo2 ["Günlük Getiri (%)"]))
haftalık_getırı = dict (zip(tablo2["Kod"], tablo2 ["Haftalık Getiri (%)"]))
aylık_getırı = dict (zip(tablo2["Kod"], tablo2 ["Aylık Getiri (%)"]))
yıllık_getırı = dict (zip(tablo2["Kod"], tablo2 ["Yıl İçi Getiri (%)"]))

tablo3 = pd.read_html(StringIO(html_string))[3]
hisse_brut_temettu = pd.DataFrame({"Hisse": tablo3["Kod"], "Hisse Başı Brüt (TL)": tablo3[("Hisse Başı Brüt (TL)")]})

hissebruttemettu = dict (zip(tablo3["Kod"], tablo3 ["Hisse Başı Brüt (TL)"]))

# Hisse Kodu ile halka açıklık oranı eşleştir
halka_acıklık_oranı = dict (zip(tablo["Kod"], tablo ["Halka Açıklık Oranı (%)"]))

# Hisse kodu-hisse ismi eşleştirme tablosunu oluştur
hisse_kod_isim_tablosu = dict(zip(tablo["Kod"], tablo["Hisse Adı"]))

# Aynı sektördeki hisseleri grupla
gruplanmış_hisseler = sektör.groupby('Sektör')['Hisse'].apply(list).reset_index()

# Sektörlere göre hisse senetlerini ayrı listelere atama
sektörler = gruplanmış_hisseler['Sektör'].tolist()
hisse_grupları = gruplanmış_hisseler['Hisse'].tolist()

# Dosya yolunu belirtin
dosya_klasoru = "C:/Users/q/Desktop/bılancolar/"

# Dosya klasöründeki tüm dosya isimlerini al
dosya_isimleri = os.listdir(dosya_klasoru)

# Web scraping ile hisse kapanış fiyatlarını al
url4 = "https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/Temel-Degerler-Ve-Oranlar.aspx#page-1"
response = requests.get(url4)
html_string = response.text
tablo = pd.read_html(StringIO(html_string))[2]

# Hisse kodlarını anahtar olarak ve kapanış fiyatlarını değer olarak içeren bir sözlük oluştur
hisse_kapanis = dict(zip(tablo["Kod"], tablo["Kapanış (TL)"]))

# DataFrame için boş bir liste oluştur
data = []

# Dönem sıralaması
donemler = ['12', '9', '6', '3']

# Sektörel F/K ve PD/DD Oranlarını hesaplamak için boş sözlükler oluştur
sektor_fk_oranlari = {}
sektor_pddd_oranlari = {}
sektor_brut_kar = {}
sektor_toplam_satıslar = {}
sektor_favok = {}


for dosya_adi in dosya_isimleri:
    dosya_yolu = os.path.join(dosya_klasoru, dosya_adi)
    
    # Dosyayı oku
    veri = pd.read_csv(dosya_yolu)

    # Ödenmiş Sermaye Bilgisi
    odenmis_sermaye = veri.iloc[59, -1]  # 62. satırın ve son sütunun indeksindeki veriyi al

    if len (veri.columns) > 5:
    # Özkaynaklar Bilgisi
        ozkaynaklar = veri.iloc[57, -1]  # 60. satırın ve son sütunun indeksindeki veriyi al
        ozkaynaklarx = veri.iloc[57, -5]
        ort_ozkaynaklar = (ozkaynaklar + ozkaynaklarx) / 2
    else:
        ort_ozkaynaklar = ozkaynaklar


    ozkaynaklar_y = veri.iloc [57]

    # Ana Ortaklığa Ait Özkaynaklar Bilgisi
    ana_ortaklik_ozkaynaklar = veri.iloc[58, -1]  # 60. satırın ve son sütunun indeksindeki veriyi al

    # Uzun Vadeli Yükümlülük
    uzun_vade_yukumluluk = veri.iloc [44, -1]

    # Ana Ortaklık Payları 
    satir_110 = veri.iloc[108]  # 0'dan başladığı için 108, aslında 110. satıra karşılık gelir

    # Net Faaliyet Karı
    netfaalıyetkar = veri.iloc [88]

    ana_ortaklık = veri.iloc [108][-1]

    # Satış gelirleri
    satısg = veri.iloc [71]

    # Toplam Varlık
    t_varlık_y = veri.iloc[28]

    # Dönen Varlıklar
    donen_varlıklar = veri.iloc [0][-1]

    # Nakit ve Nakit Benzerleri
    nakıt_benzerlerı = veri.iloc [1][-1]

    # Kısa Vadeli Yükümlülükler
    kısa_vade_yukumluluk = veri.iloc [30][-1]

    # Toplam Yükümlülük
    toplam_yukumluluk = kısa_vade_yukumluluk + uzun_vade_yukumluluk

    # Amortisman Giderleri
    amort1 = veri.iloc[113][-1]
    amort2 = veri.iloc[113][-2]
    amort = amort1 - amort2

    # Brüt Kar
    brut_kar = veri.iloc [80][-1]
    brkar = veri.iloc[80]

    # Genel Yönetim Giderleri
    genel_yonetım_gıderlerı = veri.iloc [82][-1]

    # Amortisman ve itfa giderleri
    amort_ıtfa = veri.iloc [124][-1]
    amortıtfayıl = veri.iloc[124]

    # Pazarlama Gideri
    pg = veri.iloc [81][-1]
    pgyıl = veri.iloc [81]

    # Genel Yönetim Giderleri
    gyg = veri.iloc [82][-1]
    gygyyıl = veri.iloc [82]

    # ARGE Giderleri
    arge = veri.iloc [83][-1]
    argeyıl = [83]

    # Stoklar
    stok = veri.iloc [7][-1]

    # Maddi Duran Varlıklar
    maddı_duran_varlık = veri.iloc [23, -1]

    # Kısa Vadeli Finansal Borçlar
    kısa_vade_fınansal = veri.iloc [31 , -1]

    # Uzun Vadeli Finansal Borçlar
    uzun_vade_fınansal = veri.iloc [45, -1]

    # Toplam Finansal Borçlar
    toplam_fınansal_borclar = kısa_vade_fınansal + uzun_vade_fınansal



    if len (veri.columns) > 5:
        # Toplam Varlıklar
        toplam_varlık = veri.iloc [28][-1]
        toplam_varlık2 = veri.iloc [28][-5]
        toplam_varlıkx = (toplam_varlık + toplam_varlık2) / 2
    else:
        toplam_varlıkx = toplam_varlık

    if len (veri.columns) > 5:
        # Son dönemin indeksini belirle
        son_indeks = veri.columns[-1]  # Son sütunun indeksini al
        dort_indeks = veri.columns[-2]
        uc_indeks = veri.columns[-3]
        ıkı_indeks = veri.columns[-4]
        bır_indeks = veri.columns[-5]

        brkar5 = brkar[-5:]  # Son 5 veriyi al
        aıtfa = amortıtfayıl [-5:]
        pg1 = pgyıl [-5:]
        gyg1 = gygyyıl [-5:]
        arge1 = argeyıl[-5:]

        # Yıllıklandırılmış Kar Hesabı
        if son_indeks.endswith('/3'):
            deger5bk = brkar5[-1]
            deger5aı = aıtfa[-1]
            deger5pg = pg1 [-1]
            deger5gyg = gyg1 [-1]
            deger5arge = arge1 [-1]
        else:
            deger5bk = brkar5[-1] - brkar5[-2]
            deger5aı = aıtfa [-1] - aıtfa [-2]
            deger5pg = pg1 [-1] - pg1 [-2]
            deger5gyg = gyg1 [-1] - gyg1 [-2]
            deger5arge = arge1 [-1] - arge1 [-1]

        try:
            if dort_indeks.endswith('/3'):
                deger4bk = brkar5[-2]
                deger4aı = aıtfa [-2]
                deger4pg = pg1 [-2]
                deger4gyg = gyg1 [-2]
                deger4arge = arge1 [-2]
            else:
                deger4bk = brkar5[-2] - brkar5[-3]
                deger4aı = aıtfa [-2] - aıtfa [-3]
                deger4pg = pg1 [-2] - pg1 [-3]
                deger4gyg = gyg1 [-2] - gyg1 [-3]
                deger4arge = arge1 [-2] - arge1 [-3]
        except IndexError:
            deger4arge = 0 

        try:
            if uc_indeks.endswith('/3'):
                deger3bk = brkar5[-3]
                deger3aı = aıtfa [-3]
                deger3pg = pg1 [-3]
                deger3gyg = gyg1 [-3]
                deger3arge = arge1 [-3]
            else:
                deger3bk = brkar5[-3] - brkar5[-4]
                deger3aı = aıtfa [-3] - aıtfa [-4]
                deger3pg = pg1 [-3] - pg1 [-4]
                deger3gyg = gyg1 [-3] - gyg1 [-4]
                deger3arge = arge1 [-3] - arge1 [-4]
        except IndexError:
            deger3arge = 0 

        try:
            if ıkı_indeks.endswith('/3'):
                deger2bk = brkar5[-4]
                deger2aı = aıtfa [-4]
                deger2pg = pg1 [-4]
                deger2gyg = gyg1 [-4]
                deger2arge = arge1 [-4]
            else:
                deger2bk = brkar5[-4] - brkar5[-5]
                deger2aı = aıtfa [-4] - aıtfa [-5] 
                deger2pg = pg1 [-4] - pg1 [-5]
                deger2gyg = gyg1 [-4] - gyg1 [-5]
                deger2arge = arge1 [-4] - arge1 [-5]
        except IndexError:
            deger2arge = 0 
 

        yıllık_brut_kar = deger2bk + deger3bk + deger4bk + deger5bk
        yıllık_amort_ıtfa = deger2aı +deger3aı + deger4aı + deger5aı
        yıllık_pg = deger2pg + deger3pg + deger4pg + deger5pg 
        yıllık_gyg =deger2gyg + deger3gyg + deger4gyg +deger5gyg 
        yıllık_arge = deger2arge + deger3arge + deger4arge + deger5arge

    # Eğer sütun sayısı 5'ten azsa, yıllıklandırılmış karı son sütun ile aynı yap
    else:
        yıllık_brut_kar = veri.iloc [80][-1]
        yıllık_amort_ıtfa = veri.iloc [124][-1]

    if len (veri.columns) > 5:
        # Son dönemin indeksini belirle
        son_indeks = veri.columns[-1]  # Son sütunun indeksini al
        dort_indeks = veri.columns[-2]
        uc_indeks = veri.columns[-3]
        ıkı_indeks = veri.columns[-4]
        bır_indeks = veri.columns[-5]

        son_besz_veri = ozkaynaklar_y[-5:]  # Son 5 veriyi al

        # Yıllıklandırılmış Kar Hesabı
        if son_indeks.endswith('/3'):
            deger5z = son_besz_veri.iloc[-1]
        else:
            deger5z = son_besz_veri.iloc[-1] - son_besz_veri.iloc[-2]

        if dort_indeks.endswith('/3'):
            deger4z = son_besz_veri.iloc[-2]
        else:
            deger4z = son_besz_veri.iloc[-2] - son_besz_veri.iloc[-3]

        if uc_indeks.endswith('/3'):
            deger3z = son_besz_veri.iloc[-3]
        else:
            deger3z = son_besz_veri.iloc[-3] - son_besz_veri.iloc[-4]

        if ıkı_indeks.endswith('/3'):
            deger2z = son_besz_veri.iloc[-4]
        else:
            deger2z = son_besz_veri.iloc[-4] - son_besz_veri.iloc[-5]  

        yıllık_ozsermaye = deger2z + deger3z + deger4z + deger5z
    # Eğer sütun sayısı 5'ten azsa, yıllıklandırılmış karı son sütun ile aynı yap
    else:
        yıllık_ozsermaye = veri.iloc[57][-1]

    if len (veri.columns) > 5:
        # Son dönemin indeksini belirle
        son_indeks = veri.columns[-1]  # Son sütunun indeksini al
        dort_indeks = veri.columns[-2]
        uc_indeks = veri.columns[-3]
        ıkı_indeks = veri.columns[-4]
        bır_indeks = veri.columns[-5]

        son_best_veri = t_varlık_y[-5:]  # Son 5 veriyi al

        # Yıllıklandırılmış Kar Hesabı
        if son_indeks.endswith('/3'):
            deger5t = son_best_veri.iloc[-1]
        else:
            deger5t = son_best_veri.iloc[-1] - son_best_veri.iloc[-2]

        if dort_indeks.endswith('/3'):
            deger4t = son_best_veri.iloc[-2]
        else:
            deger4t = son_best_veri.iloc[-2] - son_best_veri.iloc[-3]

        if uc_indeks.endswith('/3'):
            deger3t = son_best_veri.iloc[-3]
        else:
            deger3t = son_best_veri.iloc[-3] - son_best_veri.iloc[-4]

        if ıkı_indeks.endswith('/3'):
            deger2t = son_best_veri.iloc[-4]
        else:
            deger2t = son_best_veri.iloc[-4] - son_best_veri.iloc[-5]  

        yıllık_toplam_varlık = deger2t + deger3t + deger4t + deger5t
    # Eğer sütun sayısı 5'ten azsa, yıllıklandırılmış karı son sütun ile aynı yap
    else:
        yıllık_toplam_varlık = veri.iloc[28][-1]

    if len (veri.columns) > 5:
        # Son dönemin indeksini belirle
        son_indeks = veri.columns[-1]  # Son sütunun indeksini al
        dort_indeks = veri.columns[-2]
        uc_indeks = veri.columns[-3]
        ıkı_indeks = veri.columns[-4]
        bır_indeks = veri.columns[-5]

        son_besy_veri = satısg[-5:]  # Son 5 veriyi al

        # Yıllıklandırılmış Kar Hesabı
        if son_indeks.endswith('/3'):
            deger5y = son_besy_veri.iloc[-1]
        else:
            deger5y = son_besy_veri.iloc[-1] - son_besy_veri.iloc[-2]

        if dort_indeks.endswith('/3'):
            deger4y = son_besy_veri.iloc[-2]
        else:
            deger4y = son_besy_veri.iloc[-2] - son_besy_veri.iloc[-3]

        if uc_indeks.endswith('/3'):
            deger3y = son_besy_veri.iloc[-3]
        else:
            deger3y = son_besy_veri.iloc[-3] - son_besy_veri.iloc[-4]

        if ıkı_indeks.endswith('/3'):
            deger2y = son_besy_veri.iloc[-4]
        else:
            deger2y = son_besy_veri.iloc[-4] - son_besy_veri.iloc[-5]  

        satısgelırı = deger2y + deger3y + deger4y + deger5y
    # Eğer sütun sayısı 5'ten azsa, yıllıklandırılmış karı son sütun ile aynı yap
    else:
        satısgelırı = veri.iloc[71][-1]

    if len (veri.columns) > 5:
        # Son dönemin indeksini belirle
        son_indeks = veri.columns[-1]  # Son sütunun indeksini al
        dort_indeks = veri.columns[-2]
        uc_indeks = veri.columns[-3]
        ıkı_indeks = veri.columns[-4]
        bır_indeks = veri.columns[-5]

        son_besx_veri = netfaalıyetkar[-5:]  # Son 5 veriyi al

        # Yıllıklandırılmış Kar Hesabı
        if son_indeks.endswith('/3'):
            deger5x = son_besx_veri.iloc[-1]
        else:
            deger5x = son_besx_veri.iloc[-1] - son_besx_veri.iloc[-2]

        if dort_indeks.endswith('/3'):
            deger4x = son_besx_veri.iloc[-2]
        else:
            deger4x = son_besx_veri.iloc[-2] - son_besx_veri.iloc[-3]

        if uc_indeks.endswith('/3'):
            deger3x = son_besx_veri.iloc[-3]
        else:
            deger3x = son_besx_veri.iloc[-3] - son_besx_veri.iloc[-4]

        if ıkı_indeks.endswith('/3'):
            deger2x = son_besx_veri.iloc[-4]
        else:
            deger2x = son_besx_veri.iloc[-4] - son_besx_veri.iloc[-5]  

        YNetF = deger2x + deger3x + deger4x + deger5x

    # Eğer sütun sayısı 5'ten azsa, yıllıklandırılmış karı son sütun ile aynı yap
    else:
        YNetF = veri.iloc[88][-1]

    if len (veri.columns) > 5:
        # Son dönemin indeksini belirle
        son_indeks = veri.columns[-1]  # Son sütunun indeksini al
        dort_indeks = veri.columns[-2]
        uc_indeks = veri.columns[-3]
        ıkı_indeks = veri.columns[-4]
        bır_indeks = veri.columns[-5]

        son_bes_veri = satir_110[-5:]  # Son 5 veriyi al

        # Yıllıklandırılmış Kar Hesabı
        if son_indeks.endswith('/3'):
            deger5 = son_bes_veri.iloc[-1]
        else:
            deger5 = son_bes_veri.iloc[-1] - son_bes_veri.iloc[-2]

        if dort_indeks.endswith('/3'):
            deger4 = son_bes_veri.iloc[-2]
        else:
            deger4 = son_bes_veri.iloc[-2] - son_bes_veri.iloc[-3]

        if uc_indeks.endswith('/3'):
            deger3 = son_bes_veri.iloc[-3]
        else:
            deger3 = son_bes_veri.iloc[-3] - son_bes_veri.iloc[-4]

        if ıkı_indeks.endswith('/3'):
            deger2 = son_bes_veri.iloc[-4]
        else:
            deger2 = son_bes_veri.iloc[-4] - son_bes_veri.iloc[-5]  

        YıllıklandırılmısKar = deger2 + deger3 + deger4 + deger5

    # Eğer sütun sayısı 5'ten azsa, yıllıklandırılmış karı son sütun ile aynı yap
    else:
        YıllıklandırılmısKar = veri.iloc[108][-1]

    # Hisse adını al
    hisse_adi = dosya_adi.split('.')[0]
    hisse_adi_tam = hisse_kod_isim_tablosu.get (hisse_adi, "Bilgi yok")
    hao = halka_acıklık_oranı.get(hisse_adi, "Bilgi bulunamadı")
    gg = gunluk_getırı.get(hisse_adi, "Bilgi bulunamadı")
    hg = haftalık_getırı.get(hisse_adi, "Bilgi bulunamadı")
    ag = aylık_getırı.get(hisse_adi, "Bilgi bulunamadı")
    yg = yıllık_getırı.get(hisse_adi, "Bilgi bulunamadı")
    hbt = hissebruttemettu.get(hisse_adi, "Hisse Başı Brüt (TL)")
    if pd.isnull(hbt) or hbt == "Hisse Başı Brüt (TL)":
        hbt = 0
    else:
        try:
            hbt = float(hbt)/10000
        except ValueError:
            pass

    # Hisse adı ile kapanış fiyatını eşleştir
    kapanis_fiyati = hisse_kapanis.get(hisse_adi, "Bilgi bulunamadı")
    
    # Kapanış fiyatını sayısal formata dönüştür
    kapanis_fiyati = kapanis_fiyati.replace(".", "").replace(",", "")
    kapanis_fiyati = float(kapanis_fiyati)
    
    # Kapanış fiyatını 100'e böl
    kapanis_fiyati_100 = kapanis_fiyati / 100

    # Piyasa Değeri Hesabı İçin 
    pıyasa_degerı = kapanis_fiyati_100 * odenmis_sermaye
    
    # Hisse Başı Kar Hesabı için
    hisse_basi_kar = YıllıklandırılmısKar / odenmis_sermaye


    # F/K Oranı Hesabı için
    fk_orani = kapanis_fiyati_100 / hisse_basi_kar

    # Virgülden sonra iki basamak almak için
    fk_orani = round(fk_orani, 2)

    # PD/DD Oranı hesabı için 
    pddd_orani = pıyasa_degerı / ana_ortaklik_ozkaynaklar

    # Virgülden sonra iki basamak almak için
    pddd_orani = round(pddd_orani, 2)
    
    # Sektör bilgisini al
    hisse_sektoru = sektör[sektör['Hisse'] == hisse_adi]['Sektör'].iloc[0]
    try:
        gg = float(gg) / 100
    except ValueError:
        pass
    try:
        ag = float(ag) / 100
    except ValueError:
        pass
    try:
        hg = float(hg) / 100
    except ValueError:
        pass
    try:
        yg = float(yg) / 100
    except ValueError:
        pass

    # Karı Özsermayesinden Yüksek Olan Hissler
    ksy = YıllıklandırılmısKar / odenmis_sermaye
    ksy = round(ksy,2)

    # FAVÖK
    favok = yıllık_brut_kar + yıllık_gyg +yıllık_pg + yıllık_arge + yıllık_amort_ıtfa

    # Cari oran
    carı_oran = donen_varlıklar / kısa_vade_yukumluluk

    # Net Kar Marjı
    net_kar_marj = (YıllıklandırılmısKar / satısgelırı) * 100

    # Brüt Kar Marjı
    brut_kar_marj = (yıllık_brut_kar / satısgelırı) * 100

    # Esas Faaliyet Kar Marjı
    esas_faalıyet_marj = (YNetF / satısgelırı) * 100

    # FAVÖK Marjı
    favok_marj = (favok / satısgelırı) * 100

    # PD/Net Satış ???
    pd_netsat = pıyasa_degerı / satısgelırı

    # Kullanılan sermaye
    Kullanılan_sermaye = toplam_varlık - kısa_vade_yukumluluk

    # Net Faaliyet Çeyreklik Karı
    nfck = brut_kar - pg - gyg - arge

    # ROCE Hesaplama
    roce = nfck / Kullanılan_sermaye


    # ROA Hesaplama
    roa = (YıllıklandırılmısKar / toplam_varlıkx)*100

    # Aktif Devir Hızı
    adh = (satısgelırı / toplam_varlıkx)

    # Yıllık Özsermaye Karlılığı (ROE)
    yıllık_ozsermaye_karlılık = (ana_ortaklık / ort_ozkaynaklar) * 100

    # Temettü Verimi Hesabı
    temettu_verım = hbt / kapanis_fiyati_100

    # Asit-Test Oranı Hesaplama
    asıt_test_oranı = (donen_varlıklar-stok) / kısa_vade_yukumluluk

    # Nakit Oranı Hesaplama
    nakıt_oran = nakıt_benzerlerı / kısa_vade_yukumluluk

    # Kaldıraç Oranı
    kaldırac_oranı = toplam_yukumluluk / toplam_varlık

    # Finansman Oranı
    fınansman_oranı = ozkaynaklar / toplam_yukumluluk

    # Borç/Özsermaye Oranı
    borc_ozsermaye_oranı = toplam_yukumluluk / ozkaynaklar

    # M/Ö Oranı
    m_o_oranı = maddı_duran_varlık / ozkaynaklar

    # Finansal Borç Oranı
    fınansal_borc_oranı = toplam_fınansal_borclar / toplam_varlık

    # Stok Bağımlılık Oranı Hesaplama
    if stok !=0:
        stok_bagımlılık = (kısa_vade_yukumluluk-nakıt_benzerlerı)/stok
    else:
        stok_bagımlılık = 0

    # Satş Değişimleri
    try:
        yıllık_satıs_degısımı = ((float (veri.iloc[71, -1]) - float (veri.iloc [71, -5])) / abs (float (veri.iloc [71, -5])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        yıllık_satıs_degısımı = 0

    try:
        ceyrek_satıs_degısımı = ((float (veri.iloc[71, -1]) - float (veri.iloc [71, -2])) / abs (float (veri.iloc [71, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyrek_satıs_degısımı = 0

    # Brüt Kar Değişimi
    try:
        yıllık_bkar_degısımı = (((float (brkar[-1])) - (float (brkar[-5]))) / abs (float (brkar[-5])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        yıllık_bkar_degısımı = 0

    try:
        ceyreklık_bkar_degısmı = ((float (veri.iloc[80, -1]) - float (veri.iloc [80, -2])) / abs (float (veri.iloc [80, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyreklık_bkar_degısmı = 0

    # Net Kar Değişimi
    try:
        yıllık_nkar_degısımı = ((float (veri.iloc[108, -1]) - float (veri.iloc [108, -5])) / abs (float (veri.iloc [108, -5])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        yıllık_nkar_degısımı = 0

    try:
        ceyrek_nkar_degısımı = ((float (veri.iloc[108, -1]) - float (veri.iloc [108, -2])) / abs (float (veri.iloc [108, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyrek_nkar_degısımı = 0

    try:
        ceyreklıkdonenvarlık = ((float (veri.iloc[0, -1]) - float (veri.iloc [0, -2])) / abs (float (veri.iloc [0, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyreklıkdonenvarlık = 0 

    try:
        ceyreklıkduranvarlık = ((float (veri.iloc[12, -1]) - float (veri.iloc [12, -2])) / abs (float (veri.iloc [12, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyreklıkduranvarlık = 0 

    try:
        ceyreklıktoplamvarlık = ((float (veri.iloc[28, -1]) - float (veri.iloc [28, -2])) / abs (float (veri.iloc [28, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyreklıktoplamvarlık = 0 
        
    try:
        ceyreklıkozkaynak = ((float (veri.iloc[57, -1]) - float (veri.iloc [57, -2])) / abs (float (veri.iloc [57, -2])) * 100)
    except (ValueError, IndexError,ZeroDivisionError):
        ceyreklıkozkaynak = 0 
        
    # DataFrame'e ekle
    data.append({
        "Hisse": hisse_adi,
        "Hisse Adı": hisse_adi_tam,
        "Sektör": hisse_sektoru,
        "Hisse Başı Brüt (TL)": hbt,
        "Temettü Verimi": temettu_verım,
        "Halka Açıklık Oranı (%)": hao/10,
        "Dönemi": son_indeks,
        "Günlük Getiri (%)": gg,
        "Haftalık Getiri (%)":hg,
        "Aylık  Getiri (%)": ag,
        "Yıl İçi Getiri (%)": yg,
        "Satışlar Değişimi (Yıllık)":yıllık_satıs_degısımı,
        "Satışlar Değişimi (Çeyreklik)":ceyrek_satıs_degısımı,
        "Brüt Kar Değişimi (Yıllık)":yıllık_bkar_degısımı,
        "Brüt Kar Değişimi (Çeyreklik)":ceyreklık_bkar_degısmı,
        "Net Kar Değişimi (Yıllık)": yıllık_nkar_degısımı,
        "Net Kar Değişimi (Çeyreklik)": ceyrek_nkar_degısımı,
        "Dönen Varlıklar Değişim (Çeyreklik)": ceyreklıkdonenvarlık,
        "Duran Varlıklar Değişimi (Çeyreklik)":ceyreklıkduranvarlık,
        "Toplam Varlıklar Değişim (Çeyreklik)":ceyreklıktoplamvarlık,
        "Özkaynaklar Değişim (Çeyreklik)": ceyreklıkozkaynak,
        "Kaldıraç Oranı": kaldırac_oranı,
        "Finansman Oranı" : fınansman_oranı,
        "M/Ö Oranı": m_o_oranı,
        "Finansal Borç Oranı": fınansal_borc_oranı,
        "Borç/Özsermaye Oranı": borc_ozsermaye_oranı,
        "Aktif Karlılık ROA (%)" :roa,
        "Özsermaye Karlılığı (ROE Yıllık)": yıllık_ozsermaye_karlılık,
        "ROCE" : roce,
        "Aktif Devir Hızı": adh,
        "Karı Sermayeden Yüksekler" : ksy,
        "Cari Oran" : carı_oran,
        "Asit-Test Oranı": asıt_test_oranı,
        "Nakit Oranı": nakıt_oran,
        "Stok Bağımlılık Oranı": stok_bagımlılık,
        "Net Kar Marjı" : net_kar_marj,
        "Esas Faaliyet Kar Marjı" : esas_faalıyet_marj,
        "Brüt Kar Marjı": brut_kar_marj,
        "FAVÖK Marjı": favok_marj,
        "PD/Net Satış": pd_netsat,
        "Amortisman Giderleri": amort,
        "Ana Ortaklık Ait Özkaynaklar": ana_ortaklik_ozkaynaklar,
        "Çeyreklik Kar": ana_ortaklık,
        "Özkaynaklar": ozkaynaklar,
        "Ödenmiş Sermaye": odenmis_sermaye,
        "Piyasa Değeri" : pıyasa_degerı,
        "Yıllık Satış Geliri": satısgelırı,
        "Yıllık Esas Faaliyet Karı": YNetF,
        "Kapanış Fiyatı (TL)": kapanis_fiyati_100,
        "Yıllıklandırılmış Kar": YıllıklandırılmısKar,
        "Piyasa Değeri": pıyasa_degerı,
        "Kapanış Fiyatı (TL)": kapanis_fiyati_100,
        "Yıllıklandırılmış Kar": YıllıklandırılmısKar,
        "Piyasa Değeri": pıyasa_degerı,
        "Hisse Başına Kar": hisse_basi_kar,
        "F/K Oranı": fk_orani,
        "PD/DD Oranı": pddd_orani
    })
    
    # Sektörel F/K ve PD/DD Oranlarını hesapla ve topla
    if hisse_sektoru in sektor_fk_oranlari:
        sektor_fk_oranlari[hisse_sektoru]['toplam_fk'] += kapanis_fiyati_100
        sektor_fk_oranlari[hisse_sektoru]['toplam_pddd'] += pıyasa_degerı
        sektor_fk_oranlari[hisse_sektoru]['toplam_pd'] += ana_ortaklik_ozkaynaklar
        sektor_fk_oranlari[hisse_sektoru]['hisse_sayisi'] += hisse_basi_kar
        sektor_fk_oranlari[hisse_sektoru]["toplam_brut_kar"] += yıllık_brut_kar
        sektor_fk_oranlari[hisse_sektoru]["toplam_satıslar"] += satısgelırı
        sektor_fk_oranlari[hisse_sektoru]["toplam_favok"] += favok
        sektor_fk_oranlari[hisse_sektoru]["toplam_yıllık_netkar"] += YıllıklandırılmısKar
        sektor_fk_oranlari[hisse_sektoru]["toplam_ortvarlık"] += toplam_varlıkx
        sektor_fk_oranlari[hisse_sektoru]["toplam_ortozsermaye"] += ort_ozkaynaklar
        sektor_fk_oranlari[hisse_sektoru]["toplam_esas"] += YNetF
        sektor_fk_oranlari[hisse_sektoru]["toplam_donenvarlık"] += donen_varlıklar
        sektor_fk_oranlari[hisse_sektoru]["toplam_kısavadelı"] += kısa_vade_yukumluluk
        sektor_fk_oranlari[hisse_sektoru]["toplam_stok"] += stok
        sektor_fk_oranlari[hisse_sektoru]["toplam_nakıtbenzerı"] += nakıt_benzerlerı
    else:
        sektor_fk_oranlari[hisse_sektoru] = {'toplam_fk': kapanis_fiyati_100, 'toplam_pddd': pıyasa_degerı,'toplam_pd' :ana_ortaklik_ozkaynaklar, 'hisse_sayisi': hisse_basi_kar,'toplam_brut_kar': yıllık_brut_kar,
                                            'toplam_satıslar': satısgelırı,'toplam_favok' :favok, 'toplam_yıllık_netkar': YıllıklandırılmısKar,'toplam_ortvarlık' :toplam_varlıkx, 'toplam_ortozsermaye': ort_ozkaynaklar, 
                                            'toplam_esas':YNetF, 'toplam_donenvarlık':donen_varlıklar,"toplam_kısavadelı":kısa_vade_yukumluluk, "toplam_stok":stok,"toplam_nakıtbenzerı":nakıt_benzerlerı}

# DataFrame oluştur
df = pd.DataFrame(data)

# Sektörel F/K ve PD/DD Oranlarını hesapla
for sektor, degerler in sektor_fk_oranlari.items():
    ortalama_fk_oran = degerler['toplam_fk'] / degerler['hisse_sayisi']
    ortalama_pddd_oran = degerler['toplam_pddd'] / degerler['toplam_pd']
    ort_brutkarmarj = degerler["toplam_brut_kar"] / degerler["toplam_satıslar"]
    ort_favokmarj = degerler["toplam_favok"] / degerler["toplam_satıslar"]
    ortnetkarmarj = degerler["toplam_yıllık_netkar"] / degerler["toplam_satıslar"]
    ort_aktıfkarlılık = degerler["toplam_yıllık_netkar"] / degerler["toplam_ortvarlık"]
    ort_ozsermayekarlılık = degerler["toplam_yıllık_netkar"] / degerler["toplam_ortozsermaye"]
    ort_esas = degerler["toplam_esas"] / degerler["toplam_satıslar"]
    ort_aktıf_devır_hız = degerler["toplam_satıslar"] / degerler["toplam_ortvarlık"]

    ort_carıoran = degerler["toplam_donenvarlık"] / degerler["toplam_kısavadelı"]
    ort_asıttest = (degerler["toplam_donenvarlık"] - degerler["toplam_stok"]) / degerler["toplam_kısavadelı"]
    ort_nakıtoran = degerler["toplam_nakıtbenzerı"] / degerler["toplam_kısavadelı"]
    ort_stokbagımlılık = (degerler["toplam_kısavadelı"] - degerler["toplam_nakıtbenzerı"]) / degerler["toplam_stok"]


    df.loc[df['Sektör'] == sektor, 'Sektörel F/K Oranı'] = round(ortalama_fk_oran, 2)
    df.loc[df['Sektör'] == sektor, 'Sektörel PD/DD Oranı'] = round(ortalama_pddd_oran, 2)
    df.loc[df['Sektör'] == sektor, 'Sektörel Brüt Kar Marjı'] = ort_brutkarmarj * 100
    df.loc[df['Sektör'] == sektor, 'Sektörel FAVÖK Marjı'] = ort_favokmarj * 100
    df.loc[df['Sektör'] == sektor, 'Sektörel Net Kar Marjı'] = ortnetkarmarj * 100
    df.loc[df['Sektör'] == sektor, 'Sektörel Aktif Karlılık'] = ort_aktıfkarlılık * 100
    df.loc[df['Sektör'] == sektor, 'Sektörel Özsermaye Karlılık'] = ort_ozsermayekarlılık * 100
    df.loc[df['Sektör'] == sektor, 'Sektörel Esas Faaliyet Kar Marjı'] = ort_esas * 100
    df.loc[df['Sektör'] == sektor, 'Sektörel Cari Oran'] = ort_carıoran
    df.loc[df['Sektör'] == sektor, 'Sektörel Asit-Test Oran'] = ort_asıttest
    df.loc[df['Sektör'] == sektor, 'Sektörel Nakit Oran'] = ort_nakıtoran
    df.loc[df['Sektör'] == sektor, 'Sektörel Stok Bağımlılık Oran'] = ort_stokbagımlılık
    df.loc[df['Sektör'] == sektor, 'Sektörel Aktif Devir Hızı'] = ort_aktıf_devır_hız

# Sonuçları yazdır
print(df)

# Excel dosyası adını belirtin
excel_dosya_adi = "C:/Users/q/Desktop/bt/finansal_veriler.xlsx"

# DataFrame'i Excel dosyasına yazdır
df.to_excel(excel_dosya_adi, index=False)

print("Veriler başarıyla Excel dosyasına aktarıldı:", excel_dosya_adi)
