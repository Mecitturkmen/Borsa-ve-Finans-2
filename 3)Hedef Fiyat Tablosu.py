import pandas as pd

# Excel dosyasını oku
excel_dosya_adi = "C:/Users/q/Desktop/bt/finansal_veriler.xlsx"
veriler = pd.read_excel(excel_dosya_adi)

# Hedef fiyatları hesapla ve verilere ekle
for index, row in veriler.iterrows():
    donem = row['Dönemi']
    donemx = donem.split("/")[-1].strip()
    kapanis_fiyati = float(row['Kapanış Fiyatı (TL)'])
    sektorel_fk_orani = float(row['Sektörel F/K Oranı'])
    fk_orani = float(row['F/K Oranı'])
    odenmis_sermaye = float(row['Ödenmiş Sermaye'])
    satısg= float(row['Yıllık Satış Geliri'])
    if satısg == 0:
        satısg = 0.00000000001
    else:
        satısg = float(satısg)
    yilliklandirilmis_kar = float(row['Yıllıklandırılmış Kar'])
    netfaalıyet = float(row['Yıllık Esas Faaliyet Karı'])
    pddd_orani = float(row['PD/DD Oranı'])
    sektorel_pddd_orani = float(row['Sektörel PD/DD Oranı'])
    ozkaynaklar = float(row['Özkaynaklar'])
    ozsermayakarlılık = yilliklandirilmis_kar / ozkaynaklar
    pıyasadegerı = kapanis_fiyati * odenmis_sermaye

    if donemx == '3':
        h11 = (kapanis_fiyati * sektorel_fk_orani) / fk_orani
        gfk1 = (kapanis_fiyati * odenmis_sermaye) / (yilliklandirilmis_kar * 4)
        h12 = (kapanis_fiyati * sektorel_fk_orani) / gfk1
        h13 = (kapanis_fiyati * sektorel_pddd_orani) / pddd_orani
        h14 = (20.04 * yilliklandirilmis_kar) / odenmis_sermaye
        h15 = (14 * (yilliklandirilmis_kar / odenmis_sermaye)) + ((kapanis_fiyati*odenmis_sermaye)/(2 * pddd_orani * odenmis_sermaye))
        h16 = (20 * kapanis_fiyati * yilliklandirilmis_kar) / (kapanis_fiyati*odenmis_sermaye)
        h17 = ((ozsermayakarlılık * 10) / pddd_orani) * kapanis_fiyati
        hbk = yilliklandirilmis_kar / odenmis_sermaye
        h18 = hbk * fk_orani
        potasnıyel_pd = (netfaalıyet * 7) + (0.5 * ozkaynaklar)
        h19 = (potasnıyel_pd / pıyasadegerı) * kapanis_fiyati 
        h110 = potasnıyel_pd / odenmis_sermaye  
        hx1 = pıyasadegerı / satısg
        hx2 = (yilliklandirilmis_kar/ satısg) * 10
        h111 = (hx1/hx2) * kapanis_fiyati

        h1ç = (h11+h12+h13+h14+h15+h16+h17+h18+h19+h110+h111) / 11
        h1p = ((h1ç - kapanis_fiyati) / kapanis_fiyati) 
        h1bp = (((ozkaynaklar - odenmis_sermaye) / odenmis_sermaye) * 100)  # Beta değerini yüzde olarak hesapla

        veriler.at[index, 'Model 1'] = h11
        veriler.at[index, 'Model 2'] = h12
        veriler.at[index, 'Model 3'] = h13
        veriler.at[index, 'Model 4'] = h14
        veriler.at[index, 'Model 5'] = h15
        veriler.at[index, 'Model 6'] = h16
        veriler.at[index, 'Model 7'] = h17
        veriler.at[index, 'Model 8'] = h18
        veriler.at[index, 'Model 9'] = h19
        veriler.at[index, 'Model 10'] = h110
        veriler.at[index, 'Model 11'] = h111
        veriler.at[index, 'Hedef Fiyat'] = h1ç
        veriler.at[index, 'Potansiyel Artış'] = h1p  # Potansiyel Artışı yüzde olarak formatla
        veriler.at[index, 'Beta'] = h1bp  # Beta değerini yüzde olarak formatla

    if donemx == '6':
        h21 = (kapanis_fiyati * sektorel_fk_orani) / fk_orani
        gfk2 = (kapanis_fiyati * odenmis_sermaye) / (yilliklandirilmis_kar * 2)
        h22 = (kapanis_fiyati * sektorel_fk_orani) / gfk2
        h23 = (kapanis_fiyati * sektorel_pddd_orani) / pddd_orani
        h24 = (20.04 * yilliklandirilmis_kar) / odenmis_sermaye
        h25 = (14 * (yilliklandirilmis_kar / odenmis_sermaye)) + ((kapanis_fiyati*odenmis_sermaye)/(2 * pddd_orani * odenmis_sermaye))
        h26 = (20 * kapanis_fiyati * yilliklandirilmis_kar) / (kapanis_fiyati*odenmis_sermaye)
        h27 = ((ozsermayakarlılık * 10) / pddd_orani) * kapanis_fiyati
        hbk = yilliklandirilmis_kar / odenmis_sermaye
        h28 = hbk * fk_orani
        potasnıyel_pd = (netfaalıyet * 7) + (0.5 * ozkaynaklar)
        h29 = (potasnıyel_pd / pıyasadegerı) * kapanis_fiyati 
        h210 = potasnıyel_pd / odenmis_sermaye  
        hx1 = pıyasadegerı / satısg
        hx2 = (yilliklandirilmis_kar/ satısg) * 10
        h211 = (hx1/hx2) * kapanis_fiyati

        h2ç = (h21+h22+h23+h24+h25+h26+h27+h28+h29+h210+h211) / 11
        h2p = ((h2ç - kapanis_fiyati) / kapanis_fiyati) 
        h2bp = (((ozkaynaklar - odenmis_sermaye) / odenmis_sermaye) * 100)  # Beta değerini yüzde olarak hesapla

        veriler.at[index, 'Model 1'] = h21
        veriler.at[index, 'Model 2'] = h22
        veriler.at[index, 'Model 3'] = h23
        veriler.at[index, 'Model 4'] = h24
        veriler.at[index, 'Model 5'] = h25
        veriler.at[index, 'Model 6'] = h26
        veriler.at[index, 'Model 7'] = h27
        veriler.at[index, 'Model 8'] = h28
        veriler.at[index, 'Model 9'] = h29
        veriler.at[index, 'Model 10'] = h210
        veriler.at[index, 'Model 11'] = h211
        veriler.at[index, 'Hedef Fiyat'] = h2ç
        veriler.at[index, 'Potansiyel Artış'] = h2p  # Potansiyel Artışı yüzde olarak formatla
        veriler.at[index, 'Beta'] = h2bp  # Beta değerini yüzde olarak formatla

    if donemx == '9':
        h31 = (kapanis_fiyati * sektorel_fk_orani) / fk_orani
        gfk3 = (kapanis_fiyati * odenmis_sermaye) / ((yilliklandirilmis_kar / 3) + yilliklandirilmis_kar)
        h32 = (kapanis_fiyati * sektorel_fk_orani) / gfk3
        h33 = (kapanis_fiyati * sektorel_pddd_orani) / pddd_orani
        h34 = (20.04 * yilliklandirilmis_kar) / odenmis_sermaye
        h35 = (14 * (yilliklandirilmis_kar / odenmis_sermaye)) + ((kapanis_fiyati*odenmis_sermaye)/(2 * pddd_orani * odenmis_sermaye))
        h36 = (20 * kapanis_fiyati * yilliklandirilmis_kar) / (kapanis_fiyati*odenmis_sermaye)
        h37 = ((ozsermayakarlılık * 10) / pddd_orani) * kapanis_fiyati
        hbk = yilliklandirilmis_kar / odenmis_sermaye
        h38 = hbk * fk_orani
        potasnıyel_pd = (netfaalıyet * 7) + (0.5 * ozkaynaklar)
        h39 = (potasnıyel_pd / pıyasadegerı) * kapanis_fiyati 
        h310 = potasnıyel_pd / odenmis_sermaye  
        hx1 = pıyasadegerı / satısg
        hx2 = (yilliklandirilmis_kar/ satısg) * 10
        h311 = (hx1/hx2) * kapanis_fiyati

        h3ç = (h31+h32+h33+h34+h35+h36+h37+h38+h39+h310+h311) / 11
        h3p = ((h3ç - kapanis_fiyati) / kapanis_fiyati) 
        h3bp = (((ozkaynaklar - odenmis_sermaye) / odenmis_sermaye) * 100)  # Beta değerini yüzde olarak hesapla

        veriler.at[index, 'Model 1'] = h31
        veriler.at[index, 'Model 2'] = h32
        veriler.at[index, 'Model 3'] = h33
        veriler.at[index, 'Model 4'] = h34
        veriler.at[index, 'Model 5'] = h35
        veriler.at[index, 'Model 6'] = h36
        veriler.at[index, 'Model 7'] = h37
        veriler.at[index, 'Model 8'] = h38
        veriler.at[index, 'Model 9'] = h39
        veriler.at[index, 'Model 10'] = h310
        veriler.at[index, 'Model 11'] = h311
        veriler.at[index, 'Hedef Fiyat'] = h3ç
        veriler.at[index, 'Potansiyel Artış'] = h3p  # Potansiyel Artışı yüzde olarak formatla
        veriler.at[index, 'Beta'] = h3bp  # Beta değerini yüzde olarak formatla

    if donemx == '12':
        h41 = (kapanis_fiyati * sektorel_fk_orani) / fk_orani
        gfk4 = (kapanis_fiyati * odenmis_sermaye) / (yilliklandirilmis_kar)
        h42 = (kapanis_fiyati * sektorel_fk_orani) / gfk4
        h43 = (kapanis_fiyati * sektorel_pddd_orani) / pddd_orani
        h44 = (20.04 * yilliklandirilmis_kar) / odenmis_sermaye
        h45 = (14 * (yilliklandirilmis_kar / odenmis_sermaye)) + ((kapanis_fiyati*odenmis_sermaye)/(2 * pddd_orani * odenmis_sermaye))
        h46 = (20 * kapanis_fiyati * yilliklandirilmis_kar) / (kapanis_fiyati*odenmis_sermaye)
        h47 = ((ozsermayakarlılık * 10) / pddd_orani) * kapanis_fiyati
        hbk = yilliklandirilmis_kar / odenmis_sermaye
        h48 = hbk * fk_orani
        potasnıyel_pd = (netfaalıyet * 7) + (0.5 * ozkaynaklar)
        h49 = (potasnıyel_pd / pıyasadegerı) * kapanis_fiyati 
        h410 = potasnıyel_pd / odenmis_sermaye  
        hx1 = pıyasadegerı / satısg
        hx2 = (yilliklandirilmis_kar/ satısg) * 10
        h411 = (hx1/hx2) * kapanis_fiyati

        veriler.at[index, 'Model 1'] = h41
        veriler.at[index, 'Model 2'] = h42
        veriler.at[index, 'Model 3'] = h43
        veriler.at[index, 'Model 4'] = h44
        veriler.at[index, 'Model 5'] = h45
        veriler.at[index, 'Model 6'] = h46
        veriler.at[index, 'Model 7'] = h47
        veriler.at[index, 'Model 8'] = h48
        veriler.at[index, 'Model 9'] = h49
        veriler.at[index, 'Model 10'] = h410
        veriler.at[index, 'Model 11'] = h411
        h4ç = (h41+h42+h43+h44+h45+h46+h47+h48+h49+h410+h411) / 11
        h4p = ((h4ç - kapanis_fiyati) / kapanis_fiyati) 
        h4bp = (((ozkaynaklar - odenmis_sermaye) / odenmis_sermaye) * 100)  # Beta değerini yüzde olarak hesapla

        veriler.at[index, 'Hedef Fiyat'] = h4ç
        veriler.at[index, 'Potansiyel Artış'] = h4p  # Potansiyel Artışı yüzde olarak formatla
        veriler.at[index, 'Beta'] = h4bp  # Beta değerini yüzde olarak formatla

# Yeni Excel dosyasının adı ve yolu
yeni_excel_dosya_yolu = "C:/Users/q/Desktop/bt/finansal_veriler_hesaplanmis.xlsx"

# Verileri yeni Excel dosyasına yaz
veriler.to_excel(yeni_excel_dosya_yolu, index=False)

print("Hesaplanmış veriler başarıyla Excel dosyasına aktarıldı:", yeni_excel_dosya_yolu)
