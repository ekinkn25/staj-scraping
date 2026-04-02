import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import urllib3
import os

# Güvenlik uyarılarını gizle
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 1. Dosyadan Linkleri Okuma
dosya_yolu = "links.txt"
try:
    with open(dosya_yolu, "r", encoding="utf-8") as f:
        kaynak_linkler = [line.strip() for line in f.readlines() if line.strip()]
    
    if not kaynak_linkler:
        print(f"Uyarı: {dosya_yolu} dosyası boş görünüyor.")
        exit()
except FileNotFoundError:
    print(f"Hata: {dosya_yolu} dosyası bulunamadı! Lütfen dosyayı oluşturun.")
    exit()

toplanan_veriler = []
print(f"Toplam {len(kaynak_linkler)} link taranmaya başlanıyor...\n")

# 2. Linkleri Tarama ve Siteye Göre Ayırma
for index, url in enumerate(kaynak_linkler, start=1):
    print(f"[{index}/{len(kaynak_linkler)}] Taranıyor: {url}")
    
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        cevap = requests.get(url, headers=headers, timeout=10, verify=False)
        
        if cevap.status_code == 200:
            corba = BeautifulSoup(cevap.text, 'html.parser')
            
            # ---------------------------------------------------------
            # KURAL 1: CYBERPARK SİTESİ İÇİN (Eski Yapı)
            # ---------------------------------------------------------
            if "cyberpark" in url:
                kutular = corba.find_all('div', class_='col-md-3')
                for kutu in kutular:
                    a_etiketi = kutu.find('a')
                    if not a_etiketi:
                        continue
                    
                    sirket_linki = a_etiketi.get('href', 'Link Yok').strip()
                    h3_etiketi = a_etiketi.find('h3', class_='title')
                    # Eğer h3 etiketi hiç yoksa veya içi tamamen boşsa:
                    if not h3_etiketi or not h3_etiketi.text.strip():
                        continue  # Bu satırı tamamen es geç ve sıradaki şirkete ilerle
                        
                    # Eğer yukarıdaki if bloğuna takılmadıysa, demek ki isim var. Artık ismini alabiliriz:
                    sirket_ismi = h3_etiketi.text

                    # Karşılaşabileceğimiz tüm "boş" link ihtimallerini bir listeye koyuyoruz
                    gecersiz_link_durumlari = ["http://", "https://", "", "Link Yok", "http://-", "http:// ", "#"]

                    # Eğer çektiğimiz link bu listedeki elemanlardan biriyse, standart bir metne çeviriyoruz
                    if sirket_linki in gecersiz_link_durumlari:
                        sirket_linki = "Link Bulunamadı"

                    toplanan_veriler.append({
                        "Şirket İsmi": sirket_ismi,
                        "Şirket URL": sirket_linki,
                        "Çekildiği Kaynak": url
                    })

            # ---------------------------------------------------------
            # KURAL 2: ODTÜ TEKNOKENT SİTESİ İÇİN (Yeni Tablo Yapısı)
            # ---------------------------------------------------------
            elif "odtuteknokent" in url:
                # Tablodaki tüm satırları bul
                satirlar = corba.find_all('tr')
                
                for satir in satirlar:
                    # Satırın içindeki hücreleri (sütunları) bul
                    sutunlar = satir.find_all('td')
                    
                    # Eğer satırda en az 2 hücre varsa (biri isim, biri link için)
                    if len(sutunlar) >= 2:
                        # İlk hücre şirket ismidir (0. indeks)
                        sirket_ismi = sutunlar[0].text.replace('"', '').replace('#', '').strip()
                        
                        # İkinci hücredeki (1. indeks) <a> etiketini bul
                        a_etiketi = sutunlar[1].find('a')
                        sirket_linki = a_etiketi.get('href', 'Link Yok').strip() if a_etiketi else "Link Yok"
                        gecersiz_link_durumlari = ["http://", "https://", "", "Link Yok", "http://-", "http:// ", "http://yok", "#"]

                    # Eğer çektiğimiz link bu listedeki elemanlardan biriyse, standart bir metne çeviriyoruz
                        if sirket_linki in gecersiz_link_durumlari:
                            sirket_linki = "Link Bulunamadı"
                        
                        # Tablo başlıklarını veya boş satırları atlamak için küçük bir güvenlik kontrolü
                        if sirket_ismi: 
                            toplanan_veriler.append({
                                "Şirket İsmi": sirket_ismi,
                                "Şirket URL": sirket_linki,
                                "Çekildiği Kaynak": url
                            })
                            
            # Eğer txt dosyasına başka bir site eklenirse ama kuralı yoksa:
            else:
                print(f"Uyarı: {url} sitesi için bir kazıma kuralı yazılmamış. Bu link atlanıyor.")

        else:
            print(f"HATA: Siteye ulaşılamadı. Statü Kodu: {cevap.status_code}")
            
    except Exception as e:
        print(f"Bir hata oluştu: {url} -> Detay: {e}")
        
    time.sleep(2) 

# 3. Dinamik İsimlendirme ile Excel'e Kaydetme
if toplanan_veriler:
    print(f"\nİşlem Tamam! Toplam {len(toplanan_veriler)} adet şirket bilgisi çekildi.")
    df = pd.DataFrame(toplanan_veriler)
    
    temel_isim = "sirket_verileri"
    uzanti = ".xlsx"
    sayac = 1
    dosya_adi = f"{temel_isim}_{sayac}{uzanti}"
    
    while os.path.exists(dosya_adi):
        sayac += 1
        dosya_adi = f"{temel_isim}_{sayac}{uzanti}"
    
    df.to_excel(dosya_adi, index=False)
    print(f"Tüm veriler başarıyla '{dosya_adi}' dosyasına kaydedildi.")
else:
    print("\nUyarı: Hiç veri çekilemedi. Linkleri veya HTML yapısını kontrol et.")