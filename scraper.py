# =============================================================================
# scraper.py — Siteleri tara, mailleri bul, Excel'e kaydet
# =============================================================================

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import urllib3
import os
import re
from urllib.parse import urlparse
from config import LINKS_DOSYASI, CIKTI_EXCEL

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

EMAIL_REGEX = re.compile(
    r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}'
)

ILETISIM_YOLLARI = [
    "/iletisim", "/iletişim", "/contact", "/contact-us", "",
    "/bize-ulasin", "/bize-ulaşın", "/hakkimizda", "/hakkımızda",
    "/tr/iletisim", "/en/contact", "/contact.html", "/eng/contacts.html", "/#iletişim",
    "/#contact-section", "/TR,16/iletisim.html", "/bize-ulasin/",
]

GECERSIZ_MAILLER = {
    "example.com", "domain.com", "email.com", "test.com",
    "yoursite.com", "sentry.io", "wix.com", "wordpress.com",
    "wixpress.com", "sentry-next.wixpress.com",
}

GECERSIZ_LINKLER = ["http://", "https://", "", "Link Yok", "http://-", "http:// ", "http://yok", "#"]


def sayfa_getir(url: str, timeout: int = 10) -> str | None:
    try:
        response = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
        if response.status_code == 200:
            return response.text
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError):
        pass
    except Exception:
        pass
    return None

def mail_bul(html: str) -> set:
    bulunanlar = set(EMAIL_REGEX.findall(html))
    return {
        m.lower() for m in bulunanlar
        if m.split("@")[1].lower() not in GECERSIZ_MAILLER
        and not m.endswith((".png", ".jpg", ".gif", ".svg"))
    }


def sirket_maillerini_topla(sirket_url: str) -> list:
    if not sirket_url or sirket_url == "Link Bulunamadı":
        return []

    if not sirket_url.startswith(("http://", "https://")):
        sirket_url = "https://" + sirket_url

    tum_mailler = set()

    # 1. Ana sayfa
    ana_html = sayfa_getir(sirket_url)
    if ana_html:
        tum_mailler |= mail_bul(ana_html)
        soup = BeautifulSoup(ana_html, "html.parser")
        for a in soup.find_all("a", href=True):
            if a["href"].startswith("mailto:"):
                mail = a["href"].replace("mailto:", "").split("?")[0].strip().lower()
                if mail:
                    tum_mailler.add(mail)

    # 2. İletişim sayfaları
    base = f"{urlparse(sirket_url).scheme}://{urlparse(sirket_url).netloc}"
    for yol in ILETISIM_YOLLARI:
        html = sayfa_getir(base + yol, timeout=8)
        if html:
            tum_mailler |= mail_bul(html)
            soup2 = BeautifulSoup(html, "html.parser")
            for a in soup2.find_all("a", href=True):
                if a["href"].startswith("mailto:"):
                    mail = a["href"].replace("mailto:", "").split("?")[0].strip().lower()
                    if mail:
                        tum_mailler.add(mail)
            if tum_mailler:
                break

    return sorted(tum_mailler)

# ── Siteye göre şirket listesi çekme kuralları ───────────────────────────────
 
def cyberpark_cek(corba, url) -> list:
    sonuclar = []
    kutular = corba.find_all('div', class_='col-md-3')
    for kutu in kutular:
        a = kutu.find('a')
        if not a:
            continue
        link = a.get('href', 'Link Yok').strip()
        h3   = a.find('h3', class_='title')
        if not h3 or not h3.text.strip():
            continue
        isim = h3.text.strip()
        if link in GECERSIZ_LINKLER:
            link = "Link Bulunamadı"
        sonuclar.append({"Şirket İsmi": isim, "Şirket URL": link, "Çekildiği Kaynak": url})
    return sonuclar
 
 
def odtuteknokent_cek(corba, url) -> list:
    sonuclar = []
    for satir in corba.find_all('tr'):
        sutunlar = satir.find_all('td')
        if len(sutunlar) < 2:
            continue
        isim = sutunlar[0].text.replace('"', '').replace('#', '').strip()
        a    = sutunlar[1].find('a')
        link = a.get('href', 'Link Yok').strip() if a else "Link Yok"
        if link in GECERSIZ_LINKLER:
            link = "Link Bulunamadı"
        if isim:
            sonuclar.append({"Şirket İsmi": isim, "Şirket URL": link, "Çekildiği Kaynak": url})
    return sonuclar
 
 
# ── Ana fonksiyon ─────────────────────────────────────────────────────────────
 
def calistir():
    # Links.txt oku
    try:
        with open(LINKS_DOSYASI, "r", encoding="utf-8") as f:
            kaynak_linkler = [l.strip() for l in f if l.strip()]
        if not kaynak_linkler:
            print(f"Uyarı: {LINKS_DOSYASI} boş.")
            return
    except FileNotFoundError:
        print(f"HATA: {LINKS_DOSYASI} bulunamadı!")
        return
 
    # ── AŞAMA 1: Şirket URL'lerini topla ─────────────────────────────────────
    print(f"{'─'*50}")
    print(f"AŞAMA 1 — Şirket linkleri çekiliyor ({len(kaynak_linkler)} kaynak)")
    print(f"{'─'*50}\n")
 
    toplanan_veriler = []
    for idx, url in enumerate(kaynak_linkler, 1):
        print(f"[{idx}/{len(kaynak_linkler)}] Taranıyor: {url}")
        try:
            r = requests.get(url, headers=HEADERS, timeout=10, verify=False)
            if r.status_code == 200:
                corba = BeautifulSoup(r.text, 'html.parser')
                if "cyberpark" in url:
                    toplanan_veriler += cyberpark_cek(corba, url)
                elif "odtuteknokent" in url:
                    toplanan_veriler += odtuteknokent_cek(corba, url)
                else:
                    print(f"  Uyarı: Kural yok, atlanıyor.")
            else:
                print(f"  HATA: Statü {r.status_code}")
        except Exception as e:
            print(f"  HATA: {e}")
        time.sleep(2)
 
    print(f"\n✓ {len(toplanan_veriler)} şirket bulundu.\n")
 
    # ── AŞAMA 2: Her şirketin mailini bul ────────────────────────────────────
    print(f"{'─'*50}")
    print(f"AŞAMA 2 — Mail adresleri aranıyor")
    print(f"{'─'*50}\n")
 
    toplam = len(toplanan_veriler)
    try:
        for i, kayit in enumerate(toplanan_veriler, 1):
            print(f"[{i}/{toplam}] {kayit['Şirket İsmi']} → {kayit['Şirket URL']}")
            mailler = sirket_maillerini_topla(kayit['Şirket URL'])
            kayit['Mail Adresleri'] = " | ".join(mailler) if mailler else "Bulunamadı"
            print(f"  {'✓ ' + kayit['Mail Adresleri'] if mailler else '– Bulunamadı'}")
            time.sleep(1.5)
    except KeyboardInterrupt:
        print(f"\n⚠ Durduruldu. {i}/{toplam} şirket işlendi. Mevcut veri kaydediliyor...\n")
 
    # ── AŞAMA 3: Excel'e kaydet ───────────────────────────────────────────────
    islenenler = [k for k in toplanan_veriler if 'Mail Adresleri' in k]
    if not islenenler:
        print("Uyarı: Kaydedilecek veri yok.")
        return
 
    sayac = 1
    dosya_adi = f"{CIKTI_EXCEL}_{sayac}.xlsx"
    while os.path.exists(dosya_adi):
        sayac += 1
        dosya_adi = f"{CIKTI_EXCEL}_{sayac}.xlsx"
 
    df = pd.DataFrame(islenenler, columns=["Şirket İsmi", "Şirket URL", "Mail Adresleri", "Çekildiği Kaynak"])
    df.to_excel(dosya_adi, index=False)
 
    mail_bulunan = (df["Mail Adresleri"] != "Bulunamadı").sum()
    print(f"\n✅ Scraper tamamlandı!")
    print(f"   Toplam şirket  : {len(df)}")
    print(f"   Mail bulunan   : {mail_bulunan}")
    print(f"   Kaydedilen     : {dosya_adi}\n")
 
    return dosya_adi   # main.py bu değeri kullanır
 
 
if __name__ == "__main__":
    calistir()
 

# # ── Siteye göre şirket listesi çekme kuralları ───────────────────────────────

# kaynak_linkler = [
#     "https://odtuteknokent.com.tr/tr/firmalar/tum-firmalar.php"
# ]

# toplanan_veriler = []
# print(f"Toplam {len(kaynak_linkler)} kaynak link taranıyor...\n")

# for index, url in enumerate(kaynak_linkler, start=1):
#     print(f"[{index}/{len(kaynak_linkler)}] Taranıyor: {url}")
#     try:
#         cevap = requests.get(url, headers=HEADERS, timeout=15, verify=False)
#         if cevap.status_code == 200:
#             corba = BeautifulSoup(cevap.text, "html.parser")

#             if "odtuteknokent" in url:
#                 satirlar = corba.find_all("tr")
#                 for satir in satirlar:
#                     sutunlar = satir.find_all("td")
#                     if len(sutunlar) >= 2:
#                         sirket_ismi = sutunlar[0].text.replace('"', "").replace("#", "").strip()
#                         a_etiketi = sutunlar[1].find("a")
#                         sirket_linki = (
#                             a_etiketi.get("href", "Link Yok").strip()
#                             if a_etiketi else "Link Yok"
#                         )
#                         gecersiz = ["http://", "https://", "", "Link Yok", "http://-", "http:// ", "http://yok", "#"]
#                         if sirket_linki in gecersiz:
#                             sirket_linki = "Link Bulunamadı"
#                         if sirket_ismi:
#                             toplanan_veriler.append({
#                                 "Şirket İsmi": sirket_ismi,
#                                 "Şirket URL": sirket_linki,
#                                 "Çekildiği Kaynak": url,
#                             })
#             else:
#                 print(f"Uyarı: {url} için kural yok, atlanıyor.")
#         else:
#             print(f"HATA: Statü kodu {cevap.status_code}")
#     except Exception as e:
#         print(f"Hata: {url} -> {e}")
#     time.sleep(1)

# print(f"\nAşama 1 tamamlandı: {len(toplanan_veriler)} şirket bulundu.\n")

# # ─────────────────────────────────────────────────────────────────────────────
# # AŞAMA 2: Mail arama — KeyboardInterrupt gelirse mevcut veriyi kaydet
# # ─────────────────────────────────────────────────────────────────────────────

# toplam = len(toplanan_veriler)
# i = 0
# try:
#     for i, kayit in enumerate(toplanan_veriler, start=1):
#         sirket_adi = kayit["Şirket İsmi"]
#         sirket_url = kayit["Şirket URL"]

#         print(f"[{i}/{toplam}] Mail aranıyor: {sirket_adi} → {sirket_url}")
#         mailler = sirket_maillerini_topla(sirket_url)

#         if mailler:
#             kayit["Mail Adresleri"] = " | ".join(mailler)
#             print(f"  ✓ Bulunan: {kayit['Mail Adresleri']}")
#         else:
#             kayit["Mail Adresleri"] = "Bulunamadı"
#             print(f"  – Mail bulunamadı")

#         time.sleep(1.5)

# # --- DÜZELTME: Ctrl+C ile kesilirse o ana kadar toplanan veriyi kaydet ---
# except KeyboardInterrupt:
#     print(f"\n\n⚠ Kullanıcı tarafından durduruldu. {i}/{toplam} şirket işlendi.")
#     print("Mevcut veriler kaydediliyor...\n")

# # ─────────────────────────────────────────────────────────────────────────────
# # AŞAMA 3: Excel'e kaydet
# # ─────────────────────────────────────────────────────────────────────────────

# islenenler = [k for k in toplanan_veriler if "Mail Adresleri" in k]

# if islenenler:
#     df = pd.DataFrame(islenenler, columns=[
#         "Şirket İsmi", "Şirket URL", "Mail Adresleri", "Çekildiği Kaynak"
#     ])

#     temel_isim = "odtu_teknokent_firmalar"
#     uzanti = ".xlsx"
#     sayac = 1
#     dosya_adi = f"{temel_isim}_{sayac}{uzanti}"
#     while os.path.exists(dosya_adi):
#         sayac += 1
#         dosya_adi = f"{temel_isim}_{sayac}{uzanti}"

#     df.to_excel(dosya_adi, index=False)
#     mail_bulunan = (df["Mail Adresleri"] != "Bulunamadı").sum()
#     print(f"✅ {len(df)} şirket kaydedildi → '{dosya_adi}'")
#     print(f"   Mail bulunan: {mail_bulunan} / {len(df)}")
# else:
#     print("\nUyarı: Hiç veri işlenemedi.")