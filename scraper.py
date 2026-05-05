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
    # Türkçe
    "/iletisim", "/iletişim", "/bize-ulasin", "/bize-ulaşın",
    "/iletisim.html", "/iletisim/", "/iletisim.php",
    "/hakkimizda", "/hakkımızda", "/hakkinda", "/hakkında",
    "/tr/iletisim", "/tr/bize-ulasin",
    # İngilizce
    "/contact", "/contact-us", "/contact.html", "/contact/",
    "/contact.php", "/contactus", "/get-in-touch",
    "/en/contact", "/eng/contact", "/eng/contacts.html",
    "/about", "/about-us", "/about.html",
    # Hash/anchor tabanlı (tek sayfa siteler)
    "/#iletisim", "/#iletişim", "/#contact", "/#contact-section",
    "/#bize-ulasin", "/#iletisim-formu",
    # Karma yapılar (bazı CMS'lerin garip yolları)
    "/TR,16/iletisim.html", "/pages/contact", "/sayfalar/iletisim",

]

GIZLI_MAIL_REGEX = [
    # info(at)domain.com  /  info[at]domain.com  /  info{at}domain.com
    re.compile(
        r'[a-zA-Z0-9._%+\-]+\s*[\(\[\{]at[\)\]\}]\s*[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
        re.IGNORECASE
    ),
    # info AT domain DOT com  /  info at domain dot com
    re.compile(
        r'[a-zA-Z0-9._%+\-]+\s+[aA][tT]\s+[a-zA-Z0-9.\-]+\s+[dD][oO][tT]\s+[a-zA-Z]{2,}',
        re.IGNORECASE
    ),
    # info at domain.com  (boşluklu ama nokta normal)
    re.compile(
        r'[a-zA-Z0-9._%+\-]+\s+[aA][tT]\s+[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
        re.IGNORECASE
    ),
]

def gizli_mail_temizle(metin: str) -> str:
    """(at), [at], {at}, AT → @ dönüştürür; DOT → . dönüştürür."""
    metin = re.sub(r'\s*[\(\[\{]at[\)\]\}]\s*', '@', metin, flags=re.IGNORECASE)
    metin = re.sub(r'\s+at\s+', '@', metin, flags=re.IGNORECASE)
    metin = re.sub(r'\s*[\(\[\{]dot[\)\]\}]\s*', '.', metin, flags=re.IGNORECASE)
    metin = re.sub(r'\s+dot\s+', '.', metin, flags=re.IGNORECASE)
    return metin

GECERSIZ_MAILLER = {
    "example.com", "domain.com", "email.com", "test.com", "yoursite.com",
    # Hata takip sistemleri
    "sentry.io", "sentry-next.wixpress.com", "bugsnag.com", "rollbar.com",
    # Site oluşturma platformları
    "wix.com", "wixpress.com", "wordpress.com", "squarespace.com",
    "webflow.io", "ghost.io", "weebly.com", "jimdo.com", "strikingly.com",
    # Mail/pazarlama platformları
    "mailchimp.com", "sendgrid.com", "sendgrid.net", "mailgun.org",
    "hubspot.com", "klaviyo.com", "constantcontact.com",
    # Sosyal medya (profil URL'lerinden sızan adresler)
    "facebook.com", "twitter.com", "instagram.com", "linkedin.com",
    "youtube.com", "tiktok.com",
    # Altyapı/CDN
    "amazonaws.com", "cloudflare.com", "googletagmanager.com",
    "google-analytics.com", "googleapis.com",
    # Yorum/destek sistemleri
    "disqus.com", "zendesk.com", "intercom.io", "crisp.chat",
    # Schema/W3C meta verileri
    "schema.org", "w3.org", "ogp.me",
    # Diğer
    "gravatar.com", "shopify.com", "medium.com",
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
    """Sayfadan normal ve gizlenmiş (at) formatlı mailleri toplar."""
    temizler = set()
 
    # 1. Normal @ formatı
    for m in EMAIL_REGEX.findall(html):
        domain = m.split("@")[1].lower()
        if domain not in GECERSIZ_MAILLER and not m.endswith((".png", ".jpg", ".gif", ".svg")):
            temizler.add(m.lower())
 
    # 2. Gizlenmiş (at)/(AT) formatı — önce metni normalleştir, sonra tekrar ara
    temiz_html = gizli_mail_temizle(html)
    for m in EMAIL_REGEX.findall(temiz_html):
        domain = m.split("@")[1].lower()
        if domain not in GECERSIZ_MAILLER and not m.endswith((".png", ".jpg", ".gif", ".svg")):
            temizler.add(m.lower())
 
    return temizler

def iletisim_linkleri_bul(soup, base_url: str) -> list:
    """
    Sayfanın kendi linklerini tarayarak iletişim sayfalarını otomatik bulur.
    Sabit listeye takılmayan garip URL yapılarını da yakalar.
    """
    anahtar_kelimeler = [
        "iletisim", "iletişim", "contact", "bize-ulas", "bize ulaş",
        "ulaşın", "ulasin", "reach", "get in touch", "hakkimizda",
        "hakkında", "about",
    ]
    bulunan_linkler = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        metin = a.get_text(strip=True).lower()
        href_lower = href.lower()
 
        eslesme = any(k in href_lower or k in metin for k in anahtar_kelimeler)
        if not eslesme:
            continue
 
        # Tam URL yap
        if href.startswith("http"):
            tam_url = href
        elif href.startswith("/"):
            tam_url = base_url + href
        else:
            continue
 
        if tam_url not in bulunan_linkler:
            bulunan_linkler.append(tam_url)
 
    return bulunan_linkler[:5]  # En fazla 5 adet dene


def sirket_maillerini_topla(sirket_url: str) -> list:
    if not sirket_url or sirket_url == "Link Bulunamadı":
        return []
    if not sirket_url.startswith(("http://", "https://")):
        sirket_url = "https://" + sirket_url
 
    tum_mailler = set()
    base = f"{urlparse(sirket_url).scheme}://{urlparse(sirket_url).netloc}"
 
    def sayfadan_topla(html: str) -> set:
        """HTML'den hem regex hem mailto: linklerinden mail toplar."""
        mailler = mail_bul(html)
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            if a["href"].startswith("mailto:"):
                mail = a["href"].replace("mailto:", "").split("?")[0].strip().lower()
                if mail and "@" in mail:
                    mailler.add(mail)
        return mailler, soup
 
    # 1. Ana sayfa
    ana_html = sayfa_getir(sirket_url)
    ana_soup = None
    if ana_html:
        bulunanlar, ana_soup = sayfadan_topla(ana_html)
        tum_mailler |= bulunanlar
 
    if tum_mailler:
        return sorted(tum_mailler)
 
    # 2. Sitenin kendi navigasyonundan iletişim linki bul (otomatik keşif)
    if ana_soup:
        otomatik_linkler = iletisim_linkleri_bul(ana_soup, base)
        for link in otomatik_linkler:
            html = sayfa_getir(link, timeout=8)
            if html:
                bulunanlar, _ = sayfadan_topla(html)
                tum_mailler |= bulunanlar
            if tum_mailler:
                break
 
    if tum_mailler:
        return sorted(tum_mailler)
 
    # 3. Sabit yol listesini dene (otomatik keşif bulamazsa)
    for yol in ILETISIM_YOLLARI:
        html = sayfa_getir(base + yol, timeout=8)
        if html:
            bulunanlar, _ = sayfadan_topla(html)
            tum_mailler |= bulunanlar
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