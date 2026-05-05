# =============================================================================
# mail_gonder.py — Excel'den oku, CV ekle, mailleri at
# =============================================================================
 
import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import random
import os
from config import (
    EMAIL_ADRESIN, SIFREN, CV_DOSYA_YOLU,
    TXT_DOSYASI, KONU, GIRIS_EXCEL,
    BEKLEME_MIN, BEKLEME_MAX,
)
 
 
def calistir(excel_yolu: str = None):
    # main.py scraper çıktısını doğrudan iletebilir, yoksa config'deki yol kullanılır
    excel_yolu = excel_yolu or GIRIS_EXCEL
 
    # --- 1. TXT oku ---
    try:
        with open(TXT_DOSYASI, "r", encoding="utf-8") as f:
            sablon = f.read()
        if not sablon.strip():
            print("HATA: mail_icerik.txt boş!")
            return
        print(f"✓ Mail içeriği okundu ({len(sablon)} karakter)")
    except FileNotFoundError:
        print(f"HATA: '{TXT_DOSYASI}' bulunamadı!")
        return
 
    # --- 2. CV kontrol ---
    if not os.path.exists(CV_DOSYA_YOLU):
        print(f"HATA: '{CV_DOSYA_YOLU}' bulunamadı!")
        return
    with open(CV_DOSYA_YOLU, 'rb') as f:
        dosya_verisi = f.read()
    print(f"✓ CV yüklendi: {CV_DOSYA_YOLU}")
 
    # --- 3. Excel oku ---
    try:
        df = pd.read_excel(excel_yolu)
        df = df[df["Mail Adresleri"] != "Bulunamadı"].dropna(subset=["Mail Adresleri"])
        alici_listesi = df[["Şirket İsmi", "Mail Adresleri"]].to_dict("records")
        print(f"✓ {len(alici_listesi)} şirket yüklendi (mail adresi olan)\n")
    except Exception as e:
        print(f"HATA: Excel okunamadı → {e}")
        return
 
    # --- 4. SMTP bağlan ---
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_ADRESIN, SIFREN)
        print("✓ Gmail bağlantısı kuruldu\n")
    except Exception as e:
        print(f"HATA: Gmail'e bağlanılamadı → {e}")
        return
 
    # --- 5. Gönderim döngüsü ---
    basarili  = 0
    basarisiz = 0
    toplam    = len(alici_listesi)
 
    try:
        for i, alici in enumerate(alici_listesi, 1):
            sirket_adi = alici["Şirket İsmi"]
            alici_mail = alici["Mail Adresleri"]
 
            # Bir şirkete birden fazla mail varsa " | " ile ayrılmış olabilir
            for mail in [m.strip() for m in alici_mail.split("|")]:
                try:
                    icerik = sablon.replace("[ŞİRKET_ADI]", str(sirket_adi))
 
                    msg = EmailMessage()
                    msg['Subject'] = KONU
                    msg['From']    = EMAIL_ADRESIN
                    msg['To']      = mail
                    msg.set_content(icerik)
                    msg.add_attachment(
                        dosya_verisi,
                        maintype='application',
                        subtype='pdf',
                        filename=os.path.basename(CV_DOSYA_YOLU)
                    )
 
                    server.send_message(msg)
                    basarili += 1
                    print(f"[{i}/{toplam}] ✓ Gönderildi → {mail} ({sirket_adi})")
 
                except Exception as e:
                    basarisiz += 1
                    hata = f"[{i}/{toplam}] ✗ Hata → {sirket_adi} ({mail}) | {e}"
                    print(hata)
                    with open("hatali_gonderimler.txt", "a", encoding="utf-8") as hf:
                        hf.write(hata + "\n")
 
            if i < toplam:
                bekleme = random.uniform(BEKLEME_MIN, BEKLEME_MAX)
                print(f"  {bekleme:.1f}s bekleniyor...")
                time.sleep(bekleme)
 
    except KeyboardInterrupt:
        print(f"\n⚠ Durduruldu. {i}/{toplam} şirket işlendi.")
 
    # --- 6. Özet ---
    server.quit()
    print(f"\n{'─'*40}")
    print(f"✅ Mail gönderimi tamamlandı!")
    print(f"   Başarılı  : {basarili}")
    print(f"   Başarısız : {basarisiz}")
    print(f"   Toplam    : {toplam}")
 
 
if __name__ == "__main__":
    calistir()