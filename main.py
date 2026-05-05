# =============================================================================
# main.py — Her şeyi sırayla çalıştır
#
# Kullanım:
#   python main.py          → önce scraper, sonra mail gönder
#   python main.py scrape   → sadece scraper (excel üret)
#   python main.py mail     → sadece mail gönder (mevcut excel'den)
# =============================================================================
 
import sys
import scraper
import mail_gonder
 
 
def tam_akis():
    print("=" * 50)
    print("  ADIM 1/2 — Web Scraping Başlıyor")
    print("=" * 50 + "\n")
    excel_ciktisi = scraper.calistir()
 
    if not excel_ciktisi:
        print("\nScraper çıktı üretemedi, mail gönderimi iptal edildi.")
        return
 
    print("\n" + "=" * 50)
    print("  ADIM 2/2 — Mail Gönderimi Başlıyor")
    print("=" * 50 + "\n")
    mail_gonder.calistir(excel_yolu=excel_ciktisi)
 
 
if __name__ == "__main__":
    mod = sys.argv[1] if len(sys.argv) > 1 else "hepsi"
 
    if mod == "scrape":
        # Sadece siteleri tara, excel üret
        scraper.calistir()
 
    elif mod == "mail":
        # Mevcut excel'den direkt mail at (config.py'deki GIRIS_EXCEL kullanılır)
        mail_gonder.calistir()
 
    else:
        # İkisini sırayla çalıştır
        tam_akis()