import requests
from bs4 import BeautifulSoup
import csv
import os

# Script konumunu al
base_dir = os.path.dirname(os.path.abspath(__file__))
csv_dosyasi = os.path.join(base_dir, "kitaplar.csv")

try:
    # Web sitesine istek gönder
    url = "https://books.toscrape.com/"
    response = requests.get(url)
    response.raise_for_status()  # HTTP hatalarını kontrol et

    # HTML içeriğini çözümle
    soup = BeautifulSoup(response.text, 'html.parser')
    kitaplar = soup.select(".product_pod")  # Kitap elementlerini seç

    veriler_listesi = []

    # Her kitap için bilgileri çıkar
    for kitap in kitaplar:
        isim = kitap.h3.a['title']  # Kitap adı
        fiyat = kitap.select_one(".price_color").text  # Fiyat bilgisi
        stok = kitap.select_one(".availability").text.strip()  # Stok durumu
        
        # Veriyi listeye ekle
        veriler_listesi.append({
            "isim": isim,
            "fiyat": fiyat,
            "stok": stok
        })

    # Verileri CSV dosyasına yaz
    with open(csv_dosyasi, mode='w', encoding='utf-8', newline='') as dosya:
        writer = csv.DictWriter(dosya, fieldnames=["isim", "fiyat", "stok"])
        writer.writeheader()  # Başlıkları yaz
        writer.writerows(veriler_listesi)  # Tüm veriyi yaz

    print(f"Kitap verileri '{csv_dosyasi}' dosyasına kaydedildi.")

except requests.RequestException as e:
    print(f"Web sitesine erişim hatası: {e}")
except Exception as e:
    print(f"Bir hata oluştu: {e}")