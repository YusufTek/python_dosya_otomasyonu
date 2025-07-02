import os
import zipfile

# Jupyter Notebook ortamında çalıştırılmayacaksa bu işlemi yapmanız gerekebilir.
# Script'in bulunduğu dizini al
# __file__ = bu Python dosyasının tam yolu
# abspath() = mutlak yolu verir (C:\Users\... gibi)
# dirname() = dosyanın bulunduğu klasörü verir.
base_dir = os.path.dirname(os.path.abspath(__file__))

# Belgeler klasörünün yolu ve ZIP dosyasının adı
kaynak_klasor = os.path.join(base_dir, "Belgeler")
zip_adi = os.path.join(base_dir, "yedek.zip")

# ZIP dosyasını yazma modunda aç!
# 'w' = write (yazma) modu - yeni ZIP dosyası oluştur.
# with statement kullanarak dosya otomatik kapanır.
with zipfile.ZipFile(zip_adi, 'w') as zip_arsivi:
    for dosya in os.listdir(kaynak_klasor):
        if dosya.endswith('.txt'):
            tam_yol = os.path.join(kaynak_klasor, dosya)
            # Dosyayı ZIP arşivine ekle
            zip_arsivi.write(tam_yol, arcname=dosya)
            print(f"{dosya} -> ZIP'e eklendi.")