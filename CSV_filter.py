import csv
import os

# Script konumunu al
base_dir = os.path.dirname(os.path.abspath(__file__))

# Dosya yollarını oluştur
girdi_dosyasi = os.path.join(base_dir, "employees.csv")
cikti_dosyasi = os.path.join(base_dir, "employees_output.csv")

# CSV dosyasını oku ve filtrele
with open(girdi_dosyasi, mode='r', encoding='utf-8') as girdi:
    reader = csv.DictReader(girdi)
    # Maaşı 100.000 ve üzeri olanları filtrele
    satirlar = list([satir for satir in reader if int(satir['SALARY']) >= 100000])
    fieldnames = reader.fieldnames

# ID'ye göre sırala
siralı_satirlar = sorted(satirlar, key=lambda x: int(x["ID"]))

# Sonucu yeni dosyaya yaz
with open(cikti_dosyasi, mode='w', encoding='utf-8', newline='') as cikti:
    writer = csv.DictWriter(cikti, fieldnames=fieldnames)
    writer.writeheader()  # Başlıkları yaz
    writer.writerows(siralı_satirlar)  # Veriyi yaz

print(f"{len(siralı_satirlar)} satır yazıldı: {cikti_dosyasi}")