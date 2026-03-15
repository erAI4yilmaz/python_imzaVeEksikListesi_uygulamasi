# 📋 İmza ve Eksik Liste Otomasyonu

> Makro Muayene şirketlerinde personel imza takibini ve eksik belge yönetimini otomatikleştiren Python uygulaması.

---

## 🎯 Proje Hakkında

Bu uygulama; makine muayene şirketlerinde günlük olarak yürütülen manuel Excel işlemlerini otomatikleştirmek amacıyla geliştirilmiştir.

Uygulama sayesinde:
- Bir Excel dosyasındaki veriler, başka bir Excel dosyasına **tek tuş** ile aktarılabilmektedir.
- Personel imza işlemleri, Adobe Acrobat'ın e-imza özelliğinin iş yükü devre dışı bırakılarak **otomatik olarak** gerçekleştirilmektedir.
- Eksik belge listeleri otomatik olarak tespit edilip raporlanmaktadır.

Sonuç olarak günde birden fazla kez tekrarlanan bu işlemler için harcanan zaman önemli ölçüde azaltılmıştır.

---

## ⚙️ Özellikler

- ✅ Excel'den Excel'e tek tuşla veri aktarımı
- ✅ Otomatik e-imza işleme (Acrobat bağımsız)
- ✅ Eksik belge tespiti ve listeleme
- ✅ Sade ve kullanımı kolay arayüz

---

## 🛠️ Kullanılan Teknolojiler

| Teknoloji | Amaç |
|---|---|
| Python 3 | Ana programlama dili |
| openpyxl / pandas | Excel okuma ve yazma işlemleri |
| PyInstaller | .exe olarak derleme (spec dosyası ile) |

---

## 🚀 Kurulum ve Çalıştırma

### Gereksinimler
```bash
pip install openpyxl pandas
```

### Çalıştırma
```bash
python eklemeOtomasyonuV3.py
```

### .exe Olarak Derleme
```bash
pyinstaller eklemeOtomasyonuV3.spec
```

---

## 📁 Proje Yapısı

```
📦 python_imzaVeEksikListesi_uygulamasi
 ┣ 📄 eklemeOtomasyonuV3.py     # Ana uygulama
 ┣ 📄 eklemeOtomasyonuV3.spec   # PyInstaller derleme ayarları
 ┣ 🖼️  logo.ico                  # Uygulama ikonu
 ┣ 🖼️  logo.png                  # Logo görseli
 ┗ 📄 README.md
```

> ⚠️ Personellere ait imza görselleri (PNG formatında) gizlilik nedeniyle repoya dahil edilmemiştir.

---

## 💡 Motivasyon

Bu proje, staj sürecimde çalıştığım şirkette gözlemlediğim tekrarlayan manuel işlemleri yazılım ile çözme fikrinden doğmuştur. Gerçek bir iş problemini tespit edip çözüme kavuşturmak amacıyla kişisel inisiyatifimle geliştirilmiştir.

---

## 👨‍💻 Geliştirici

**erAI4yilmaz**  
Bilgisayar Programcılığı Mezunu  
[GitHub Profilim](https://github.com/erAI4yilmaz)
