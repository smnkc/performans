# e-Okul Not Çekici & Ölçek Dağıtıcı 🚀

Bu uygulama, öğretmenlerin e-Okul **Hızlı Not Girişi** ekranındaki performans veya proje notlarını otomatik olarak çekip, MEB yönetmeliklerine %100 uyumlu, yatay hizalanmış ve yazdırma için A4 boyutuna sığdırılmış detaylı ölçek Excel belgeleri oluşturmasını sağlar.

---

## 🌟 Öne Çıkan Özellikler

* **Otomatik Veri Çekme:** e-Okul'daki tüm sınıf ve dersleri saniyeler içinde tarar ve öğrencilerin girilmiş olan notlarını okur.
* **Dinamik Ölçek Dağıtımı:** Öğrencinin genel notunu (0-100) belirlediğiniz kriter sayısına göre dengeli ve adil bir şekilde dağıtır.
* **Deterministik Karıştırma (Shuffle):** Aynı notu alan farklı öğrencilerin kriter puanları, öğrenci numaralarına göre benzersiz şekillerde karıştırılır. Böylece tüm ölçekler özgün ve organik görünür.
* **Yazıcı Dostu Sayfa Düzeni:** Oluşturulan Excel sekmeleri otomatik olarak **Yatay (Landscape)** olarak ayarlanır, A4 boyutuna göre ölçeklenir ve tek bir kağıda tam sığacak şekilde yapılandırılır.
* **Gelişmiş Arayüz İçi Editör:** Kendi ölçek kriterlerinizi (performans veya proje) uygulama içerisindeki sekmeden satır satır kolayca düzenleyip kaydedebilirsiniz. Excel şablonu kriter sayınıza göre kendini otomatik olarak genişletir.
* **Hatırlama Sistemi:** Öğretmen adı, branş ve okul müdürü gibi imza alanları sonraki açılışlarda otomatik olarak hatırlanır.
* **Taşınabilir (Standalone) EXE:** Bilgisayara Python veya harici bir kütüphane kurulmasına gerek yoktur. Sadece `.exe` dosyasını çalıştırmanız yeterlidir.

---

## 🛠️ Nasıl Kullanılır?

1. Klasör içerisindeki **`eOkul_Not_Cekici.exe`** dosyasına çift tıklayarak uygulamayı başlatın.
2. Arayüzde öğretmen adı, branşı, müdür adı ve değerlendirme tipini (Performans/Proje) seçin.
3. Excel dosyasının kaydedileceği konumu seçin.
4. **"1. Tarayıcıyı Aç ve e-Okul'a Git"** butonuna tıklayın. Açılan Chrome penceresinde e-Okul'a giriş yapın ve sol menüden **Hızlı Not Girişi** ekranına gidin. (Sınıf seçimi yapmayın, sadece sayfaya ulaşın).
5. Uygulamaya dönüp **"2. Verileri Çek ve Excel Üret"** butonuna basın. İşlem başladığında program tüm sınıfları sırayla tarayacak ve Excel belgesini üretecektir.

---

## 📝 Ölçek Kriterlerini Değiştirme

Uygulamanın içindeki **"Ölçek Kriterleri Düzenle"** sekmesini açarak performans veya proje kriterlerinizi satır satır düzenleyip kaydedebilirsiniz. Yazdığınız kriter sayısına göre Excel'deki formüller ve sütunlar otomatik olarak güncellenecektir.

---

## 👨‍💻 Geliştiriciler İçin

Kodları bilgisayarınızda python ile çalıştırmak isterseniz aşağıdaki kütüphanelerin yüklü olduğundan emin olun:

```bash
pip install selenium openpyxl ttkbootstrap webdriver-manager
```

Ardından uygulamayı başlatabilirsiniz:
```bash
python app.py
```

---

## 📄 Lisans
Bu proje eğitim ve öğretmenlerin iş yükünü hafifletme amacıyla açık kaynak olarak geliştirilmiştir.
İyi çalışmalar dileriz! 🎓
