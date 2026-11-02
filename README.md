# 📊 Haftalık Otomasyon Raporu — Streamlit Web Uygulaması

Mevcut Excel tabanlı haftalık otomasyon raporunun (`Sheet1`, `Sheet2`, `Sheet3`, `Sheet4`) ekipçe web üzerinden girilebilmesini, otomatik formül hesaplamalarını ve her bir sayfanın (veya tüm raporun) **Outlook / Gmail uyumlu renkli tablo formatında tek tıkla kopyalanıp maile yapıştırılabilmesini** sağlayan Streamlit uygulamasıdır.

## 🚀 Hızlı Başlangıç (Yerelde Çalıştırma)

```bash
pip install -r requirements.txt
streamlit run app.py
```

- **Varsayılan Yerel Erişim Parolası:** `Otomasyon2026!` (`.streamlit/secrets.toml` dosyasından değiştirilebilir).

## ☁️ Streamlit Community Cloud (`*.streamlit.app`) Kurulumu

1. Projeyi GitHub deponuza yükleyin (`.gitignore` sayesinde `.streamlit/secrets.toml`, `*.xlsx` ve `data/` klasörü otomatik olarak gizlenir ve repo içerisine hassas veri sızmaz).
2. [share.streamlit.io](https://share.streamlit.io) üzerinden deponuzu bağlayın ve ana dosya olarak `app.py` seçin.
3. **App Settings -> Secrets** bölümüne giderek giriş parolanızı tanımlayın:
   ```toml
   APP_PASSWORD = "Ekip_Guclu_Parolaniz"
   ```

## 🛡️ Uygulanan Veri Güvenliği Prensipleri

- **Kimlik Doğrulama & Brute-Force Koruması (`security.py`):** `st.secrets["APP_PASSWORD"]` ile parola doğrulaması (`hmac.compare_digest` ile timing-attack korumalı), 5 hatalı denemede 5 dakika geçici kilitleme ve 4 saatlik oturum zaman aşımı.
- **XSS Koruması:** Mail tabloları üretilirken tüm kullanıcı girdileri `html.escape` ile temizlenir.
- **CSV / Excel Formül Enjeksiyonu Koruması (CWE-1236):** `=`, `+`, `-`, `@` ile başlayan metin girdileri Excel çıktısında otomatik olarak güvenli metne dönüştürülür.
- **Dosya Yükleme Güvenliği:** Yüklenen `.xlsx` dosyalarında 5 MB boyut sınırı, uzantı kontrolü ve ZIP/XLSX magic byte (`PK\x03\x04`) doğrulaması yapılır.
- **Git / Repo Güvenliği (`.gitignore`):** Gerçek Jira kayıtları içeren `.xlsx` dosyaları, `data/` klasörü ve `secrets.toml` dosyası Git takibinden hariç tutulmuştur.

