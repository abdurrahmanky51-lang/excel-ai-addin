# Excel AI Asistanı — Kurulum Rehberi (v2, ücretsiz & sınırsız)

Bu sürüm **hiçbir API anahtarı istemez**. AI motoru olarak Puter.js
kullanıyor — geliştirici (yayınlayan kişi) için tamamen ücretsiz ve
sınırsızdır. Kullanıcının ilk AI işleminde, Puter tarafından **ücretsiz,
tek seferlik bir giriş penceresi** çıkabilir (sadece e-posta, kart bilgisi
istemez); bir daha karşısına çıkmaz.

---

## A) Senin yapman gereken (bir kere, yayıncı olarak)

### 1. Dosyaları GitHub'a yükle

Bu klasördeki tüm dosyaları (`manifest.xml`, `taskpane.html`, `taskpane.css`,
`taskpane.js`, `assets/` klasörü) `abdurrahmanky51-lang/excel-ai-addin`
reponuza yükleyin — **mevcut eski dosyaların üzerine yazacak şekilde**
(aynı isimde olduğu için otomatik değişecek).

GitHub'da: repo sayfası → **Add file → Upload files** → tüm dosyaları
sürükle bırak → alta **Commit changes** yaz, onayla.

### 2. Yayının güncellendiğini doğrula

1-2 dakika bekleyip şu adresi **Ctrl+F5** ile sert yenile:
`https://abdurrahmanky51-lang.github.io/excel-ai-addin/manifest.xml`

İçinde `<AppDomain>https://js.puter.com</AppDomain>` ve
`<Version>2.0.0.0</Version>` görmen lazım.

### 3. Kendi Excel'inde test et

**Ekle → Eklentilerim → Eklentimi Yükle** ile şu dosyayı seç:
`manifest.xml` (bilgisayarına indirdiğin hali). Şeritte "AI Asistan"
grubu çıkmalı.

---

## B) Başka birinin bu eklentiyi kurması için

Eklentini herkesin kolayca ekleyebilmesi için **en güvenilir ve evrensel
yöntem**, manifest dosyasını doğrudan yüklemektir (Office Store/AppSource
üzerinden resmi yayın yapmadıkça, "tek tıkla ekle" linki gibi bir şey
Microsoft tarafında yok — ama bu yöntem 30 saniye sürer):

1. Şu linke git ve dosyayı indir (sağ tık → Farklı Kaydet, veya tarayıcıda
   aç → Ctrl+S):
   `https://abdurrahmanky51-lang.github.io/excel-ai-addin/manifest.xml`
2. Excel'i aç → **Ekle** sekmesi → **Eklentilerim**
3. **Eklentimi Yükle / Upload My Add-in** → indirdiği `manifest.xml`
   dosyasını seç
4. Şeritte "AI Asistan" grubu belirir, tıklayıp panel açılır — hiçbir
   kayıt, kart bilgisi veya API key istemeden kullanmaya başlar

> Not: "Güvenilen Eklenti Katalogları" (network/SharePoint kataloğu)
> yöntemi kurumsal ortamlar için tasarlanmıştır ve genelde tek bir GitHub
> Pages adresiyle güvenilir biçimde çalışmaz — bu yüzden yukarıdaki
> "Eklentimi Yükle" yöntemini kullan, tek dosyayla anında çalışır.

### Excel Web (office.com) için

Excel Online'da bir dosya aç → **Ekle → Eklentiler → Eklentimi Yükle** →
aynı `manifest.xml` dosyasını seç.

---

## Kullanım

- **Formül**: Ne yapmak istediğini yaz → önerilen formülü tek tıkla
  seçili hücreye uygula
- **Hata Bul**: Bir aralık seç → formüllerdeki hataları ve düzeltmesini gör
- **Tablo**: Veri aralığı seç → otomatik biçimli Excel tablosuna çevir
- **Grafik**: Veri aralığı seç → en uygun grafik türü otomatik oluşturulur
- **Soru Sor**: Serbest metinle her türlü Excel sorusu sor

## Bilmen gerekenler

- **Tamamen ücretsiz ve sınırsız** — Puter.js'in "kullanıcı kendi
  kullanımını karşılar" modeli sayesinde ne sen ne de kullanıcı bir
  API faturası görmez.
- İlk AI isteğinde Puter'ın ücretsiz giriş penceresi çıkabilir — bu
  Puter.com'un kendi sistemi, senin sitenle ilgisi yok, veriler
  Puter'ın altyapısı üzerinden AI modeline gider.
- İnternet bağlantısı gerektirir.
- `taskpane.js` içindeki `MODEL` sabitini değiştirerek farklı bir model
  kullanabilirsin (Puter 400'den fazla model destekliyor).
