# KiloHesap (CorelDRAW)

**CorelDRAW için kilo hesaplama makrosu** — Çizimdeki şekilleri seçin, malzemeyi seçin, makro toplam kiloyu (kg) hesaplar. Açık kaynak VBA makrosudur, CorelDRAW 2021 için hazırlanmıştır; 2019–2024 sürümlerinde de çalışması gerekir.

**Açık kaynak VBA makrosu, CorelDRAW 2019–2024 (2021 için hazırlanmıştır) için ağırlık hesaplayıcı.**

## Özellikler / Features

- CorelDRAW'da seçili şekillerin toplam alanını (cm²) ölçer
- `Alan × Kalınlık × Yoğunluk` formülüyle kilo hesabı yapar
- Gruplar, iç içe gruplar, metinler ve PowerClip içerikleri desteklenir
- Birikimli toplam: farklı malzemeler için tekrar tekrar hesaplayabilirsiniz
  (örn. aynı işte hem Pleksi 5 mm hem Pleksi 8 mm)
- Satır silme, sıfırlama, özeti çizim sayfasına yazma
- Malzeme listesi `malzemeler.json` dosyasından düzenlenebilir
- Modeless pencere: makro penceresi açıkken CorelDRAW'da çalışmaya devam edersiniz

## Kurulum (Windows)

> **Ön koşul:** CorelDRAW kurulumunda VBA özelliği kurulu olmalı (normal kurulumda varsayılan olarak gelir). Kontrol: CorelDRAW'da `Alt+F11` VBA editörünü açıyorsa hazırsınız.

1. Bu repoyu indirin veya klonlayın:
   ```
   git clone https://github.com/Volagejs/KiloHesap-CDR.git
   ```
2. `install\Kurulum.bat` dosyasına çift tıklayın.
   Betik CorelDRAW sürümünüzü otomatik bulur, dosyaları GMS klasörüne kopyalar
   ve CorelDRAW COM üzerinden modülleri **otomatik import etmeyi dener**.
3. Kurulumun sonunda açılan **KurulumRehberi.txt** size durumu söyler:
   - **Otomatik kayıt başarılıysa** hiçbir şey yapmanız gerekmez.
   - **Tek adım kaldıysa:** CorelDRAW'da `Alt+F11` → VBA editöründe `Ctrl+S` →
     GlobalMacros.gms kaydedilir. Bu ~10 saniyelik tek elle adımdır.
   - **Otomatik import olmadıysa:** rehberdeki elle import adımlarını uygulayın
     (`File > Import File...` ile `KiloHesap.bas` ve `KiloHesapForm.frm`).
4. CorelDRAW'da `Window > Dockers > Macro Manager` (veya `Alt+Shift+F11`) açın,
   listede **GlobalMacros > KiloHesapModul > KiloHesapGoster** makrosunu çift tıklayın.
   İsterseniz makroyu araç çubuğuna sürükleyip buton da yapabilirsiniz.

### Manuel kurulum (betiksiz)

`src` klasöründeki 3 dosyayı kopyalamanız yeterli:
- `KiloHesap.bas` ve `KiloHesapForm.frm` → VBA editöründen import edilir (yukarıdaki adımlar)
- `malzemeler.json` → `%APPDATA%\Corel\CorelDRAW Graphics Suite 2021\Draw\GMS\KiloHesap\` klasörüne kopyalanır
  (klasör yoksa oluşturun; CorelDRAW'ı bir kez açıp kapatmak klasörün oluşmasını sağlar)

### Kaldırma

`install\Kaldir.bat` dosyasına çift tıklayın.

## Kullanım

1. Pencerede malzeme/kalınlık seçin (örn. `Pleksi 5 mm`).
2. CorelDRAW'da kilosunu hesaplamak istediğiniz şekilleri seçin.
3. **Hesapla**'ya basın — sonuç listeye eklenir ve toplam birikir.
4. Farklı malzemeler için seçimi değiştirip tekrar hesaplayın.
5. **Sayfaya yaz** ile özeti çizime metin olarak ekleyebilirsiniz.

## Nasıl çalışır

```
kilo (g) = alan (cm²) × kalınlık (cm) × yoğunluk (g/cm³)
```

Alan ölçümü, şeklin eğri (curve) alanı üzerinden yapılır: düz şekiller (dikdörtgen,
elips, metin vb.) geçici bir kopya üzerinden curve'e çevrilip alanı okunur ve kopya
silinir. Çizimde kalıcı bir değişiklik olmaz (tek adımda geri alınabilir).

## Varsayılan malzemeler

| Malzeme | Kalınlık (mm) | Yoğunluk (g/cm³) |
|---|---|---|
| Pleksi | 3 / 5 / 8 / 10 | 1,19 |
| Dekota (MDF) | 9 / 14 / 18 | 0,75 |
| Krom | 1 | 7,90 |
| Aluminyum | 2 | 2,70 |
| Krom Yan Sac | 0,60 | 7,90 |
| Aluminyum Yan Sac | 0,70 | 2,70 |

Bu liste `src\malzemeler.json` (kurulum sonrası GMS klasöründeki kopyası) düzenlenerek
değiştirilebilir: yeni malzeme, yeni kalınlık veya farklı yoğunluk ekleyebilirsiniz.

## Sınırlamalar / Bilinen durumlar

- **Üst üste binen şekiller:** binen alan iki kez sayılır (lazer kesim teklifi
  mantığına uyar). Kesişen çizimlerde dikkatli olun.
- **Kontur (çizgi) kalınlığı** hesaba katılmaz; kapalı şekillerin iç alanı ölçülür.
- **Açık eğriler (çizgiler)** alanı olmadığından kiloya katkı vermez.
- Yoğunluk değerleri yaklaşıktır; malzemenizin gerçek yoğunluğunu JSON'a yazabilirsiniz.
- CorelDRAW "Home and Student" sürümünde VBA bulunmaz.

## Katkı

PR'lar ve sorun bildirimleri açıktır. Özellikle farklı malzemeler için yoğunluk
değerleri ve CorelDRAW sürüm uyumluluk geri bildirimleri değerli.

## Lisans

[MIT](LICENSE) © 2026 Volagejs
