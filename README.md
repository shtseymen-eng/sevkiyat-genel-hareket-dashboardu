# Sevkiyat Genel Hareket Dashboardu

Araç çıkışı, gemi girişi, gemi çıkışı ve boru hattı raporlarını tek ekranda analiz eder.

## Kullanım

Rapor ekle düğmesiyle birden fazla Excel dosyasını birlikte seçin. Hareket türü dosya içeriğinden tanınır. Dönem, firma, ürün ve tank filtreleriyle analiz yapabilir, seçili kapsamı renkli Excel raporu olarak indirebilirsiniz.

- Pivot gerekmez; ham hareket kayıtları kullanılır.
- Aynı plaka, tank ve kantar çıkış zamanı tek araç seferi sayılır; miktarlar toplanır.
- Eksik hareket türleri gösterilmez.
- Raporlar tarayıcı belleğinde işlenir; sunucuya veya GitHub deposuna gönderilmez. Sayfa yenilenince dosyaları tekrar seçmeniz gerekir.
- Miktarlar kaynakta kg kabul edilerek ekranda tona çevrilir.
- Farklı dosyalardaki örtüşen tarih aralıkları otomatik ayıklanmaz. Aynı içerikli dosya tekrar eklenmez.
- Türü tanınamayan veya geçersiz satırlar sayılarıyla belirtilir, hesaplara dahil edilmez.

## Yerel çalıştırma

Bu klasörü bir statik HTTP sunucusuyla açın. Örneğin: `python3 -m http.server 8765`.

## GitHub Pages

Settings → Pages → Deploy from a branch → main → /(root).

## Test

`node --test tests/*.test.mjs`

Harici kurulum gerektirmez. SheetJS ve JSZip tarayıcı paketleri ve lisansları vendor klasöründedir.
