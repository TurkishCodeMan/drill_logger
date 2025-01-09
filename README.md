# Excel Tablet App

Bu uygulama, büyük Excel dosyalarını tablet üzerinden daha kolay doldurmak için geliştirilmiş bir web uygulamasıdır.

## 🎯 Özellikler

- Excel dosyalarını (.xlsx, .xls) yükleme ve görüntüleme
- Tablet uyumlu kullanıcı arayüzü
- Akıllı veri tipi algılama ve doğrulama
- Otomatik tamamlama önerileri
- Tasarım özelliklerini koruma (renkler, yazı tipleri, kenarlıklar)
- Hızlı veri girişi için optimize edilmiş form arayüzü

## 🚀 Teknolojiler

- React + TypeScript
- Material-UI (MUI)
- XLSX.js
- React Hook Form

## 📋 Gereksinimler

- Node.js 18+
- npm veya yarn

## 🛠️ Kurulum

```bash
# Projeyi klonlayın
git clone [repo-url]

# Proje dizinine gidin
cd excel-tablet-app

# Bağımlılıkları yükleyin
npm install

# Geliştirme sunucusunu başlatın
npm run dev
```

## 📱 Kullanım

1. **Excel Dosyası Yükleme**
   - "Excel Dosyası Yükle" butonuna tıklayın
   - `.xlsx` veya `.xls` formatındaki dosyanızı seçin

2. **Veri Girişi**
   - Form alanlarını doldurun
   - "Satır Ekle" butonu ile yeni veri ekleyin
   - Otomatik tamamlama önerilerinden faydalanın

3. **Veri Doğrulama**
   - Zorunlu alanlar kırmızı ile işaretlenir
   - Sayısal alanlar için min/max kontrolleri
   - Tarih alanları için takvim seçici

4. **Kaydetme**
   - "Kaydet" butonuna tıklayın
   - Excel dosyası orijinal formatında indirilecektir

## 🎨 Tasarım Özellikleri

- Başlık satırları için özel renkler ve yazı tipleri
- Alternatif satır renkleri
- Kenarlık stilleri
- Tablet ekranına optimize edilmiş butonlar ve form alanları

## 🔄 Veri Dönüşümü

### Excel'den JSON'a
```typescript
interface ExcelSheet {
  name: string;
  columns: ExcelColumn[];
  rows: ExcelRow[];
  design: {
    headerColor: string;
    fontStyle: string;
    rowColors?: {
      even: string;
      odd: string;
    };
    borders?: {
      color: string;
      style: string;
    };
  };
}
```

### Veri Tipleri
- Text
- Number
- Date
- Select (dropdown)

## 🛡️ Veri Doğrulama

- Zorunlu alan kontrolü
- Sayısal değer aralığı kontrolü
- Tarih formatı kontrolü
- Özel regex pattern kontrolü

## 📊 Performans

- Lazy loading ile büyük dosya desteği
- Optimize edilmiş render işlemleri
- Önbelleğe alınmış form değerleri

## 🔜 Planlanan Özellikler

1. **Çoklu Sayfa Desteği**
   - Sayfalar arası kolay geçiş
   - Sayfa başına özel tasarım ayarları

2. **Gelişmiş Veri Doğrulama**
   - Özel formül desteği
   - Koşullu biçimlendirme

3. **Otomatik Kaydetme**
   - Yerel depolama desteği
   - Değişiklik geçmişi

4. **Veri Analizi**
   - Basit istatistikler
   - Veri görselleştirme

## 🤝 Katkıda Bulunma

1. Fork'layın
2. Feature branch oluşturun (`git checkout -b feature/amazing-feature`)
3. Commit'leyin (`git commit -m 'feat: Add amazing feature'`)
4. Push'layın (`git push origin feature/amazing-feature`)
5. Pull Request açın

## 📝 Lisans

MIT License - daha fazla detay için [LICENSE.md](LICENSE.md) dosyasına bakın.
