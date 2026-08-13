# TSB HD Kârlılık Analizi — Paylaşım Sürümü

Türkiye Sigorta Birliği (TSB) finansal tabloları kullanılarak hayat dışı sigorta sektörünün kârlılık dinamiklerini incelemek için hazırlanmış Streamlit demo uygulamasıdır.

## Kapsam

- 2025H1 ve 2026H1 karşılaştırması
- Hayat dışı (HD) şirketler
- 18 ana branş
- Sektör Özeti
- Branş Analizi
- Şirket Detayı
- Metodoloji

Bu paket **salt okunur paylaşım sürümüdür**. Veri Güncelleme ekranı bilerek kaldırılmıştır.

## Sunum yaklaşımı

Uygulama iki ayrı görünüm kullanır:

1. **Brüt Teknik Görünüm:** Brüt Yazılan Prim, Brüt Kazanılmış Prim, Brüt H/P, Masraf Oranı, Brüt Bileşik Oran.
2. **Teknik Sonuç Görünümü:** Brüt Yazılan Prim (hacim referansı), Mali Gelir Aktarımı Hariç Teknik Sonuç, Mali Gelir Aktarımı (603), Aktarım Dahil Teknik Sonuç.

Şirket ekranlarında niteliksel sınıflama yapılmaz; yalnızca sayısal değerler, dönemsel değişimler ve benchmark farkları gösterilir.

## Streamlit Community Cloud ile yayınlama

Repository kökünde şu dosyalar bulunmalıdır:

```text
app.py
initial_history.json
tsb_engine.py
requirements.txt
.streamlit/config.toml
README.md
```

1. GitHub'da yeni bir repository oluştur.
2. Bu klasördeki dosyaları repository'nin köküne yükle ve commit et.
3. Streamlit Community Cloud'a GitHub hesabınla giriş yap.
4. **Create app** seçeneğini aç.
5. Repository, branch (`main`) ve entrypoint olarak `app.py` seç.
6. İstersen özel bir `*.streamlit.app` alt alan adı belirle.
7. Deploy et.
8. Uygulama açıldıktan sonra paylaşım ayarlarından public/private erişimi kontrol et.

Not: `requirements.txt` dosyası bu pakete dahildir ve uygulamanın Python bağımlılıklarını tanımlar.

## Yerelde test

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Veri kaynağı ve metodoloji notu

Kaynak: Türkiye Sigorta Birliği (TSB) finansal tabloları. Uygulamadaki “Mali Gelir Aktarımı Hariç Teknik Sonuç” analitik olarak `Teknik Kâr/Zarar - 603` şeklinde hesaplanır ve brüt bileşik oranla aynı muhasebe tabanında bir underwriting sonucu olarak yorumlanmamalıdır.
