# WhatsApp Toplu Gönderim Aracı

Excel dosyasındaki alıcılara WhatsApp Web üzerinden kişiselleştirilmiş mesajlar gönderen Python masaüstü uygulamasıdır. Gönderim durumu arayüzden izlenir; işlem sonunda Excel ve CSV raporları oluşturulur.

> [!IMPORTANT]
> Bu proje WhatsApp'ın resmi API'sini kullanmaz. Yalnızca mesaj almayı kabul etmiş kişilere, WhatsApp'ın kullanım koşullarına ve yürürlükteki mevzuata uygun şekilde kullanın. Hesap kısıtlaması ve veri işleme sorumluluğu kullanıcıya aittir.

## Özellikler

- Excel'den `phone`, `name` ve `message` alanlarını okuma
- `{name}` yer tutucusuyla kişiselleştirme
- SAFE, FAST ve TURBO gönderim hızları
- Gerçek zamanlı ilerleme ve hata günlüğü
- Excel ve CSV sonuç raporları
- Açık, koyu ve sistem teması

## Gereksinimler

- Python 3.10 veya üzeri
- Google Chrome
- WhatsApp hesabı

## Kurulum

Depoyu klonlayın ve proje klasörüne geçin:

```bash
git clone <depo-adresi>
cd Sendify_WhatsApp_Toplu_Mesaj_Otomasyonu-main
```

Sanal ortam oluşturup bağımlılıkları kurun:

```bash
python -m venv .venv
```

Windows:

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

macOS/Linux:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Uygulamayı başlatın:

```bash
python main.py
```

İlk çalıştırmada açılan WhatsApp Web sayfasındaki QR kodunu telefonunuzla okutmanız gerekebilir.

## Excel biçimi

`phone` sütunu zorunludur. `name` ve `message` isteğe bağlıdır.

| phone | name | message |
| --- | --- | --- |
| 905551112233 | Örnek Kullanıcı | Merhaba {name}, duyurumuz var. |

- Telefon numarasını ülke koduyla, `+` işareti olmadan yazın.
- Arayüzde mesaj şablonu girilmezse satırdaki `message` değeri kullanılır.
- İsim yoksa uygulama varsayılan bir hitap kullanır.

## Raporlar ve yerel veriler

Raporlar kullanıcının `Documents/WhatsAppBroadcastRuns` klasöründe oluşturulur. Kalıcı WhatsApp Web oturumu kullanıcı dizinindeki `whatsapp_profile` klasöründe tutulur. Bu dosyalar kişisel veri veya oturum bilgisi içerebileceğinden Git deposuna eklenmemelidir.

## Proje yapısı

```text
.
├── main.py                 # Uygulama giriş noktası
├── gui.py                  # Masaüstü arayüzü
├── broadcaster_logic.py    # Selenium, veri işleme ve raporlama
├── requirements.txt        # Python bağımlılıkları
└── assets/
    └── logo.png            # Arayüz logosu
```

`app/` klasörü geliştirme sırasında tutulan eski sürüm arşivlerini içerir; çalıştırılan güncel uygulama kök dizindeki dosyalardır.

## Geliştirme

Değişiklik göndermeden önce en azından sözdizimi kontrolünü çalıştırın:

```bash
python -m compileall main.py gui.py broadcaster_logic.py
```

Katkılar için ayrı bir dal açın, değişikliğinizi test edin ve açıklayıcı bir pull request oluşturun.

## Lisans

Bu depoda henüz bir lisans dosyası bulunmamaktadır. Bir lisans eklenene kadar kaynak kodun yeniden kullanımı veya dağıtımı için açık izin verilmiş sayılmaz.
