import pandas as pd
import time
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import xlsxwriter # Pandas'ın Excel raporlama için kullandığı kütüphane

# --- Global Yapılandırma ve Sabitler ---
# Chrome kullanıcı verilerinin saklanacağı yol. QR kodunu tekrar okutmamak için.
CHROME_PROFILE_PATH = os.path.join(os.path.expanduser("~"), "whatsapp_profile") 

# Raporların kaydedileceği ana klasör yolu.
REPORT_BASE_DIR = os.path.join(os.path.expanduser("~"), "Documents", "WhatsAppBroadcastRuns")
os.makedirs(REPORT_BASE_DIR, exist_ok=True)


class BroadcasterLogic:
    """
    Tüm Selenium otomasyonu, veri işleme (Pandas) ve raporlama mantığını içerir.
    Arayüz (GUI) sınıfı, bu sınıfın metotlarını çağırarak arka plan işlemlerini yönetir.
    """
    def __init__(self, gui_app):
        # GUI (Arayüz) referansını tutar, böylece durum güncellemeleri arayüze gönderilebilir.
        self.gui_app = gui_app 
        self.driver = None             # Selenium WebDriver nesnesi
        self.is_running = False        # Gönderim sürecinin aktif olup olmadığını tutar
        self.df_data = None            # Excel'den okunan Pandas DataFrame
        self.failed_log = []           # Başarısız gönderim kayıtları
        self.sent_log = []             # Başarılı gönderim kayıtları
        self.total_recipients = 0      # Toplam alıcı sayısı
        self.current_run_dir = None    # Mevcut çalıştırma için oluşturulan rapor klasörü

    # --- Yardımcı Fonksiyonlar ---

    def _get_delays(self, mode):
        """Seçilen hıza göre bekleme sürelerini (saniye) döndürür."""
        # Dönüş sırası: WA_OPEN_DELAY, SEND_DELAY, SEARCH_SUCCESS_DELAY, SEARCH_FAIL_DELAY
        if mode == "SAFE":
            return 8, 15, 4, 7
        elif mode == "FAST":
            return 5, 8, 2, 4
        elif mode == "TURBO":
            return 3, 5, 1, 2
        return 5, 8, 2, 4

    def _clean_phone_number(self, phone):
        """Telefon numarasını sadece rakamları içerecek ve '90' ile başlayacak şekilde temizler."""
        # Önceki kodda zaten string'e dönüştürüyordu ancak daha güvenli hale getirildi.
        phone = re.sub(r'[^0-9]', '', str(phone))
        if not phone.startswith('90'):
            return '90' + phone
        return phone

    def _log_to_gui(self, message, tag="info"):
        """Mesajı GUI'nin terminal alanına iletir."""
        self.gui_app._log_to_terminal(message, tag)
    
    # --- Veri Yükleme ---

    def load_data(self, file_path):
        """Seçilen dosyayı Pandas ile okur, sütunları kontrol eder."""
        try:
            # Excel dosyasını okurken 'phone' sütununu String olarak okumasını zorlar (önemli düzeltme)
            self.df_data = pd.read_excel(file_path, dtype={'phone': str})
            
            # Zorunlu sütun kontrolü
            required_cols = ['phone']
            missing_cols = [col for col in required_cols if col not in self.df_data.columns]
            
            if missing_cols:
                raise ValueError(f"Excel dosyasında zorunlu sütun(lar) eksik: {', '.join(missing_cols)}")
            
            # DataFrame'deki NaN (boş) değerleri boş dize ile doldurur. (Kişiselleştirme için önemli)
            self.df_data = self.df_data.fillna('') 
            
            self.total_recipients = len(self.df_data)
            return True, None
            
        except Exception as e:
            self.df_data = None
            self.total_recipients = 0
            return False, f"Dosya okunamadı veya biçim hatası: {e}"

    # --- Otomasyon Çekirdeği ---

    def _init_browser(self):
        """Chrome tarayıcısını başlatır ve WhatsApp Web oturumunu kontrol eder."""
        options = webdriver.ChromeOptions()
        options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}") 
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        try:
            self._log_to_gui("Chrome sürücüsü indiriliyor ve başlatılıyor...", "info")
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.get("https://web.whatsapp.com")

            # Başarılı giriş kontrolü için arama kutusunun yüklenmesini bekler.
            SEARCH_INPUT_XPATH = '//*[@id="side"]//div[@role="textbox"]'
            self._log_to_gui("Tarayıcı başlatıldı. Oturum kontrol ediliyor (60sn bekleniyor)...")
            
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.XPATH, SEARCH_INPUT_XPATH)) 
            )
            
            self._log_to_gui("WhatsApp Web oturumu başarıyla açıldı. Gönderim başlıyor...", "success")
            return True

        except Exception as e:
            if self.driver:
                self.driver.quit()
                self.driver = None
            self._log_to_gui(f"WhatsApp Web oturumu açılamadı. Hata: {e}", "error")
            return False

    def _check_number_invalid(self):
        """Geçersiz numara pop-up'ını kontrol eder."""
        try:
            invalid_popup_xpath = '//*[contains(text(), "telefon numarası geçersiz")] | //*[contains(text(), "WhatsApp kullanıcısı değil")]'
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, invalid_popup_xpath))
            )
            return True
        except:
            return False

    def _send_message(self, index, row, message_template, delays):
        """Belirli bir kişiye mesajı gönderir (Çift yazma sorununu temizleme ile çözer)."""
        
        WA_OPEN_DELAY, SEND_DELAY, SEARCH_SUCCESS_DELAY, SEARCH_FAIL_DELAY = delays
        
        phone_raw = row['phone']
        
        # Düzeltme: 'name' sütunu boş olsa bile (load_data'da doldurulduğu için) güvenle alınır.
        # Ayrıca, strip() ile gereksiz boşlukları temizleriz.
        name = str(row['name']).strip() if 'name' in row and row['name'] else 'kişi'
        
        excel_message = row['message'] if 'message' in row and row['message'] else None
        
        # Mesaj içeriği seçimi ve kişiselleştirme
        msg_template_to_use = excel_message if excel_message else message_template
        # Kişiselleştirme sırasında 'name' değişkeninin boş dize olması sorun teşkil etmez, sadece {name} kalkar.
        message_content = str(msg_template_to_use).replace('{name}', name).strip()

        phone_clean = self._clean_phone_number(phone_raw)
        self.gui_app._update_list_status(index, "Gönderiliyor...", "sending")
        
        link = f"https://web.whatsapp.com/send?phone={phone_clean}"
        
        try:
            self.driver.get(link)
            time.sleep(WA_OPEN_DELAY) 
            
            try:
                # Mesaj kutusunu bulma
                message_box_xpath = '//*[@id="main"]//footer//*[@contenteditable="true"]'
                message_box = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, message_box_xpath))
                )
                
                # 1. Kutu içeriğini temizle (Çift yazma sorununu çözer)
                message_box.click() 
                message_box.send_keys(Keys.CONTROL, 'a') 
                message_box.send_keys(Keys.BACKSPACE) 
                
                # 2. Mesajı satır satır yaz ve gönder
                lines = message_content.split('\n')
                for i, line in enumerate(lines):
                    message_box.send_keys(line)
                    if i < len(lines) - 1:
                        message_box.send_keys(Keys.SHIFT, Keys.ENTER)
                
                message_box.send_keys(Keys.ENTER)
                
                self._log_success(index, phone_raw, name, message_content)
                time.sleep(SEND_DELAY) 
            
            except Exception as e:
                if self._check_number_invalid():
                    self._log_fail(index, phone_raw, name, "Numara WhatsApp kullanıcısı değil veya geçersiz.", None)
                    time.sleep(SEARCH_FAIL_DELAY)
                else:
                    self._log_fail(index, phone_raw, name, f"Mesaj kutusu bulunamadı/Gönderim hatası: {e}", None)
                    time.sleep(SEARCH_FAIL_DELAY)
                    
        except Exception as e:
            self._log_fail(index, phone_raw, name, f"Genel Gönderim Hatası: {e}", e)
            time.sleep(SEARCH_FAIL_DELAY)

    # --- Ana Çalıştırma Döngüsü ---

    def start_broadcast(self, message_template, speed_mode):
        """Ana gönderim döngüsünü çalıştırır."""
        self.is_running = True
        self.failed_log = []
        self.sent_log = []
        self.gui_app._reset_list_colors() # GUI'deki listeyi sıfırla

        # Tarayıcıyı başlat (QR kod kontrolü burada yapılır)
        if not self._init_browser():
            self.gui_app.cancel_broadcast(hard_stop=True)
            return 

        delays = self._get_delays(speed_mode)
        
        try:
            for index, (_, row) in enumerate(self.df_data.iterrows()):
                if not self.is_running:
                    break # Kullanıcı iptal ettiyse döngüyü kır
                
                self._send_message(index, row, message_template, delays)
                
        except Exception as e:
            self._log_to_gui(f"Beklenmedik bir hata oluştu: {e}", "error")
            
        finally:
            if self.driver:
                self.driver.quit()
            self.gui_app._finish_broadcast(cancelled=(not self.is_running))

    def cancel_broadcast(self):
        """Gönderimi durdurur ve tarayıcıyı kapatır."""
        self.is_running = False
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

    # --- Loglama ve Raporlama ---

    def _log_success(self, index, phone, name, message):
        """Başarılı gönderimi kaydeder ve GUI'ye bilgi gönderir."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sent_log.append({
            'timestamp': now, 
            'phone': phone, 
            'name': name, 
            'message': message, 
            'status': 'SENT'
        })
        self.gui_app.update_progress()
        self.gui_app._update_list_status(index, "BAŞARILI", "sent")
        self._log_to_gui(f"BAŞARILI: {name} ({phone}) kişisine mesaj gönderildi.", "success")

    def _log_fail(self, index, phone, name, reason, exception):
        """Başarısız gönderimi kaydeder ve GUI'ye bilgi gönderir."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        error_detail = str(exception) if exception else reason
        self.failed_log.append({
            'timestamp': now, 
            'phone': phone, 
            'name': name, 
            'reason': reason, 
            'error_detail': error_detail
        })
        self.gui_app.update_progress()
        self.gui_app._update_list_status(index, f"BAŞARISIZ ({reason})", "failed")
        self._log_to_gui(f"BAŞARISIZ: {name} ({phone}). Sebep: {reason}", "error")

    def generate_reports(self):
        """Gönderim sonuçlarını rapor dosyalarına (Excel, CSV) kaydeder."""
        if not self.sent_log and not self.failed_log:
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.current_run_dir = os.path.join(REPORT_BASE_DIR, f"run_{timestamp}")
        os.makedirs(self.current_run_dir, exist_ok=True)
        
        # ... (Raporlama Mantığı aynı kalır) ...
        final_df = self.df_data.copy()
        
        sent_df = pd.DataFrame(self.sent_log)
        failed_df = pd.DataFrame(self.failed_log)
        
        final_df['status'] = 'PENDING'
        final_df['log_time'] = ''
        final_df['reason'] = ''

        for index, row in final_df.iterrows():
            phone_raw = row['phone']
            
            match_sent = sent_df[sent_df['phone'] == phone_raw]
            if not match_sent.empty:
                final_df.loc[index, 'status'] = 'SENT'
                final_df.loc[index, 'log_time'] = match_sent.iloc[0]['timestamp']
                continue

            match_failed = failed_df[failed_df['phone'] == phone_raw]
            if not match_failed.empty:
                final_df.loc[index, 'status'] = 'FAILED'
                final_df.loc[index, 'log_time'] = match_failed.iloc[0]['timestamp']
                final_df.loc[index, 'reason'] = match_failed.iloc[0]['reason']

        report_file_path = os.path.join(self.current_run_dir, "results.xlsx")
        
        try:
            writer = pd.ExcelWriter(report_file_path, engine='xlsxwriter')
            final_df.to_excel(writer, sheet_name='Rapor', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Rapor']

            sent_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            failed_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            worksheet.conditional_format('A1:Z' + str(len(final_df) + 1), 
                                        {'type': 'text', 'criteria': 'containing', 'value': 'SENT', 'format': sent_format})

            worksheet.conditional_format('A1:Z' + str(len(final_df) + 1), 
                                        {'type': 'text', 'criteria': 'containing', 'value': 'FAILED', 'format': failed_format})
            
            writer.close()

            sent_csv_path = os.path.join(self.current_run_dir, "sent_log.csv")
            if not sent_df.empty:
                sent_df.to_csv(sent_csv_path, index=False, encoding='utf-8')

            failed_csv_path = os.path.join(self.current_run_dir, "failed_log.csv")
            if not failed_df.empty:
                failed_df.to_csv(failed_csv_path, index=False, encoding='utf-8')

            return True, self.current_run_dir
        
        except Exception as e:
            self._log_to_gui(f"Rapor oluşturulurken hata oluştu: {e}", "error")
            return False, str(e)
