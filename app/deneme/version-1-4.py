import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import pandas as pd
import time
import os
import re
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from threading import Thread

# --- Global Yapılandırma ---
VERSION = "Sürüm 1.4 (Gönderim & QR Kontrolü Düzeltildi)"
# Bu, Chrome'u bir kullanıcı profiliyle başlatmak için kullanılacak. 
CHROME_PROFILE_PATH = os.path.join(os.path.expanduser("~"), "whatsapp_profile") 

# Raporların kaydedileceği ana klasör
REPORT_BASE_DIR = os.path.join(os.path.expanduser("~"), "Documents", "WhatsAppBroadcastRuns")
os.makedirs(REPORT_BASE_DIR, exist_ok=True)


class WhatsAppBroadcaster(ctk.CTk):
    """
    WhatsApp Web üzerinden toplu ve kişiselleştirilmiş mesaj gönderen masaüstü uygulaması.
    """
    def __init__(self):
        super().__init__()
        
        # --- Arayüz Ayarları ---
        self.title("WhatsApp Toplu Gönderim Aracı") # Uygulama adı güncellendi
        self.geometry("1100x900") # Terminale yer açmak için yükseklik artırıldı
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1) 
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1) # Terminal alanı için yeni satır
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- Durum Değişkenleri ---
        self.is_running = False
        self.driver = None
        self.df_data = None
        self.failed_log = []
        self.sent_log = []
        self.total_recipients = 0
        self.current_run_dir = None
        self.current_thread = None
        
        # --- Arayüz Bileşenlerini Oluşturma ---
        self._create_sidebar()
        self._create_main_frames()
        self._create_list_frame() 
        self._create_controls_frame()
        self._create_progress_frame()
        self._create_log_frame() # Yeni: Log terminali
        
        # Kişi listesi için arayüz öğeleri (gönderim öncesi ve sırası)
        self.recipient_widgets = {}

    def _create_sidebar(self):
        """Sol taraftaki navigasyon ve yapılandırma çubuğunu oluşturur."""
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=5, sticky="nsew") # rowspan 5 olarak güncellendi
        # Alt bileşenlerin yerleşimi için row yapılandırması güncellendi
        self.sidebar_frame.grid_rowconfigure(5, weight=1) 

        # Başlık
        ctk.CTkLabel(self.sidebar_frame, text="Kodlama Desteği", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Versiyon Bilgisi
        ctk.CTkLabel(self.sidebar_frame, text=VERSION, font=ctk.CTkFont(size=12)).grid(row=1, column=0, padx=20, pady=(0, 5))
        

        # Hız Modu Seçimi
        ctk.CTkLabel(self.sidebar_frame, text="Hız Modu:", anchor="w").grid(row=3, column=0, padx=20, pady=(10, 0))
        self.speed_mode = ctk.StringVar(value="FAST")
        self.mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["SAFE", "FAST", "TURBO"], command=self._show_delay_info)
        self.mode_optionemenu.grid(row=4, column=0, padx=20, pady=(0, 5))
        
        self.delay_info_label = ctk.CTkLabel(self.sidebar_frame, text="Gecikme: 5-8 sn", font=ctk.CTkFont(size=10), text_color="gray")
        self.delay_info_label.grid(row=5, column=0, padx=20, pady=(0, 10), sticky="n")
        
        # Tema Seçimi
        ctk.CTkLabel(self.sidebar_frame, text="Tema:", anchor="w").grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=7, column=0, padx=20, pady=(0, 20))
        
        self._show_delay_info(self.speed_mode.get()) # Başlangıç bilgisini göster

    def _show_delay_info(self, mode):
        """Hız moduna göre gecikme bilgisini gösterir."""
        if mode == "SAFE":
            info = "Gecikme: 8-15 sn (Güvenli)"
        elif mode == "FAST":
            info = "Gecikme: 5-8 sn (Dengeli)"
        elif mode == "TURBO":
            info = "Gecikme: 3-5 sn (Riskli)"
        self.delay_info_label.configure(text=info)
        self.speed_mode.set(mode)

    def _create_main_frames(self):
        """Ana içerik çerçevelerini (Dosya Yolu, Mesaj Şablonu) oluşturur."""
        
        # --- Dosya Yolu Çerçevesi ---
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=0, column=1, padx=(20, 10), pady=(20, 10), sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.file_frame, text="1. Excel Dosyası Seç (phone, name, message olmalı):", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.file_path_entry = ctk.CTkEntry(self.file_frame, placeholder_text="Dosya yolu...")
        self.file_path_entry.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        
        self.file_button = ctk.CTkButton(self.file_frame, text="Dosya Seç", command=self.select_file)
        self.file_button.grid(row=1, column=1, padx=10, pady=(0, 10))

        # --- Mesaj Şablonu Çerçevesi ---
        self.message_frame = ctk.CTkFrame(self)
        self.message_frame.grid(row=1, column=1, padx=(20, 10), pady=(10, 10), sticky="nsew")
        self.message_frame.grid_columnconfigure(0, weight=1)
        self.message_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.message_frame, text="2. Mesaj Şablonu (Kişiselleştirme için {name} kullanın):", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.message_textbox = ctk.CTkTextbox(self.message_frame, height=150)
        self.message_textbox.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.message_textbox.insert("0.0", "Merhaba {name},\n\nBu, toplu mesaj gönderim aracımızın bir testidir. İyi günler!")

    def _create_list_frame(self):
        """Excel'den çekilen kişileri gösteren çerçeveyi oluşturur."""
        self.list_container_frame = ctk.CTkFrame(self)
        self.list_container_frame.grid(row=0, column=2, rowspan=2, padx=(10, 20), pady=(20, 10), sticky="nsew")
        self.list_container_frame.grid_columnconfigure(0, weight=1)
        self.list_container_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.list_container_frame, text="3. Kişi Listesi ve Anlık Durum:", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.list_scroll_frame = ctk.CTkScrollableFrame(self.list_container_frame, label_text="Yüklenen Kişiler (0 Kişi)")
        self.list_scroll_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.list_scroll_frame.grid_columnconfigure(0, weight=1)

    def _create_controls_frame(self):
        """Başlatma/İptal etme düğmelerini içeren çerçeveyi oluşturur."""
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=2, column=1, columnspan=2, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.controls_frame.grid_columnconfigure((0, 1), weight=1)

        self.start_button = ctk.CTkButton(self.controls_frame, text="Gönderimi Başlat", command=self.start_broadcast_thread, height=40, fg_color="green", hover_color="#006400")
        self.start_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.cancel_button = ctk.CTkButton(self.controls_frame, text="Gönderimi İptal Et", command=self.cancel_broadcast, height=40, fg_color="red", hover_color="#8B0000", state="disabled")
        self.cancel_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
    
    def _create_progress_frame(self):
        """İlerleme çubuğunu ve sayaçları içeren çerçeveyi oluşturur."""
        self.progress_frame = ctk.CTkFrame(self)
        self.progress_frame.grid(row=3, column=1, columnspan=2, padx=(20, 20), pady=(10, 20), sticky="ew")
        self.progress_frame.grid_columnconfigure((0, 1), weight=1)
        
        # İlerleme Çubuğu
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, orientation="horizontal")
        self.progress_bar.grid(row=0, column=0, columnspan=2, padx=10, pady=(15, 10), sticky="ew")
        self.progress_bar.set(0)

        # Sayaçlar
        self.counter_label = ctk.CTkLabel(self.progress_frame, text="Hazır | Toplam Kişi: 0", anchor="w")
        self.counter_label.grid(row=1, column=0, padx=10, pady=(0, 15), sticky="w")
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="Durum: Bekleniyor...", anchor="e")
        self.status_label.grid(row=1, column=1, padx=10, pady=(0, 15), sticky="e")

    def _create_log_frame(self):
        """Uygulama mesajlarını ve durumunu gösteren terminal alanını oluşturur."""
        self.log_frame = ctk.CTkFrame(self)
        # Yeni satır 4, tüm sütunları kapsar.
        self.log_frame.grid(row=4, column=1, columnspan=2, padx=(20, 20), pady=(10, 20), sticky="nsew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(self.log_frame, text="Terminal Çıktısı:", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        self.terminal_textbox = ctk.CTkTextbox(self.log_frame, height=150, activate_scrollbars=True)
        self.terminal_textbox.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.terminal_textbox.configure(state="disabled") # Terminal başlangıçta sadece okuma modunda
        self._log_to_terminal(f"[{self.title()}] Uygulama başlatıldı. Lütfen Excel dosyasını seçin.", "info")

    def _log_to_terminal(self, message, tag="info"):
        """Mesajı terminal alanına ekler."""
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        
        # Renkli etiketler ekleyebiliriz (opsiyonel, şu an CTkTextbox ile sadece metin rengini değiştirebiliriz)
        if tag == "error":
            prefix = "[HATA]"
            # Konsola yazmayı kaldırdık, sadece arayüze logluyoruz
        elif tag == "success":
            prefix = "[BAŞARILI]"
        else:
            prefix = "[BİLGİ]"
            
        full_message = f"{timestamp} {prefix} {message}\n"
        
        # Terminale yazmak için durumu geçici olarak etkinleştir
        self.terminal_textbox.configure(state="normal")
        self.terminal_textbox.insert(ctk.END, full_message)
        
        # Otomatik kaydırma
        self.terminal_textbox.see(ctk.END)
        self.terminal_textbox.configure(state="disabled")

    # --- Arayüz İşleyicileri ---

    def change_appearance_mode_event(self, new_appearance_mode: str):
        """Tema değişimini yönetir."""
        ctk.set_appearance_mode(new_appearance_mode)

    def select_file(self):
        """Kullanıcının Excel dosyasını seçmesini sağlar ve veriyi önizler."""
        file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyaları", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        )
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)
            self._preview_data(file_path)

    def _preview_data(self, file_path):
        """Seçilen dosyayı okur, kişi sayısını günceller ve listeyi doldurur."""
        try:
            self.df_data = pd.read_excel(file_path)
            
            # Gerekli sütunların kontrolü
            required_cols = ['phone']
            missing_cols = [col for col in required_cols if col not in self.df_data.columns]
            
            if missing_cols:
                messagebox.showerror("Hata", f"Excel dosyasında zorunlu sütun(lar) eksik: {', '.join(missing_cols)}. Lütfen 'phone' sütununun bulunduğundan emin olun.")
                self.df_data = None
                self.total_recipients = 0
                self._log_to_terminal(f"Zorunlu sütun(lar) eksik: {', '.join(missing_cols)}", "error")
            else:
                self.total_recipients = len(self.df_data)
                self.counter_label.configure(text=f"Hazır | Toplam Kişi: {self.total_recipients}")
                self.status_label.configure(text="Durum: Veri Yüklendi.")
                self._populate_list()
                self._log_to_terminal(f"Excel verisi başarıyla yüklendi. Toplam {self.total_recipients} kişi.", "info")
                
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya okunamadı veya biçim hatası: {e}")
            self.df_data = None
            self.total_recipients = 0
            self.counter_label.configure(text="Hazır | Toplam Kişi: 0")
            self._clear_list()
            self._log_to_terminal(f"Dosya okuma hatası: {e}", "error")

    def _clear_list(self):
        """Kişi listesi görünümünü temizler."""
        for widget in self.list_scroll_frame.winfo_children():
            widget.destroy()
        self.recipient_widgets = {}
        self.list_scroll_frame.configure(label_text="Yüklenen Kişiler (0 Kişi)")

    def _populate_list(self):
        """DataFrame'den kişileri okur ve arayüze listeler."""
        self._clear_list()
        
        self.list_scroll_frame.configure(label_text=f"Yüklenen Kişiler ({self.total_recipients} Kişi)")
        
        # DataFrame satır indekslerini kaybetmemek için iterrows kullanılır.
        for index, row in self.df_data.iterrows():
            
            phone_raw = row['phone']
            name = row['name'] if 'name' in row and pd.notna(row['name']) else '(İsimsiz Kişi)'
            status = row['status'] if 'status' in row and pd.notna(row['status']) else 'Bekliyor'
            
            # Bu listenin düzgün sıralanması için, 0'dan başlayan basit bir sayaç kullanalım:
            list_row_index = len(self.recipient_widgets) 
            
            person_frame = ctk.CTkFrame(self.list_scroll_frame, border_width=1, corner_radius=5)
            person_frame.grid(row=list_row_index, column=0, padx=5, pady=3, sticky="ew")
            person_frame.grid_columnconfigure(0, weight=1)
            
            # İsim ve Telefon
            name_label = ctk.CTkLabel(person_frame, text=f"{name} ({phone_raw})", anchor="w", font=ctk.CTkFont(weight="bold"))
            name_label.grid(row=0, column=0, padx=10, pady=(5, 0), sticky="w")
            
            # Durum göstergesi (gönderim sırasında güncellenecek)
            status_label = ctk.CTkLabel(person_frame, text=f"Durum: {status}", anchor="e", text_color="gray")
            status_label.grid(row=0, column=1, padx=10, pady=(5, 0), sticky="e")
            
            # Bu widget'ları list_row_index (0'dan başlayan sıra numarası) ile saklıyoruz.
            self.recipient_widgets[list_row_index] = {
                'frame': person_frame,
                'status_label': status_label
            }


    def _update_list_status(self, index, status_text, color_key="pending"):
        """Kişi listesindeki bir öğenin durumunu ve rengini günceller."""
        # Gönderim sırasında kullanılan index (0'dan başlayan sıra numarası) kullanılır.
        if index in self.recipient_widgets:
            status_widget = self.recipient_widgets[index]['status_label']
            frame_widget = self.recipient_widgets[index]['frame']
            
            # Durum rengini ayarla
            if color_key == "sent":
                color = "green"
            elif color_key == "failed":
                color = "red"
            elif color_key == "sending":
                color = "orange"
            else: # pending
                color = "gray"
            
            status_widget.configure(text=f"Durum: {status_text}", text_color=color)
            
            # Gönderilen veya başarısız olan çerçeveyi hafifçe renklendir
            if color_key in ["sent", "failed"]:
                if color_key == "sent":
                    frame_widget.configure(fg_color="#3A533E") # Koyu yeşil tonu
                elif color_key == "failed":
                    frame_widget.configure(fg_color="#533A3A") # Koyu kırmızı tonu
            else:
                # Varsayılan arka plan rengine dön
                frame_widget.configure(fg_color=self.list_scroll_frame.cget("fg_color"))


    # --- Mesaj Gönderim Mantığı (Önceki Koddan Devam Eden Fonksiyonlar) ---

    def _get_delays(self):
        """Seçilen hıza göre bekleme sürelerini (saniye) döndürür."""
        mode = self.speed_mode.get()
        if mode == "SAFE":
            return 8, 15, 4, 7
        elif mode == "FAST":
            return 5, 8, 2, 4
        elif mode == "TURBO":
            return 3, 5, 1, 2
        return 5, 8, 2, 4

    def _clean_phone_number(self, phone):
        """Telefon numarasını WhatsApp formatına uygun olarak temizler."""
        phone = re.sub(r'[^0-9]', '', str(phone))
        if not phone.startswith('90'):
            return '90' + phone
        return phone

    def _init_browser(self):
        """Chrome tarayıcısını (Selenium) başlatır ve WhatsApp Web'e gider."""
        
        options = webdriver.ChromeOptions()
        options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}") 
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        
        try:
            self._log_to_terminal("Chrome sürücüsü indiriliyor ve başlatılıyor...", "info")
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.get("https://web.whatsapp.com")

            # BAŞARILI GİRİŞ KONTROLÜ: Yan çubuktaki arama kutusunu bekleriz.
            # Bu öğe sadece başarılı giriş yapıldıktan sonra görünür.
            # Oturum zaten açıksa, bu bekleme çok hızlı tamamlanır.
            SEARCH_INPUT_XPATH = '//*[@id="side"]//div[@role="textbox"]'
            
            self.status_label.configure(text="Durum: QR Kodu Taranıyor... (Lütfen Tarayıcıya Bakın)")
            self._log_to_terminal("Tarayıcı başlatıldı. Oturum kontrol ediliyor (60sn bekleniyor)...")
            
            WebDriverWait(self.driver, 60).until(
                EC.presence_of_element_located((By.XPATH, SEARCH_INPUT_XPATH)) 
            )
            
            self.status_label.configure(text="Durum: WhatsApp Web Hazır.")
            self._log_to_terminal("WhatsApp Web oturumu başarıyla açıldı. Gönderim başlıyor...")
            return True

        except Exception as e:
            self._log_error(f"Tarayıcı başlatılırken veya WhatsApp Web yüklenirken hata: {e}")
            self._log_to_terminal("WhatsApp Web oturumu açılamadı. Süre aşımı (QR kod taranmadı) veya tarayıcı hatası.", "error")
            return False

    def _send_message(self, index, row, message_template, delays):
        """Belirli bir kişiye mesajı gönderir."""
        
        WA_OPEN_DELAY, SEND_DELAY, SEARCH_SUCCESS_DELAY, SEARCH_FAIL_DELAY = delays
        
        phone_raw = row['phone']
        name = row['name'] if 'name' in row and pd.notna(row['name']) else 'kişi'
        
        excel_message = row['message'] if 'message' in row and pd.notna(row['message']) else None
        
        if excel_message:
            message_content = str(excel_message).replace('{name}', name).strip()
        else:
            message_content = message_template.replace('{name}', name).strip()

        phone_clean = self._clean_phone_number(phone_raw)
        
        self._update_list_status(index, "Gönderiliyor...", "sending")
        
        # WhatsApp API linki
        link = f"https://web.whatsapp.com/send?phone={phone_clean}&text={message_content}"
        
        try:
            self.driver.get(link)
            time.sleep(WA_OPEN_DELAY)
            
            try:
                # MESAJ KUTUSU GÜNCEL XPATH DÜZELTMESİ: contenteditable="true" olan mesaj kutusu
                message_box_xpath = '//*[@id="main"]//footer//*[@contenteditable="true"]'
                
                message_box = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, message_box_xpath))
                )
                
                # Mesajı yazma ve gönderme
                message_box.send_keys(message_content)
                message_box.send_keys(Keys.ENTER)
                
                self._log_success(index, phone_raw, name, message_content)
                time.sleep(SEND_DELAY)

            except Exception as e:
                # Eğer sohbet bulunamazsa 'Numara geçersiz' gibi bir hata çıkar
                if self._check_number_invalid():
                    self._log_fail(index, phone_raw, name, "Numara WhatsApp kullanıcısı değil veya geçersiz.", None)
                    time.sleep(SEARCH_FAIL_DELAY)
                else:
                    self._log_fail(index, phone_raw, name, f"Mesaj kutusu bulunamadı/Gönderim hatası: {e}", None)
                    time.sleep(SEARCH_FAIL_DELAY)
                    
        except Exception as e:
            self._log_fail(index, phone_raw, name, f"Genel Gönderim Hatası: {e}", e)
            time.sleep(SEARCH_FAIL_DELAY)

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

    def _log_success(self, index, phone, name, message):
        """Başarılı gönderimi kaydeder ve arayüzü günceller."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.sent_log.append({
            'timestamp': now, 
            'phone': phone, 
            'name': name, 
            'message': message, 
            'status': 'SENT'
        })
        self.update_progress()
        self._update_list_status(index, "BAŞARILI", "sent")
        self._log_to_terminal(f"BAŞARILI: {name} ({phone}) kişisine mesaj gönderildi.", "success")

    def _log_fail(self, index, phone, name, reason, exception):
        """Başarısız gönderimi kaydeder ve arayüzü günceller."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        error_detail = str(exception) if exception else reason
        self.failed_log.append({
            'timestamp': now, 
            'phone': phone, 
            'name': name, 
            'reason': reason, 
            'error_detail': error_detail
        })
        self.update_progress()
        self._update_list_status(index, f"BAŞARISIZ ({reason})", "failed")
        self.status_label.configure(text=f"Durum: {name} kişisine gönderim BAŞARISIZ.")
        self._log_to_terminal(f"BAŞARISIZ: {name} ({phone}). Sebep: {reason}", "error")

    def _log_error(self, message):
        """Uygulama düzeyinde hataları kaydeder."""
        self._log_to_terminal(f"Uygulama Hatası: {message}", "error")
        messagebox.showerror("Uygulama Hatası", message)
        self.status_label.configure(text=f"Durum: HATA! İşlem durduruldu. ({message})")

    def update_progress(self):
        """Arayüzdeki ilerleme çubuğunu ve sayaçları günceller."""
        sent_count = len(self.sent_log)
        failed_count = len(self.failed_log)
        processed_count = sent_count + failed_count
        
        progress_value = processed_count / self.total_recipients if self.total_recipients > 0 else 0
        
        self.progress_bar.set(progress_value)
        self.counter_label.configure(text=f"İşlendi: {processed_count}/{self.total_recipients} | Başarılı: {sent_count} | Başarısız: {failed_count}")
        
        if processed_count < self.total_recipients:
            self.status_label.configure(text=f"Durum: Gönderiliyor... (Kişi {processed_count + 1}/{self.total_recipients})")

    # --- Başlatma/Kontrol Fonksiyonları ---

    def start_broadcast_thread(self):
        """Gönderim işlemini arayüzü kilitlememesi için ayrı bir iş parçacığında başlatır."""
        if self.is_running:
            messagebox.showwarning("Uyarı", "Gönderim zaten devam ediyor.")
            return

        file_path = self.file_path_entry.get()
        message_template = self.message_textbox.get("0.0", "end-1c")

        # HATA DÜZELTMESİ: ValueError'ı önlemek için 'not self.df_data' yerine 'self.df_data is None' kullanıldı.
        if not file_path or self.df_data is None or self.total_recipients == 0: 
            messagebox.showerror("Hata", "Lütfen geçerli bir Excel dosyası seçin ve veriyi yükleyin.")
            self._log_to_terminal("HATA: Gönderim başlatılamadı. Dosya seçimi veya veri yüklemesi eksik.", "error")
            return
        
        if not message_template.strip() and ('message' not in self.df_data.columns or self.df_data['message'].isna().all()):
             messagebox.showerror("Hata", "Lütfen bir mesaj şablonu girin veya Excel dosyanızdaki 'message' sütununu doldurun.")
             self._log_to_terminal("HATA: Gönderim başlatılamadı. Mesaj içeriği eksik.", "error")
             return
            
        self.is_running = True
        self.failed_log = []
        self.sent_log = []
        self.start_button.configure(state="disabled", text="Gönderim Başladı")
        self.cancel_button.configure(state="normal")
        self.progress_bar.set(0)
        self.update_progress()
        self._reset_list_colors() # Liste durumunu sıfırla
        self._log_to_terminal(f"Gönderim işlemi başlatılıyor. Hız Modu: {self.speed_mode.get()}", "info")

        # Ana gönderim işlevini başlat
        self.current_thread = Thread(target=self.start_broadcast, args=(message_template,))
        self.current_thread.start()

    def _reset_list_colors(self):
        """Gönderim başlamadan listeyi 'Bekliyor' durumuna getirir."""
        for index, widgets in self.recipient_widgets.items():
            widgets['status_label'].configure(text="Durum: Bekliyor", text_color="gray")
            # Çerçeve rengini varsayılana döndür
            widgets['frame'].configure(fg_color=self.list_scroll_frame.cget("fg_color"))


    def start_broadcast(self, message_template):
        """Ana gönderim döngüsünü ve Selenium kontrolünü yönetir."""
        
        if not self._init_browser():
            self.cancel_broadcast(hard_stop=True)
            return

        delays = self._get_delays()
        
        try:
            # enumerate() ile 0'dan başlayan bir index değeri alıyoruz
            for index, (df_index, row) in enumerate(self.df_data.iterrows()):
                if not self.is_running:
                    break # İptal düğmesine basıldı
                
                # index değeri arayüzdeki widget'ı bulmak için kullanılır.
                self._send_message(index, row, message_template, delays)
                
        except Exception as e:
            self._log_error(f"Beklenmedik bir hata oluştu: {e}")
            
        finally:
            if self.driver:
                self.driver.quit()
            self._finish_broadcast()

    def cancel_broadcast(self, hard_stop=False):
        """Gönderimi iptal eder."""
        if self.is_running:
            self.is_running = False
            self.status_label.configure(text="Durum: İptal Ediliyor...")
            self._log_to_terminal("Kullanıcı isteği üzerine iptal ediliyor. Tarayıcı kapatılıyor...", "info")
            
            if self.driver:
                 try:
                    self.driver.quit()
                 except:
                    pass
                 self.driver = None

            if not hard_stop and self.current_thread and self.current_thread.is_alive():
                self.current_thread.join(timeout=5)
            
            self._finish_broadcast(cancelled=True)
        else:
             if hard_stop:
                self._finish_broadcast(cancelled=True)
             else:
                messagebox.showinfo("Bilgi", "Gönderim zaten durdurulmuş.")


    def _finish_broadcast(self, cancelled=False):
        """Gönderim tamamlandığında veya iptal edildiğinde raporları oluşturur ve arayüzü sıfırlar."""
        self.is_running = False
        self.start_button.configure(state="normal", text="Gönderimi Başlat")
        self.cancel_button.configure(state="disabled")

        if cancelled:
            self.status_label.configure(text="Durum: GÖNDERİM İPTAL EDİLDİ.")
            self._log_to_terminal("Gönderim kullanıcı tarafından İPTAL EDİLDİ.", "info")
        else:
            self.status_label.configure(text="Durum: GÖNDERİM TAMAMLANDI.")
            self._log_to_terminal("Gönderim başarıyla TAMAMLANDI.", "success")

        if self.sent_log or self.failed_log:
            self._generate_reports()
            self._log_to_terminal(f"Raporlar oluşturuldu ve '{self.current_run_dir}' klasörüne kaydedildi.", "info")
            messagebox.showinfo("Tamamlandı", f"Gönderim tamamlandı.\nRaporlar şu klasöre kaydedildi: \n{self.current_run_dir}")
        elif not cancelled:
             messagebox.showinfo("Bilgi", "İşlenecek veri yok veya işlem başlatılamadı.")
             self._log_to_terminal("İşlenecek veri kalmadı, işlem sonlandı.", "info")

    # --- Raporlama Mantığı ---

    def _generate_reports(self):
        """Gönderim sonuçlarını rapor dosyalarına kaydeder."""
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.current_run_dir = os.path.join(REPORT_BASE_DIR, f"run_{timestamp}")
        os.makedirs(self.current_run_dir, exist_ok=True)
        
        # 1. Genel Rapor (results.xlsx - Renkli Rapor)
        final_df = self.df_data.copy()
        
        # Logları DataFramelere dönüştür
        sent_df = pd.DataFrame(self.sent_log)
        failed_df = pd.DataFrame(self.failed_log)
        
        # Ana tabloyu güncelle
        final_df['status'] = 'PENDING'
        final_df['log_time'] = ''
        final_df['reason'] = ''

        for index, row in final_df.iterrows():
            phone_raw = row['phone']
            
            # Başarılı logları eşleştir
            match_sent = sent_df[sent_df['phone'] == phone_raw]
            if not match_sent.empty:
                final_df.loc[index, 'status'] = 'SENT'
                final_df.loc[index, 'log_time'] = match_sent.iloc[0]['timestamp']
                continue

            # Başarısız logları eşleştir
            match_failed = failed_df[failed_df['phone'] == phone_raw]
            if not match_failed.empty:
                final_df.loc[index, 'status'] = 'FAILED'
                final_df.loc[index, 'log_time'] = match_failed.iloc[0]['timestamp']
                final_df.loc[index, 'reason'] = match_failed.iloc[0]['reason']

        report_file_path = os.path.join(self.current_run_dir, "results.xlsx")
        
        try:
            # Excel'e kaydet ve renklendirme uygula
            writer = pd.ExcelWriter(report_file_path, engine='xlsxwriter')
            final_df.to_excel(writer, sheet_name='Rapor', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Rapor']

            # Renk formatları
            sent_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            failed_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            # Durum sütununa koşullu biçimlendirme
            worksheet.conditional_format('A1:Z' + str(len(final_df) + 1), 
                                        {'type': 'text',
                                         'criteria': 'containing',
                                         'value': 'SENT',
                                         'format': sent_format})

            worksheet.conditional_format('A1:Z' + str(len(final_df) + 1), 
                                        {'type': 'text',
                                         'criteria': 'containing',
                                         'value': 'FAILED',
                                         'format': failed_format})
            
            writer.close()

            # 2. Sent Log (sent_log.csv)
            sent_csv_path = os.path.join(self.current_run_dir, "sent_log.csv")
            if not sent_df.empty:
                sent_df.to_csv(sent_csv_path, index=False, encoding='utf-8')

            # 3. Failed Log (failed_log.csv)
            failed_csv_path = os.path.join(self.current_run_dir, "failed_log.csv")
            if not failed_df.empty:
                failed_df.to_csv(failed_csv_path, index=False, encoding='utf-8')

        except Exception as e:
            self._log_error(f"Rapor oluşturulurken hata oluştu: {e}")

if __name__ == "__main__":
    app = WhatsAppBroadcaster()
    app.mainloop()
