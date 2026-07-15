import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image
from threading import Thread
from datetime import datetime
import pandas as pd # Sadece type hint'ler ve pd.notna kontrolü için
from pathlib import Path

# BroadcasterLogic sınıfını import ediyoruz.
from broadcaster_logic import BroadcasterLogic

# --- Global Yapılandırma ---
VERSION = "Sürüm 2.3 Beta"
LOGO_PATH = Path(__file__).resolve().parent / "assets" / "logo.png"


class WhatsAppGUI(ctk.CTk):
    """
    Ana Arayüz Sınıfı.
    Görsel bileşenleri oluşturur, kullanıcı etkileşimlerini (düğme tıklamaları) yakalar
    ve BroadcasterLogic sınıfı ile iletişim kurar.
    """
    def __init__(self):
        super().__init__()
        
        # --- Arayüz Temel Ayarları ---
        self.title("WhatsApp Toplu Gönderim Aracı")
        self.geometry("1100x900")
        
        self.grid_columnconfigure(1, weight=1) 
        self.grid_columnconfigure(2, weight=1) 
        self.grid_rowconfigure(0, weight=1)    
        self.grid_rowconfigure(4, weight=1)    

        ctk.set_appearance_mode("System") 
        ctk.set_default_color_theme("blue")

        # Broadcaster mantığını başlatır ve GUI'ye referansını verir.
        self.logic = BroadcasterLogic(self)
        
        # --- Durum ve Veri Değişkenleri ---
        self.current_thread = None     # Arka planda çalışan gönderim iş parçacığı
        self.recipient_widgets = {}    # Kişi listesi arayüz öğeleri
        
        # --- Arayüz Bileşenlerini Oluşturma Sırası ---
        self._create_sidebar()
        self._create_main_frames()
        self._create_list_frame() 
        self._create_controls_frame()
        self._create_progress_frame()
        self._create_log_frame()
        
        self._log_to_terminal(f"[{self.title()}] Uygulama başlatıldı. Lütfen Excel dosyasını seçin.", "info")

    def _create_sidebar(self):
        """Sol taraftaki navigasyon ve yapılandırma çubuğunu oluşturur."""
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0, fg_color=("gray85", "gray15"))
        self.sidebar_frame.grid(row=0, column=0, rowspan=5, sticky="nsew") 
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        # --- LOGO VE BAŞLIK ALANI ---
        logo_frame = ctk.CTkFrame(self.sidebar_frame, fg_color=("gray80", "gray18"), corner_radius=0, height=80)
        logo_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        logo_frame.grid_columnconfigure(0, weight=1)

        logo_label = None
        try:
            pil_image = Image.open(LOGO_PATH)
            pil_image = pil_image.resize((300,300)) 
            self.logo_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(300, 300))
            logo_label = ctk.CTkLabel(logo_frame, text=" ", image=self.logo_image, compound="left", 
                                     font=ctk.CTkFont(family="Inter", size=1, weight="bold"))
        except FileNotFoundError:
            self._log_to_terminal(f"UYARI: Logo dosyası bulunamadı: {LOGO_PATH}", "error")

        if logo_label is not None:
            logo_label.grid(row=0, column=0, padx=10, pady=20)
        # ---------------------------------------------

        ctk.CTkLabel(self.sidebar_frame, text="WhatsApp Toplu Mesaj Otomasyonu", 
                     font=ctk.CTkFont(family="Inter", size=16, weight="bold")).grid(row=1, column=0, padx=20, pady=(15, 5))
        
        # Hız Modu Seçimi Kontrolleri
        ctk.CTkLabel(self.sidebar_frame, text="Hız Modu:", anchor="w", 
                     font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=20, pady=(10, 0), sticky="w")
        self.speed_mode = ctk.StringVar(value="FAST")
        self.mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["SAFE", "FAST", "TURBO"], command=self._show_delay_info)
        self.mode_optionemenu.grid(row=3, column=0, padx=20, pady=(5, 5), sticky="ew")
        
        # Hız moduna ilişkin açıklayıcı bilgi etiketi
        self.delay_info_label = ctk.CTkLabel(self.sidebar_frame, text="Gecikme: 5-8 sn", font=ctk.CTkFont(size=10), text_color="gray")
        self.delay_info_label.grid(row=4, column=0, padx=20, pady=(0, 15), sticky="n") 
        
        # Tema Seçimi Kontrolleri
        ctk.CTkLabel(self.sidebar_frame, text="Tema:", anchor="w", 
                     font=ctk.CTkFont(weight="bold")).grid(row=5, column=0, padx=20, pady=(10, 0), sticky="w")
        
        initial_appearance = ctk.get_appearance_mode().capitalize()
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, 
                                                            values=["Light", "Dark", "System"], 
                                                            command=self.change_appearance_mode_event,
                                                            variable=ctk.StringVar(value=initial_appearance)) 
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(5, 5), sticky="ew")
        
        ctk.CTkLabel(self.sidebar_frame, text=VERSION, 
                     font=ctk.CTkFont(size=10)).grid(row=8, column=0, padx=20, pady=(20, 20), sticky="s")
        
        self._show_delay_info(self.speed_mode.get())

    def _show_delay_info(self, mode):
        """Seçilen hız moduna göre gecikme bilgisini günceller."""
        if mode == "SAFE":
            info = "Gecikme: 8-15 sn (Güvenli, Önerilir)"
        elif mode == "FAST":
            info = "Gecikme: 5-8 sn (Dengeli)"
        elif mode == "TURBO":
            info = "Gecikme: 3-5 sn (Çok Hızlı, Riskli)"
        self.delay_info_label.configure(text=info)
        self.speed_mode.set(mode)

    def _create_main_frames(self):
        """Ana içerik çerçevelerini (Dosya Yolu ve Mesaj Şablonu) oluşturur."""
        self.file_frame = ctk.CTkFrame(self, corner_radius=10, border_width=2)
        self.file_frame.grid(row=0, column=1, padx=(20, 10), pady=(20, 10), sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.file_frame, text="1. Excel Dosyası Seç (phone, name, message sütunları zorunludur):", anchor="w", 
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 5), sticky="w")
        
        self.file_path_entry = ctk.CTkEntry(self.file_frame, placeholder_text="Excel dosyasının tam yolu...", height=30)
        self.file_path_entry.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="ew")
        
        self.file_button = ctk.CTkButton(self.file_frame, text="Dosya Seç", command=self.select_file, 
                                          fg_color="#3498db", hover_color="#2980b9")
        self.file_button.grid(row=1, column=1, padx=15, pady=(0, 15))

        self.message_frame = ctk.CTkFrame(self, corner_radius=10, border_width=2)
        self.message_frame.grid(row=1, column=1, padx=(20, 10), pady=(10, 10), sticky="nsew")
        self.message_frame.grid_columnconfigure(0, weight=1)
        self.message_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.message_frame, text="2. Mesaj Şablonu (Kişiselleştirme için {name} kullanın):", anchor="w", 
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 5), sticky="w")
        
        self.message_textbox = ctk.CTkTextbox(self.message_frame, height=150, corner_radius=8)
        self.message_textbox.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        
        self.message_textbox.insert("0.0", "Merhaba {name},\n\nBu, toplu mesaj gönderim aracımızın bir testidir. İyi günler!")

    def _create_list_frame(self):
        """Kişi listesi görüntüleme çerçevesini oluşturur."""
        self.list_container_frame = ctk.CTkFrame(self, corner_radius=10, border_width=2)
        self.list_container_frame.grid(row=0, column=2, rowspan=2, padx=(10, 20), pady=(20, 10), sticky="nsew")
        self.list_container_frame.grid_columnconfigure(0, weight=1)
        self.list_container_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.list_container_frame, text="3. Kişi Listesi ve Anlık Durum:", anchor="w", 
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 5), sticky="w")
        
        self.list_scroll_frame = ctk.CTkScrollableFrame(self.list_container_frame, label_text="Yüklenen Kişiler (0 Kişi)", label_font=ctk.CTkFont(weight="bold"))
        self.list_scroll_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.list_scroll_frame.grid_columnconfigure(0, weight=1)

    def _create_controls_frame(self):
        """Başlatma/İptal etme düğmelerini içeren kontrol çerçevesini oluşturur."""
        self.controls_frame = ctk.CTkFrame(self, corner_radius=10)
        self.controls_frame.grid(row=2, column=1, columnspan=2, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.controls_frame.grid_columnconfigure((0, 1), weight=1)

        self.start_button = ctk.CTkButton(self.controls_frame, text="Gönderimi BAŞLAT", command=self.start_broadcast_thread, height=45, 
                                          font=ctk.CTkFont(size=15, weight="bold"),
                                          fg_color="#2ecc71", hover_color="#27ae60")
        self.start_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        self.cancel_button = ctk.CTkButton(self.controls_frame, text="Gönderimi İPTAL ET", command=self.cancel_broadcast, height=45, 
                                           font=ctk.CTkFont(size=15, weight="bold"),
                                           fg_color="#e74c3c", hover_color="#c0392b", state="disabled")
        self.cancel_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
    
    def _create_progress_frame(self):
        """İlerleme çubuğunu ve sayaçları içeren durumu gösteren çerçeveyi oluşturur."""
        self.progress_frame = ctk.CTkFrame(self, corner_radius=10)
        self.progress_frame.grid(row=3, column=1, columnspan=2, padx=(20, 20), pady=(10, 10), sticky="ew")
        self.progress_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, orientation="horizontal", height=10)
        self.progress_bar.grid(row=0, column=0, columnspan=2, padx=20, pady=(15, 10), sticky="ew")
        self.progress_bar.set(0)

        self.counter_label = ctk.CTkLabel(self.progress_frame, text="Hazır | Toplam Kişi: 0", anchor="w", font=ctk.CTkFont(weight="bold"))
        self.counter_label.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
        
        self.status_label = ctk.CTkLabel(self.progress_frame, text="Durum: Bekleniyor...", anchor="e", font=ctk.CTkFont(weight="bold"))
        self.status_label.grid(row=1, column=1, padx=20, pady=(0, 15), sticky="e")

    def _create_log_frame(self):
        """Uygulama mesajlarını ve durumunu gösteren terminal alanını oluşturur."""
        self.log_frame = ctk.CTkFrame(self, corner_radius=10, border_width=2)
        self.log_frame.grid(row=4, column=1, columnspan=2, padx=(20, 20), pady=(10, 20), sticky="nsew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(self.log_frame, text="Terminal Çıktısı (Anlık İşlem Günlükleri):", anchor="w", 
                     font=ctk.CTkFont(size=14, weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 5), sticky="w")
        
        self.terminal_textbox = ctk.CTkTextbox(self.log_frame, height=150, activate_scrollbars=True, 
                                               corner_radius=8, font=ctk.CTkFont(family="Consolas", size=12))
        self.terminal_textbox.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.terminal_textbox.configure(state="disabled")

    def _log_to_terminal(self, message, tag="info"):
        """Mesajı terminal alanına ekler."""
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        if tag == "error":
            prefix = "[HATA]"
        elif tag == "success":
            prefix = "[BAŞARILI]"
        else:
            prefix = "[BİLGİ]"
            
        full_message = f"{timestamp} {prefix} {message}\n"
        
        self.terminal_textbox.configure(state="normal")
        self.terminal_textbox.insert(ctk.END, full_message)
        self.terminal_textbox.see(ctk.END)
        self.terminal_textbox.configure(state="disabled")

    # --- İşleyiciler (Handlers) ---

    def change_appearance_mode_event(self, new_appearance_mode: str):
        """Uygulamanın tema değişimini yönetir."""
        ctk.set_appearance_mode(new_appearance_mode)

    def select_file(self):
        """Dosya seçimi ve veriyi yükleme işlemini tetikler."""
        file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyaları", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        )
        if file_path:
            self.file_path_entry.delete(0, tk.END)
            self.file_path_entry.insert(0, file_path)
            self.preview_data_ui(file_path)

    def preview_data_ui(self, file_path):
        """Veri okuma mantığını çağırır ve arayüzü günceller."""
        success, error_msg = self.logic.load_data(file_path)
        
        if success:
            total = self.logic.total_recipients
            self.counter_label.configure(text=f"Hazır | Toplam Kişi: {total}")
            self.status_label.configure(text="Durum: Veri Yüklendi.")
            self._populate_list(self.logic.df_data)
            self._log_to_terminal(f"Excel verisi başarıyla yüklendi. Toplam {total} kişi.", "info")
        else:
            messagebox.showerror("Hata", error_msg)
            self.logic.df_data = None
            self.counter_label.configure(text="Hazır | Toplam Kişi: 0")
            self._clear_list()
            self._log_to_terminal(f"Dosya okuma hatası: {error_msg}", "error")

    def _clear_list(self):
        """Kişi listesi görünümündeki tüm öğeleri temizler."""
        for widget in self.list_scroll_frame.winfo_children():
            widget.destroy()
        self.recipient_widgets = {}
        self.list_scroll_frame.configure(label_text="Yüklenen Kişiler (0 Kişi)")

    def _populate_list(self, df):
        """DataFrame'den kişileri okur ve arayüzdeki listeye doldurur."""
        self._clear_list()
        total_recipients = len(df)
        self.list_scroll_frame.configure(label_text=f"Yüklenen Kişiler ({total_recipients} Kişi)")
        
        for list_row_index, (_, row) in enumerate(df.iterrows()):
            phone_raw = row['phone']
            # pandas.notna kontrolünü kullanmadan güvenli erişim (logic dosyasında fillna yapıldı)
            name = row['name'] if 'name' in row and row['name'] else '(İsimsiz Kişi)'
            status = row['status'] if 'status' in row and row['status'] else 'Bekliyor'
            
            person_frame = ctk.CTkFrame(self.list_scroll_frame, border_width=1, corner_radius=5)
            person_frame.grid(row=list_row_index, column=0, padx=5, pady=3, sticky="ew")
            person_frame.grid_columnconfigure(0, weight=1)
            
            name_label = ctk.CTkLabel(person_frame, text=f"{name} ({phone_raw})", anchor="w", 
                                       font=ctk.CTkFont(weight="bold", size=13))
            name_label.grid(row=0, column=0, padx=10, pady=(5, 5), sticky="w")
            
            status_label = ctk.CTkLabel(person_frame, text=f"Durum: {status}", anchor="e", text_color="gray", 
                                         font=ctk.CTkFont(weight="bold"))
            status_label.grid(row=0, column=1, padx=10, pady=(5, 5), sticky="e")
            
            self.recipient_widgets[list_row_index] = {
                'frame': person_frame,
                'status_label': status_label
            }

    def _update_list_status(self, index, status_text, color_key="pending"):
        """Kişi listesindeki bir öğenin durumunu ve rengini anlık olarak günceller (Logic tarafından çağrılır)."""
        if index in self.recipient_widgets:
            status_widget = self.recipient_widgets[index]['status_label']
            frame_widget = self.recipient_widgets[index]['frame']
            
            if color_key == "sent":
                color = "#2ecc71"
                frame_bg = ("#C6EFCE", "#3A533E")
            elif color_key == "failed":
                color = "#e74c3c"
                frame_bg = ("#FFC7CE", "#533A3A")
            elif color_key == "sending":
                color = "orange"
                frame_bg = self.list_scroll_frame.cget("fg_color")
            else: 
                color = "gray"
                frame_bg = self.list_scroll_frame.cget("fg_color")
            
            status_widget.configure(text=f"Durum: {status_text}", text_color=color)
            frame_widget.configure(fg_color=frame_bg)

    def _reset_list_colors(self):
        """Gönderim başlamadan listeyi 'Bekliyor' durumuna ve rengine sıfırlar (Logic tarafından çağrılır)."""
        for index, widgets in self.recipient_widgets.items():
            widgets['status_label'].configure(text="Durum: Bekliyor", text_color="gray")
            widgets['frame'].configure(fg_color=self.list_scroll_frame.cget("fg_color"))

    def update_progress(self):
        """İlerleme çubuğunu ve sayaçları anlık olarak günceller (Logic tarafından çağrılır)."""
        sent_count = len(self.logic.sent_log)
        failed_count = len(self.logic.failed_log)
        processed_count = sent_count + failed_count
        total = self.logic.total_recipients
        
        progress_value = processed_count / total if total > 0 else 0
        
        self.progress_bar.set(progress_value)
        self.counter_label.configure(text=f"İşlendi: {processed_count}/{total} | Başarılı: {sent_count} | Başarısız: {failed_count}")
        
        if processed_count < total:
            self.status_label.configure(text=f"Durum: Gönderiliyor... (Kişi {processed_count + 1}/{total})")

    # --- Kontrol Fonksiyonları ---

    def start_broadcast_thread(self):
        """Gönderim işlemini ayrı bir Thread'de başlatır."""
        if self.logic.is_running:
            messagebox.showwarning("Uyarı", "Gönderim zaten devam ediyor.")
            return

        file_path = self.file_path_entry.get()
        message_template = self.message_textbox.get("0.0", "end-1c")

        if not file_path or self.logic.df_data is None or self.logic.total_recipients == 0: 
            messagebox.showerror("Hata", "Lütfen geçerli bir Excel dosyası seçin ve veriyi yükleyin.")
            self._log_to_terminal("HATA: Gönderim başlatılamadı. Dosya seçimi veya veri yüklemesi eksik.", "error")
            return
        
        # logic.df_data zaten fillna() ile temizlendiği için doğrudan kontrol edilebilir.
        has_message_col = 'message' in self.logic.df_data.columns and not self.logic.df_data['message'].eq('').all()
        if not message_template.strip() and not has_message_col:
             messagebox.showerror("Hata", "Lütfen bir mesaj şablonu girin veya Excel dosyanızdaki 'message' sütununu doldurun.")
             self._log_to_terminal("HATA: Gönderim başlatılamadı. Mesaj içeriği eksik.", "error")
             return
            
        # Arayüz durumunu güncelle
        self.start_button.configure(state="disabled", text="Gönderim Başladı")
        self.cancel_button.configure(state="normal")
        self.progress_bar.set(0)
        self.update_progress()
        self._log_to_terminal(f"Gönderim işlemi başlatılıyor. Hız Modu: {self.speed_mode.get()}", "info")

        # Logic'i ayrı bir Thread'de başlat
        self.current_thread = Thread(target=self.logic.start_broadcast, args=(message_template, self.speed_mode.get()))
        self.current_thread.start()

    def cancel_broadcast(self, hard_stop=False):
        """Gönderimi durdurur ve Logic'i çağırır."""
        if self.logic.is_running or hard_stop:
            self.status_label.configure(text="Durum: İptal Ediliyor...")
            self._log_to_terminal("Kullanıcı isteği üzerine iptal ediliyor. Tarayıcı kapatılıyor...", "info")
            
            self.logic.cancel_broadcast()
            
            if not hard_stop and self.current_thread and self.current_thread.is_alive():
                self.current_thread.join(timeout=5)
            
            if hard_stop:
                self._finish_broadcast(cancelled=True)
        else:
            messagebox.showinfo("Bilgi", "Gönderim zaten durdurulmuş.")

    def _finish_broadcast(self, cancelled=False):
        """Gönderim tamamlandığında veya iptal edildiğinde son işlemleri yapar (Logic tarafından çağrılır)."""
        self.start_button.configure(state="normal", text="Gönderimi BAŞLAT")
        self.cancel_button.configure(state="disabled")

        if cancelled:
            self.status_label.configure(text="Durum: GÖNDERİM İPTAL EDİLDİ.")
            self._log_to_terminal("Gönderim kullanıcı tarafından İPTAL EDİLDİ.", "info")
        else:
            self.status_label.configure(text="Durum: GÖNDERİM TAMAMLANDI.")
            self._log_to_terminal("Gönderim başarıyla TAMAMLANDI.", "success")

        if self.logic.sent_log or self.logic.failed_log:
            success, report_path = self.logic.generate_reports()
            if success:
                 self._log_to_terminal(f"Raporlar oluşturuldu ve '{report_path}' klasörüne kaydedildi.", "info")
                 messagebox.showinfo("Tamamlandı", f"Gönderim tamamlandı.\nRaporlar şu klasöre kaydedildi: \n{report_path}")
            else:
                 self._log_to_terminal(f"Rapor oluşturulurken kritik hata oluştu.", "error")
        elif not cancelled:
             messagebox.showinfo("Bilgi", "İşlenecek veri yok veya işlem başlatılamadı.")
             self._log_to_terminal("İşlenecek veri kalmadı, işlem sonlandı.", "info")
