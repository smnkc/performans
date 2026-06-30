import os
import sys
import json
import time
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# Selenium imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

# Excel generator import
import excel_generator

class EOkulApp(tb.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("e-Okul Not Çekici & Ölçek Dağıtıcı")
        self.geometry("800x700")
        self.resizable(True, True)
        
        self.driver = None
        self.log_queue = queue.Queue()
        self.config_file = "config.json"
        
        self.init_variables()
        self.load_config()
        
        # Kriter dosyalarının ilk açılışta oluşturulmasını sağla
        excel_generator.load_criteria_files()
        
        self.create_widgets()
        
        # Kriterleri ve kılavuzu doldur
        self.load_criteria_into_ui()
        self.load_help_guide()
        
        # Log kuyruğunu izleme döngüsü
        self.after(100, self.process_logs)
        
    def init_variables(self):
        self.teacher_name_var = tk.StringVar()
        self.branch_var = tk.StringVar()
        self.principal_name_var = tk.StringVar()
        self.eval_type_var = tk.StringVar(value="Performans")
        self.output_path_var = tk.StringVar()
        
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.teacher_name_var.set(config.get("teacher_name", ""))
                    self.branch_var.set(config.get("branch", ""))
                    self.principal_name_var.set(config.get("principal_name", ""))
                    self.eval_type_var.set(config.get("eval_type", "Performans"))
                    self.output_path_var.set(config.get("output_path", ""))
            except Exception as e:
                print(f"Konfigürasyon yükleme hatası: {e}")
                
    def save_config(self):
        config = {
            "teacher_name": self.teacher_name_var.get(),
            "branch": self.branch_var.get(),
            "principal_name": self.principal_name_var.get(),
            "eval_type": self.eval_type_var.get(),
            "output_path": self.output_path_var.get()
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.write_log(f"[HATA] Ayarlar kaydedilemedi: {e}\n")

    def create_widgets(self):
        # Ana Konteyner
        main_frame = tb.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=YES)
        
        # Başlık Banner
        title_lbl = tb.Label(
            main_frame, 
            text="e-Okul Not Çekici & Ölçek Dağıtıcı", 
            font=("Arial", 16, "bold"), 
            bootstyle="primary"
        )
        title_lbl.pack(pady=(0, 10))
        
        # Tab Kontrolü (Notebook)
        notebook = tb.Notebook(main_frame, bootstyle="primary")
        notebook.pack(fill=BOTH, expand=YES, pady=5)
        
        # Tab'ler
        tab_operations = tb.Frame(notebook, padding=10)
        notebook.add(tab_operations, text="  Not Çekme İşlemleri  ")
        
        tab_criteria = tb.Frame(notebook, padding=10)
        notebook.add(tab_criteria, text="  Ölçek Kriterleri Düzenle  ")
        
        tab_help = tb.Frame(notebook, padding=10)
        notebook.add(tab_help, text="  Kullanım Kılavuzu  ")
        
        # ----------------------------------------------------
        # TAB 1: İŞLEMLER İÇERİĞİ
        # ----------------------------------------------------
        # Frame 1: Kullanıcı ve İmza Bilgileri
        info_lf = tb.LabelFrame(tab_operations, text=" Öğretmen ve Okul Bilgileri (İmzalar İçin) ", padding=10)
        info_lf.pack(fill=X, pady=5)
        
        tb.Label(info_lf, text="Öğretmen Adı Soyadı:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        tb.Entry(info_lf, textvariable=self.teacher_name_var, width=30).grid(row=0, column=1, padx=5, pady=5, sticky=W)
        
        tb.Label(info_lf, text="Öğretmen Branşı:").grid(row=0, column=2, sticky=W, padx=15, pady=5)
        tb.Entry(info_lf, textvariable=self.branch_var, width=25).grid(row=0, column=3, padx=5, pady=5, sticky=W)
        
        tb.Label(info_lf, text="Okul Müdürü Adı:").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        tb.Entry(info_lf, textvariable=self.principal_name_var, width=30).grid(row=1, column=1, padx=5, pady=5, sticky=W)
        
        # Frame 2: Değerlendirme ve Dosya Seçimi
        options_lf = tb.LabelFrame(tab_operations, text=" Seçenekler ve Kayıt Yeri ", padding=10)
        options_lf.pack(fill=X, pady=5)
        
        # Değerlendirme Tipi Seçimi
        tb.Label(options_lf, text="Değerlendirme Tipi:").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        
        radio_frame = tb.Frame(options_lf)
        radio_frame.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky=W)
        
        tb.Radiobutton(
            radio_frame, 
            text="Performans Notları (Ölçekli Dağıtım)", 
            variable=self.eval_type_var, 
            value="Performans",
            bootstyle="primary"
        ).pack(side=LEFT, padx=(0, 20))
        
        tb.Radiobutton(
            radio_frame, 
            text="Proje Notları (Ölçekli Dağıtım)", 
            variable=self.eval_type_var, 
            value="Proje",
            bootstyle="primary"
        ).pack(side=LEFT)
        
        # Kayıt Yeri Seçimi
        tb.Label(options_lf, text="Excel Kayıt Dosyası:").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        tb.Entry(options_lf, textvariable=self.output_path_var, width=60).grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=EW)
        
        browse_btn = tb.Button(options_lf, text="Gözat...", command=self.browse_file, bootstyle="secondary-outline")
        browse_btn.grid(row=1, column=3, padx=5, pady=5, sticky=W)
        
        options_lf.columnconfigure(1, weight=1)
        
        # Frame 3: İşlem Butonları
        btn_frame = tb.Frame(tab_operations)
        btn_frame.pack(fill=X, pady=10)
        
        self.btn_open_browser = tb.Button(
            btn_frame, 
            text="1. Tarayıcıyı Aç ve e-Okul'a Git", 
            command=self.open_browser, 
            bootstyle="info",
            width=25
        )
        self.btn_open_browser.pack(side=LEFT, padx=5)
        
        self.btn_start_scrape = tb.Button(
            btn_frame, 
            text="2. Verileri Çek ve Excel Üret", 
            command=self.start_scraping_thread, 
            bootstyle="success",
            state=DISABLED,
            width=25
        )
        self.btn_start_scrape.pack(side=LEFT, padx=5)
        
        self.btn_close_browser = tb.Button(
            btn_frame, 
            text="Tarayıcıyı Kapat", 
            command=self.close_browser, 
            bootstyle="danger-outline",
            state=DISABLED,
            width=15
        )
        self.btn_close_browser.pack(side=RIGHT, padx=5)
        
        # Frame 4: Log Ekranı
        log_lf = tb.LabelFrame(tab_operations, text=" İşlem Günlüğü ", padding=10)
        log_lf.pack(fill=BOTH, expand=YES, pady=5)
        
        self.log_text = tb.Text(log_lf, wrap=WORD, height=10, font=("Consolas", 9))
        self.log_text.pack(fill=BOTH, expand=YES)
        
        # Başlangıç logu
        self.write_log("Uygulama başlatıldı.\n")
        self.write_log("Adım 1: Bilgileri girip 'Tarayıcıyı Aç' butonuna basın.\n")
        self.write_log("Adım 2: Açılan Chrome tarayıcısında e-Okul'a giriş yapıp Hızlı Not Girişi ekranına gelin.\n")
        self.write_log("Adım 3: 'Verileri Çek ve Excel Üret' butonu aktifleştiğinde tıklayarak işlemi başlatın.\n\n")

        # ----------------------------------------------------
        # TAB 2: KRİTER DÜZENLEME İÇERİĞİ
        # ----------------------------------------------------
        desc_lbl = tb.Label(
            tab_criteria, 
            text="Ölçek kriterlerini düzenleyebilirsiniz. Her satıra tek bir kriter yazın.",
            font=("Arial", 9, "italic")
        )
        desc_lbl.pack(anchor=W, pady=(0, 10))
        
        text_frames = tb.Frame(tab_criteria)
        text_frames.pack(fill=BOTH, expand=YES)
        
        perf_lf = tb.LabelFrame(text_frames, text=" Performans Kriterleri ", padding=10)
        perf_lf.pack(side=LEFT, fill=BOTH, expand=YES, padx=(0, 5))
        self.perf_text_box = tb.Text(perf_lf, wrap=WORD, font=("Arial", 9))
        self.perf_text_box.pack(fill=BOTH, expand=YES)
        
        proj_lf = tb.LabelFrame(text_frames, text=" Proje Kriterleri ", padding=10)
        proj_lf.pack(side=RIGHT, fill=BOTH, expand=YES, padx=(5, 0))
        self.proj_text_box = tb.Text(proj_lf, wrap=WORD, font=("Arial", 9))
        self.proj_text_box.pack(fill=BOTH, expand=YES)
        
        save_btn_frame = tb.Frame(tab_criteria)
        save_btn_frame.pack(fill=X, pady=(10, 0))
        
        tb.Button(
            save_btn_frame, 
            text="Değişiklikleri Kaydet", 
            command=self.save_criteria_from_ui, 
            bootstyle="success",
            width=20
        ).pack(side=LEFT, padx=5)
        
        tb.Button(
            save_btn_frame, 
            text="Metin Dosyasından Geri Yükle", 
            command=self.load_criteria_into_ui, 
            bootstyle="secondary-outline",
            width=28
        ).pack(side=LEFT, padx=5)

        # ----------------------------------------------------
        # TAB 3: KULLANIM KILAVUZU İÇERİĞİ
        # ----------------------------------------------------
        help_lf = tb.LabelFrame(tab_help, text=" Adım Adım Kullanım Kılavuzu ", padding=10)
        help_lf.pack(fill=BOTH, expand=YES)
        
        self.help_text_box = tb.Text(help_lf, wrap=WORD, font=("Arial", 10))
        self.help_text_box.pack(fill=BOTH, expand=YES)

    def load_criteria_into_ui(self):
        # Kriter dosyalarını yeniden oku ve metin alanlarına doldur
        excel_generator.load_criteria_files()
        
        # Performans
        self.perf_text_box.delete("1.0", END)
        self.perf_text_box.insert(END, "\n".join(excel_generator.PERFORMANCE_CRITERIA))
        
        # Proje
        self.proj_text_box.delete("1.0", END)
        self.proj_text_box.insert(END, "\n".join(excel_generator.PROJECT_CRITERIA))
        
        self.write_log("[BİLGİ] Kriterler arayüze yüklendi.\n")

    def save_criteria_from_ui(self):
        perf_content = self.perf_text_box.get("1.0", END).strip()
        proj_content = self.proj_text_box.get("1.0", END).strip()
        
        perf_lines = [line.strip() for line in perf_content.split("\n") if line.strip()]
        proj_lines = [line.strip() for line in proj_content.split("\n") if line.strip()]
        
        if not perf_lines or not proj_lines:
            messagebox.showerror("Hata", "Kriter alanları boş bırakılamaz!")
            return
            
        try:
            with open("performans_kriterleri.txt", "w", encoding="utf-8") as f:
                f.write("\n".join(perf_lines))
            with open("proje_kriterleri.txt", "w", encoding="utf-8") as f:
                f.write("\n".join(proj_lines))
                
            excel_generator.PERFORMANCE_CRITERIA.clear()
            excel_generator.PERFORMANCE_CRITERIA.extend(perf_lines)
            
            excel_generator.PROJECT_CRITERIA.clear()
            excel_generator.PROJECT_CRITERIA.extend(proj_lines)
            
            self.write_log("[BAŞARILI] Yeni kriterler kaydedildi ve etkinleştirildi.\n")
            messagebox.showinfo("Başarılı", "Kriterler başarıyla kaydedildi!")
        except Exception as e:
            messagebox.showerror("Hata", f"Kriterler kaydedilirken hata oluştu: {e}")

    def load_help_guide(self):
        help_text = """==========================================================================
              e-Okul Not Çekici & Ölçek Dağıtıcı Kullanım Kılavuzu
==========================================================================

Bu uygulama, e-Okul Hızlı Not Girişi ekranındaki performans veya proje notlarını
çekerek, resmî MEB şablonuna uygun yatay ve tek sayfaya sığdırılmış bir ölçek
Excel belgesi oluşturur.

--------------------------------------------------------------------------
Kullanım Adımları:
--------------------------------------------------------------------------
1. Uygulama arayüzündeki gerekli bilgileri (Öğretmen, Branş, Müdür) doldurun.

2. Seçenekleri belirleyin:
   - Değerlendirme Tipi: Performans Notları mı yoksa Proje Notları mı çekilecek?
   - Excel Kayıt Dosyası: "Gözat..." butonuna tıklayıp dosyanın kaydedileceği 
     konumu (örneğin Masaüstü) seçin ve dosyaya bir isim verip kaydedin.

3. "1. Tarayıcıyı Aç ve e-Okul'a Git" butonuna basın.
   - Ayrı bir Google Chrome tarayıcı penceresi açılacaktır.

4. Açılan Chrome penceresinde e-Okul şifreniz veya MEBBİS girişiniz ile giriş yapın.
   - Kurum İşlemleri -> Not İşlemleri -> Hızlı Not Girişi (veya Ders Notu Girişi)
     ekranına gidin.
   * ÖNEMLİ: e-Okul'da bu ekrana ulaştığınızda başka hiçbir işlem yapmanıza gerek
     yoktur. Sınıf veya ders seçmeyin, sadece o sayfaya gidin.

5. Uygulama arayüzüne geri dönün ve aktifleşen "2. Verileri Çek ve Excel Üret"
   butonuna tıklayın.
   - Uygulama, e-Okul'da sorumlu olduğunuz tüm sınıf ve dersleri sırayla gezecek,
     notları okuyacak ve arka planda Excel belgesine işleyecektir.
   - İşlem log ekranında onay kutusu göründüğünde süreç tamamlanmış demektir.

--------------------------------------------------------------------------
Ölçek Kriterlerini Değiştirme (Özelleştirme):
--------------------------------------------------------------------------
Uygulamanın "Ölçek Kriterleri Düzenle" sekmesini kullanarak kriterlerinizi
doğrudan bu uygulama içerisinden satır satır düzenleyip kaydedebilirsiniz.
- Kriterleri kaydettiğinizde, Excel sütunları ve formüller yazdığınız kriter
  sayısına göre dinamik olarak kendini ayarlayacak ve notları buna göre
  dağıtacaktır.
"""
        self.help_text_box.delete("1.0", END)
        self.help_text_box.insert(END, help_text)
        self.help_text_box.config(state=DISABLED)

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Excel Dosyasını Kaydet"
        )
        if file_path:
            self.output_path_var.set(file_path)
            self.save_config()

    def write_log(self, text):
        self.log_queue.put(text)

    def process_logs(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.log_text.insert(END, msg)
            self.log_text.see(END)
        self.after(100, self.process_logs)

    def open_browser(self):
        self.save_config()
        self.btn_open_browser.config(state=DISABLED)
        self.write_log("[BİLGİ] Tarayıcı başlatılıyor, lütfen bekleyin...\n")
        
        def run():
            try:
                chrome_options = webdriver.ChromeOptions()
                # Bot algılamasını zorlaştırmak için ayarlar
                chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
                chrome_options.add_experimental_option('useAutomationExtension', False)
                chrome_options.add_argument("--disable-blink-features=AutomationControlled")
                
                self.driver = webdriver.Chrome(options=chrome_options)
                
                self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                })
                
                self.driver.get("https://e-okul.meb.gov.tr/")
                
                self.write_log("[BAŞARILI] Tarayıcı açıldı. e-Okul giriş sayfası yüklendi.\n")
                self.write_log("[UYARI] Lütfen e-Okul'a giriş yapın ve 'Hızlı Not Girişi' ekranına gidin.\n")
                
                self.btn_start_scrape.config(state=NORMAL)
                self.btn_close_browser.config(state=NORMAL)
            except Exception as e:
                self.write_log(f"[HATA] Tarayıcı açılamadı: {e}\n")
                self.btn_open_browser.config(state=NORMAL)
                
        threading.Thread(target=run, daemon=True).start()

    def close_browser(self):
        if self.driver:
            try:
                self.driver.quit()
                self.write_log("[BİLGİ] Tarayıcı kapatıldı.\n")
            except Exception:
                pass
            self.driver = None
        self.btn_open_browser.config(state=NORMAL)
        self.btn_start_scrape.config(state=DISABLED)
        self.btn_close_browser.config(state=DISABLED)

    def wait_for_page_load(self, old_element, timeout=12):
        """Postback sonrası sayfanın yüklenmesini bekler."""
        try:
            WebDriverWait(self.driver, timeout).until(EC.staleness_of(old_element))
        except Exception:
            pass
        WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

    def check_and_dismiss_alert(self):
        """Hata uyarı pencerelerini kapatır ve mesajı döndürür."""
        try:
            alert = self.driver.switch_to.alert
            alert_text = alert.text
            alert.accept()
            return alert_text
        except Exception:
            return None

    def find_listele_btn(self):
        selectors = [
            "//input[@id='btnListele']",
            "//input[@name='btnListele']",
            "//input[contains(@id, 'Listele')]",
            "//input[contains(@name, 'Listele')]",
            "//input[@value='Listele']",
            "//a[contains(text(), 'Listele')]",
            "//button[contains(text(), 'Listele')]"
        ]
        for xpath in selectors:
            try:
                elem = self.driver.find_element(By.XPATH, xpath)
                if elem.is_displayed():
                    return elem
            except Exception:
                continue
        return None

    def start_scraping_thread(self):
        if not self.output_path_var.get():
            messagebox.showerror("Hata", "Lütfen Excel kayıt dosyasını seçin!")
            return
            
        self.save_config()
        self.btn_start_scrape.config(state=DISABLED)
        self.btn_close_browser.config(state=DISABLED)
        
        # Arka planda taramayı başlat
        threading.Thread(target=self.scrape_process, daemon=True).start()

    def scrape_process(self):
        self.write_log("\n[BAŞLADI] Not çekme ve Excel oluşturma süreci başlatıldı.\n")
        eval_type = self.eval_type_var.get()
        
        try:
            # Doğru sayfada mıyız kontrolü
            # Sınıf/Şube dropdown'ı aranıyor
            class_dropdown_xpath = "//select[contains(@id, 'ddlSinifSube')]"
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, class_dropdown_xpath))
                )
            except Exception:
                self.write_log("[HATA] 'Hızlı Not Girişi' sayfası bulunamadı. Lütfen e-Okul'da not giriş ekranına gidin!\n")
                self.btn_start_scrape.config(state=NORMAL)
                self.btn_close_browser.config(state=NORMAL)
                return
                
            # e-Okul sayfasından okul adı ayıklanmaya çalışılır
            school_name = "BİLİNMEYEN OKUL"
            try:
                school_el = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Lisesi') or contains(text(), 'Okulu') or contains(text(), 'Müdürlüğü')]")
                school_name = school_el.text.strip().split("\n")[0]
                if "(" in school_name:
                    school_name = school_name.split("(")[0].strip()
            except Exception:
                pass
            self.write_log(f"[BİLGİ] Okul Adı: {school_name}\n")
            
            # Dönem/Eğitim Yılı ayıklanır
            academic_year = "2025-2026"
            term = "II. DÖNEM"
            try:
                donem_select = Select(self.driver.find_element(By.XPATH, "//select[contains(@id, 'ddlDonem') or contains(@id, 'Donem')]"))
                donem_text = donem_select.first_selected_option.text.strip()
                if " " in donem_text:
                    parts = donem_text.split(" ")
                    academic_year = parts[0]
                    term = " ".join(parts[1:]).upper()
            except Exception:
                pass
            self.write_log(f"[BİLGİ] Dönem: {academic_year} {term}\n")
            
            # Ana Veri Yapısı
            scraped_data = {}
            
            # Sınıf dropdown'ını bul
            class_select_elem = self.driver.find_element(By.XPATH, class_dropdown_xpath)
            class_select = Select(class_select_elem)
            
            class_options = [opt.get_attribute("value") for opt in class_select.options if opt.get_attribute("value") != "-1"]
            class_texts = {opt.get_attribute("value"): opt.text.strip() for opt in class_select.options}
            
            self.write_log(f"[BİLGİ] Toplam {len(class_options)} adet sınıf tespit edildi.\n")
            
            for class_val in class_options:
                class_text = class_texts[class_val]
                self.write_log(f"\n[BİLGİ] Sınıf taranıyor: {class_text}\n")
                
                # Sınıfı Seç
                class_elem = self.driver.find_element(By.XPATH, class_dropdown_xpath)
                Select(class_elem).select_by_value(class_val)
                self.wait_for_page_load(class_elem)
                time.sleep(0.5)
                
                # Alert kontrolü (yetki uyarısı vs varsa)
                alert_msg = self.check_and_dismiss_alert()
                if alert_msg:
                    self.write_log(f"[UYARI] e-Okul Uyarısı: {alert_msg} (Bu sınıf atlanıyor)\n")
                    continue
                
                # Sınıf program detayını ayır
                # Genelde format: AİHL - 12. Sınıf / D Şubesi (İMAM HATİP PROGRAMI UYGULANAN ALAN)
                program_detail = ""
                if "(" in class_text:
                    program_detail = class_text[class_text.find("("):]
                    class_clean = class_text[:class_text.find("(")].strip()
                else:
                    class_clean = class_text
                    
                # Sınıf adı kısaltması Excel sekmeleri için (örn: "12D")
                class_short = class_clean
                prefix = ""
                if " - " in class_clean:
                    prefix = class_clean.split(" - ")[0].strip()
                    class_clean_no_prefix = class_clean.split(" - ")[1].strip()
                else:
                    class_clean_no_prefix = class_clean
                
                # Parantez içindeki program detayından ayırt edici kelimeyi al (örn: İMAM HATİP, FEN)
                prog_suffix = ""
                if "(" in class_text and ")" in class_text:
                    try:
                        content = class_text[class_text.find("(")+1 : class_text.find(")")]
                        words = [w.strip() for w in content.split(" ") if w.strip()]
                        if words:
                            first_word = words[0].upper()
                            if first_word == "İMAM" and len(words) > 1:
                                prog_suffix = f" {first_word} {words[1].upper()}"
                            else:
                                prog_suffix = f" {first_word}"
                    except Exception:
                        pass
                
                try:
                    if "/" in class_clean_no_prefix:
                        parts = class_clean_no_prefix.split("/")
                        s_parts = parts[0].replace(".", "").replace("Sınıf", "").strip().split(" ")
                        sh_parts = parts[1].replace("Şubesi", "").strip().split(" ")
                        class_num = s_parts[-1]
                        sube_letter = sh_parts[0]
                        class_short = f"{prefix} {class_num}{sube_letter}{prog_suffix}".strip()
                except Exception:
                    pass
                
                # Dersler dropdown'ını bul
                course_dropdown_xpath = "//select[contains(@id, 'ddlDersler')]"
                try:
                    course_select_elem = self.driver.find_element(By.XPATH, course_dropdown_xpath)
                    course_options = [opt.get_attribute("value") for opt in Select(course_select_elem).options if opt.get_attribute("value") != "-1"]
                    course_texts = {opt.get_attribute("value"): opt.text.strip() for opt in Select(course_select_elem).options}
                except Exception:
                    self.write_log(f"[UYARI] {class_text} için ders seçeneği bulunamadı, geçiliyor.\n")
                    continue
                    
                self.write_log(f"[BİLGİ] Bu sınıf için {len(course_options)} ders bulundu.\n")
                
                for course_val in course_options:
                    course_text = course_texts[course_val]
                    self.write_log(f"  -> Ders taranıyor: {course_text}\n")
                    
                    # Dersi Seç
                    course_elem = self.driver.find_element(By.XPATH, course_dropdown_xpath)
                    Select(course_elem).select_by_value(course_val)
                    self.wait_for_page_load(course_elem)
                    time.sleep(0.5)
                    
                    alert_msg = self.check_and_dismiss_alert()
                    if alert_msg:
                        self.write_log(f"  [UYARI] e-Okul Uyarısı: {alert_msg} (Atlanıyor)\n")
                        continue
                        
                    # Listele butonuna bas
                    btn_listele = self.find_listele_btn()
                    if not btn_listele:
                        self.write_log("  [HATA] 'Listele' butonu bulunamadı! Geçiliyor.\n")
                        continue
                        
                    btn_listele.click()
                    self.wait_for_page_load(btn_listele)
                    time.sleep(0.8)
                    
                    alert_msg = self.check_and_dismiss_alert()
                    if alert_msg:
                        self.write_log(f"  [UYARI] e-Okul Uyarısı: {alert_msg} (Atlanıyor)\n")
                        continue
                        
                    # Öğrenci listesi tablosunu bul
                    try:
                        table = WebDriverWait(self.driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, "//table[@id='dgListem']"))
                        )
                    except Exception:
                        self.write_log("  [UYARI] Öğrenci listesi tablosu yüklenemedi. Geçiliyor.\n")
                        continue
                        
                    # Tablo satırlarını çek
                    rows = table.find_elements(By.XPATH, ".//tr[td]")
                    students_list = []
                    
                    for row in rows:
                        tds = row.find_elements(By.TAG_NAME, "td")
                        if len(tds) < 3:
                            continue
                            
                        # Okul no ve Ad Soyad dinamik bulma
                        student_no = ""
                        student_name = ""
                        name_cell_idx = -1
                        
                        for cell_idx, td in enumerate(tds):
                            txt = td.text.strip()
                            if txt.isdigit():
                                student_no = txt
                                name_cell_idx = cell_idx + 1
                                break
                                
                        if not student_no or name_cell_idx >= len(tds):
                            continue
                            
                        student_name = tds[name_cell_idx].text.strip()
                        
                        # Not alanlarını tara (hem input hem span etiketleri için)
                        elements = row.find_elements(By.XPATH, ".//input | .//span")
                        perf1, perf2, perf3, project = "", "", "", ""
                        
                        for el in elements:
                            el_id = el.get_attribute("id") or el.get_attribute("name") or ""
                            val = el.get_attribute("value") if el.tag_name == "input" else el.text.strip()
                            
                            if "SZL1" in el_id:
                                perf1 = val
                            elif "SZL2" in el_id:
                                perf2 = val
                            elif "SZL3" in el_id:
                                perf3 = val
                            elif "TPU" in el_id:
                                project = val
                            elif "ODV" in el_id:
                                if not project:
                                    project = val
                                    
                        student_entry = {
                            "no": student_no,
                            "name": student_name,
                            "perf1": perf1,
                            "perf2": perf2,
                            "perf3": perf3,
                            "project": project
                        }
                        students_list.append(student_entry)
                        
                    self.write_log(f"  [BAŞARILI] {len(students_list)} öğrenci bilgisi çekildi.\n")
                    
                    if students_list:
                        # Sayfa Sekme Adı: "12D SOSYOLOJİ" veya "12D MANTIK" gibi
                        sheet_key = f"{class_short} {course_text}"
                        scraped_data[sheet_key] = {
                            "school_name": school_name,
                            "academic_year": academic_year,
                            "term": term,
                            "program": class_clean + (" " + program_detail if program_detail else ""),
                            "course_name": course_text,
                            "students": students_list
                        }
                        
            # Veri çekme bitti, Excel üret
            if not scraped_data:
                self.write_log("\n[HATA] Taranan sınıflardan hiçbir öğrenci verisi çekilemedi!\n")
                self.btn_start_scrape.config(state=NORMAL)
                self.btn_close_browser.config(state=NORMAL)
                return
                
            self.write_log(f"\n[BİLGİ] Toplam {len(scraped_data)} sekme için Excel dosyası hazırlanıyor...\n")
            
            config = {
                "teacher_name": self.teacher_name_var.get(),
                "branch": self.branch_var.get(),
                "principal_name": self.principal_name_var.get()
            }
            
            output_file = self.output_path_var.get()
            
            excel_generator.create_excel_report(
                data=scraped_data,
                config=config,
                file_path=output_file,
                eval_type=eval_type
            )
            
            self.write_log(f"[BAŞARILI] Excel belgesi başarıyla oluşturuldu!\nDosya Yolu: {output_file}\n")
            messagebox.showinfo("Başarılı", f"Excel dosyası başarıyla oluşturuldu!\n\nDosya Yolu: {output_file}")
            
        except Exception as e:
            self.write_log(f"\n[SİSTEM HATASI] İşlem sırasında beklenmedik bir hata oluştu: {e}\n")
            import traceback
            self.write_log(traceback.format_exc())
            
        finally:
            self.btn_start_scrape.config(state=NORMAL)
            self.btn_close_browser.config(state=NORMAL)

if __name__ == "__main__":
    app = EOkulApp()
    app.mainloop()
