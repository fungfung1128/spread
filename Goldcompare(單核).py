import sys
import os
import json
import time
import threading
import re
import datetime
import winsound

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QTextEdit, QTabWidget, QGroupBox, QGridLayout,
                             QFileDialog, QMessageBox)
from PyQt6.QtCore import pyqtSignal, QThread, Qt, QTimer, QTime, pyqtSlot
from PyQt6.QtGui import QFont

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# --- 設定檔名稱 ---
CONFIG_FILE = "monitor_config_v2.json"

# --- 輔助函數 ---
def parse_price(price_str):
    """移除逗號與非數字字符，轉換為 float"""
    try:
        clean_str = re.sub(r'[^\d.]', '', price_str)
        return float(clean_str)
    except:
        return 0.0

# ==========================================
#   統一爬蟲執行緒 (Unified Crawler Thread)
#   核心改變：只用一個 Chrome，開啟 4 個分頁
# ==========================================

class UnifiedMonitorThread(QThread):
    log_signal = pyqtSignal(str)
    price_signal = pyqtSignal(str, float, float, str)  # (Source, Bid, Ask, Time)
    status_signal = pyqtSignal(str, str)  # (Source, Status Msg)
    finished_signal = pyqtSignal() 

    def __init__(self):
        super().__init__()
        self.running = True
        self.driver = None
        # 定義所有要監控的網站與對應的標籤頁 ID
        self.sites = {
            "WF": {"url": "https://www.wfbullion.com/", "handle": None},
            "IG": {"url": "https://www.ig.com/cn/commodities/markets-commodities/gold", "handle": None},
            "Oanda": {"url": "https://www.oanda.com/bvi-en/cfds/metals/", "handle": None},
            "Forex": {"url": "https://www.forex.com/cn/markets-to-trade/precious-metals/", "handle": None},
        }

    def setup_driver(self):
        """設定單一瀏覽器實例"""
        chrome_options = Options()
        # 為了省資源，建議使用 headless=new，但如果 IG 抓不到，可暫時註解掉這行變成有頭模式
        chrome_options.add_argument("--headless=new") 
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080") # IG 需要大視窗
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--mute-audio")
        
        # 注意：為了確保 IG 能跑，這裡不禁用 CSS 和 圖片
        chrome_options.add_argument(
            "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

        self.driver = webdriver.Chrome(options=chrome_options)

    def run(self):
        try:
            self.log_signal.emit("正在啟動整合瀏覽器引擎 (單核心模式)...")
            self.setup_driver()
            wait = WebDriverWait(self.driver, 10) # 設定等待時間

            # --- 1. 初始化分頁 (開啟 4 個 Tabs) ---
            site_keys = list(self.sites.keys())
            
            # 開啟第一個網站 (主分頁)
            first_key = site_keys[0]
            self.log_signal.emit(f"正在載入: {first_key} ...")
            self.driver.get(self.sites[first_key]["url"])
            self.sites[first_key]["handle"] = self.driver.current_window_handle
            
            # 開啟其餘網站 (新分頁)
            for key in site_keys[1:]:
                if not self.running: break
                self.log_signal.emit(f"正在載入: {key} ...")
                # JavaScript 開新分頁
                self.driver.execute_script(f"window.open('{self.sites[key]['url']}', '_blank');")
                # 切換到新分頁並記錄 handle
                self.driver.switch_to.window(self.driver.window_handles[-1])
                self.sites[key]["handle"] = self.driver.current_window_handle
                time.sleep(2) # 稍作等待讓網頁載入

            self.log_signal.emit("所有網站載入完成，開始輪詢監控...")

            # --- 2. 輪詢監控迴圈 ---
            while self.running:
                for key in site_keys:
                    if not self.running: break
                    
                    try:
                        # 切換到該網站的分頁
                        target_handle = self.sites[key]["handle"]
                        self.driver.switch_to.window(target_handle)
                        
                        # 執行對應的爬蟲邏輯
                        self.scrape_site(key, wait)
                        
                    except Exception as e:
                        self.status_signal.emit(key, "讀取錯誤")
                        # print(f"Error scraping {key}: {e}") # Debug用

                    # 每個網站之間稍微休息一下，減少 CPU 瞬間峰值
                    time.sleep(0.5)

                # 每一輪結束後，休息較長時間 (這是省電的關鍵)
                # 建議設定 3~5 秒
                for _ in range(30): # 30 * 0.1 = 3秒
                    if not self.running: break
                    time.sleep(0.1)

        except Exception as e:
            self.log_signal.emit(f"系統核心錯誤: {str(e)}")
        finally:
            self.stop_driver()
            self.finished_signal.emit()

    def scrape_site(self, key, wait):
        """根據不同的網站 key 執行對應的抓取邏輯"""
        now_str = time.strftime("%H:%M:%S")

        if key == "WF":
            # 永豐金邏輯
            try:
                price_element = wait.until(EC.presence_of_element_located((By.ID, "pm-llg")))
                raw_text = price_element.text.strip()
                lines = raw_text.split('\n')
                if len(lines) > 3:
                    bid = parse_price(lines[2].strip().replace(',', ''))
                    ask_str = lines[3].strip().split(' ')[0].replace(',', '')
                    ask = parse_price(ask_str)
                    self.price_signal.emit(key, bid, ask, now_str)
                    self.status_signal.emit(key, "監控中")
            except:
                self.status_signal.emit(key, "等待數據")

        elif key == "IG":
            # IG 邏輯
            try:
                # 使用 visibility 確保 IG 載入完成
                sell_el = wait.until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, ".price-ticket__button--sell .price-ticket__price")))
                buy_el = wait.until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, ".price-ticket__button--buy .price-ticket__price")))
                
                bid_text = sell_el.text.strip()
                ask_text = buy_el.text.strip()

                if bid_text and ask_text and bid_text != "-":
                    bid = parse_price(bid_text)
                    ask = parse_price(ask_text)
                    self.price_signal.emit(key, bid, ask, now_str)
                    self.status_signal.emit(key, "監控中")
                else:
                    self.status_signal.emit(key, "載入中...")
            except:
                self.status_signal.emit(key, "等待IG")

        elif key == "Oanda":
            # Oanda 邏輯
            try:
                gold_row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[.//span[contains(text(), 'Gold')]]")))
                cells = gold_row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 3:
                    bid = parse_price(cells[1].text)
                    ask = parse_price(cells[2].text)
                    self.price_signal.emit(key, bid, ask, now_str)
                    self.status_signal.emit(key, "監控中")
            except:
                 self.status_signal.emit(key, "等待Oanda")

        elif key == "Forex":
            # Forex.com 邏輯
            try:
                product_row = wait.until(EC.presence_of_element_located((By.XPATH, "//tr[.//a[@title='XAU USD']]")))
                bid = parse_price(product_row.find_element(By.CSS_SELECTOR, ".mp__td--Bid").text)
                ask = parse_price(product_row.find_element(By.CSS_SELECTOR, ".mp__td--Offer").text)
                self.price_signal.emit(key, bid, ask, now_str)
                self.status_signal.emit(key, "監控中")
            except:
                 self.status_signal.emit(key, "等待Forex")

    def stop(self):
        self.running = False

    def stop_driver(self):
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None


# ==========================================
#   UI 元件 (保持不變)
# ==========================================

class BrokerPanel(QGroupBox):
    def __init__(self, title, color_theme):
        super().__init__(title)
        self.setStyleSheet(f"""
            QGroupBox {{ font-weight: bold; font-size: 16px; border: 2px solid {color_theme}; margin-top: 10px; background-color: #fafafa; }}
            QGroupBox::title {{ color: {color_theme}; subcontrol-origin: margin; left: 10px; padding: 0 3px; }}
        """)
        layout = QGridLayout()
        self.lbl_bid = QLabel("0.00")
        self.lbl_bid.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        self.lbl_bid.setStyleSheet("color: blue;")
        self.lbl_ask = QLabel("0.00")
        self.lbl_ask.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        self.lbl_ask.setStyleSheet("color: red;")
        self.lbl_spread = QLabel("0.00")
        self.lbl_spread.setFont(QFont("Arial", 22, QFont.Weight.Bold))
        self.lbl_spread.setStyleSheet("color: purple; background-color: #e0e0e0; border-radius: 5px; padding: 5px;")
        self.lbl_spread.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_time = QLabel("--:--:--")
        self.lbl_status = QLabel("等待啟動")
        self.lbl_status.setStyleSheet("color: gray; font-size: 11px;")

        layout.addWidget(QLabel("Bid (買入):"), 0, 0)
        layout.addWidget(self.lbl_bid, 0, 1)
        layout.addWidget(QLabel("Ask (賣出):"), 1, 0)
        layout.addWidget(self.lbl_ask, 1, 1)
        layout.addWidget(QLabel("點差 (Spread):"), 2, 0)
        layout.addWidget(self.lbl_spread, 2, 1)
        layout.addWidget(QLabel("更新時間:"), 3, 0)
        layout.addWidget(self.lbl_time, 3, 1)
        layout.addWidget(self.lbl_status, 4, 0, 1, 2, Qt.AlignmentFlag.AlignCenter)
        self.setLayout(layout)

    def update_data(self, bid, ask, time_str):
        self.lbl_bid.setText(f"{bid:.2f}")
        self.lbl_ask.setText(f"{ask:.2f}")
        spread = abs(ask - bid)
        self.lbl_spread.setText(f"{spread:.2f}")
        self.lbl_time.setText(time_str)
        return spread

    def update_status(self, msg):
        self.lbl_status.setText(msg)


class GoldMonitorApp(QMainWindow):
    audio_log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("XAUUSD黃金監控系統")
        self.resize(1000, 800)

        self.monitor_thread = None # 變更：只需要一個 thread 變數
        self.setting_inputs = {}
        self.alert_status_labels = {}
        self.last_triggered_levels = {}

        self.init_ui()
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_realtime_clock)
        self.clock_timer.start(1000)

        self.audio_log_signal.connect(self.log_message)
        self.load_settings()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        top_bar = QHBoxLayout()
        self.btn_start = QPushButton("全部開始")
        self.btn_start.setStyleSheet("background-color: green; color: white; font-weight: bold; padding: 8px;")
        self.btn_start.clicked.connect(self.start_monitor)

        self.btn_stop = QPushButton("全部停止")
        self.btn_stop.setStyleSheet("background-color: red; color: white; font-weight: bold; padding: 8px;")
        self.btn_stop.clicked.connect(self.stop_monitor)
        self.btn_stop.setEnabled(False)

        self.lbl_clock = QLabel("--:--:--")
        self.lbl_clock.setFont(QFont("Arial", 20, QFont.Weight.Bold))

        top_bar.addWidget(self.btn_start)
        top_bar.addWidget(self.btn_stop)
        top_bar.addStretch()
        top_bar.addWidget(QLabel("系統時間:"))
        top_bar.addWidget(self.lbl_clock)
        main_layout.addLayout(top_bar)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        self.tab_monitor = QWidget()
        self.setup_monitor_tab()
        self.tabs.addTab(self.tab_monitor, "監控儀表板")

        self.tab_settings = QWidget()
        self.setup_settings_tab()
        self.tabs.addTab(self.tab_settings, "各平台警報設定")

        self.tab_log = QWidget()
        self.setup_log_tab()
        self.tabs.addTab(self.tab_log, "系統日誌")

    def setup_monitor_tab(self):
        layout = QGridLayout(self.tab_monitor)
        self.panel_wf = BrokerPanel("永豐金 (Wing Fung)", "#aa0000")
        self.panel_ig = BrokerPanel("IG Markets", "#ff9900")
        self.panel_oanda = BrokerPanel("Oanda", "#000000")
        self.panel_forex = BrokerPanel("Forex.com", "#00cc00")

        layout.addWidget(self.panel_wf, 0, 0)
        layout.addWidget(self.panel_ig, 0, 1)
        layout.addWidget(self.panel_oanda, 1, 0)
        layout.addWidget(self.panel_forex, 1, 1)

    def setup_settings_tab(self):
        layout = QVBoxLayout(self.tab_settings)
        layout.addWidget(QLabel("請在此為每一間公司分別設定 3 層級的點差警報。"))
        setting_tabs = QTabWidget()
        self.create_settings_page(setting_tabs, "WF", "永豐金設定")
        self.create_settings_page(setting_tabs, "IG", "IG 設定")
        self.create_settings_page(setting_tabs, "Oanda", "Oanda 設定")
        self.create_settings_page(setting_tabs, "Forex", "Forex 設定")
        layout.addWidget(setting_tabs)
        btn_save = QPushButton("手動儲存設定")
        btn_save.clicked.connect(self.save_settings)
        layout.addWidget(btn_save)

    def create_settings_page(self, parent_tab, key, title):
        page = QWidget()
        layout = QGridLayout(page)
        layout.addWidget(QLabel("層級"), 0, 0)
        layout.addWidget(QLabel("點差門檻 (>X 觸發)"), 0, 1)
        layout.addWidget(QLabel("音效路徑"), 0, 2)
        layout.addWidget(QLabel("警報狀態"), 0, 4)

        self.setting_inputs[key] = []
        for i in range(3):
            lbl_level = QLabel(f"層級 {i + 1}")
            txt_diff = QLineEdit()
            txt_diff.setPlaceholderText("例如 0.5")
            txt_sound = QLineEdit()
            txt_sound.setPlaceholderText("未選擇音效...")
            btn_browse = QPushButton("...")
            btn_browse.setFixedWidth(30)
            btn_browse.clicked.connect(lambda chk, t=txt_sound: self.browse_file(t))
            lbl_status = QLabel("正常")
            lbl_status.setStyleSheet("color: green")

            self.alert_status_labels[(key, i)] = lbl_status
            layout.addWidget(lbl_level, i + 1, 0)
            layout.addWidget(txt_diff, i + 1, 1)
            layout.addWidget(txt_sound, i + 1, 2)
            layout.addWidget(btn_browse, i + 1, 3)
            layout.addWidget(lbl_status, i + 1, 4)

            self.setting_inputs[key].append({"diff": txt_diff, "sound": txt_sound})
        layout.setRowStretch(4, 1)
        parent_tab.addTab(page, title)

    def setup_log_tab(self):
        layout = QVBoxLayout(self.tab_log)
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        layout.addWidget(self.txt_log)
        btn_clear = QPushButton("清除介面日誌")
        btn_clear.clicked.connect(self.txt_log.clear)
        layout.addWidget(btn_clear)

    def update_realtime_clock(self):
        self.lbl_clock.setText(QTime.currentTime().toString("HH:mm:ss"))

    def browse_file(self, line_edit):
        f, _ = QFileDialog.getOpenFileName(self, "選取音效", "", "WAV Audio (*.wav);;All (*.*)")
        if f: line_edit.setText(f)

    @pyqtSlot(str)
    def log_message(self, msg):
        ts = time.strftime("%H:%M:%S")
        full_log_text = f"[{ts}] {msg}"
        self.txt_log.append(full_log_text)
        today_date = datetime.datetime.now().strftime("%Y-%m-%d")
        log_filename = f"monitor_log_{today_date}.txt"
        try:
            with open(log_filename, "a", encoding="utf-8") as f:
                f.write(full_log_text + "\n")
        except: pass

    def start_monitor(self):
        self.save_settings()
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.last_triggered_levels = {}
        self.log_message("--- 系統啟動 (單核心輪詢模式) ---")

        # 啟動單一整合執行緒
        self.monitor_thread = UnifiedMonitorThread()
        self.monitor_thread.log_signal.connect(self.log_message)
        self.monitor_thread.price_signal.connect(self.on_price_update)
        self.monitor_thread.status_signal.connect(self.on_status_update)
        self.monitor_thread.finished_signal.connect(self.on_thread_finished)
        self.monitor_thread.start()

    def stop_monitor(self):
        self.log_message("正在停止監控...")
        self.btn_stop.setEnabled(False)
        if self.monitor_thread:
            self.monitor_thread.stop()

    def on_price_update(self, source, bid, ask, time_str):
        spread = 0.0
        if source == "WF": spread = self.panel_wf.update_data(bid, ask, time_str)
        elif source == "IG": spread = self.panel_ig.update_data(bid, ask, time_str)
        elif source == "Oanda": spread = self.panel_oanda.update_data(bid, ask, time_str)
        elif source == "Forex": spread = self.panel_forex.update_data(bid, ask, time_str)
        self.check_alert(source, spread)

    def on_status_update(self, source, msg):
        if source == "WF": self.panel_wf.update_status(msg)
        elif source == "IG": self.panel_ig.update_status(msg)
        elif source == "Oanda": self.panel_oanda.update_status(msg)
        elif source == "Forex": self.panel_forex.update_status(msg)

    def on_thread_finished(self):
        self.log_message("--- 監控執行緒已完全停止 ---")
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.monitor_thread = None

    def check_alert(self, source, spread):
        inputs = self.setting_inputs.get(source, [])
        current_highest_level = -1
        current_sound_path = None
        current_threshold_str = "0.00"

        for i, item in enumerate(inputs):
            try:
                threshold_text = item['diff'].text()
                threshold = float(threshold_text) if threshold_text else 0.0
            except: threshold = 0.0

            status_lbl = self.alert_status_labels.get((source, i))
            if threshold > 0 and spread >= threshold:
                status_lbl.setText("觸發中")
                status_lbl.setStyleSheet("color: red; font-weight: bold;")
                current_highest_level = i
                current_sound_path = item['sound'].text()
                current_threshold_str = threshold_text
            else:
                status_lbl.setText("正常")
                status_lbl.setStyleSheet("color: green;")

        last_level = self.last_triggered_levels.get(source, -1)
        if current_highest_level > last_level:
            formatted_msg = f"[{source}] 點差:{spread:.2f} 大於 層級點差:{current_threshold_str}"
            self.log_message(formatted_msg)
            if current_sound_path:
                threading.Thread(target=self.play_sound, args=(current_sound_path,), daemon=True).start()
            self.last_triggered_levels[source] = current_highest_level
        elif current_highest_level < last_level:
            self.last_triggered_levels[source] = current_highest_level

    def play_sound(self, path):
        normalized_path = os.path.normpath(path)
        if os.path.exists(normalized_path):
            try: winsound.PlaySound(normalized_path, winsound.SND_FILENAME | winsound.SND_NODEFAULT)
            except Exception as e: self.audio_log_signal.emit(f"音效播放失敗: {str(e)}")

    def save_settings(self):
        data = {}
        for key, inputs in self.setting_inputs.items():
            data[key] = []
            for item in inputs:
                data[key].append({"diff": item['diff'].text(), "sound": item['sound'].text()})
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            self.log_message("設定已儲存")
        except Exception as e: self.log_message(f"儲存失敗: {e}")

    def load_settings(self):
        if not os.path.exists(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f: data = json.load(f)
            for key, tiers in data.items():
                if key in self.setting_inputs:
                    ui_inputs = self.setting_inputs[key]
                    for i, t_data in enumerate(tiers):
                        if i < len(ui_inputs):
                            ui_inputs[i]['diff'].setText(t_data.get('diff', ''))
                            ui_inputs[i]['sound'].setText(t_data.get('sound', ''))
        except Exception as e: self.log_message(f"讀取設定失敗: {e}")

    def closeEvent(self, event):
        if self.monitor_thread and self.monitor_thread.isRunning():
            reply = QMessageBox.question(self, '確認', '程式正在執行，確定要關閉嗎？',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.stop_monitor()
                event.accept()
            else: event.ignore()
        else:
            self.save_settings()
            event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft JhengHei", 10))
    window = GoldMonitorApp()
    window.show()
    sys.exit(app.exec())
